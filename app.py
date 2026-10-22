# -*- coding: utf-8 -*-
"""Edge 阅读器 —— 粘贴文字，享受沉浸式阅读 + Windows 系统语音朗读。

运行依赖: pywebview（界面）、pywin32（SAPI 语音朗读）。
数据保存在 %APPDATA%\\EdgeReader\\articles.json。
"""
import base64
import ctypes
import hashlib
import json
import os
import subprocess
import sys
import threading
import time
import uuid

import webview

APP_VERSION = '1.0.4'
REPO = 'VINK77aaaa/edge-reader'

try:
    import pythoncom
    import win32com.client
    HAS_SAPI = True
except ImportError:
    HAS_SAPI = False

APP_NAME = 'EdgeReader'
DATA_DIR = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), APP_NAME)
DB_PATH = os.path.join(DATA_DIR, 'articles.json')
TTS_CACHE_DIR = os.path.join(DATA_DIR, 'tts_cache')

# SAPI ISpVoice::Speak 标志
SPF_ASYNC = 1
SPF_PURGEBEFORESPEAK = 2
# ISpVoice Status.RunningState 实测取值：0=未开始/已暂停，2=朗读中，1=完成
SPRS_DONE = 1


def clamp_rate(rate):
    try:
        rate = int(rate)
    except (TypeError, ValueError):
        rate = 1
    return max(-5, min(8, rate))


# ---- 全局朗读状态 ----
# 注意：pywebview 会递归扫描 js_api 实例的所有属性来生成 JS 桥接对象，
# 因此 API 实例上不能挂窗口对象、锁、线程等复杂属性（会递归卡死），
# 全部状态放在模块级变量里。
_stop_evt = threading.Event()
_pause_evt = threading.Event()
_tts_thread = None
_rate = 1
_voice_name = ''
_db_lock = threading.Lock()
# 朗读事件由前端轮询取走，避免后台线程直接 evaluate_js 造成死锁
_events = []
_ev_lock = threading.Lock()
_edge_lock = threading.Lock()
_edge_segs = []
_edge_paths = []
_edge_gen_done = False
_edge_error = None
_tts_session = 0      # 朗读会话编号：每次 tts_start 递增，旧线程发现过时立即退出
_window = None  # webview 窗口引用（放模块级，避免被 pywebview 扫描 API 对象）
_fs = False     # 当前是否处于全屏


class ReaderApi:
    """暴露给前端 JS 的接口（window.pywebview.api.xxx）。只放方法，不放状态。"""

    def __init__(self):
        pass

    # ---------------- 数据存储 ----------------

    def _load_db(self):
        with _db_lock:
            try:
                with open(DB_PATH, 'r', encoding='utf-8') as f:
                    db = json.load(f)
            except (OSError, ValueError):
                return {'articles': [], 'prefs': {}}
        db.setdefault('articles', [])
        db.setdefault('prefs', {})
        return db

    def _save_db(self, db):
        with _db_lock:
            os.makedirs(DATA_DIR, exist_ok=True)
            tmp = DB_PATH + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(db, f, ensure_ascii=False)
            os.replace(tmp, DB_PATH)

    def _find(self, db, article_id):
        for a in db['articles']:
            if a['id'] == article_id:
                return a
        return None

    def save_article(self, title, text):
        db = self._load_db()
        article = {
            'id': uuid.uuid4().hex[:12],
            'title': (title or '').strip() or '未命名文章',
            'text': text,
            'created': time.strftime('%Y-%m-%d %H:%M'),
            'progress': 0.0,
            'para': 0,
        }
        db['articles'].insert(0, article)
        self._save_db(db)
        return article

    def get_articles(self):
        db = self._load_db()
        return [
            {k: a[k] for k in ('id', 'title', 'created', 'progress')}
            for a in db['articles']
        ]

    def get_article(self, article_id):
        return self._find(self._load_db(), article_id)

    def delete_article(self, article_id):
        db = self._load_db()
        db['articles'] = [a for a in db['articles'] if a['id'] != article_id]
        self._save_db(db)
        return True

    def save_progress(self, article_id, progress, para):
        db = self._load_db()
        a = self._find(db, article_id)
        if a:
            a['progress'] = float(progress or 0)
            a['para'] = int(para or 0)
            self._save_db(db)
        return True

    def get_prefs(self):
        return self._load_db()['prefs']

    def save_prefs(self, prefs):
        db = self._load_db()
        db['prefs'].update(prefs or {})
        self._save_db(db)
        return True

    # ---------------- 文档导入（docx / doc / txt / md） ----------------

    @staticmethod
    def _parse_docx(data):
        import io
        from docx import Document
        doc = Document(io.BytesIO(data))
        paras = [p.text.strip() for p in doc.paragraphs]
        return '\n\n'.join(p for p in paras if p)

    @staticmethod
    def _parse_doc_via_word(path):
        """老式 .doc 二进制格式：调用本机 Word 转出文本（需安装 Office Word）。"""
        import tempfile
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            text = doc.Content.Text
            doc.Close(False)
            text = text.replace('\r\x07', '\n').replace('\r', '\n')
            return '\n\n'.join(p.strip() for p in text.split('\n') if p.strip())
        finally:
            if word:
                word.Quit()
            pythoncom.CoUninitialize()

    def load_document_bytes(self, b64, filename):
        """前端拖入文件：内容以 base64 传入，解析出正文文本。"""
        try:
            data = base64.b64decode(b64)
        except Exception:
            return {'error': '文件内容读取失败'}
        ext = os.path.splitext(filename)[1].lower()
        title = os.path.splitext(filename)[0]
        try:
            if ext == '.docx':
                text = self._parse_docx(data)
            elif ext == '.doc':
                import tempfile
                fd, tmp = tempfile.mkstemp(suffix='.doc')
                with os.fdopen(fd, 'wb') as f:
                    f.write(data)
                try:
                    text = self._parse_doc_via_word(tmp)
                finally:
                    os.remove(tmp)
            elif ext in ('.txt', '.md'):
                text = data.decode('utf-8', errors='ignore').replace('\r\n', '\n')
            else:
                return {'error': '仅支持 docx / doc / txt / md 文档'}
        except Exception as exc:
            if ext == '.doc':
                return {'error': '解析 .doc 需要本机安装 Word。可先另存为 .docx 再拖入（%s）' % exc}
            return {'error': '文档解析失败: %s' % exc}
        if not text.strip():
            return {'error': '文档里没有可读取的文字'}
        return {'title': title, 'text': text}

    def pick_document(self):
        """弹文件选择框，返回解析后的正文。"""
        global _window
        if not _window:
            return {'error': '窗口未就绪'}
        result = _window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=('文档 (*.docx;*.doc;*.txt;*.md)', '所有文件 (*.*)'))
        if not result:
            return None
        path = result[0]
        ext = os.path.splitext(path)[1].lower()
        title = os.path.splitext(os.path.basename(path))[0]
        try:
            with open(path, 'rb') as f:
                data = f.read()
            if ext == '.docx':
                text = self._parse_docx(data)
            elif ext == '.doc':
                text = self._parse_doc_via_word(path)
            elif ext in ('.txt', '.md'):
                text = data.decode('utf-8', errors='ignore').replace('\r\n', '\n')
            else:
                return {'error': '仅支持 docx / doc / txt / md 文档'}
        except Exception as exc:
            if ext == '.doc':
                return {'error': '解析 .doc 需要本机安装 Word。可先另存为 .docx 再打开（%s）' % exc}
            return {'error': '文档解析失败: %s' % exc}
        if not text.strip():
            return {'error': '文档里没有可读取的文字'}
        return {'title': title, 'text': text}

    # ---------------- 语音朗读（SAPI，即 Edge“大声朗读”同款系统引擎） ----------------

    def _notify(self, payload):
        with _ev_lock:
            _events.append(payload)

    def tts_poll(self):
        """前端定时调用，取走朗读事件（替代后台线程推送）。"""
        global _events
        with _ev_lock:
            events, _events = _events, []
        return events

    def tts_get_audio(self, index):
        """在线语音模式：前端轮询取已合成好的分段音频。"""
        try:
            i = int(index)
        except (TypeError, ValueError):
            return {'done': True}
        with _edge_lock:
            if i >= len(_edge_paths):
                if _edge_gen_done:
                    return {'done': True, 'error': _edge_error}
                return {}
            path = _edge_paths[i]
            seg = _edge_segs[i] if i < len(_edge_segs) else None
        if path is None or not os.path.exists(path):
            # 段落尚未合成；若合成已整体结束仍未产出，视为提前终止
            with _edge_lock:
                if _edge_gen_done:
                    return {'done': True, 'error': _edge_error}
            return {}
        try:
            with open(path, 'rb') as f:
                b64 = base64.b64encode(f.read()).decode('ascii')
        except OSError:
            return {}
        return {'p': seg['p'] if seg else 0, 'b64': b64}

    def toggle_fullscreen(self):
        """原生窗口级全屏（WebView2 的网页全屏只填客户区，标题栏还在，故用窗口 API）。"""
        global _window, _fs
        try:
            if _window:
                _window.toggle_fullscreen()
                _fs = not _fs
                return _fs
        except Exception:
            pass
        return _fs

    # ---------------- 自动更新 ----------------

    def check_update(self):
        """查询 GitHub 最新 Release 并比较版本。"""
        try:
            import urllib.request
            req = urllib.request.Request(
                'https://api.github.com/repos/%s/releases/latest' % REPO,
                headers={'User-Agent': 'EdgeReader', 'Accept': 'application/vnd.github+json'})
            with urllib.request.urlopen(req, timeout=15) as r:
                data = json.loads(r.read().decode('utf-8'))
        except Exception as exc:
            return {'error': '检查更新失败（请检查网络）: %s' % exc}
        tag = (data.get('tag_name') or '').lstrip('v')
        try:
            latest = tuple(int(x) for x in tag.split('.'))
        except ValueError:
            return {'error': '无法解析版本号: %s' % tag}
        current = tuple(int(x) for x in APP_VERSION.split('.'))
        if latest <= current:
            return {'update': False, 'current': APP_VERSION}
        assets = data.get('assets', [])
        url = next((a['browser_download_url'] for a in assets if a['name'] == 'EdgeReader.exe'), None)
        size = next((a.get('size', 0) for a in assets if a['name'] == 'EdgeReader.exe'), 0)
        if not url:
            return {'error': '新版本 %s 缺少 EXE 附件' % tag}
        return {'update': True, 'current': APP_VERSION, 'version': tag,
                'notes': (data.get('body') or '')[:600], 'url': url, 'size': size}

    def download_update(self, url):
        """后台下载新 EXE 到 <自身>.new，进度通过轮询事件推送。"""
        if not getattr(sys, 'frozen', False):
            return {'error': '源码运行模式请用 git pull 更新'}
        target = sys.executable + '.new'

        def work():
            import urllib.request
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'EdgeReader'})
                with urllib.request.urlopen(req, timeout=30) as r, open(target, 'wb') as f:
                    total = int(r.headers.get('Content-Length', 0)) or 1
                    got = 0
                    while True:
                        chunk = r.read(65536)
                        if not chunk:
                            break
                        f.write(chunk)
                        got += len(chunk)
                        self._notify({'type': 'dl_progress', 'pct': min(99, got * 100 // total)})
                if os.path.getsize(target) < 1024 * 1024:
                    raise RuntimeError('下载的文件不完整')
                self._notify({'type': 'dl_done'})
            except Exception as exc:
                try:
                    os.remove(target)
                except OSError:
                    pass
                self._notify({'type': 'error', 'message': '下载失败: %s' % exc})

        threading.Thread(target=work, daemon=True).start()
        return True

    def apply_update(self):
        """替换自身并重启新版本。"""
        if not getattr(sys, 'frozen', False):
            return {'error': '源码运行模式请用 git pull 更新'}
        exe = sys.executable
        new, old = exe + '.new', exe + '.old'
        if not os.path.exists(new):
            return {'error': '尚未下载新版本'}
        try:
            if os.path.exists(old):
                os.remove(old)
            os.rename(exe, old)
            os.rename(new, exe)
        except Exception as exc:
            return {'error': '替换文件失败: %s' % exc}
        subprocess.Popen([exe, '--updated'], cwd=os.path.dirname(exe),
                         creationflags=0x00000008 | 0x200)  # DETACHED | NEW_PROCESS_GROUP
        time.sleep(0.4)
        if _window:
            try:
                _window.destroy()
            except Exception:
                pass
        os._exit(0)
        return {'ok': True}

    def set_title(self, title):
        """窗口标题跟随文章标题。"""
        global _window
        try:
            if _window:
                _window.set_title((str(title) or 'Edge 阅读器')[:60])
        except Exception:
            pass

    def tts_voices(self):
        if not HAS_SAPI:
            return []
        pythoncom.CoInitialize()
        try:
            voice = win32com.client.Dispatch('SAPI.SpVoice')
            voices = voice.GetVoices()
            return [voices.Item(i).GetAttribute('Name') for i in range(voices.Count)]
        except Exception as exc:
            self._notify({'type': 'error', 'message': '读取语音列表失败: %s' % exc})
            return []
        finally:
            pythoncom.CoUninitialize()

    def tts_start(self, segments, voice_name, rate):
        """segments: [{'p': 段落序号, 't': 文本}, ...]

        voice_name 带 'edge:' 前缀走 Edge 在线神经网络语音（如晓晓、云健），
        否则走本机 SAPI 语音。"""
        global _rate, _voice_name, _tts_thread, _tts_session
        global _edge_segs, _edge_paths, _edge_gen_done, _edge_error
        if voice_name and voice_name.startswith('edge:'):
            try:
                import edge_tts  # noqa: F401
            except ImportError:
                self._notify({'type': 'error',
                              'message': '使用在线语音需先安装 edge-tts: py -m pip install edge-tts'})
                return False
            self.tts_stop()
            _tts_session += 1
            sess = _tts_session
            _stop_evt.clear()
            _pause_evt.clear()
            _rate = clamp_rate(rate)
            segs = [{'p': int(s.get('p', 0)), 't': str(s.get('t', '')).strip()}
                    for s in segments or [] if str(s.get('t', '')).strip()]
            if not segs:
                return False
            with _edge_lock:
                _edge_segs = segs
                _edge_paths = [None] * len(segs)
                _edge_gen_done = False
                _edge_error = None
            _tts_thread = threading.Thread(
                target=self._edge_gen_loop, args=(segs, voice_name[5:], _rate, sess), daemon=True)
            _tts_thread.start()
            return True

        if not HAS_SAPI:
            self._notify({'type': 'error',
                          'message': '缺少 pywin32，无法朗读。请运行: py -m pip install pywin32'})
            return False
        self.tts_stop()
        _tts_session += 1
        sess = _tts_session
        _stop_evt.clear()
        _pause_evt.clear()
        _rate = clamp_rate(rate)
        _voice_name = voice_name or ''
        segs = [{'p': int(s.get('p', 0)), 't': str(s.get('t', '')).strip()}
                for s in segments or [] if str(s.get('t', '')).strip()]
        if not segs:
            return False
        _tts_thread = threading.Thread(target=self._speak_loop, args=(segs, sess), daemon=True)
        _tts_thread.start()
        return True

    def tts_pause(self):
        _pause_evt.set()

    def tts_resume(self):
        _pause_evt.clear()

    def tts_stop(self):
        _stop_evt.set()
        _pause_evt.clear()

    def tts_set_rate(self, rate):
        global _rate
        _rate = clamp_rate(rate)

    def _apply_voice(self, voice):
        if not _voice_name:
            return
        try:
            voices = voice.GetVoices()
            for i in range(voices.Count):
                if voices.Item(i).GetAttribute('Name') == _voice_name:
                    voice.Voice = voices.Item(i)
                    break
        except Exception:
            pass

    @staticmethod
    def _synth_once(edge_tts_mod, text, path, voice_id, rate_str):
        import asyncio
        com = edge_tts_mod.Communicate(text, voice_id, rate=rate_str)
        asyncio.run(asyncio.wait_for(com.save(path), timeout=90))

    def _upgrade_edge_tts(self):
        """源码运行：pip 升级 edge-tts（微软改鉴权导致 403 时的自动自救）。"""
        if getattr(sys, 'frozen', False):
            return False
        try:
            r = subprocess.run(
                [sys.executable, '-m', 'pip', 'install', '-U', 'edge-tts',
                 '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple'],
                capture_output=True, timeout=300)
            return r.returncode == 0
        except Exception:
            return False

    def _synth_heal(self, text, path, voice_id, rate_str):
        """合成一段；若遇 403 鉴权拒绝：源码模式自动升级 edge-tts 后重试，EXE 模式抛出并提示更新程序。"""
        import importlib
        try:
            import edge_tts
            self._synth_once(edge_tts, text, path, voice_id, rate_str)
            return
        except Exception as exc:
            first = exc
        msg = str(first)
        if '403' not in msg and 'Invalid response status' not in msg:
            raise first
        # 403：微软鉴权变更，尝试自动升级 edge-tts
        if self._upgrade_edge_tts():
            self._notify({'type': 'info', 'message': '检测到微软接口变更，已自动升级 edge-tts，正在重试…'})
            import edge_tts as fresh
            importlib.reload(fresh)
            self._synth_once(fresh, text, path, voice_id, rate_str)
            return
        raise RuntimeError('微软在线语音接口鉴权已变更（403），请更新程序后重试')

    def _edge_gen_loop(self, segments, voice_id, rate, sess):
        """Edge 在线语音模式：后台逐段合成 mp3（带缓存），前端轮询 tts_get_audio 取走播放。"""
        global _edge_gen_done, _edge_error

        os.makedirs(TTS_CACHE_DIR, exist_ok=True)
        pct = rate * 10
        rate_str = ('+%d%%' % pct) if pct >= 0 else ('%d%%' % pct)

        if sess == _tts_session:
            self._notify({'type': 'state', 'state': 'speaking'})
        try:
            for idx, seg in enumerate(segments):
                if _stop_evt.is_set() or sess != _tts_session:
                    break
                key = hashlib.md5((voice_id + rate_str + seg['t']).encode('utf-8')).hexdigest()
                path = os.path.join(TTS_CACHE_DIR, key + '.mp3')
                if not os.path.exists(path):
                    self._synth_heal(seg['t'], path, voice_id, rate_str)
                if _stop_evt.is_set() or sess != _tts_session:
                    break
                with _edge_lock:
                    if sess == _tts_session:
                        _edge_paths[idx] = path
            if not _stop_evt.is_set() and sess == _tts_session:
                self._notify({'type': 'gen_done'})
        except Exception as exc:
            if sess != _tts_session:
                return
            with _edge_lock:
                _edge_gen_done = True
                _edge_error = '在线语音合成失败: %s' % exc
            self._notify({'type': 'error',
                          'message': '在线语音合成失败: %s' % exc})
            if '鉴权' in str(exc):
                self._notify({'type': 'update_suggest'})
        finally:
            with _edge_lock:
                if sess == _tts_session:
                    _edge_gen_done = True

    def _speak_loop(self, segments, sess):
        pythoncom.CoInitialize()
        try:
            voice = win32com.client.Dispatch('SAPI.SpVoice')
            self._apply_voice(voice)
            voice.Rate = _rate
            if sess == _tts_session:
                self._notify({'type': 'state', 'state': 'speaking'})
            for idx, seg in enumerate(segments):
                if _stop_evt.is_set() or sess != _tts_session:
                    break
                voice.Speak(seg['t'], SPF_ASYNC)
                self._notify({'type': 'segment', 'index': idx, 'p': seg['p']})
                was_paused = False
                while True:
                    if _stop_evt.is_set() or sess != _tts_session:
                        voice.Speak('', SPF_ASYNC | SPF_PURGEBEFORESPEAK)
                        voice.Resume()
                        break
                    # 暂停/恢复按事件切换一次，不能依赖轮询状态值（暂停时状态是 0，与未开始相同）
                    if _pause_evt.is_set() and not was_paused:
                        voice.Pause()
                        was_paused = True
                    elif not _pause_evt.is_set() and was_paused:
                        voice.Resume()
                        was_paused = False
                    if not _pause_evt.is_set() and voice.Status.RunningState == SPRS_DONE:
                        break
                    if voice.Rate != _rate:
                        voice.Rate = _rate
                    time.sleep(0.05)
            if not _stop_evt.is_set() and sess == _tts_session:
                self._notify({'type': 'state', 'state': 'finished'})
        except Exception as exc:
            self._notify({'type': 'error', 'message': '朗读出错: %s' % exc})
        finally:
            pythoncom.CoUninitialize()


def resource_path(rel):
    """兼容 PyInstaller 打包后的资源路径（onefile 解压到 sys._MEIPASS）。"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def _acquire_single_instance():
    """用 Windows 互斥量保证同一时间只有一个实例运行。

    自动更新重启时带 --updated 参数：旧进程退出互斥量需要一点时间，此时最多等 10 秒。"""
    wait_until = time.time() + 10 if '--updated' in sys.argv else 0
    while True:
        ctypes.windll.kernel32.CreateMutexW(None, False, 'EdgeReader_SingleInstance_Mutex')
        if ctypes.windll.kernel32.GetLastError() != 183:  # 183 = ERROR_ALREADY_EXISTS
            return True
        if time.time() >= wait_until:
            return False
        time.sleep(0.3)


def _read_prefs():
    try:
        with open(DB_PATH, 'r', encoding='utf-8') as f:
            return json.load(f).get('prefs', {})
    except (OSError, ValueError):
        return {}


def _foreground_existing():
    """已有一个实例在运行时，把它的窗口拉到前台（解决最小化/被遮挡时“打不开”的错觉）。"""
    user32 = ctypes.windll.user32
    found = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def _cb(hwnd, _lparam):
        buf = ctypes.create_unicode_buffer(64)
        user32.GetWindowTextW(hwnd, buf, 64)
        if 'Edge 阅读器' in buf.value:
            found.append(hwnd)
        return True

    user32.EnumWindows(_cb, None)
    if found:
        user32.ShowWindow(found[0], 9)  # SW_RESTORE
        user32.SetForegroundWindow(found[0])
        return True
    return False


def main():
    if not _acquire_single_instance():
        if _foreground_existing():
            return
        ctypes.windll.user32.MessageBoxW(0, 'Edge 阅读器已经在运行了。', '提示', 0x40)
        return
    api = ReaderApi()
    geom = _read_prefs().get('window') or {}
    # 屏幕外 / 超屏的坏几何（历史遗留的最小化坐标等）回退默认
    sw, sh = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)
    if (geom.get('w', 0) > sw or geom.get('h', 0) > sh
            or geom.get('x', 0) < -20000 or geom.get('y', 0) < -20000):
        geom = {}
    window = webview.create_window(
        'Edge 阅读器', resource_path(os.path.join('web', 'index.html')), js_api=api,
        width=geom.get('w', 1040), height=geom.get('h', 780),
        x=geom.get('x'), y=geom.get('y'),
        min_size=(780, 560),
        text_select=True,  # 允许选中正文复制（pywebview 默认全页禁选）
    )
    global _window
    _window = window
    try:
        old = sys.executable + '.old'
        if getattr(sys, 'frozen', False) and os.path.exists(old):
            os.remove(old)
    except Exception:
        pass

    def _save_geometry():
        if _fs:
            return  # 全屏中不保存
        try:
            # 最小化(1)/最大化(2)状态下不保存（否则会把 -32000 这类假坐标存下来）
            native = getattr(window, 'native', None)
            if native is not None and int(native.WindowState) in (1, 2):
                return
        except Exception:
            pass
        try:
            if window.x is None or window.y is None or window.x < -20000 or window.y < -20000:
                return
            db = api._load_db()
            db.setdefault('prefs', {})['window'] = {
                'x': window.x, 'y': window.y, 'w': window.width, 'h': window.height,
            }
            api._save_db(db)
        except Exception:
            pass

    window.events.closing += _save_geometry
    webview.start(debug=False)


if __name__ == '__main__':
    main()
