# -*- coding: utf-8 -*-
"""Edge 阅读器 —— 粘贴文字，享受沉浸式阅读 + Windows 系统语音朗读。

运行依赖: pywebview（界面）、pywin32（SAPI 语音朗读）。
数据保存在 %APPDATA%\\EdgeReader\\articles.json。
"""
import base64
import hashlib
import json
import os
import threading
import time
import uuid

import webview

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
        global _rate, _voice_name, _tts_thread
        global _edge_segs, _edge_paths, _edge_gen_done, _edge_error
        if voice_name and voice_name.startswith('edge:'):
            try:
                import edge_tts  # noqa: F401
            except ImportError:
                self._notify({'type': 'error',
                              'message': '使用在线语音需先安装 edge-tts: py -m pip install edge-tts'})
                return False
            self.tts_stop()
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
                target=self._edge_gen_loop, args=(segs, voice_name[5:], _rate), daemon=True)
            _tts_thread.start()
            return True

        if not HAS_SAPI:
            self._notify({'type': 'error',
                          'message': '缺少 pywin32，无法朗读。请运行: py -m pip install pywin32'})
            return False
        self.tts_stop()
        _stop_evt.clear()
        _pause_evt.clear()
        _rate = clamp_rate(rate)
        _voice_name = voice_name or ''
        segs = [{'p': int(s.get('p', 0)), 't': str(s.get('t', '')).strip()}
                for s in segments or [] if str(s.get('t', '')).strip()]
        if not segs:
            return False
        _tts_thread = threading.Thread(target=self._speak_loop, args=(segs,), daemon=True)
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

    def _edge_gen_loop(self, segments, voice_id, rate):
        """Edge 在线语音模式：后台逐段合成 mp3（带缓存），前端轮询 tts_get_audio 取走播放。"""
        global _edge_gen_done, _edge_error
        import asyncio

        import edge_tts

        os.makedirs(TTS_CACHE_DIR, exist_ok=True)
        pct = rate * 10
        rate_str = ('+%d%%' % pct) if pct >= 0 else ('%d%%' % pct)

        async def synth(text, path):
            com = edge_tts.Communicate(text, voice_id, rate=rate_str)
            await asyncio.wait_for(com.save(path), timeout=90)

        self._notify({'type': 'state', 'state': 'speaking'})
        try:
            for idx, seg in enumerate(segments):
                if _stop_evt.is_set():
                    break
                key = hashlib.md5((voice_id + rate_str + seg['t']).encode('utf-8')).hexdigest()
                path = os.path.join(TTS_CACHE_DIR, key + '.mp3')
                if not os.path.exists(path):
                    asyncio.run(synth(seg['t'], path))
                if _stop_evt.is_set():
                    break
                with _edge_lock:
                    _edge_paths[idx] = path
            if not _stop_evt.is_set():
                self._notify({'type': 'gen_done'})
        except Exception as exc:
            with _edge_lock:
                _edge_gen_done = True
                _edge_error = '在线语音合成失败（请检查网络）: %s' % exc
            self._notify({'type': 'error',
                          'message': '在线语音合成失败（请检查网络）: %s' % exc})
        finally:
            with _edge_lock:
                _edge_gen_done = True

    def _speak_loop(self, segments):
        pythoncom.CoInitialize()
        try:
            voice = win32com.client.Dispatch('SAPI.SpVoice')
            self._apply_voice(voice)
            voice.Rate = _rate
            self._notify({'type': 'state', 'state': 'speaking'})
            for idx, seg in enumerate(segments):
                if _stop_evt.is_set():
                    break
                voice.Speak(seg['t'], SPF_ASYNC)
                self._notify({'type': 'segment', 'index': idx, 'p': seg['p']})
                was_paused = False
                while True:
                    if _stop_evt.is_set():
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
            if not _stop_evt.is_set():
                self._notify({'type': 'state', 'state': 'finished'})
            else:
                self._notify({'type': 'state', 'state': 'stopped'})
        except Exception as exc:
            self._notify({'type': 'error', 'message': '朗读出错: %s' % exc})
        finally:
            pythoncom.CoUninitialize()


def _acquire_single_instance():
    """用 Windows 互斥量保证同一时间只有一个实例运行。"""
    import ctypes
    ctypes.windll.kernel32.CreateMutexW(None, False, 'EdgeReader_SingleInstance_Mutex')
    return ctypes.windll.kernel32.GetLastError() != 183  # 183 = ERROR_ALREADY_EXISTS


def main():
    if not _acquire_single_instance():
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, 'Edge 阅读器已经在运行了。', '提示', 0x40)
        return
    api = ReaderApi()
    window = webview.create_window(
        'Edge 阅读器', 'web/index.html', js_api=api,
        width=1040, height=780, min_size=(780, 560),
    )
    webview.start(debug=False)


if __name__ == '__main__':
    main()
