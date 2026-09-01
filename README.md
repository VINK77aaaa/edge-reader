# Edge 阅读器 📖

一个 Windows 本地小工具：把任意文字粘贴进去，用舒服的排版沉浸式阅读，并用 **Edge「大声朗读」同款在线神经网络语音**（晓晓、云健等）读出来。

![Python](https://img.shields.io/badge/Python-3.9%2B-blue) ![平台](https://img.shields.io/badge/平台-Windows%2010%2F11-lightgrey) ![UI](https://img.shields.io/badge/界面-pywebview%20%2B%20WebView2-green)

## ✨ 功能

- **粘贴即读**：从网页、Word、PDF 等任意地方复制文字，粘贴后一键进入阅读视图
- **段落智能整理**：自动合并被硬换行打断的句子，网页/PDF 复制的文字排版也正常
- **沉浸式排版**：字号调节、行距三档、浅色 / 米黄 / **GitHub Dark** 三种主题
- **语音朗读**：
  - 🌐 在线语音（默认，需联网）：晓晓（女·温柔）、云健（男·浑厚）、云希、云扬、晓伊，基于 [edge-tts](https://github.com/rany2/edge-tts) 调用微软神经网络语音，带本地缓存
  - 💻 本机语音（离线）：Microsoft Huihui / Zira 等 SAPI 语音，断网可用
  - **逐句跟读高亮**：读到哪句亮哪句，自动滚动跟随；点击任意句子可从该处开始朗读
  - 语速可调（50%–180%），暂停 / 继续 / 停止
- **历史记录**：文章自动保存，支持删除
- **继续上次阅读**：恢复上次阅读的位置和进度百分比

## 🚀 快速开始

```bat
:: 1. 克隆本仓库
git clone https://github.com/VINK77aaaa/edge-reader.git
cd edge-reader

:: 2. 双击「启动阅读器.bat」
::    首次运行会自动安装 pywebview / pywin32 / edge-tts
```

或手动安装依赖后运行：

```bat
py -m pip install pywebview pywin32 edge-tts
py app.py
```

要求：Windows 10/11（需自带 WebView2 运行时，装了 Edge 就有）+ Python 3.9+。

## 📁 项目结构

```
├── app.py              # 主程序：pywebview 窗口、文章存储、SAPI/在线语音引擎
├── web/index.html      # 界面：首页 + 阅读视图（内联 CSS/JS，无前端框架）
├── 启动阅读器.bat       # 一键启动（自动装依赖）
└── 使用说明.md          # 详细使用说明
```

## 🔒 隐私

文章与设置只保存在本机 `%APPDATA%\EdgeReader\`，除在线语音合成请求外不联网、不上传任何数据。

## 📄 许可

MIT
