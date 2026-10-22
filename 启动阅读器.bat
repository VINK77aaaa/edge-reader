@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -c "import webview, win32com.client, docx" 2>nul
if errorlevel 1 (
    echo 首次运行或缺少依赖，正在安装 pywebview / pywin32 / python-docx ...
    py -m pip install pywebview pywin32 python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple
)
py -c "import edge_tts, importlib.metadata as md; assert tuple(map(int, md.version('edge_tts').split('.')[:2])) >= (7, 3)" 2>nul
if errorlevel 1 (
    echo 升级 edge-tts（旧版本会报 403）...
    py -m pip install -U edge-tts -i https://pypi.tuna.tsinghua.edu.cn/simple
)
py app.py
if errorlevel 1 pause
