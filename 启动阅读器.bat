@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -c "import webview, win32com.client, edge_tts" 2>nul
if errorlevel 1 (
    echo 首次运行或缺少依赖，正在安装 pywebview / pywin32 / edge-tts ...
    py -m pip install pywebview pywin32 edge-tts -i https://pypi.tuna.tsinghua.edu.cn/simple
)
py app.py
if errorlevel 1 pause
