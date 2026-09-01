@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在安装构建依赖 ...
py -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
echo 正在构建 EdgeReader.exe ...
py -m PyInstaller --noconfirm --clean --windowed --onefile --name EdgeReader ^
    --icon EdgeReader.ico ^
    --add-data "web;web" ^
    --hidden-import webview.platforms.winforms ^
    --hidden-import edge_tts --hidden-import aiohttp --hidden-import docx ^
    app.py
if exist dist\EdgeReader.exe (
    echo 构建完成: dist\EdgeReader.exe
) else (
    echo 构建失败，请检查上方错误信息。
)
pause
