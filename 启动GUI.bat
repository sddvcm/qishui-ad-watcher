@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
set "PYTHON=C:\Users\Administrator\AppData\Roaming\TRAE SOLO CN\ModularData\ai-agent\vm\tools\python\python.exe"
title 汽水音乐自动看广告 - GUI
"%PYTHON%" -X utf8 qishui_gui.py
if errorlevel 1 (
    echo.
    echo 运行出错，请检查依赖是否安装完整
    echo 可运行: pip install PyQt5 rapidocr-onnxruntime pywin32 keyboard mss Pillow numpy
    pause
)