@echo off
chcp 65001 >nul
title 启动 Excel 智能朗读核对助手...
cd /d "%~dp0"
echo 正在启动 Excel 智能朗读核对助手...
start pythonw excel_reader.py
exit
