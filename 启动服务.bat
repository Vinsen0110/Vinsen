@echo off
chcp 65001 >nul
title Nano Banana PS Bridge Server

echo =======================================================
echo 正在为您检查并安装必要的 Python 运行库（如已安装会自动跳过）...
echo =======================================================
python -m pip install flask flask-cors pywin32 requests

echo.
echo =======================================================
echo 正在启动 Photoshop 桥接服务...
echo （如果报错说找不到 python，请确保您在安装 Python 时勾选了 "Add Python to PATH"）
echo =======================================================
python ps_server.py

echo.
echo =======================================================
echo ⚠️ 服务已停止或发生错误。请查看上方红字/错字。
echo =======================================================
pause
