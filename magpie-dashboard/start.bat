@echo off
chcp 65001 >nul
cd /d %~dp0
echo 看板: http://127.0.0.1:8788
start "" http://127.0.0.1:8788
python server.py
