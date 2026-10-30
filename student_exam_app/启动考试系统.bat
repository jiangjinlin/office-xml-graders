@echo off
chcp 65001 >nul
echo ============================================
echo    🦁 小狮子 Office 2010 操作题考试系统
echo    Excel 1-5 + Word 1-5 + PPT 1-5
echo ============================================
echo.
echo 启动中...
echo 浏览器打开: http://localhost:5000
echo 按 Ctrl+C 停止服务器
echo.
cd /d "%~dp0"
python app.py
pause
