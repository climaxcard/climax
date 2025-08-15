@echo off
setlocal
cd /d "%~dp0"

set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "EXCEL_PATH=buylist.xlsx"
set "OUT_DIR=docs"
set "PER_PAGE=80"
set "BUILD_THUMBS=0"

"%PY%" gen_buylist.py || goto :fail
start "" ".\docs\default\p1.html"
goto :eof

:fail
echo [NG] 生成失敗。ログを確認してください。
pause
exit /b 1
