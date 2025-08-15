@echo off
setlocal
cd /d "%~dp0"

rem ===== Settings =====
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "OUT_DIR=docs"
set "PER_PAGE=80"
set "BUILD_THUMBS=0"

rem オプション: 第1引数にExcelのフルパスを渡せます（未指定なら自動検出）
if not "%~1"=="" set "EXCEL_PATH=%~1"

"%PY%" gen_buylist.py || goto :fail

start "" ".\docs\default\index.html"
goto :eof

:fail
echo [NG] 生成失敗。ログを確認してください。
pause
exit /b 1
