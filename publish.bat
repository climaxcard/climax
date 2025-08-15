@echo off
setlocal
REM ========= カレントをこのbatの場所へ =========
cd /d "%~dp0"

REM ========= 設定 =========
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "EXCEL_DEFAULT=buylist.xlsx"
set "EXCEL_FALLBACK=C:\Users\user\Desktop\デュエマ買取表\buylist.xlsx"
set "OUT_DIR=docs"
set "PER_PAGE=80"

REM 引数: fast でサムネOFF（高速） / 既定はサムネON
set "BUILD_THUMBS=1"
if /I "%~1"=="fast" set "BUILD_THUMBS=0"

echo [*] Mode: BUILD_THUMBS=%BUILD_THUMBS%  PER_PAGE=%PER_PAGE%
echo.

REM ========= Gitロック掃除 =========
if exist ".git\index.lock" del /f /q ".git\index.lock"
for %%L in (".git\shallow.lock" ".git\packed-refs.lock" ".git\logs\HEAD.lock") do if exist "%%~L" del /f /q "%%~L"
git gc --prune=now >nul 2>&1

REM ========= Excelの場所を自動判定 =========
set "EXCEL_PATH="
if exist "%EXCEL_DEFAULT%" set "EXCEL_PATH=%EXCEL_DEFAULT%"
if not defined EXCEL_PATH if exist "%EXCEL_FALLBACK%" set "EXCEL_PATH=%EXCEL_FALLBACK%"
if not defined EXCEL_PATH (
  echo [NG] Excelが見つかりません。
  echo      %CD%\%EXCEL_DEFAULT%
  echo      %EXCEL_FALLBACK%
  pause
  exit /b 1
)

REM ========= 生成 =========
echo [*] Generate docs...
set "EXCEL_PATH=%EXCEL_PATH%"
set "OUT_DIR=%OUT_DIR%"
set "PER_PAGE=%PER_PAGE%"
set "BUILD_THUMBS=%BUILD_THUMBS%"
"%PY%" gen_buylist.py || goto :fail

REM ========= Git 反映 =========
echo [*] Commit and push...
git add -A || goto :fail
git diff --cached --quiet && echo [i] 変更なし（commitスキップ） || git commit -m "update buylist %date% %time:~0,5%"
git pull --rebase --autostash origin main || goto :fail
git push origin main || goto :fail

echo [OK] 公開反映 完了
echo.
REM ローカル表示
start "" ".\docs\default\p1.html"
REM 公開URL（使うならコメント解除）
REM start "" "https://climaxcard.github.io/climax/default/p1.html"

pause
goto :eof

:fail
echo [NG] エラーが発生しました。上のログを確認してください。
pause
exit /b 1
