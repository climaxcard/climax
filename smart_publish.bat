@echo off
setlocal
cd /d "%~dp0"

REM ===== 設定 =====
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "EXCEL_PATH=buylist.xlsx"
set "OUT_DIR=docs"
set "PER_PAGE=80"

REM 既定：サムネあり＆Gitあり
set "BUILD_THUMBS=1"
set "DO_GIT=1"

REM オプション:
REM   fast  = サムネOFF（高速）
REM   nogit = Git処理スキップ
for %%A in (%*) do (
  if /I "%%~A"=="fast"  set "BUILD_THUMBS=0"
  if /I "%%~A"=="nogit" set "DO_GIT=0"
)

echo [*] Mode: BUILD_THUMBS=%BUILD_THUMBS% PER_PAGE=%PER_PAGE% DO_GIT=%DO_GIT%

REM === Excel変更チェック（同じなら生成スキップ） ===
set "NEW_HASH="
for /f "tokens=1,*" %%H in ('certutil -hashfile "%EXCEL_PATH%" SHA256 ^| findstr /R "^[0-9A-F]"') do set "NEW_HASH=%%H"
if not defined NEW_HASH (
  echo [NG] Excelが見つからないか、ハッシュ取得失敗: %EXCEL_PATH%
  pause & exit /b 1
)

set "OLD_HASH="
if exist ".publish.hash" set /p OLD_HASH=<.publish.hash

if /I "%NEW_HASH%"=="%OLD_HASH%" (
  echo [i] Excelに変更なし → 生成スキップ
) else (
  echo %NEW_HASH%>.publish.hash
  echo [*] Generate docs...
  set "BUILD_THUMBS=%BUILD_THUMBS%"
  set "OUT_DIR=%OUT_DIR%"
  set "PER_PAGE=%PER_PAGE%"
  set "EXCEL_PATH=%EXCEL_PATH%"
  "%PY%" gen_buylist.py || goto :fail
)

if "%DO_GIT%"=="1" (
  echo [*] Commit and push...
  if exist ".git\index.lock" del /f /q ".git\index.lock"
  for %%L in (".git\shallow.lock" ".git\packed-refs.lock" ".git\logs\HEAD.lock") do if exist "%%~L" del /f /q "%%~L"
  git gc --prune=now >nul 2>&1

  git add -A || goto :fail
  git diff --cached --quiet && echo [i] 変更なし（commitスキップ） || git commit -m "update buylist %date% %time:~0,5%"
  git pull --rebase --autostash origin main || goto :fail
  git push origin main || goto :fail
  echo [OK] 公開反映 完了
) else (
  echo [i] Git処理スキップ（nogit）
)

pause
exit /b 0

:fail
echo [NG] エラーが発生しました。上のログを確認してください。
pause
exit /b 1
