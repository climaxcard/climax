@echo off
setlocal
cd /d "%~dp0"

REM ===== 設定 =====
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "OUT_DIR=docs"
set "PER_PAGE=80"

REM Excel 候補（どれか存在するものを採用）
set "EXCEL_DEFAULT=buylist.xlsx"
set "EXCEL_ALT=買取読み込みファイル.xlsx"
set "EXCEL_FALLBACK=C:\Users\user\Desktop\デュエマ買取表\買取読み込みファイル.xlsx"

REM 既定：サムネあり＆Gitあり
set "BUILD_THUMBS=1"
set "DO_GIT=1"
set "FORCE=0"

REM オプション:
REM   fast  = サムネOFF（高速）
REM   nogit = Git処理スキップ
REM   force = Excelの変更有無に関わらず生成を実行
for %%A in (%*) do (
  if /I "%%~A"=="fast"  set "BUILD_THUMBS=0"
  if /I "%%~A"=="nogit" set "DO_GIT=0"
  if /I "%%~A"=="force" set "FORCE=1"
)

REM === Excel 自動検出 ===
set "EXCEL_PATH="
if exist "%EXCEL_DEFAULT%"  set "EXCEL_PATH=%EXCEL_DEFAULT%"
if not defined EXCEL_PATH if exist "%EXCEL_ALT%"      set "EXCEL_PATH=%EXCEL_ALT%"
if not defined EXCEL_PATH if exist "%EXCEL_FALLBACK%" set "EXCEL_PATH=%EXCEL_FALLBACK%"

if not defined EXCEL_PATH (
  echo [NG] Excelが見つかりません。
  echo      %CD%\%EXCEL_DEFAULT%
  echo      %CD%\%EXCEL_ALT%
  echo      %EXCEL_FALLBACK%
  pause & exit /b 1
)

echo [*] Mode: BUILD_THUMBS=%BUILD_THUMBS% PER_PAGE=%PER_PAGE% DO_GIT=%DO_GIT% FORCE=%FORCE%
echo [*] Excel: %EXCEL_PATH%

REM === 変更チェック（force=1 ならスキップせず生成） ===
set "NEW_HASH="
for /f "tokens=1,*" %%H in ('certutil -hashfile "%EXCEL_PATH%" SHA256 ^| findstr /R "^[0-9A-F]"') do set "NEW_HASH=%%H"
if not defined NEW_HASH (
  echo [NG] ハッシュ取得失敗: %EXCEL_PATH%
  pause & exit /b 1
)

set "OLD_HASH="
if exist ".publish.hash" set /p OLD_HASH=<.publish.hash

if "%FORCE%"=="1" (
  echo [!] FORCE有効 → 生成を実行
) else if /I "%NEW_HASH%"=="%OLD_HASH%" (
  echo [i] Excelに変更なし → 生成スキップ
  goto :maybe_git
)

echo %NEW_HASH%>.publish.hash
echo [*] Generate docs...
set "BUILD_THUMBS=%BUILD_THUMBS%"
set "OUT_DIR=%OUT_DIR%"
set "PER_PAGE=%PER_PAGE%"
set "EXCEL_PATH=%EXCEL_PATH%"
"%PY%" gen_buylist.py || goto :fail

:maybe_git
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
