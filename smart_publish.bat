@echo off
setlocal
cd /d "%~dp0"

set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "OUT_DIR=docs"
set "PER_PAGE=80"
set "BUILD_THUMBS=0"
set "DO_GIT=1"

rem 引数1: Excelのフルパス（未指定なら既存の buylist.xlsx を自動検出）
set "EXCEL_PATH="
if not "%~1"=="" if exist "%~1" if /I "%~x1"==".xlsx" set "EXCEL_PATH=%~1"
if not defined EXCEL_PATH if exist "buylist.xlsx" set "EXCEL_PATH=%CD%\buylist.xlsx"
if not defined EXCEL_PATH if exist "data\buylist.xlsx" set "EXCEL_PATH=%CD%\data\buylist.xlsx"

if not defined EXCEL_PATH (
  echo [NG] Excelが見つかりません。フルパスを渡すか、buylist.xlsx を置いてください。
  echo 例: publish_simple.bat "C:\full\path\buylist.xlsx"
  pause & exit /b 1
)

echo [*] Excel: "%EXCEL_PATH%"
echo [*] PER_PAGE=%PER_PAGE% BUILD_THUMBS=%BUILD_THUMBS%

rem --- 毎回ビルド（変更検知しない） ---
set "OUT_DIR=%OUT_DIR%"
set "PER_PAGE=%PER_PAGE%"
set "BUILD_THUMBS=%BUILD_THUMBS%"
"%PY%" gen_buylist.py "%EXCEL_PATH%" || goto :fail

rem --- ページ向けスタンプ（軽い差分を常に作る） ---
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"
> "%OUT_DIR%\.build_stamp.txt" echo built at %date% %time% from "%EXCEL_PATH%"

rem --- Git（任意） ---
where git >nul 2>&1 || set "DO_GIT=0"
if "%DO_GIT%"=="1" (
  echo [*] Git pull/commit/push...
  if exist ".git\index.lock" del /f /q ".git\index.lock"
  for %%L in (".git\shallow.lock" ".git\packed-refs.lock" ".git\logs\HEAD.lock") do if exist "%%~L" del /f /q "%%~L"

  git fetch origin || goto :fail
  git pull --rebase --autostash origin main || goto :fail

  git add -A || goto :fail
  git diff --cached --quiet && (
    echo [i] 変更なし（commit省略）
  ) || (
    git commit -m "update buylist %date% %time:~0,5%" || goto :fail
  )

  git push origin main || goto :fail
  echo [OK] publish done
) else (
  echo [i] Git skipped
)

start "" ".\docs\default\index.html"
exit /b 0

:fail
echo [NG] 生成失敗。上のログを確認してください。
pause
exit /b 1
