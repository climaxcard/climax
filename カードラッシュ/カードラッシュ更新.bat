@echo off
REM 日本語パス対策：コンソールを CP932 に
chcp 932 >nul

REM ===== 固定パス =====
set "BASEDIR=C:\Users\user\OneDrive\Desktop\デュエマ買取表\カードラッシュ"
set "PYFILE=cardrush_to_sheet.py"
set "CREDS=%BASEDIR%\credentials.json"
set "SHEET_URL=https://docs.google.com/spreadsheets/d/1gYYmzLkrtAgNZB6dlwzFEqg0VXT7o0QKlFVPS2ln5Xs/edit?gid=0"
set "SHEET_NAME=CardRush_DM"
set "LIMIT=100"
set "MAX_PAGES=200"
set "SLEEP_MS=350"

echo === CardRush（DM）→ Googleスプレッドシート更新 ===

REM パス存在チェック & 8.3短縮名取得（日本語安全対策）
if not exist "%BASEDIR%" (
  echo [ERROR] フォルダが見つかりません: "%BASEDIR%"
  pause & exit /b 1
)
for %%I in ("%BASEDIR%") do set "BASEDIR_S=%%~sI"

REM 作業ディレクトリ移動（短縮名で安全に）
pushd "%BASEDIR_S%" || (echo [ERROR] cd に失敗しました & pause & exit /b 1)

REM python 存在確認
where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] python が見つかりません。PATHを通すかフルパス指定にしてください。
  pause & exit /b 1
)

REM スクリプト確認
if not exist "%PYFILE%" (
  echo [ERROR] スクリプトがありません: "%CD%\%PYFILE%"
  echo このフォルダの .py 一覧:
  dir /b *.py
  popd
  pause & exit /b 1
)

REM 認証JSONは警告のみ
if not exist "%CREDS%" (
  echo [WARN] 認証JSONが見つかりません: "%CREDS%"
  echo そのまま実行を試みます…
)

REM 実行
python "%PYFILE%" ^
  --sheet-url "%SHEET_URL%" ^
  --sheet-name "%SHEET_NAME%" ^
  --creds "%CREDS%" ^
  --limit %LIMIT% ^
  --max-pages %MAX_PAGES% ^
  --sleep-ms %SLEEP_MS%

echo.
echo ==== 完了 ====
popd
pause
