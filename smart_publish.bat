@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

rem === (必要なら) UTF-8表示に切替。文字化けするならこの行を残して、ファイルはUTF-8で保存。
rem === 逆にANSI(CP932)保存するなら、この行は消してください。
chcp 65001 >nul

rem === Python のパス ===
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"

rem === 出力やビルド設定 ===
set "OUT_DIR=docs"
set "PER_PAGE=80"
set "BUILD_THUMBS=0"
set "DO_GIT=1"

rem === Excelパス 受け取り & 自動検出 ===
set "EXCEL_PATH="

if not "%~1"=="" (
  rem 引数があれば正規化（ドラッグ&ドロップやスペース対応）
  for %%F in ("%~1") do (
    if exist "%%~fF" set "EXCEL_PATH=%%~fF"
  )
)

if not defined EXCEL_PATH (
  rem よくある配置を探索（拡張子も許容）
  for %%E in (xlsx xlsm csv) do (
    if not defined EXCEL_PATH if exist "%CD%\buylist.%%E" set "EXCEL_PATH=%CD%\buylist.%%E"
    if not defined EXCEL_PATH if exist "%CD%\data\buylist.%%E" set "EXCEL_PATH=%CD%\data\buylist.%%E"
  )
)

if not defined EXCEL_PATH (
  echo [NG] Excelが見つかりません。フルパスを渡すか、buylist.xlsx/xlsm/csv を配置してください。
  echo 例: publish_simple.bat "C:\full\path\buylist.xlsx"
  pause & exit /b 1
)

if not exist "%EXCEL_PATH%" (
  echo [NG] 指定ファイルが存在しません:
  echo   "%EXCEL_PATH%"
  pause & exit /b 1
)

echo [*] Excel: "%EXCEL_PATH%"
echo [*] PER_PAGE=%PER_PAGE%  BUILD_THUMBS=%BUILD_THUMBS%

rem === 生成 ===
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%" >nul 2>&1
set "OUT_DIR=%OUT_DIR%"
set "PER_PAGE=%PER_PAGE%"
set "BUILD_THUMBS=%BUILD_THUMBS%"

rem gen_buylist.py と同じフォルダ想定。場所が違うならパス修正してください。
"%PY%" "%~dp0gen_buylist.py" "%EXCEL_PATH%" || goto :fail

rem === ビルドスタンプ ===
> "%OUT_DIR%\.build_stamp.txt" echo built at %date% %time% from "%EXCEL_PATH%"

rem === Git（任意） ===
where git >nul 2>&1 || set "DO_GIT=0"
if "%DO_GIT%"=="1" (
  echo [*] Git pull/commit/push...
  for %%L in (".git\index.lock" ".git\shallow.lock" ".git\packed-refs.lock" ".git\logs\HEAD.lock") do (
    if exist "%%~L" del /f /q "%%~L"
  )

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

rem === 完成ページを開く ===
if exist ".\docs\default\index.html" (
  start "" ".\docs\default\index.html"
) else if exist ".\docs\index.html" (
  start "" ".\docs\index.html"
)

exit /b 0

:fail
echo [NG] 生成失敗。上のログを確認してください。
pause
exit /b 1
