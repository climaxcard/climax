@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

REM === 設定（必要に応じて編集） ==========================
set "REPO_DIR=%~dp0"
set "PY=python"
set "EXCEL_PATH=C:\Users\user\OneDrive\Desktop\デュエマ買取表\buylist.xlsx"
set "OUT_DIR=docs"
REM set "SHEET_NAME=シート1"  REM 先頭シートでOKならコメントのままで
REM set "PER_PAGE=80"
REM set "BUILD_THUMBS=1"
REM =======================================================

cd /d "%REPO_DIR%" || (echo リポジトリに移動できません & exit /b 1)

echo [1/4] 依存確認（初回だけ少し時間がかかります）
"%PY%" -m pip install -q --upgrade pip
"%PY%" -m pip install -q pandas openpyxl requests xlrd==2.0.1 pillow

echo [2/4] ビルド実行
set "OUT_DIR=%OUT_DIR%"
set "EXCEL_PATH=%EXCEL_PATH%"
if defined SHEET_NAME set "SHEET_NAME=%SHEET_NAME%"
if defined PER_PAGE   set "PER_PAGE=%PER_PAGE%"
if defined BUILD_THUMBS set "BUILD_THUMBS=%BUILD_THUMBS%"

"%PY%" generate_buylist.py
if errorlevel 1 (
  echo ビルドに失敗しました。ログを確認してください。
  exit /b 1
)

echo [3/4] 変更検出
git add docs
git diff --cached --quiet
if not errorlevel 1 (
  echo 変更なし（docsに差分はありません）
  goto :END
)

echo [4/4] コミット & プッシュ
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a %%b %%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
git commit -m "build: buylist pages (%TODAY% %NOW%)"
git push

:END
echo 完了
endlocal
