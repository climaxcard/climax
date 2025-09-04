@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

REM === 設定 ==========================
set "REPO_DIR=%~dp0"
set "PY=python"
set "EXCEL_PATH=C:\Users\user\OneDrive\Desktop\デュエマ買取表\buylist.xlsx"
set "OUT_DIR=docs"
REM set "SHEET_NAME=シート1"
REM set "PER_PAGE=80"
REM set "BUILD_THUMBS=1"
REM ===================================

cd /d "%REPO_DIR%" || (echo リポジトリに移動できません & exit /b 1)

echo [1/5] 依存確認
"%PY%" -m pip install -q --upgrade pip
"%PY%" -m pip install -q pandas openpyxl requests xlrd==2.0.1 pillow

echo [2/5] ビルド実行
set "OUT_DIR=%OUT_DIR%"
set "EXCEL_PATH=%EXCEL_PATH%"
if defined SHEET_NAME set "SHEET_NAME=%SHEET_NAME%"
if defined PER_PAGE   set "PER_PAGE=%PER_PAGE%"
if defined BUILD_THUMBS set "BUILD_THUMBS=%BUILD_THUMBS%"

"%PY%" generate_buylist.py
if errorlevel 1 (
  echo ✖ ビルド失敗。ログを確認してください。
  exit /b 1
)

echo [3/5] 変更検出（削除も含めてステージ）
REM 削除も拾うため -A を使う
git add -A "%OUT_DIR%"

echo -- git status (short) --
git status -s "%OUT_DIR%"
echo --------------------------------

REM 差分の有無を確認（0=差分なし, 1=差分あり）
git diff --cached --quiet
if not errorlevel 1 (
  echo 変更なし（%OUT_DIR% に差分がありません）
  goto :END
)

echo [4/5] コミット
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a %%b %%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
set "MSG=build: buylist pages (%TODAY% %NOW%)"
git commit -m "%MSG%"

echo [5/5] プッシュ
for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do set CURBR=%%b
REM 追跡ブランチが無い場合に備える
git rev-parse --abbrev-ref --symbolic-full-name "@{u}" >nul 2>&1
if errorlevel 1 (
  echo 追跡ブランチ未設定のため upstream を設定して push します: origin %CURBR%
  git push -u origin "%CURBR%"
) else (
  git push
)

:END
echo 完了
endlocal
