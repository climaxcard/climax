@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM ---- Git の実体を特定（フルパス）----
for /f "delims=" %%G in ('where git 2^>nul') do set "GIT=%%G"
if not defined GIT (
  REM 代表的な既定パス（必要なら追記）
  if exist "C:\Program Files\Git\cmd\git.exe" set "GIT=C:\Program Files\Git\cmd\git.exe"
)
if not defined GIT (
  echo [ERR] git が見つかりません。Git for Windows をインストールしてください。
  goto :END
)

REM === 設定 ==========================
set "PY=python"
set "EXCEL_PATH=C:\Users\user\OneDrive\Desktop\デュエマ買取表\buylist.xlsx"
set "OUT_DIR=docs"
REM set "SHEET_NAME=シート1"
REM set "PER_PAGE=80"
REM set "BUILD_THUMBS=1"
REM ===================================

REM ▼ リポジトリルートへ
for /f "delims=" %%A in ('"%GIT%" rev-parse --show-toplevel 2^>nul') do set "REPO_ROOT=%%A"
if not defined REPO_ROOT (
  pushd "%~dp0" >nul 2>&1
  for /f "delims=" %%A in ('"%GIT%" rev-parse --show-toplevel 2^>nul') do set "REPO_ROOT=%%A"
  popd >nul 2>&1
)
if not defined REPO_ROOT (
  echo [ERR] Gitリポジトリが見つかりません。
  goto :END
)
cd /d "%REPO_ROOT%" || (echo [ERR] ルートに移動できません & goto :END)

REM ▼ ガード
"%GIT%" rebase --abort >nul 2>&1

for /f "delims=" %%A in ('"%GIT%" config --get user.email 2^>nul') do set GITEMAIL=%%A
if not defined GITEMAIL (
  "%GIT%" config user.name  "CLIMAX"
  "%GIT%" config user.email "you@example.com"
)

echo [1/6] 依存確認
"%PY%" -m pip install -q --upgrade pip || goto :FAIL_PIP_UP
"%PY%" -m pip install -q pandas openpyxl requests xlrd==2.0.1 pillow || goto :FAIL_PIP_PKGS

echo [2/6] ビルド実行
set "OUT_DIR=%OUT_DIR%"
set "EXCEL_PATH=%EXCEL_PATH%"
if defined SHEET_NAME set "SHEET_NAME=%SHEET_NAME%"
if defined PER_PAGE   set "PER_PAGE=%PER_PAGE%"
if defined BUILD_THUMBS set "BUILD_THUMBS=%BUILD_THUMBS%"

"%PY%" generate_buylist.py || goto :FAIL_BUILD

REM ▼ Pages再デプロイ用の強制差分
mkdir "%OUT_DIR%" 2>nul
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a %%b %%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
> "%OUT_DIR%\.buildstamp" echo %TODAY% %NOW%

REM ▼ サムネ掃除
rmdir /s /q "%OUT_DIR%\assets\thumbs" 2>nul

echo [3/6] 変更検出（削除も含めてステージ）
"%GIT%" add -A "%OUT_DIR%"

echo -- git status (short) --
"%GIT%" status -s "%OUT_DIR%"
echo --------------------------------

"%GIT%" diff --cached --quiet
if not errorlevel 1 (
  echo [INFO] 変更なし（%OUT_DIR% に差分がありません）
  goto :END
)

echo [4/6] コミット
set "MSG=build: buylist pages (%TODAY% %NOW%)"
"%GIT%" commit -m "%MSG%" || goto :FAIL_COMMIT

echo [5/6] リモート同期（pull --rebase）
"%GIT%" pull --rebase --autostash || goto :FAIL_PULL

echo [6/6] プッシュ
for /f "delims=" %%b in ('"%GIT%" rev-parse --abbrev-ref HEAD') do set CURBR=%%b
"%GIT%" rev-parse --abbrev-ref --symbolic-full-name "@{u}" >nul 2>&1
if errorlevel 1 (
  echo [INFO] 追跡ブランチ未設定: origin %CURBR% を upstream にします
  "%GIT%" push -u origin "%CURBR%" || goto :FAIL_PUSH
) else (
  "%GIT%" push || goto :FAIL_PUSH
)
echo [OK] 完了。GitHub Pages: Settings→Pages が「main /docs」か要確認。
goto :END

:FAIL_PIP_UP
echo [ERR] pip upgrade 失敗
goto :END
:FAIL_PIP_PKGS
echo [ERR] 依存パッケージのインストール失敗
goto :END
:FAIL_BUILD
echo [ERR] ✖ ビルド失敗。generate_buylist.py の実行やパスを確認してください。
goto :END
:FAIL_COMMIT
echo [ERR] コミット失敗
goto :END
:FAIL_PULL
echo [ERR] pull 失敗（競合時は手動解決が必要）
goto :END
:FAIL_PUSH
echo [ERR] push 失敗（先に pull が必要 or 権限/ネットワーク）
goto :END

:END
echo.
echo ==== 実行終了 ====
echo Git: %GIT%
echo Repo: %REPO_ROOT%
echo 日時: %TODAY% %NOW%
echo （このウィンドウは閉じません。何かキーを押すと閉じます）
pause
endlocal
exit /b 0
