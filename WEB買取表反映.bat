@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM ============================================================
REM 0) Git の実体を特定（元の自動検出版）
REM ============================================================
for /f "delims=" %%G in ('where git 2^>nul') do set "GIT=%%G"
if not defined GIT (
  if exist "C:\Program Files\Git\cmd\git.exe" set "GIT=C:\Program Files\Git\cmd\git.exe"
)
if not defined GIT (
  echo [ERR] git が見つかりません
  goto :END
)

REM ============================================================
REM リポジトリルートへ移動
REM ============================================================
for /f "delims=" %%A in ('"%GIT%" rev-parse --show-toplevel 2^>nul') do set "REPO_ROOT=%%A"
if not defined REPO_ROOT (
  pushd "%~dp0" >nul 2>&1
  for /f "delims=" %%A in ('"%GIT%" rev-parse --show-toplevel 2^>nul') do set "REPO_ROOT=%%A"
  popd >nul 2>&1
)

if not defined REPO_ROOT (
  echo [ERR] Gitリポジトリが見つかりません
  goto :END
)

cd /d "%REPO_ROOT%" || (echo [ERR] ルート移動失敗 & goto :END)

REM ============================================================
REM 設定 ※REPO_ROOT確定後にセットする
REM ============================================================
set "PY=python"
set "EXCEL_PATH=%REPO_ROOT%\buylist.xlsm"
set "OUT_DIR=docs"
set "BUILD_THUMBS=0"

REM 日時（commitメッセージ・buildstamp用）
set "TODAY=%DATE%"
set "NOW=%TIME%"

REM TIMEは小数秒などが入る場合があるため、見やすく短縮
set "NOW=%NOW:~0,8%"

REM commitメッセージで使いにくい文字を置換
set "TODAY_SAFE=%TODAY:/=-%"
set "NOW_SAFE=%NOW::=-%"


echo [INFO] REPO_ROOT=%REPO_ROOT%
echo [INFO] EXCEL_PATH=%EXCEL_PATH%

if not exist "%EXCEL_PATH%" (
  echo [ERR] buylist.xlsm が見つかりません: %EXCEL_PATH%
  goto :END
)

REM ============================================================
REM [0/6] 依存確認 ※Python実行前に必ず行う
REM ============================================================
echo.
echo [0/6] 依存確認
"%PY%" -m pip install -q --upgrade pip
if errorlevel 1 goto :FAIL_PIP_UP

REM bs4エラー対策：beautifulsoup4 を追加
"%PY%" -m pip install -q beautifulsoup4 pandas openpyxl requests xlrd==2.0.1 pillow lxml
if errorlevel 1 goto :FAIL_PIP_PKGS

REM ============================================================
REM [1/6] CardRush取得
REM ============================================================
echo.
echo [1/6] CardRush取得
"%PY%" "%REPO_ROOT%\cardrush_to_excel.py" --file-path "%EXCEL_PATH%"
if errorlevel 1 (
  echo [ERR] cardrush_to_excel.py 失敗
  goto :END
)

REM ============================================================
REM [2/6] 買取価格更新
REM ============================================================
echo.
echo [2/6] 買取価格更新
"%PY%" "%REPO_ROOT%\値段更新.py"
if errorlevel 1 (
  echo [ERR] 値段更新.py 失敗
  goto :END
)

REM ============================================================
REM [3/6] WEBビルド
REM ============================================================
echo.
echo [3/6] WEBビルド
"%PY%" "%REPO_ROOT%\generate_buylist.py"
if errorlevel 1 goto :FAIL_BUILD

REM ============================================================
REM [4/6] Git準備（変更検出）
REM ============================================================
echo.
echo [4/6] 変更検出

mkdir "%OUT_DIR%" 2>nul

> "%OUT_DIR%\.buildstamp" echo %TODAY% %NOW%

REM 不要フォルダを削除（GitHub Pagesの容量肥大化防止）
REM 過去に誤生成された重複ディレクトリ
rmdir /s /q "%OUT_DIR%\climax" 2>nul
rmdir /s /q "%OUT_DIR%\default\default" 2>nul

REM 未使用サムネイル。BUILD_THUMBS=0運用なら不要
rmdir /s /q "%OUT_DIR%\default\assets\thumbs" 2>nul
rmdir /s /q "%OUT_DIR%\assets\thumbs" 2>nul

REM 念のため古い誤出力
rmdir /s /q "%OUT_DIR%\ドキュメント" 2>nul

"%GIT%" add -A "%OUT_DIR%"

echo -- git status --
"%GIT%" status -s "%OUT_DIR%"
echo -----------------

REM 変更があるかチェック
"%GIT%" diff --cached --quiet
if not errorlevel 1 (
  echo [INFO] 変更なし（commit/pushはスキップ）
  goto :EXPORT_CSV
)

REM ============================================================
REM [5/6] commit
REM ============================================================
echo.
echo [5/6] commit
set "MSG=build: buylist pages (%TODAY_SAFE% %NOW_SAFE%)"
"%GIT%" commit -m "%MSG%"
if errorlevel 1 goto :FAIL_COMMIT

REM ============================================================
REM [6/6] push
REM ============================================================
echo.
echo [6/6] push

for /f "delims=" %%b in ('"%GIT%" rev-parse --abbrev-ref HEAD') do set "CURBR=%%b"

echo.
echo [5.5/6] pull --rebase
"%GIT%" pull --rebase origin "%CURBR%"
if errorlevel 1 goto :FAIL_PULL

"%GIT%" rev-parse --abbrev-ref --symbolic-full-name "@{u}" >nul 2>&1
if errorlevel 1 (
  "%GIT%" push -u origin "%CURBR%"
  if errorlevel 1 goto :FAIL_PUSH
) else (
  "%GIT%" push
  if errorlevel 1 goto :FAIL_PUSH
)

REM ============================================================
REM [CSV] Mycaアップロード用CSV 出力（Python）
REM ============================================================
:EXPORT_CSV
echo.
echo [CSV] Mycaアップロード用CSV 出力（Python）

REM export_myca_csv.py が repo 直下にある前提
if not exist "%REPO_ROOT%\export_myca_csv.py" (
  echo [ERR] export_myca_csv.py が見つかりません: %REPO_ROOT%\export_myca_csv.py
  goto :END
)

"%PY%" "%REPO_ROOT%\export_myca_csv.py"
if errorlevel 1 (
  echo [ERR] export_myca_csv.py 失敗
  goto :END
)

echo.
echo [OK] 完了
goto :END

:FAIL_PIP_UP
echo [ERR] pip upgrade 失敗
goto :END

:FAIL_PIP_PKGS
echo [ERR] pip install 失敗
goto :END

:FAIL_BUILD
echo [ERR] WEB生成失敗
goto :END

:FAIL_COMMIT
echo [ERR] commit失敗
goto :END

:FAIL_PUSH
echo [ERR] push失敗
goto :END

:END
echo.
echo ==== 実行終了 ====
pause
endlocal
exit /b 0
