@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM ============================================================
REM 0) Git の実体を特定
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
REM 0.5) Python 実行コマンドを特定（py優先）
REM ============================================================
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY (
  where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo [ERR] Pythonが見つかりません（py / python）
  goto :END
)

REM ============================================================
REM 設定
REM ============================================================
set "OUT_DIR=docs"
set "EXCEL_FILE=buylist.xlsm"

REM ★追加：最初と最後に実行するPython
set "GEN_PY=C:\Users\user\OneDrive\Desktop\デュエマ買取表\repo\generate_buylist.py"
set "MYCA_PY=C:\Users\user\OneDrive\Desktop\デュエマ買取表\repo\export_myca_csv.py"

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
REM ★[0/3] 生成スクリプト実行（最初）
REM ============================================================
echo.
echo [0/3] run generate_buylist.py

if not exist "%GEN_PY%" (
  echo [ERR] generate_buylist.py が見つかりません: %GEN_PY%
  goto :END
)

"%PY%" "%GEN_PY%"
if errorlevel 1 (
  echo [ERR] generate_buylist.py 失敗（中断）
  goto :END
)

REM ============================================================
REM 対象存在チェック
REM ============================================================
if not exist "%REPO_ROOT%\%OUT_DIR%\" (
  echo [ERR] 対象フォルダが見つかりません: %REPO_ROOT%\%OUT_DIR%
  goto :END
)

if not exist "%REPO_ROOT%\%EXCEL_FILE%" (
  echo [ERR] 対象Excelが見つかりません: %REPO_ROOT%\%EXCEL_FILE%
  goto :END
)

REM ============================================================
REM [1/3] add（docs と buylist.xlsm のみ）
REM ============================================================
echo.
echo [1/3] git add
"%GIT%" add -A "%OUT_DIR%" "%EXCEL_FILE%"
if errorlevel 1 (
  echo [ERR] git add 失敗
  goto :END
)

echo -- git status --
"%GIT%" status -s "%OUT_DIR%" "%EXCEL_FILE%"
echo -----------------

REM 変更があるかチェック（ステージ済み差分）
"%GIT%" diff --cached --quiet
if not errorlevel 1 (
  echo [INFO] 変更なし（commit/pushはスキップ）
  REM ★変更がなくてもMyca CSVは作りたいなら、ここで export を実行してからENDへ
  goto :RUN_MYCA
)

REM ============================================================
REM [2/3] commit（日時入り）
REM ============================================================
echo.
echo [2/3] commit

for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set "TODAY=%%a %%b %%c"
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set "NOW=%%a:%%b"

set "MSG=build: update docs+excel (%TODAY% %NOW%)"
"%GIT%" commit -m "%MSG%"
if errorlevel 1 (
  echo [ERR] commit失敗
  goto :END
)

REM ============================================================
REM [3/3] push（upstream無ければ -u）
REM ============================================================
echo.
echo [3/3] push

for /f "delims=" %%b in ('"%GIT%" rev-parse --abbrev-ref HEAD') do set "CURBR=%%b"

"%GIT%" rev-parse --abbrev-ref --symbolic-full-name "@{u}" >nul 2>&1
if errorlevel 1 (
  "%GIT%" push -u origin "%CURBR%"
  if errorlevel 1 (
    echo [ERR] push失敗
    goto :END
  )
) else (
  "%GIT%" push
  if errorlevel 1 (
    echo [ERR] push失敗
    goto :END
  )
)

echo.
echo [OK] commit/push 完了

REM ============================================================
REM ★追加：最後に Myca CSV 出力
REM ============================================================
:RUN_MYCA
echo.
echo [LAST] run export_myca_csv.py

if not exist "%MYCA_PY%" (
  echo [ERR] export_myca_csv.py が見つかりません: %MYCA_PY%
  goto :END
)

"%PY%" "%MYCA_PY%"
if errorlevel 1 (
  echo [ERR] export_myca_csv.py 失敗
  goto :END
)

echo.
echo [OK] export_myca_csv.py 完了

:END
echo.
echo ==== 実行終了 ====
pause
endlocal
exit /b 0