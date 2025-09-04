@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
chcp 65001 >nul

rem ====== 設定 ======
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "OUT_DIR=docs\default"
set "PER_PAGE=80"
set "BUILD_THUMBS=0"
set "DO_GIT=1"

rem ====== Excel 検出 ======
set "EXCEL_PATH="
if not "%~1"=="" for %%F in ("%~1") do if exist "%%~fF" set "EXCEL_PATH=%%~fF"
if not defined EXCEL_PATH (
  for %%E in (xlsx xlsm csv) do (
    if not defined EXCEL_PATH if exist "%CD%\buylist.%%E" set "EXCEL_PATH=%CD%\buylist.%%E"
    if not defined EXCEL_PATH if exist "%CD%\data\buylist.%%E" set "EXCEL_PATH=%CD%\data\buylist.%%E"
  )
)
if not defined EXCEL_PATH (
  echo [NG] Excelが見つかりません。publish_simple.bat "C:\full\path\buylist.xlsx"
  pause & exit /b 1
)
if not exist "%EXCEL_PATH%" (
  echo [NG] 指定ファイルが存在しません: "%EXCEL_PATH%"
  pause & exit /b 1
)

echo [*] Excel: "%EXCEL_PATH%"
echo [*] PER_PAGE=%PER_PAGE%  BUILD_THUMBS=%BUILD_THUMBS%

rem ====== 出力先作成 ======
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%" >nul 2>&1
set "DEFAULT_INDEX=%CD%\docs\default\index.html"
set "ROOT_INDEX=%CD%\docs\index.html"

rem ====== 生成 ======
set "OUT_DIR=%OUT_DIR%"
set "PER_PAGE=%PER_PAGE%"
set "BUILD_THUMBS=%BUILD_THUMBS%"
"%PY%" "%~dp0gen_buylist.py" "%EXCEL_PATH%" || goto FAIL

rem ====== 出力確認＆同期（保険）======
if exist "%DEFAULT_INDEX%" goto HAVE_DEFAULT
if exist "%ROOT_INDEX%" goto SYNC_FROM_ROOT
echo [NG] 生成物が見つかりません。gen_buylist.py の出力先を確認してください。
goto FAIL

:HAVE_DEFAULT
echo [OK] generated -> "%DEFAULT_INDEX%"
goto STAMP

:SYNC_FROM_ROOT
echo [!] gen_buylist.py が "docs" に出力 → default に同期
copy /y "%ROOT_INDEX%" "%DEFAULT_INDEX%" >nul
for %%D in (assets static dist js css img images fonts) do (
  if exist "%CD%\docs\%%D" robocopy "%CD%\docs\%%D" "%CD%\docs\default\%%D" /E /NFL /NDL /NJH /NJS /NP >nul
)
goto STAMP

:STAMP
rem ====== build_stamp を一時ファイル→移動（ロックに強い）======
set "TMPSTAMP=%TEMP%\stamp_%RANDOM%%RANDOM%.txt"
> "%TMPSTAMP%" echo built at %date% %time% from "%EXCEL_PATH%"
for /l %%i in (1,1,5) do (
  move /y "%TMPSTAMP%" "%OUT_DIR%\.build_stamp.txt" >nul 2>&1 && goto GIT
  timeout /t 1 >nul
)
if exist "%TMPSTAMP%" del /f /q "%TMPSTAMP%" >nul 2>&1
echo [!] stamp write skipped (locked?)
goto GIT

:GIT
where git >nul 2>&1 || set "DO_GIT=0"
if not "%DO_GIT%"=="1" goto OPEN

echo [*] Git commit/pull/push...

rem ==== 強力ロック解除（最大5回リトライ）====
set "LOCKS=.git\index.lock .git\shallow.lock .git\packed-refs.lock .git\logs\HEAD.lock"
for /l %%i in (1,1,5) do (
  tasklist | findstr /i git.exe >nul && taskkill /f /im git.exe >nul 2>&1
  git rebase --abort 1>nul 2>nul
  if exist ".git\rebase-merge" rmdir /s /q ".git\rebase-merge" 2>nul
  for %%L in (%LOCKS%) do if exist "%%~L" del /f /q "%%~L"
  if exist ".git\index.lock" (
    echo [!] index.lock detected... retry %%i/5
    timeout /t 1 >nul
  ) else (
    goto GIT_UNLOCKED
  )
)
echo [NG] Could not remove Git lock files. Aborting...
goto FAIL

:GIT_UNLOCKED
rem ==== pullで引っかかる生成サブフォルダを事前削除（静音）====
for %%D in (price_asc price_desc search dist) do (
  if exist "%OUT_DIR%\%%D" (
    attrib -r -s -h /s /d "%OUT_DIR%\%%D" >nul 2>&1
    rmdir /s /q "%OUT_DIR%\%%D" >nul 2>&1
  )
)

rem ==== 生成物＆Excelのみステージ ====
git add "%OUT_DIR%" || goto FAIL
if exist "buylist.xlsx" git add "buylist.xlsx"

rem ==== 変更があればコミット ====
git diff --cached --quiet
if errorlevel 1 (
  set "TS=%date%_%time%"
  set "TS=%TS::=%"
  set "TS=%TS:/=%"
  set "TS=%TS:.=%"
  set "TS=%TS: =_%"
  git commit -m "update buylist %TS%" || goto FAIL
) else (
  echo [i] 変更なし（commit省略）
)

rem ==== まず FF だけ試す → ダメでも push 優先 ====
git fetch origin || goto FAIL
git pull --ff-only origin main
if errorlevel 1 (
  echo [!] fast-forward 不可 → pull をスキップして push
)

git push origin main || git push --force-with-lease origin main || goto FAIL
echo [OK] publish done
goto OPEN

:OPEN
rem ====== 完成ページを開く ======
if exist ".\docs\default\index.html" (
  start "" ".\docs\default\index.html"
  echo [i] GitHub Pages URL: https://climaxcard.github.io/climax/default/?v=%date%_%time%
) else if exist ".\docs\index.html" (
  start "" ".\docs\index.html"
  echo [i] GitHub Pages URL: https://climaxcard.github.io/climax/?v=%date%_%time%
)
exit /b 0


:FAIL
echo [NG] 生成失敗。上のログを確認してください。
pause
exit /b 1
