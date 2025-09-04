@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
chcp 65001 >nul

rem ====== 設定 ======
set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "OUT_DIR=docs"     rem ★必ず docs 固定（スクリプト側が default/price_* を自動付与する）
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
  echo [NG] Excelが見つかりません。WEB買取表反映.bat "C:\full\path\buylist.xlsx"
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
set "DEFAULT_DIR=%CD%\docs\default"
set "DEFAULT_INDEX=%DEFAULT_DIR%\index.html"
set "ROOT_INDEX=%CD%\docs\index.html"

rem ====== 生成 ======
set "OUT_DIR=%OUT_DIR%"
set "PER_PAGE=%PER_PAGE%"
set "BUILD_THUMBS=%BUILD_THUMBS%"
"%PY%" "%~dp0gen_buylist.py" "%EXCEL_PATH%" || goto FAIL

rem ====== 出力確認（PowerShellで確実にサイズ取得＆中身チェック）======
set "DEFAULT_DIR=%CD%\docs\default"
set "DEFAULT_INDEX=%DEFAULT_DIR%\index.html"
set "ROOT_INDEX=%CD%\docs\index.html"

for /f "usebackq delims=" %%S in (`powershell -NoP -C "(Get-Item '%DEFAULT_INDEX%' -EA SilentlyContinue).Length ^| ForEach-Object { $_ -as [int] } ; if(-not $?) {0} ; if($null -eq (Get-Item '%DEFAULT_INDEX%' -EA SilentlyContinue)){0}"`) do set "SZ=%%S"

echo [dbg] index=%DEFAULT_INDEX%
echo [dbg] size=%SZ% bytes

rem === index.htmlが小さい or <html>タグが無い → 自動修復 ===
for /f "usebackq delims=" %%H in (`powershell -NoP -C "if(Test-Path '%DEFAULT_INDEX%'){ $c=Get-Content '%DEFAULT_INDEX%' -Raw; if($c -match '<html' ){ '1' } else { '0' } } else { '0' }"`) do set "HAS_HTML=%%H"
echo [dbg] has_html=%HAS_HTML%

if "%SZ%"=="" set "SZ=0"
if %SZ% LSS 2048 goto FIX_INDEX
if "%HAS_HTML%"=="0" goto FIX_INDEX

echo [OK] generated -> "%DEFAULT_INDEX%"
goto INJECT_BASE

:FIX_INDEX
echo [!] index.html が不正（size=%SZ%, html=%HAS_HTML%）→ 自動修復します
if exist "%DEFAULT_DIR%\p1.html" (
  copy /y "%DEFAULT_DIR%\p1.html" "%DEFAULT_INDEX%" >nul || goto FAIL
  echo [fix] index.html <- p1.html へ差し替え
) else if exist "%ROOT_INDEX%" (
  copy /y "%ROOT_INDEX%" "%DEFAULT_INDEX%" >nul
  for %%D in (assets static dist js css img images fonts) do (
    if exist "%CD%\docs\%%D" robocopy "%CD%\docs\%%D" "%DEFAULT_DIR%\%%D" /E /NFL /NDL /NJH /NJS /NP >nul
  )
  echo [fix] root docs から default に同期
) else (
  echo [NG] 有効な index.html が見つかりません（p1.html も root も無し）
  goto FAIL
)

:INJECT_BASE
rem ====== <base> 自動挿入（head がある時だけ、安全に）======
powershell -NoP -C ^
  "$p='%DEFAULT_INDEX%'; if(Test-Path $p){$h=Get-Content $p -Raw; if($h -match '</head>' -and $h -notmatch '<base href='){ $h=$h -replace '</head>','<base href=\"/climax/default/\" /></head>'; Set-Content -Encoding UTF8 $p $h; '[/] base inserted' } else { '[/] base ok/skip' }}"


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
git rev-parse --is-inside-work-tree >nul 2>&1 || set "DO_GIT=0"
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
rem ==== 衝突しがちな生成サブフォルダを事前削除（静音）====
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
  for /f "delims=" %%B in ('git rev-parse --abbrev-ref HEAD') do set "BRANCH=%%B"
  if not defined BRANCH set "BRANCH=main"
  set "TS=%date%_%time%"
  set "TS=%TS::=%"
  set "TS=%TS:/=%"
  set "TS=%TS:.=%"
  set "TS=%TS: =_%"
  git commit -m "update buylist %TS%" || goto FAIL
) else (
  echo [i] 変更なし（commit省略）
)

rem ==== まず FF-only を試す → ダメでも push 優先 ====
git fetch origin
git pull --ff-only origin main
if errorlevel 1 (
  echo [!] fast-forward 不可 → pull をスキップして push
)

git push origin main || git push --force-with-lease origin main || goto FAIL
echo [OK] publish done
goto OPEN

:OPEN
rem ====== 完成ページを開く ======
if exist "%DEFAULT_INDEX%" (
  start "" "%DEFAULT_INDEX%"
  echo [i] GitHub Pages URL: https://climaxcard.github.io/climax/default/?v=%date%_%time%
) else if exist "%ROOT_INDEX%" (
  start "" "%ROOT_INDEX%"
  echo [i] GitHub Pages URL: https://climaxcard.github.io/climax/?v=%date%_%time%
)
exit /b 0

:FAIL
echo [NG] 処理失敗。上のログを確認してください。
pause
exit /b 1
