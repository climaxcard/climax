@echo off
setlocal
cd /d "%~dp0"

set "PY=C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
set "OUT_DIR=docs"
set "PER_PAGE=80"

rem defaults
set "BUILD_THUMBS=1"
set "DO_GIT=1"
set "FORCE=0"
set "EXCEL_PATH="

rem Options: fast(thumb off) / nogit / force / <xlsx path>
for %%A in (%*) do (
  if /I "%%~A"=="fast"  set "BUILD_THUMBS=0"
  if /I "%%~A"=="nogit" set "DO_GIT=0"
  if /I "%%~A"=="force" set "FORCE=1"
  if exist "%%~fA" if /I "%%~xA"==".xlsx" set "EXCEL_PATH=%%~fA"
)

rem Excel autodetect（buylist.xlsx 優先）
if not defined EXCEL_PATH if exist "buylist.xlsx" set "EXCEL_PATH=%CD%\buylist.xlsx"
if not defined EXCEL_PATH if exist "data\buylist.xlsx" set "EXCEL_PATH=%CD%\data\buylist.xlsx"
if not defined EXCEL_PATH (
  for /f "delims=" %%F in ('dir /b /a:-d /o:-d "*.xlsx" 2^>nul') do (
    set "EXCEL_PATH=%CD%\%%~F"
    goto :got_excel
  )
)
:got_excel
if not defined EXCEL_PATH (
  echo [NG] No Excel found. Put buylist.xlsx next to this bat,
  echo      or pass a full path: smart_publish.bat fast "C:\full\path\buylist.xlsx"
  pause & exit /b 1
)

echo [*] Excel: %EXCEL_PATH%
echo [*] Mode: BUILD_THUMBS=%BUILD_THUMBS% PER_PAGE=%PER_PAGE% DO_GIT=%DO_GIT% FORCE=%FORCE%

rem change detection (Excel)
set "NEW_HASH="
for /f "tokens=1" %%H in ('certutil -hashfile "%EXCEL_PATH%" SHA256 ^| findstr /R "^[0-9A-F]"') do set "NEW_HASH=%%H"
if not defined NEW_HASH (
  echo [NG] Hash failed: %EXCEL_PATH%
  pause & exit /b 1
)
set "OLD_HASH="
if exist ".publish.hash" set /p OLD_HASH=<.publish.hash

if "%FORCE%"=="1" (
  echo [!] FORCE=1 -> regenerate
) else if /I "%NEW_HASH%"=="%OLD_HASH%" (
  echo [i] Excel unchanged -> may skip generate
) else (
  > ".publish.hash" echo %NEW_HASH%
)

rem ===== Build =====
if "%FORCE%"=="1" (
  set "NEED_BUILD=1"
) else if /I "%NEW_HASH%"=="%OLD_HASH%" (
  set "NEED_BUILD=0"
) else (
  set "NEED_BUILD=1"
)

if "%NEED_BUILD%"=="1" (
  echo [*] Generate docs...
  set "EXCEL_PATH=%EXCEL_PATH%"
  set "OUT_DIR=%OUT_DIR%"
  set "PER_PAGE=%PER_PAGE%"
  set "BUILD_THUMBS=%BUILD_THUMBS%"
  "%PY%" gen_buylist.py || goto :fail
) else (
  echo [i] Excel unchanged -> build step skipped
)

rem Build stamp to force a tiny diff so Pages redeploys
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"
> "%OUT_DIR%\.build_stamp.txt" echo built at %date% %time% from "%EXCEL_PATH%"

rem Basic sanity: docs/index.html should exist for Pages(doc) publish
if not exist "%OUT_DIR%\index.html" (
  echo [WARN] %OUT_DIR%\index.html not found. If GitHub Pages is set to 'docs', site won't update.
)

rem ===== Git =====
where git >nul 2>&1 || set "DO_GIT=0"
if "%DO_GIT%"=="1" (
  echo [*] Git pull/commit/push...
  if exist ".git\index.lock" del /f /q ".git\index.lock"
  for %%L in (".git\shallow.lock" ".git\packed-refs.lock" ".git\logs\HEAD.lock") do if exist "%%~L" del /f /q "%%~L"
  git gc --prune=now >nul 2>&1

  git fetch origin || goto :fail
  git pull --rebase --autostash origin main || goto :fail

  git add -A || goto :fail
  git diff --cached --quiet && (
    echo [i] no staged changes -> nothing to commit
  ) || (
    git commit -m "update buylist %date% %time:~0,5%" || goto :fail
  )

  echo [i] Staged files:
  git diff --cached --name-status

  git push origin main || goto :fail

  echo [OK] publish done
  echo [i] Last commit:
  git log -1 --oneline --name-status
) else (
  echo [i] Git skipped (not found or nogit)
)

exit /b 0

:fail
echo [NG] Error. Check the log above.
pause
exit /b 1
