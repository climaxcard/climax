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

rem --- calc SHA256 ---
set "NEW_HASH="
for /f "tokens=1" %%H in ('certutil -hashfile "%EXCEL_PATH%" SHA256 ^| findstr /R "^[0-9A-Fa-f]"') do set "NEW_HASH=%%H"
if not defined NEW_HASH (
  echo [NG] Hash failed: %EXCEL_PATH%
  pause & exit /b 1
)

rem --- load previous (2-line format: PATH then HASH). Old 1-line format also許容 ---
set "OLD_PATH="
set "OLD_HASH="
if exist ".publish.hash" (
  for /f "usebackq tokens=* delims=" %%L in (".publish.hash") do (
    if not defined OLD_PATH (
      set "OLD_PATH=%%L"
    ) else if not defined OLD_HASH (
      set "OLD_HASH=%%L"
    )
  )
  rem 旧形式（1行＝HASHのみ）なら補正
  if defined OLD_PATH if not defined OLD_HASH (
    set "OLD_HASH=%OLD_PATH%"
    set "OLD_PATH="
  )
)

rem --- mtime check helper ---
set "NEED_BUILD="
set "STAMP_FILE=%OUT_DIR%\.build_stamp.txt"

rem 条件1: 強制
if "%FORCE%"=="1" set "NEED_BUILD=1"

rem 条件2: Excelパスが変わった
if not defined NEED_BUILD if /I not "%EXCEL_PATH%"=="%OLD_PATH%" set "NEED_BUILD=1"

rem 条件3: ハッシュが変わった
if not defined NEED_BUILD if /I not "%NEW_HASH%"=="%OLD_HASH%" set "NEED_BUILD=1"

rem 条件4: ハッシュ同じでも、Excelがビルドスタンプより新しければ再ビルド
if not defined NEED_BUILD if exist "%STAMP_FILE%" (
  for %%A in ("%EXCEL_PATH%") do set "EXCEL_MTIME=%%~tA"
  for %%B in ("%STAMP_FILE%") do set "STAMP_MTIME=%%~tB"
  if "%EXCEL_MTIME%" GTR "%STAMP_MTIME%" set "NEED_BUILD=1"
)

if not defined NEED_BUILD set "NEED_BUILD=0"

if "%NEED_BUILD%"=="1" (
  echo [*] Generate docs...
  rem 新フォーマットで保存（1行目=PATH, 2行目=HASH）
  > ".publish.hash" (echo %EXCEL_PATH%&echo %NEW_HASH%)

  set "OUT_DIR=%OUT_DIR%"
  set "PER_PAGE=%PER_PAGE%"
  set "BUILD_THUMBS=%BUILD_THUMBS%"

  rem ★ EXCELのフルパスをPythonに引数で渡す ★
  "%PY%" gen_buylist.py "%EXCEL_PATH%" || goto :fail
) else (
  echo [i] Excel unchanged -> build step skipped
)

rem Build stamp to force tiny diff so Pages redeploys
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"
> "%STAMP_FILE%" echo built at %date% %time% from "%EXCEL_PATH%"

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
