@echo off
REM ���{��p�X�΍�F�R���\�[���� CP932 ��
chcp 932 >nul

REM ===== �Œ�p�X =====
set "BASEDIR=C:\Users\user\OneDrive\Desktop\�f���G�}����\\�J�[�h���b�V��"
set "PYFILE=cardrush_to_sheet.py"
set "CREDS=%BASEDIR%\credentials.json"
set "SHEET_URL=https://docs.google.com/spreadsheets/d/1gYYmzLkrtAgNZB6dlwzFEqg0VXT7o0QKlFVPS2ln5Xs/edit?gid=0"
set "SHEET_NAME=CardRush_DM"
set "LIMIT=100"
set "MAX_PAGES=200"
set "SLEEP_MS=350"

echo === CardRush�iDM�j�� Google�X�v���b�h�V�[�g�X�V ===

REM �p�X���݃`�F�b�N & 8.3�Z�k���擾�i���{����S�΍�j
if not exist "%BASEDIR%" (
  echo [ERROR] �t�H���_��������܂���: "%BASEDIR%"
  pause & exit /b 1
)
for %%I in ("%BASEDIR%") do set "BASEDIR_S=%%~sI"

REM ��ƃf�B���N�g���ړ��i�Z�k���ň��S�Ɂj
pushd "%BASEDIR_S%" || (echo [ERROR] cd �Ɏ��s���܂��� & pause & exit /b 1)

REM python ���݊m�F
where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] python ��������܂���BPATH��ʂ����t���p�X�w��ɂ��Ă��������B
  pause & exit /b 1
)

REM �X�N���v�g�m�F
if not exist "%PYFILE%" (
  echo [ERROR] �X�N���v�g������܂���: "%CD%\%PYFILE%"
  echo ���̃t�H���_�� .py �ꗗ:
  dir /b *.py
  popd
  pause & exit /b 1
)

REM �F��JSON�͌x���̂�
if not exist "%CREDS%" (
  echo [WARN] �F��JSON��������܂���: "%CREDS%"
  echo ���̂܂܎��s�����݂܂��c
)

REM ���s
python "%PYFILE%" ^
  --sheet-url "%SHEET_URL%" ^
  --sheet-name "%SHEET_NAME%" ^
  --creds "%CREDS%" ^
  --limit %LIMIT% ^
  --max-pages %MAX_PAGES% ^
  --sleep-ms %SLEEP_MS%

echo.
echo ==== ���� ====
popd
pause
