@echo off
chcp 932 >nul
set "BASEDIR=C:\Users\user\OneDrive\Desktop\�f���G�}����\\�J�[�h���b�V��"
pushd "%BASEDIR%" || (echo [ERROR] �t�H���_������܂���: "%BASEDIR%" & pause & exit /b 1)

if not exist compare_sheets_partial.py (
  echo [ERROR] compare_sheets_partial.py ��������܂���B & dir /b *.py & popd & pause & exit /b 1
)
if not exist "%BASEDIR%\credentials.json" (
  echo [WARN] credentials.json ��������܂���: "%BASEDIR%\credentials.json"
  echo ���̂܂܎��s�����݂܂��c
)

where python >nul 2>nul || (echo [ERROR] python ��������܂���BPATH��ʂ����t���p�X�w�肵�Ă��������B & popd & pause & exit /b 1)

python compare_sheets_partial.py --sheet-url "https://docs.google.com/spreadsheets/d/1gYYmzLkrtAgNZB6dlwzFEqg0VXT7o0QKlFVPS2ln5Xs/edit?gid=0" --sheet1 "�V�[�g1" --sheet2 "CardRush_DM" --sheet1-name-col C --sheet1-exp-col E --sheet1-model-col F --sheet1-price-col O --sheet2-name-col "�J�[�h��" --sheet2-model-col "�^��" --sheet2-price-col C --out-sheet "������r" --creds "%BASEDIR%\credentials.json"

echo.
echo ==== �I�� ====
popd
pause
