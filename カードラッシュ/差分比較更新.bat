@echo off
chcp 932 >nul
set "BASEDIR=C:\Users\user\OneDrive\Desktop\デュエマ買取表\カードラッシュ"
pushd "%BASEDIR%" || (echo [ERROR] フォルダがありません: "%BASEDIR%" & pause & exit /b 1)

if not exist compare_sheets_partial.py (
  echo [ERROR] compare_sheets_partial.py が見つかりません。 & dir /b *.py & popd & pause & exit /b 1
)
if not exist "%BASEDIR%\credentials.json" (
  echo [WARN] credentials.json が見つかりません: "%BASEDIR%\credentials.json"
  echo そのまま実行を試みます…
)

where python >nul 2>nul || (echo [ERROR] python が見つかりません。PATHを通すかフルパス指定してください。 & popd & pause & exit /b 1)

python compare_sheets_partial.py --sheet-url "https://docs.google.com/spreadsheets/d/1gYYmzLkrtAgNZB6dlwzFEqg0VXT7o0QKlFVPS2ln5Xs/edit?gid=0" --sheet1 "シート1" --sheet2 "CardRush_DM" --sheet1-name-col C --sheet1-exp-col E --sheet1-model-col F --sheet1-price-col O --sheet2-name-col "カード名" --sheet2-model-col "型番" --sheet2-price-col C --out-sheet "差分比較" --creds "%BASEDIR%\credentials.json"

echo.
echo ==== 終了 ====
popd
pause
