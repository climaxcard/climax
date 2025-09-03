@echo off
cd C:\Users\user\OneDrive\Desktop\デュエマ買取表

:: 実行してログを残す
python C:\Users\user\OneDrive\Desktop\cardrush_to_sheet.py ^
  --sheet-url "https://docs.google.com/spreadsheets/d/1gYYmzLkrtAgNZB6dlwzFEqg0VXT7o0QKlFVPS2ln5Xs/edit?gid=0" ^
  --sheet-name "CardRush_DM" >> update_cardrush.log 2>&1

echo [%date% %time%] 更新完了 >> update_cardrush.log
