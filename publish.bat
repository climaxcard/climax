@echo off
setlocal
cd /d "%~dp0"

rem --- Git ロック掃除 ---
if exist ".git\index.lock" del /f /q ".git\index.lock"
for %%L in (".git\shallow.lock" ".git\packed-refs.lock" ".git\logs\HEAD.lock") do (
  if exist "%%~L" del /f /q "%%~L"
)
git gc --prune=now >nul 2>&1

rem --- ここからあなたの生成処理 ---
rem pythonや生成コマンドなど（既存の “[*] Installing deps...” 以降）

rem --- コミット＆push（リモートが進んでても自動追従） ---
git add -A || goto :gitfail
git commit -m "update buylist %date% %time:~0,5%" || echo [i] 変更なし or commit失敗
git pull --rebase --autostash origin main || goto :gitfail
git push origin main || goto :gitfail
echo [OK] push 完了
goto :eof

:gitfail
echo [NG] git 処理でエラー。ロックや競合を確認してください。
exit /b 1

)

REM ===== Git 初期化（初回のみ）=====
for /f %%A in ('git rev-parse --is-inside-work-tree 2^>NUL ^| findstr /i true') do set GIT=1
if not defined GIT (
  echo [*] init git repo...
  git init || (echo [NG] git init 失敗 & pause & exit /b 1)
  git branch -M main
  git config user.name "climax-local"
  git config user.email "climax@example.com"
)

REM .gitignore（Excel等をコミットしない）
if not exist ".gitignore" (
  > .gitignore echo *.xlsx
  >> .gitignore echo data/
  >> .gitignore echo buylist_pages_offline/
)

REM ===== リモート設定（毎回検証）=====
git remote remove origin 2>nul
git remote add origin "%REPO_URL%" || (
  echo [NG] git remote add 失敗（REPO_URLを確認）
  pause & exit /b 1
)

REM ===== コミット & プッシュ =====
for /f %%i in ('powershell -NoProfile -Command "(Get-Date).ToString(\"yyyy-MM-dd HH:mm\")"') do set NOW=%%i
git add docs .gitignore gen_buylist.py publish.bat || (
  echo [NG] git add 失敗
  pause & exit /b 1
)
git commit -m "update buylist !NOW!" || echo [i] 変更なし

git push -u origin main || (
  echo [NG] git push 失敗（URL/認証/ネットワークを確認）
  echo  - URL: %REPO_URL%
  echo  - 認証: GitHubのユーザー名＋PAT（Personal Access Token）
  pause & exit /b 1
)

echo [OK] push 完了。公開URL: https://climaxcard.github.io/climax/
echo 反映には数十秒〜数分かかることがあります（GitHub Pages）。
pause
