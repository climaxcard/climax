@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

REM ===== 設定 =====
set "PY=%LocalAppData%\Programs\Python\Python313\python.exe"
set "REPO_URL=https://github.com/climaxcard/climax.git"

REM ===== Pythonチェック & 依存導入 =====
if not exist "%PY%" (
  echo [NG] Python not found: %PY%
  echo Python 3.13 をインストールしてください。
  pause & exit /b 1
)
echo [*] Installing deps...
"%PY%" -m pip install -q --disable-pip-version-check pandas openpyxl || (
  echo [NG] pip install failed
  pause & exit /b 1
)

REM ===== 生成 =====
echo [*] Generate docs...
"%PY%" gen_buylist.py || (
  echo [NG] 生成失敗。上のエラーを確認してください。
  pause & exit /b 1
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
