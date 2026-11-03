@echo off
rem Commits everything in this folder and pushes it to GitHub.
cd /d "%~dp0"
git add -A
if not exist commit_message.txt goto ask
git reset -q commit_message.txt
git commit -F commit_message.txt
if not errorlevel 1 del commit_message.txt
goto push
:ask
set /p MSG=Describe your changes: 
git commit -m "%MSG%"
:push
git push
if errorlevel 1 goto failed
echo.
echo Pushed to GitHub.
pause
exit /b 0
:failed
echo.
echo Push failed - see the messages above.
echo If GitHub has newer changes, run:  git pull --rebase   then run this script again.
pause
exit /b 1
