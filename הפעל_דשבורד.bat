@echo off
title ועד בית — Dashboard Generator
cd /d D:\Claude_projects\Home_tech
set PATH=%PATH%;C:\Program Files\Git\bin
echo Generating dashboard...
echo.
C:\Python\python.exe -B -c "import sys; sys.path.insert(0,'D:/Claude_projects/Home_tech'); import vaad_bayit_generator as g; g.run_once()"
if errorlevel 1 (
    echo.
    echo ERROR: Generator failed. Check output above.
    pause
    exit /b 1
)
echo.
echo Pushing to GitHub...
git config --global --add safe.directory D:/Claude_projects/Home_tech
git checkout main 2>nul
git add index.html
git commit -m "Update dashboard %date% %time%"
git pull --rebase --autostash origin main
git push origin main
echo.
echo Done! Site will update at https://emeqhashalom84-jpg.github.io/Vaad-bayit/ in ~1 minute.
pause
