@echo off
chcp 65001 >nul
color 0A
title Ops Weekly Output Report - Workflow Menu

:: Locate GitHub Desktop's git.exe if available
set GIT_CMD=git
FOR /D %%G IN ("%LocalAppData%\GitHubDesktop\app-*") DO (
    IF EXIST "%%G\resources\app\git\cmd\git.exe" set GIT_CMD="%%G\resources\app\git\cmd\git.exe"
)

:MENU
cls
echo =======================================================
echo     OPS WEEKLY OUTPUT REPORT - WORKFLOW MENU
echo =======================================================
echo.
echo Please select a step to run:
echo.
echo [1] Download Data (Run Data_loader.py)
echo [2] Generate Dashboard (Run generate_dashboard.py)
echo [3] Deploy Github (Commit and Push changes)
echo [4] Run All (Download -^> Generate -^> Deploy)
echo [5] Exit
echo.

set /p choice="Enter your choice (1/2/3/4/5): "

if "%choice%"=="1" goto DOWNLOAD
if "%choice%"=="2" goto GENERATE
if "%choice%"=="3" goto DEPLOY
if "%choice%"=="4" goto ALL
if "%choice%"=="5" goto EOF

echo Invalid choice, please try again.
pause
goto MENU

:DOWNLOAD
cls
echo -------------------------------------------------------
echo  STEP 1: DOWNLOADING DATA...
echo -------------------------------------------------------
cd scripts
python Data_loader.py
cd ..
echo.
echo Data download complete.
pause
goto MENU

:GENERATE
cls
echo -------------------------------------------------------
echo  STEP 2: GENERATING DASHBOARD...
echo -------------------------------------------------------
cd scripts
python generate_dashboard.py
cd ..
echo.
echo Dashboard generation complete.
pause
goto MENU

:DEPLOY
cls
echo -------------------------------------------------------
echo  STEP 3: DEPLOYING TO GITHUB...
echo -------------------------------------------------------
echo Committing changes...
%GIT_CMD% add .
%GIT_CMD% commit -m "Auto-deploy via Workflow Menu"
echo Pushing to GitHub...
%GIT_CMD% push origin main
echo.
echo Deploy complete.
pause
goto MENU

:ALL
cls
echo -------------------------------------------------------
echo  RUNNING FULL WORKFLOW...
echo -------------------------------------------------------
cd scripts
echo [1/3] Downloading Data...
python Data_loader.py
echo.
echo [2/3] Generating Dashboard...
python generate_dashboard.py
cd ..
echo.
echo [3/3] Deploying to GitHub...
%GIT_CMD% add .
%GIT_CMD% commit -m "Auto-deploy via Workflow Menu"
%GIT_CMD% push origin main
echo.
echo Full workflow complete.
pause
goto MENU

:EOF
exit
