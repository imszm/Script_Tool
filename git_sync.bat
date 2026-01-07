@echo off
setlocal enabledelayedexpansion

:: Force UTF-8 encoding
chcp 65001 > nul

title Git Auto-Sync Tool

echo ===========================================
echo       Git Repository Auto-Sync Tool
echo ===========================================
echo.

:: 1. Check Status
echo [1/4] Checking file status...
git status
echo.

:: 2. Add Changes
echo [2/4] Adding changes to stage...
git add .
echo Done.
echo.

:: 3. Commit
set /p msg="Enter commit message (Press Enter for default): "

if "%msg%"=="" (
    set msg=Routine_update_%date%_%time%
    set msg=!msg: =_!
    set msg=!msg:/=-!
    set msg=!msg::=-!
)

echo.
echo [3/4] Committing changes...
git commit -m "!msg!"
echo.

:: 4. Push
echo [4/4] Pushing to GitHub (origin main)...
git push origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===========================================
    echo       SUCCESS: Synced to GitHub.
    echo ===========================================
) else (
    echo.
    echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    echo       ERROR: Push failed. 
    echo       Check Proxy or Token status.
    echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
)

echo.
echo Press any key to exit...
pause > nul