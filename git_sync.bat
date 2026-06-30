@echo off
setlocal enabledelayedexpansion

:: Force UTF-8 encoding
chcp 65001 > nul

title Git Auto-Sync Tool

echo ===========================================
echo       Git Repository Auto-Sync Tool
echo ===========================================
echo.

:: ── Pre-flight Checks ──────────────────────────────────────────────

git --version > nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Git not found. Please install Git and add it to PATH.
    goto :exit
)

git rev-parse --is-inside-work-tree > nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Not a Git repository. Please run from your project root.
    goto :exit
)

:: ── .gitignore Guard ───────────────────────────────────────────────

echo [Check] Verifying .gitignore rules...
set NEED_UPDATE=0

if not exist ".gitignore" (
    echo [WARN]  No .gitignore found. Creating with recommended exclusions...
    call :write_gitignore
    set NEED_UPDATE=1
) else (
    findstr /i "Tool\\logs" .gitignore > nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo [WARN]  Log patterns missing from .gitignore. Appending...
        call :append_gitignore
        set NEED_UPDATE=1
    ) else (
        echo [OK]    .gitignore already covers log files.
    )
)

if !NEED_UPDATE! equ 1 (
    echo.
    echo   Exclusion patterns added to .gitignore:
    echo     Tool/logs/            
    echo     Tool/Test_Logs/  
    echo     *.log / *_raw_*.txt / *_error_*.txt / *_full_*.txt ...
    echo.
    echo [IMPORTANT] .gitignore only prevents FUTURE tracking.
    echo   If log files are already committed, run the following ONCE to untrack them:
    echo.
    echo     git rm -r --cached Tool/logs/ Tool/Test_Logs/
    echo     git commit -m "chore: stop tracking log files"
    echo.
    set /p "CONT=Continue with sync now? (Y/N): "
    if /i "!CONT!" neq "Y" (
        echo Aborted.
        goto :exit
    )
)
echo.

:: ── Step 1: Status ─────────────────────────────────────────────────

echo [1/4] Repository status:
git status --short
echo.

:: Exit early if nothing to commit
set HAS_CHANGES=
for /f "usebackq" %%i in (`git status --porcelain`) do set HAS_CHANGES=1
if not defined HAS_CHANGES (
    echo Nothing to commit. Working tree is already clean.
    goto :exit
)

:: ── Step 2: Stage ──────────────────────────────────────────────────

echo [2/4] Staging changes (respecting .gitignore)...
git add .
echo.
echo Staged file summary:
git --no-pager diff --cached --stat
echo.

:: ── Step 3: Commit ─────────────────────────────────────────────────

set /p "msg=Enter commit message (Press Enter for auto-generated): "
echo.

if "!msg!"=="" (
    for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format 'yyyyMMdd_HHmm'"') do set TIMESTAMP=%%i
    set msg=Routine_update_!TIMESTAMP!
)

echo [3/4] Committing: "!msg!"
git commit -m "!msg!"
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Commit failed. Please check the output above.
    goto :exit
)
echo.

:: ── Step 4: Push ───────────────────────────────────────────────────

echo [4/4] Pushing to origin/main...
git push origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===========================================
    echo       SUCCESS: Synced to GitHub.
    echo ===========================================
) else (
    echo.
    echo ===========================================
    echo   ERROR: Push failed.
    echo   Possible causes:
    echo     - Network / proxy issue
    echo     - Expired or invalid token
    echo     - Wrong branch name (main vs master)
    echo ===========================================
)

goto :exit

:: ── Subroutines ────────────────────────────────────────────────────

:write_gitignore
(
    echo # === Log and Test Output Files ===
    echo Tool/logs/
    echo Tool/Test_Logs/
    echo *.log
    echo *_raw_*.txt
    echo *_error_*.txt
    echo *_full_*.txt
    echo *_summary_*.txt
    echo.
    echo # === Python Cache ===
    echo __pycache__/
    echo *.pyc
    echo *.pyo
    echo.
    echo # === Build Artifacts ===
    echo build/
    echo dist/
    echo *.spec
    echo.
    echo # === OS Metadata ===
    echo .DS_Store
    echo Thumbs.db
    echo desktop.ini
) > .gitignore
goto :eof

:append_gitignore
(
    echo.
    echo # === Log and Test Output Files ^(auto-added by git_sync.bat^) ===
    echo Tool/logs/
    echo Tool/Test_Logs/
    echo *.log
    echo *_raw_*.txt
    echo *_error_*.txt
    echo *_full_*.txt
    echo *_summary_*.txt
) >> .gitignore
goto :eof

:exit
echo.
echo Press any key to exit...
pause > nul
endlocal