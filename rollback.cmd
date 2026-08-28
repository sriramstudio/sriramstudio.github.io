@echo off
setlocal
cd /d "%~dp0"

REM ============================================================
REM  Sriram Studio - put the code back to a known-good state
REM
REM    rollback.cmd            restore from the pre-fee-split tag
REM    rollback.cmd <tag>      restore from another tag
REM    rollback.cmd --list     show the tags available
REM
REM  This restores Code.gs and sriramstudio_admin.html from a tag
REM  and stops. It does NOT push and does NOT rewrite history, so
REM  nothing is lost and it can itself be undone. It tells you the
REM  two commands to make the restore live.
REM
REM  It does not touch the Google Sheet. Data already in the sheet
REM  is not affected by any of this - the only sheet change the
REM  fee split makes is ADDING columns at the end, which is
REM  harmless to leave in place.
REM ============================================================

set "TAG=%~1"
if "%TAG%"=="" set "TAG=pre-fee-split"

if /i "%TAG%"=="--list" (
  echo Tags you can roll back to:
  echo.
  git tag -n1
  goto :done
)

git rev-parse -q --verify "refs/tags/%TAG%" >nul
if errorlevel 1 (
  echo No tag called "%TAG%".
  echo.
  echo Available:
  git tag -n1
  goto :failed
)

echo.
echo Restoring Code.gs and sriramstudio_admin.html from tag: %TAG%
echo   ^(that tag is commit^)
git log --oneline -1 "%TAG%"
echo.

git checkout "%TAG%" -- Code.gs sriramstudio_admin.html
if errorlevel 1 (
  echo Could not restore the files. Nothing has been changed.
  goto :failed
)

echo Files restored. Checking them...
call node tests\auth.test.js Code.gs
if errorlevel 1 (
  echo.
  echo The restored Code.gs FAILED its own tests. That should not happen.
  echo Undo this with:  git checkout HEAD -- Code.gs sriramstudio_admin.html
  goto :failed
)
call node tests\check-refs.js sriramstudio_admin.html
if errorlevel 1 goto :failed

echo.
echo ============================================================
echo  Restored, tested, and sitting in your folder. NOT yet live.
echo.
echo  To make it live:
echo      .\deploy.cmd            ^(the backend^)
echo      git add -A
echo      git commit -m "Roll back to %TAG%"
echo      git push                ^(the admin panel^)
echo.
echo  To change your mind and keep the newer code instead:
echo      git checkout HEAD -- Code.gs sriramstudio_admin.html
echo ============================================================

:done
echo.
exit /b 0

REM find.exe in full: a Git Bash PATH puts a different 'find' first.
:failed
echo.
echo %cmdcmdline% | "%SystemRoot%\System32\find.exe" /i "%~nx0" >nul && pause
exit /b 1
