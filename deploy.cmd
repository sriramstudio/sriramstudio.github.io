@echo off
setlocal
cd /d "%~dp0"

REM ============================================================
REM  Sriram Studio - deploy the Apps Script backend
REM
REM    deploy.cmd                 push Code.gs only
REM    deploy.cmd "what changed"  push, then redeploy the web app
REM
REM  Push alone is enough for anything run from the Apps Script
REM  editor or the Analytics menu. Redeploy only when the admin
REM  panel or the public form needs the change - that is what
REM  makes it live at the SAME url.
REM
REM  Why .cmd and not the bare 'clasp': PowerShell's execution
REM  policy blocks npm's clasp.ps1 shim. clasp.cmd is the same
REM  program and is not blocked. Nothing here needs that policy
REM  changed.
REM
REM  Set SS_DEPLOY_DRYRUN=1 to print the clasp commands instead
REM  of running them.
REM ============================================================

REM The live deployment. Redeploying to THIS id keeps the url that
REM register.html has hardcoded. Never 'clasp deploy' - that mints
REM a new url and breaks the public form.
set "DEPLOY_ID=AKfycbzRB6XQrT8lx2lgcgv9nPQ18akiRCU8x58ritJIgbjNRkKxE3FZAifMREpZ4f7ZgSWD"

set "CLASP=clasp.cmd"
where /q clasp.cmd || set "CLASP=%APPDATA%\npm\clasp.cmd"
if not exist "%CLASP%" if "%CLASP%" NEQ "clasp.cmd" (
  echo clasp was not found. Install it with:  npm.cmd install -g @google/clasp
  goto :failed
)

echo.
echo [1/3] Running the tests...
call node tests\auth.test.js Code.gs
if errorlevel 1 (
  echo.
  echo TESTS FAILED - nothing was pushed. Fix them first.
  goto :failed
)

echo.
echo [2/3] Pushing Code.gs to Apps Script...
if defined SS_DEPLOY_DRYRUN (
  echo   [dry run] "%CLASP%" push
) else (
  call "%CLASP%" push
  if errorlevel 1 (
    echo.
    echo PUSH FAILED. If it asked you to log in, run:  clasp.cmd login
    goto :failed
  )
)

if "%~1"=="" (
  echo.
  echo [3/3] No description given, so the live web app was NOT redeployed.
  echo       That is correct for anything you run from the Apps Script
  echo       editor or the Analytics menu - the editor always runs the
  echo       code you just pushed.
  echo.
  echo       To redeploy the admin panel / public form as well:
  echo           deploy.cmd "what changed"
  goto :done
)

echo.
echo [3/3] Redeploying the web app to the existing deployment...
if defined SS_DEPLOY_DRYRUN (
  echo   [dry run] "%CLASP%" redeploy %DEPLOY_ID% -d "%~1"
) else (
  call "%CLASP%" redeploy %DEPLOY_ID% -d "%~1"
  if errorlevel 1 (
    echo.
    echo REDEPLOY FAILED. The push succeeded, so the editor already has
    echo the new code - only the live web app is still on the old version.
    goto :failed
  )
)
echo       Done. The url is unchanged, so no HTML needs touching.

:done
echo.
echo Finished.
exit /b 0

REM On failure only, hold the window open so the reason is readable even if
REM this was double-clicked from Explorer rather than run from a prompt.
REM find.exe is named in full: a Git Bash PATH puts a different 'find' first.
:failed
echo.
echo %cmdcmdline% | "%SystemRoot%\System32\find.exe" /i "%~nx0" >nul && pause
exit /b 1
