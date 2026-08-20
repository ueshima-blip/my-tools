@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  Task Board launcher (one-click)
rem  - Keep this file in the SAME folder as index.html
rem  - Double-click to open Edge/Chrome as an app window (no tabs)
rem  - Uses a dedicated profile so your normal browser is unaffected
rem  - The profile is kept, so your tasks remain next time
rem ============================================================

set "SRC=%~dp0"
if not exist "%SRC%index.html" goto :nofiles

rem Build a file:// URL (backslash -> slash)
set "URLPATH=%SRC%index.html"
set "URLPATH=%URLPATH:\=/%"
set "APPURL=file:///%URLPATH%"

rem Dedicated, persistent profile (not deleted)
set "PROFILE=%LOCALAPPDATA%\TaskBoard\profile"
if not exist "%PROFILE%" mkdir "%PROFILE%" 2>nul

rem 1) Find Microsoft Edge
set "BROWSER="
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"

rem 2) Otherwise find Google Chrome
if not defined BROWSER if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "BROWSER=%LocalAppData%\Google\Chrome\Application\chrome.exe"

if defined BROWSER (
  start "" "%BROWSER%" --app="%APPURL%" --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check --window-size=410,700
  exit /b 0
)

rem 3) No Edge/Chrome found: open with the default app
start "" "%SRC%index.html"
exit /b 0

:nofiles
echo.
echo [Error] Put this launcher in the SAME folder as index.html.
echo Current folder: %SRC%
echo.
pause
exit /b 1
