@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  Genetics Simulator Launcher (Chrome only)
rem  - Place this file next to index.html / app.js / style.css
rem  - Double click to launch
rem  - Copies files to %TEMP%, runs Chrome with a dedicated profile,
rem    waits for Chrome to close, then deletes the local copy.
rem ============================================================

set "SRC=%~dp0"
set "DST=%TEMP%\GeneticsApp"
set "PROFILE=%TEMP%\GeneticsAppProfile"

if not exist "%SRC%index.html" goto :nofiles
if not exist "%SRC%app.js"     goto :nofiles
if not exist "%SRC%style.css"  goto :nofiles

if not exist "%DST%" mkdir "%DST%" 2>nul

copy /Y "%SRC%index.html" "%DST%\" >nul
copy /Y "%SRC%app.js"     "%DST%\" >nul
copy /Y "%SRC%style.css"  "%DST%\" >nul

if not exist "%DST%\index.html" goto :copyfail

set "CHROME="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"

if not defined CHROME goto :nochrome

start "" /wait "%CHROME%" --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check "%DST%\index.html"

if exist "%DST%"     rd /s /q "%DST%"     2>nul
if exist "%PROFILE%" rd /s /q "%PROFILE%" 2>nul

exit /b 0

:nochrome
echo.
echo [エラー] Google Chrome が見つかりません。
echo Chrome をインストールしてから、もう一度ダブルクリックしてください。
echo   https://www.google.com/chrome/
echo.
pause
exit /b 1

:nofiles
echo.
echo [エラー] この起動ファイルは index.html / app.js / style.css と
echo 同じフォルダに置いてください。
echo 現在の場所: %SRC%
echo.
pause
exit /b 1

:copyfail
echo.
echo [エラー] ローカルへのコピーに失敗しました。
echo 共有フォルダへのアクセス権限と、%TEMP% の書き込み権限を確認してください。
echo.
pause
exit /b 1
