@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  タスクボード ランチャー（ワンクリック起動）
rem  - このファイルは index.html と同じフォルダに置いてください
rem  - ダブルクリックすると Edge / Chrome を「アプリ表示（枠なし）」で開きます
rem  - データ保存用の専用プロファイルを使うので、ふだんのブラウザに影響しません
rem  - 保存先プロファイルは消さないため、タスクは次回も残ります
rem ============================================================

set "SRC=%~dp0"
if not exist "%SRC%index.html" goto :nofiles

rem file:// 形式の URL を組み立て（円記号→スラッシュ）
set "URLPATH=%SRC%index.html"
set "URLPATH=%URLPATH:\=/%"
set "APPURL=file:///%URLPATH%"

rem データを残すための専用プロファイル（削除しない）
set "PROFILE=%LOCALAPPDATA%\TaskBoard\profile"
if not exist "%PROFILE%" mkdir "%PROFILE%" 2>nul

rem 1) Microsoft Edge を探す
set "BROWSER="
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"

rem 2) なければ Google Chrome を探す
if not defined BROWSER if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "BROWSER=%LocalAppData%\Google\Chrome\Application\chrome.exe"

if defined BROWSER (
  start "" "%BROWSER%" --app="%APPURL%" --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check --window-size=1100,720
  exit /b 0
)

rem 3) Edge も Chrome も無ければ、既定のアプリで普通に開く
start "" "%SRC%index.html"
exit /b 0

:nofiles
echo.
echo [エラー] この起動ファイルは index.html と同じフォルダに置いてください。
echo 現在の場所: %SRC%
echo.
pause
exit /b 1
