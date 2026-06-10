@echo off
rem ===================================================================
rem  Word の STARTUP フォルダをエクスプローラーで開きます。
rem  バッチでの自動コピーがうまくいかないとき、
rem  ここに AutoRuby.dotm を手でドラッグして入れてください。
rem ===================================================================
set "DEST=%APPDATA%\Microsoft\Word\STARTUP"
if not exist "%DEST%" mkdir "%DEST%"
start "" "%DEST%"
