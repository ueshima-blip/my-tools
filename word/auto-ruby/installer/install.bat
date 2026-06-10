@echo off
chcp 65001 >nul
rem ===================================================================
rem  自動ルビ振り インストーラー
rem  このバッチと同じフォルダにある AutoRuby.dotm を
rem  Word の STARTUP フォルダにコピーします。
rem  → 次回 Word を開くと、すべての文書で「ルビ振り」が使えます。
rem  （管理者権限は不要です）
rem ===================================================================

set "SRC=%~dp0AutoRuby.dotm"
set "DEST=%APPDATA%\Microsoft\Word\STARTUP"

echo.
echo  自動ルビ振りをインストールします。
echo.

if not exist "%SRC%" (
  echo  [エラー] このバッチと同じ場所に AutoRuby.dotm が見つかりません。
  echo          AutoRuby.dotm を install.bat と同じフォルダに置いてから、
  echo          もう一度ダブルクリックしてください。
  echo.
  pause
  exit /b 1
)

if not exist "%DEST%" mkdir "%DEST%"

copy /Y "%SRC%" "%DEST%\AutoRuby.dotm" >nul
if errorlevel 1 (
  echo  [エラー] コピーに失敗しました。Word を一度すべて閉じてから、
  echo          もう一度お試しください。
  echo.
  pause
  exit /b 1
)

echo  インストールが完了しました。
echo.
echo  次の手順で使えます：
echo    1. 開いている Word をすべて閉じる
echo    2. Word をもう一度開く
echo    3. リボンの「ルビ振り」タブ（または 表示 → マクロ）から実行
echo.
pause
