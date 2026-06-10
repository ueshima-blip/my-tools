@echo off
chcp 65001 >nul
rem ===================================================================
rem  自動ルビ振り アンインストーラー
rem  Word の STARTUP フォルダから AutoRuby.dotm を取り除きます。
rem  （元に戻したいときに使います。管理者権限は不要です）
rem ===================================================================

set "TARGET=%APPDATA%\Microsoft\Word\STARTUP\AutoRuby.dotm"

echo.
if not exist "%TARGET%" (
  echo  自動ルビ振りは入っていません（すでに削除済みです）。
  echo.
  pause
  exit /b 0
)

del "%TARGET%"
if errorlevel 1 (
  echo  [エラー] 削除に失敗しました。Word を一度すべて閉じてから、
  echo          もう一度お試しください。
  echo.
  pause
  exit /b 1
)

echo  自動ルビ振りを削除しました。
echo  開いている Word を閉じて、もう一度開くと反映されます。
echo.
pause
