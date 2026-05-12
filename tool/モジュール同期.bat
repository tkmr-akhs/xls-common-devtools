@ECHO OFF
SET CURRENT_PATH=%~dp0
SET ARG1=%1

powershell -ExecutionPolicy bypass -NonInteractive -NoLogo -File "%CURRENT_PATH%sync_modules_main.ps1" "%ARG1%" "%CURRENT_PATH%"

ECHO このウィンドウは 1 分後に自動で終了します。処理は終了しているため、ウィンドウを閉じても、CTRL+C で終了しても問題ありません。
TIMEOUT /T 60 /NOBREAK