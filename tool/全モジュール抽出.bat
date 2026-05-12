@ECHO OFF

powershell -ExecutionPolicy bypass -NonInteractive -NoLogo -File "%~dp0\export_all_main.ps1" "%~dp0\.."

ECHO このウィンドウは 1 分後に自動で終了します。処理は終了しているため、ウィンドウを閉じても、CTRL+C で終了しても問題ありません。
TIMEOUT /T 60 /NOBREAK