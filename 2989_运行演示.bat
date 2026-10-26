@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "runtimePython=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%runtimePython%" (
  "%runtimePython%" "2989_合并引擎.py" --source "2989_自造月表.xlsx" --output "2989_新演示结果" --bare-currency USD
) else (
  py -3.12 "2989_合并引擎.py" --source "2989_自造月表.xlsx" --output "2989_新演示结果" --bare-currency USD
)
echo 请查看上方处理结果。重复运行前请使用新的结果目录名称，已有目录不会被覆盖。
pause
endlocal
