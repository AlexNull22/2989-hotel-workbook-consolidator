@echo off
chcp 65001 >nul
setlocal
set "projectUrl=https://www.upwork.com/jobs/Clean-and-merge-messy-hotel-booking-data-from-Excel-sheets-into-one-table_~022113027544652278345/"
if exist "%ProgramFiles%\Mozilla Firefox\firefox.exe" (
  start "" "%ProgramFiles%\Mozilla Firefox\firefox.exe" "%projectUrl%"
) else if exist "%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe" (
  start "" "%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe" "%projectUrl%"
) else if exist "%LOCALAPPDATA%\Mozilla Firefox\firefox.exe" (
  start "" "%LOCALAPPDATA%\Mozilla Firefox\firefox.exe" "%projectUrl%"
) else (
  echo 请先安装Firefox，或自行用Firefox打开以下公开项目链接：
  echo %projectUrl%
  pause
)
endlocal
