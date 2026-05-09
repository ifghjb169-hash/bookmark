@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  python bookmark_tool.py
) else (
  py -3 bookmark_tool.py
)
if errorlevel 1 (
  pause
)
