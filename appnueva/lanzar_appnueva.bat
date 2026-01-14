@echo off
setlocal
pushd "%~dp0.."

set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"
set "PY=python"
if exist "%VENV_PY%" set "PY=%VENV_PY%"

"%PY%" appnueva\main.py %*

pause
