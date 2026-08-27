@echo off
python "%~dp0..\ci_pyinstaller.py" %*
exit /b %ERRORLEVEL%
