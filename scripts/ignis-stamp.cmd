@echo off
"%~dp0\..\bin\python.cmd" "%~dp0\ignis-stamp.py" %*
exit /b %errorlevel%