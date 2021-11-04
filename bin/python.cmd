@echo off
setlocal
set "PATH=%~dp0\pdftk;%~dp0\python;%PATH%"
"%~dp0\python\python.exe" %*
exit /b %ERRORLEVEL%