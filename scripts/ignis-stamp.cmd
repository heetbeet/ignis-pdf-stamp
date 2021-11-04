@echo off
call "%~dp0\..\bin\python.cmd" "%~dp0\ignis-stamp.py" %*
if "%errorlevel%" neq "0" pause