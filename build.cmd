@echo off

set app=xlstext.exe

del "%app%" 1>nul 2>nul
call gcc libxls/src/*.c src/*.c -l:libiconv.a -O3 -o %app%
if %errorlevel% == 1 goto:failed

goto:eof

:failed
cd %dp0
echo failed.
pause >nul