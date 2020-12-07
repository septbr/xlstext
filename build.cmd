@echo off

set app=xlstext.exe

call gcc libxls/src/*.c src/*.c -l:libiconv.a -O3 -o %app%
if %errorlevel% == 1 goto:failed

echo Ok.
goto:eof

:failed
echo Failed.
pause > nul