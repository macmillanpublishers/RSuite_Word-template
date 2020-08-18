@echo off 
SETLOCAL EnableExtensions
set EXE=WINWORD.EXE
FOR /F %%x IN ('tasklist /NH /FI "IMAGENAME eq %EXE%"') DO IF %%x == %EXE% goto FOUND
FOR /F %%x IN ('tasklist /NH /FI "IMAGENAME eq %EXE%"') DO echo %%x
echo not running
goto FIN
:FOUND
echo running
:FIN