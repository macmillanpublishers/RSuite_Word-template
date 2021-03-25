@echo off
@echo off&setlocal
for %%i in ("%~dp0..") do set "folder=%%~fi"
echo %folder%
SET mypath=%~dp0
echo.

echo "Applying versions to files"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "%mypath%runMacro_Wd_NoSave.ps1 word %folder%\devSetup.docm zz_updateVersionsForRepoTemplates_silent"
if %errorlevel%==0 (
	echo "Finished applying versions"
	) else (
	echo "error encountered"
	exit 1
	)
rem ^ without rethrowing the exitcode above, it isn't surfacing for the gradle job; we want the gradle job to fail on err.
rem echo Exit Code is %errorlevel% 
