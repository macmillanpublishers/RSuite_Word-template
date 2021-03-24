#------------------------VARIABLES
param([string]$application, [string]$working_filename, [string]$macroName)

# FOR TESTING
#$application = "excel"
#$working_filename = "WordTemplateStyles.xlsm"
#$macroName="autorun_ToJsonNew"
#$application = "word"
#$working_filename = "styleTemplateCreator.docm"
#$macroName="WriteTemplatefromJson"

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$logfilename = "gradleAddVersions_logfile.txt"
# \/ other version of this file expects host word file to be in same dir as this file; here we pass the full path
$working_file = "$($working_filename)"
$logfile="$($scriptPath)\$($logfilename)"
$workfile_fixed=$working_file -replace '/','\'

#-------------------- LOGGING
$TimestampA=(Get-Date).tostring("yyyy-MM-dd hh:mm:ss")
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $logfile -value "$logstring"
}
LogWrite "$($TimestampA)      : run_macro -- macro: ""$($macroName)."" Received file ""$($workfile_fixed)"", checking filetype."


#--------------------- RUN THE MACRO


If ($application -eq "word") {
	$word = new-object -comobject word.application # create a com object interface (word application)
	$word.visible = $false
	# added try/catch block so we can surface an exit code to the batch file; was exiting 0 even on fail
	# without try/catch and explicit 'exit 1'
	# for more detailed error message during debug; comment/remove the 'exit 1' line
	try {
	$doc = $word.documents.open($workfile_fixed)
	$word.run($macroName)
#	$word.run($macroName, [ref]$workfile_fixed, [ref]$logfile)	#this one for running via batch (deploy) script
#	$word.run($macroName, $workfile_fixed, $logfile) 				#this one for calling direct from cmd line
	$doc.close([ref]$word.WdSaveOptions.wdDoNotSaveChanges)
	$word.quit()
	}
	catch {
	$doc.close([ref]$word.WdSaveOptions.wdDoNotSaveChanges)
	$word.quit()
	exit 1
	}
}
Elseif ($application -eq "excel") {
    $excel = new-object -comobject excel.application
    $excel.visible = $false
    $workbook = $excel.workbooks.open($workfile_fixed)
    $excel.Run($macroName)
    $workbook.close($false)
    $excel.quit()
}
Start-Sleep 1
$TimestampB=(Get-Date).tostring("yyyy-MM-dd hh:mm:ss")
LogWrite "$($TimestampB)      : run_macro -- Macro ""$($macroName)"" completed, exiting .ps1"
