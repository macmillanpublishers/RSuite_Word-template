Attribute VB_Name = "SharedFileInstaller"
' created by Erica Warren - erica.warren@macmillan.com

' ======== PURPOSE ===================================
' Downloads and installs an array of template files & logs the downloads

' ======== DEPENDENCIES =======================================
' This is Part 2 of 2. Must be called from a sub in another module that declares file names and locations.
' The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
' If this is an installer file, The Part 1 code needs to reside in the ThisDocument module as a sub called
' Documents_Open in a .docm file so that it will launch when users open the file.
' Requires the SharedMacros module be installed in the same template.

Option Explicit
Option Base 1

Public Enum TemplatesList
  updaterTemplates = 1
  toolsTemplates = 2
  stylesTemplates = 3
  installTemplates = 4
  allTemplates = 5
End Enum

Public Sub Installer(Installer As Boolean, templateName As String, ByRef _
  TemplatesToInstall As Collection)

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current
'                       file and only updates if out of date.

' Remove file names from collection if they don't need to be updated rignt now
  Dim a As Long
  Dim strFileName As String
  Dim dictFileInfo As Dictionary
  Dim logString As String

  If Installer = False Then
' Not using For Each b/c we need index number to remove from collection
' Counting backwards since removing reassigns index numbers
    For a = TemplatesToInstall.Count To 1 Step -1
      'Need a variant to loop through Collection, but these functions only
      'accept strings as arguments.
      strFileName = TemplatesToInstall(a)
      Set dictFileInfo = FileInfo(strFileName)
      
    ' If exists, check if it's been checked today
      If IsTemplateThere(FileName:=strFileName) = True Then
        logString = Now & " -- " & strFileName & " already exists."
  ' Can't update Log here, b/c CheckLog uses log updates to determine when checked
      ' If HASN'T been checked today, run NeedUpdate
        If CheckLog(LogPath:=dictFileInfo("Log")) = False Then
        ' If DON'T need update, remove from collection
          If NeedUpdate(FileName:=strFileName) = False Then
            TemplatesToInstall.Remove strFileName
          End If
        Else ' Has been checked today, don't update
          TemplatesToInstall.Remove (a)
        End If
      Else
        logString = Now & " -- " & strFileName & " doesn't exist in " & dictFileInfo("Final")
      End If
    Next a
    LogInformation dictFileInfo("Log"), logString
  End If

' If everything is OK, quit sub
  If TemplatesToInstall.Count = 0 Then
    Exit Sub
  End If

  ' Alert user that installation is happening
  Dim strWelcome As String
  If Installer = True Then
    strWelcome = "Welcome to the " & templateName & " Installer!" & vbNewLine _
    & vbNewLine & "Please click OK to begin the installation. It should only" _
    & " take a few seconds."
  Else
    strWelcome = "Your " & templateName & " is out of date. Click OK to " & _
      "update automatically."
  End If
  
  If MsgBox(strWelcome, vbOKCancel, templateName) = vbCancel Then
    MsgBox "Please try to install the files at a later time."
  
    If Installer = True Then
      activeDoc.Close (wdDoNotSaveChanges)
    End If
    Exit Sub
  End If

' ---------------- Close any open docs (with prompt) --------------------------
  Call CloseOpenDocs

' --------- DOWNLOAD FILES ----------------------------------------------------
'If False, error in download; user was notified in DownloadFromGithub function
  Dim varItem As Variant
  For Each varItem In TemplatesToInstall
    strFileName = CStr(varItem)
    If DownloadFromGithub(FileName:=strFileName) = False Then
      If Installer = True Then
        #If Mac Then    ' because application.quit generates error on Mac
          activeDoc.Close (wdDoNotSaveChanges)
        #Else
          Application.Quit (wdDoNotSaveChanges)
        #End If
      Else
        Exit Sub
      End If
    End If

  ' If we just updated the main template, delete the old toolbar
  ' Will be added again by Word-template AutoExec when it's launched
    #If Mac Then
      Dim Bar As CommandBar
      If strFileName = "Word-template.dotm" Then
        For Each Bar In CommandBars
          If Bar.Name = "Macmillan Tools" Then
            Bar.Delete
          End If
        Next
      End If
    #End If
  Next varItem
  
'------Display installation complete message   ---------------------------
  Dim strComplete As String
  Dim strInstallType As String
  
' Quit if it's an installer, but not if it's an updater
' (updater was causing conflicts between Word-template.dotm and GtUpdater)
  If Installer = True Then
    strInstallType = "installed"
  Else
    strInstallType = "updated"
  End If
  strComplete = "The " & templateName & " has been " & strInstallType & " on your computer."
  MsgBox strComplete, vbOKOnly, "Installation Successful"

End Sub


' ===== DownloadJson ==========================================================
' Downloads JSON, if download fails can still continue if a previous version
' is available.
' PARAMS
' FileName: name and extension of JSON file (not path)

' RETURNS:
' Full path to downloaded JSON file.
' If download failed and we don't already have a copy locally, returns null string

Public Function DownloadJson(FileName As String) As String
  Dim dictJsonInfo As Dictionary
  Dim strFinalPath As String
  
  Set dictJsonInfo = FileInfo(FileName)
  strFinalPath = dictJsonInfo("Final")

  If DownloadFromGithub(FileName) = False Then
    If Utils.IsItThere(strFinalPath) = True Then
      DownloadJson = strFinalPath
    Else
      DownloadJson = vbNullString
    End If
  Else
    DownloadJson = strFinalPath
  End If

End Function


Public Function DownloadCSV(FileName As String) As Variant
'---------Download CSV with design specs from Confluence site-------
  Dim dictCsvInfo As Dictionary
  Set dictCsvInfo = FileInfo(FileName)
  Dim strMessage As String
  
' Always download, so don't return CheckLog resutl
  CheckLog LogPath:=dictCsvInfo("Log")

  If DownloadFromGithub(FileName:=FileName) = False Then
    ' If download fails, check if we have an older version of the CSV
    If IsItThere(dictCsvInfo("Final")) = False Then
      strMessage = "Looks like we can't download the design info from the " & _
        "internet right now. Please check your internet connection, or " & _
        "contact workflows@macmillan.com."
      MsgBox strMessage, vbCritical, "Error 5: Download failed, no design file."
      Exit Function
    Else
      strMessage = "Looks like we can't download the most up-to-date design " _
        & "info from the internet right now, so we'll just use the info we " & _
        "have on file for your castoff."
      MsgBox strMessage, vbInformation, "Let's do this thing!"
    End If
  End If
    
' Heading row/col different based on different InfoTypes
  Dim blnRemoveHeaderRow As Boolean
  Dim blnRemoveHeaderCol As Boolean
  
' Because castoff CSV has header row and col, but Spine CSV only has header row
  If InStr(1, FileName, "Castoff") <> 0 Then
    blnRemoveHeaderRow = True
    blnRemoveHeaderCol = True
  ElseIf InStr(1, FileName, "Spine") <> 0 Then
    blnRemoveHeaderRow = True
    blnRemoveHeaderCol = False
  ElseIf InStr(1, FileName, "Styles_Bookmaker") <> 0 Then
    blnRemoveHeaderRow = True
    blnRemoveHeaderCol = False
  End If

  'Double check that CSV is there
  Dim arrFinal() As Variant
  If IsItThere(dictCsvInfo("Final")) = False Then
    strMessage = "The macro is unable to access the data file right now." _
      & " Please check your internet connection and try again, or " & _
      "contact workflows@macmillan.com."
    MsgBox strMessage, vbCritical, "Error 3: CSV doesn't exist"
    Exit Function
  Else
  ' Load CSV into an array
    Dim strPath As String
    strPath = dictCsvInfo("Final")
    Debug.Print strPath
    arrFinal = LoadCSVtoArray(Path:=strPath, RemoveHeaderRow:= _
      blnRemoveHeaderRow, RemoveHeaderCol:=blnRemoveHeaderCol)
  End If

  DownloadCSV = arrFinal

End Function


Public Function GetTemplatesList(TemplatesYouWant As TemplatesList) As _
  Collection
' returns a Collection of template file names to download

  Dim colTemplates As Collection
  Set colTemplates = New Collection

 'get the template switcher file
  If TemplatesYouWant = updaterTemplates Or _
    TemplatesYouWant = installTemplates Or _
    TemplatesYouWant = allTemplates Then
      colTemplates.Add "template_switcher.dotm"
  End If

' get the tools file for these requests
  If TemplatesYouWant = toolsTemplates Or _
    TemplatesYouWant = installTemplates Or _
    TemplatesYouWant = allTemplates Then
      colTemplates.Add "RSuite_Word-template.dotm"
  End If

  ' get the styles files for these requests
  If TemplatesYouWant = stylesTemplates Or _
    TemplatesYouWant = installTemplates Or _
    TemplatesYouWant = allTemplates Then
    colTemplates.Add "RSuite.dotx"
    colTemplates.Add "RSuite_NoColor.dotx"
    'colTemplates.Add "macmillan_CoverCopy.dotm"
  End If

  ' also get the installer file
'  If TemplatesYouWant = allTemplates Then
'    colTemplates.Add "MacmillanTemplateInstaller.docm"
'  End If

  Set GetTemplatesList = colTemplates

End Function


' ===== DownloadFromGithub ================================================
' Actually now it downloads from Github but don't want to mess with things, we're
' going to be totally refacroting soon.

' DEPENDENCIES:
' Add file and download URL info to FullURL function.

Public Function DownloadFromGithub(FileName As String) As Boolean

  Dim dictFullPaths As Dictionary
  Set dictFullPaths = FileInfo(FileName)
  
  Dim strErrMsg As String
  Dim myURL As String
  Dim logString As String
  Dim isMacPre15 As Boolean
  Dim isWindows As Boolean
  Dim tmpPath As String
  
' Get URL to download from.
  myURL = FullURL(FileName:=FileName)
  DebugPrint "Attempting to download: " & myURL
  
  #If Mac Then
    isMacPre15 = IsOldMac
  #Else
    isWindows = True
  #End If
  
  'Run separate download routine if this is Office 2016 or later
  If isMacPre15 = False And isWindows = False Then
    tmpPath = DownloadMac2016.DownloadFileMac(myURL)
  Else:
    tmpPath = dictFullPaths("Tmp")
  End If
  
  If isMacPre15 = True Or isWindows Then
    'Get temp dir based on OS, then download file.
    If isMacPre15 = True Then
        Dim strBashTmp As String
        strBashTmp = Replace(Right(dictFullPaths("Tmp"), Len(dictFullPaths("Tmp")) - _
          (InStr(dictFullPaths("Tmp"), ":") - 1)), ":", "/")
        
        Dim canConnect
        'check for network
        If canConnect <> 0 Then   'can't connect to internet
          logString = Now & " -- Tried update; unable to connect to network."
          LogInformation dictFullPaths("Log"), logString
          strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
                      "Please check your internet connection or contact workflows@macmillan.com for help."
          MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
          DownloadFromGithub = False
          Exit Function
        Else 'internet is working, download file
          'Make sure file is there
          Dim httpStatus As Long
          httpStatus = ShellAndWaitMac("curl -L -s -o /dev/null -w '%{http_code}' " & myURL)
          
          If httpStatus = 200 Then                    ' File is there
            'Now delete file if already there, then download new file
            ShellAndWaitMac ("rm -f " & strBashTmp & " ; curl -L -o " & strBashTmp & " " & myURL)
          ElseIf httpStatus = 404 Then            ' 404 = page not found
            logString = Now & " -- 404 File not found. Cannot download file."
            LogInformation dictFullPaths("Log"), logString
            strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
                        "Please contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FileName & ")"
            DownloadFromGithub = False
            Exit Function
          Else
            logString = Now & " -- Http status is " & httpStatus & ". Cannot download file."
            LogInformation dictFullPaths("Log"), logString
            strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
                "Please check your internet connection or contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 2: Http status " & httpStatus & " (" & FileName & ")"
            DownloadFromGithub = False
            Exit Function
          End If
        End If
      ElseIf isWindows = True Then
        'Check if file is already in tmp dir, delete if yes
        If IsItThere(dictFullPaths("Tmp")) = True Then
          Kill dictFullPaths("Tmp")
        End If
        
        'try to download the file from Public Confluence page
        Dim WinHttpReq As Object
        Dim oStream As Object
        
        'Attempt to download file
        On Error Resume Next
          ' Set WinHttpReq = CreateObject("MSXML2.XMLHTTP.3.0")
          Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
          WinHttpReq.Open "GET", myURL, False
          WinHttpReq.send
    
          ' Exit sub if error in connecting to website
          If Err.Number <> 0 Then 'HTTP request is not OK
            logString = Now & " -- could not connect to website: Error " & Err.Number
            LogInformation dictFullPaths("Log"), logString
            strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
                "Please check your internet connection or contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
            DownloadFromGithub = False
            On Error GoTo 0
            Exit Function
          End If
    '    On Error GoTo 0
      
    'DebugPrint "Http status for " & FileName & ": " & WinHttpReq.Status
    If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
  
      'if connection OK, download file to temp dir
      myURL = WinHttpReq.responseBody
      Set oStream = CreateObject("ADODB.Stream")
      oStream.Open
      oStream.Type = 1
      oStream.Write WinHttpReq.responseBody
      oStream.SaveToFile dictFullPaths("Tmp"), 2 ' 1 = no overwrite, 2 = overwrite
      oStream.Close
      Set oStream = Nothing
      Set WinHttpReq = Nothing
    ElseIf WinHttpReq.Status = 404 Then ' 404 = file not found
      logString = Now & " -- 404 File not found. Cannot download file."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
          "Please contact workflows@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FileName & ")"
      DownloadFromGithub = False
      Exit Function
    Else
      logString = Now & " -- Http status is " & WinHttpReq.Status & ". Cannot download file."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
          "Please check your internet connection or contact workflows@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 2: Http status " & WinHttpReq.Status & " (" & FileName & ")"
      DownloadFromGithub = False
      Exit Function
    End If
  End If
End If

  'Error if download was not successful
  If IsItThere(tmpPath) = False Then
    logString = Now & " -- " & FileName & " file download to Temp was not successful."
    LogInformation dictFullPaths("Log"), logString
    strErrMsg = "There was an error downloading the Macmillan template." & vbNewLine & _
        "Please contact workflows@macmillan.com for assitance."
    MsgBox strErrMsg, vbCritical, "Error 3: Download failed (" & FileName & ")"
    DownloadFromGithub = False
    On Error GoTo 0
    Exit Function
  Else
    logString = Now & " -- " & FileName & " file download to Temp was successful."
    LogInformation dictFullPaths("Log"), logString
  End If

  'If file exists already, log it and delete it
  If IsItThere(dictFullPaths("Final")) = True Then

    logString = Now & " -- Previous version file in final directory."
    LogInformation dictFullPaths("Log"), logString
    
    ' get file extension
    Dim strExt As String
    strExt = Utils.GetFileExtension(dictFullPaths("Final"))
    
    ' can't delete template if it's installed as an add-in
    If InStr(strExt, "dot") > 0 Then
      Utils.UnloadAddIn (dictFullPaths("Final"))
    End If

    ' Test if dir is read only
    Dim strFinalDir As String
    strFinalDir = Utils.GetPath(dictFullPaths("Final"))
    If IsReadOnly(strFinalDir) = True Then ' Dir is read only
      logString = Now & " -- old " & FileName & " file is read only, can't delete/replace. " _
          & "Alerting user."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "The installer doesn't have permission. Please conatct workflows" & _
          "@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 8: Permission denied (" & FileName & ")"
      DownloadFromGithub = False
      On Error GoTo 0
      Exit Function
    Else
      On Error Resume Next
        Kill dictFullPaths("Final")
        
        If Err.Number = 70 Then         'File is open and can't be replaced
          logString = Now & " -- old " & FileName & " file is open, can't delete/replace. Alerting user."
          LogInformation dictFullPaths("Log"), logString
          strErrMsg = "Please close all other Word documents and try again."
          MsgBox strErrMsg, vbCritical, "Error 4: Previous version removal failed (" & FileName & ")"
          DownloadFromGithub = False
          On Error GoTo 0
          Exit Function
        End If
      On Error GoTo 0
    End If
  Else
    logString = Now & " -- No previous version file in final directory."
    LogInformation dictFullPaths("Log"), logString
  End If
      
  'If delete was successful, move downloaded file to final directory
  Dim there As Boolean
  there = IsItThere(dictFullPaths("Final"))
  If there = False Then
    logString = Now & " -- Final directory clear of " & FileName & " file."
    LogInformation dictFullPaths("Log"), logString
    
    ' move template to final directory
    Name dictFullPaths("Tmp") As dictFullPaths("Final")
    
    'Mac won't load macros from a template downloaded from the internet to Startup.
    'Need to send these commands for it to work, see Confluence
    ' Do NOT use open/save as option, this removes customUI which creates Mac Tools toolbar later
    #If Mac Then
      If InStr(1, FileName, ".dotm") Then
        Dim strCommand As String
        strCommand = "do shell script " & Chr(34) & "xattr -wx com.apple.FinderInfo \" & Chr(34) & _
            "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00\" & _
            Chr(34) & Chr(32) & Chr(34) & " & quoted form of POSIX path of " & Chr(34) & dictFullPaths("Final") & Chr(34)
            'DebugPrint strCommand
            MacScript (strCommand)
      End If
    #End If
  
  Else
    logString = Now & " -- old " & FileName & " file not cleared from Final directory."
    LogInformation dictFullPaths("Log"), logString
    strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
        "Please close all other Word documents and try again, or contact workflows@macmillan.com."
    MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & FileName & ")"
    DownloadFromGithub = False
    On Error GoTo 0
    Exit Function
  End If
  
  'If move was successful, yay! Else, :(
  If IsItThere(dictFullPaths("Final")) = True Then
    logString = Now & " -- " & FileName & " file successfully saved to final directory."
    LogInformation dictFullPaths("Log"), logString
  Else
    logString = Now & " -- " & FileName & " file not saved to final directory."
    LogInformation dictFullPaths("Log"), logString
    strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
        "Please cotact workflows@macmillan.com for assistance."
    MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & FileName & ")"
    DownloadFromGithub = False
    On Error GoTo 0
    Exit Function
  End If
  
  'Cleanup: Get rid of temp file if downloaded correctly
  If IsItThere(dictFullPaths("Tmp")) = True Then
    Kill dictFullPaths("Tmp")
  End If
  
  ' Disable Startup add-ins so they don't launch right away and mess of the code that's running
  If InStr(1, LCase(dictFullPaths("Final")), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
    On Error Resume Next                                        'Error = add-in not available, don't need to uninstall
      AddIns(dictFullPaths("Final")).Installed = False
    On Error GoTo 0
  End If
  
  DownloadFromGithub = True

End Function


Private Sub LogInformation(LogFile As String, LogMessage As String)

' Create parent dir if it doesn't exist yet
  If Utils.ParentDirExists(LogFile) = False Then
    MkDir Utils.GetPath(LogFile)
  End If

  Dim FileNum As Integer
  FileNum = FreeFile ' next file number
  Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
  Print #FileNum, LogMessage ' write information at the end of the text file
  Close #FileNum ' close the file
End Sub

Public Function LoadCSVtoArray(Path As String, RemoveHeaderRow As Boolean, _
  RemoveHeaderCol As Boolean) As Variant

'------Load CSV into 2d array, NOTE!!: base 0---------
' But also note that this now removes the header row and column too
    Dim fnum As Integer
    Dim whole_file As String
    Dim lines As Variant
    Dim one_line As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim the_array() As Variant
    Dim R As Long
    Dim c As Long
    
        If IsItThere(Path) = False Then
            MsgBox "There was a problem with your Castoff.", vbCritical, "Error: CSV not available"
            Exit Function
        End If
        'DebugPrint Path
        
        ' Do we need to remove a header row?
        Dim lngHeaderRow As Long
        If RemoveHeaderRow = True Then
            lngHeaderRow = 1
        Else
            lngHeaderRow = 0
        End If
        
        ' Do we need to remove a header column?
        Dim lngHeaderCol As Long
        If RemoveHeaderCol = True Then
            lngHeaderCol = 1
        Else
            lngHeaderCol = 0
        End If
        
        ' Load the csv file.
        fnum = FreeFile
        Open Path For Input As fnum
        whole_file = Input$(LOF(fnum), #fnum)
        Close fnum
        
        ' Break the file into lines (trying to capture whichever line break is used)
        If InStr(1, whole_file, vbCrLf) <> 0 Then
            lines = Split(whole_file, vbCrLf)
        ElseIf InStr(1, whole_file, vbCr) <> 0 Then
            lines = Split(whole_file, vbCr)
        ElseIf InStr(1, whole_file, vbLf) <> 0 Then
            lines = Split(whole_file, vbLf)
        Else
            MsgBox "There was an error with your castoff.", vbCritical, "Error parsing CSV file"
        End If

        ' Dimension the array.
        num_rows = UBound(lines)
        one_line = Split(lines(0), ",")
        num_cols = UBound(one_line)
        
        num_cols = 1
        lngHeaderCol = 1

        ReDim the_array(num_rows - lngHeaderRow, num_cols - lngHeaderCol) ' -1 if we are not using header row/col
        
        ' Copy the data into the array.
        For R = lngHeaderRow To num_rows           ' start at 1 (not 0) if we are not using the header row
            If Len(lines(R)) > 0 Then
                one_line = Split(lines(R), ",")
                For c = lngHeaderCol To num_cols   ' start at 1 (not 0) if we are not using the header column
                    'DebugPrint one_line(c)
                    the_array((R - lngHeaderRow), (c - lngHeaderCol)) = one_line(c)   ' -1 because if are not using header row/column from CSV
                Next c
            End If
        Next R
    
        ' Prove we have the data loaded.
'         DebugPrint LBound(the_array)
'         DebugPrint UBound(the_array)
'         For R = 0 To (num_rows - 1)          ' -1 again if we removed the header row
'             For c = 0 To num_cols      ' -1 again if we removed the header column
'                 DebugPrint the_array(R, c) & " | ";
'             Next c
'             DebugPrint
'         Next R
'         DebugPrint "======="
    
    LoadCSVtoArray = the_array
 
End Function



Private Function CheckLog(LogPath As String) As Boolean
' LogPath = full path to log file we're checking
' REturns TRUE if file has been updated today.
  Dim logString As String
  Dim strLogDir As String
  Dim strStylesDir As String
  
  strLogDir = Utils.GetPath(LogPath)
  strStylesDir = WT_Settings.StyleDir

' Have to create "log" directory here, bc creation elsewhere marks as "updated"
  If IsItThere(LogPath) = False Then
    CheckLog = False
    logString = Now & " -- Creating logfile."
    If IsItThere(strLogDir) = False Then
      If IsItThere(strStylesDir) = False Then
        MkDir (strStylesDir)
        MkDir (strLogDir)
        logString = logString & vbNewLine & Now & " -- Creating RsuiteStyleTemplate directory."
      Else
        MkDir (strLogDir)
        logString = logString & vbNewLine & Now & " -- Creating log directory."
      End If
    End If
  Else    'logfile exists, so check last modified date
    Dim lastModDate As Date
    lastModDate = FileDateTime(LogPath)
    If DateDiff("d", lastModDate, Date) < 1 Then       'i.e. 1 day
      CheckLog = True
      logString = logString & vbNewLine & Now & " -- Already checked less than 1 day ago."
    Else
      CheckLog = False
      logString = logString & vbNewLine & Now & " -- >= 1 day since last update check."
    End If
  End If
  
  'Log that info!
  LogInformation LogPath, logString
    
End Function


Private Function IsTemplateThere(FileName As String) As Boolean
  Dim dictTemplateInfo As Dictionary
  Set dictTemplateInfo = FileInfo(FileName)

  Dim logString As String
  Dim strDir As String
  strDir = Utils.GetPath(dictTemplateInfo("Final"))

  If IsItThere(strDir) = False Then
    MkDir (strDir)
    IsTemplateThere = False
  Else
  ' Check if template file exists
    If IsItThere(dictTemplateInfo("Final")) = False Then
      IsTemplateThere = False
    Else
      IsTemplateThere = True
    End If
  End If
End Function

Private Function NeedUpdate(FileName As String) As Boolean
' FileName should be to TEMPLATE (.dotm) file
  Dim dictTemplateFile As Dictionary
  Dim strFilePath As String
  Set dictTemplateFile = FileInfo(FileName)
  strFilePath = dictTemplateFile("Final")

'----- Get local version number -----------------------------------
  Dim logString As String
  Dim strLocalVersion As String

  If IsItThere(strFilePath) = True Then
    #If Mac Then
      Call OpenDocMac(strFilePath)
    #Else
      Call OpenDocPC(strFilePath)
    #End If

    strLocalVersion = Documents(strFilePath).CustomDocumentProperties("Version")
    Documents(strFilePath).Close SaveChanges:=wdDoNotSaveChanges
    logString = Now & " -- Local version is " & strLocalVersion
    LogInformation dictTemplateFile("Log"), logString
  Else
    NeedUpdate = True
    logString = Now & " -- No template installed, update required."
    LogInformation dictTemplateFile("Log"), logString
    Exit Function
  End If

'------------------------- Get latest version from config files ---------------
  Dim dictConfigData As Dictionary
  Dim strLatestVersion As String

  Set dictConfigData = GetWorkingData(FileName)
  ' If branch is listed, that takes precedence over latest_release
  If dictConfigData.Exists("branch") = True Then
    NeedUpdate = True
    logString = Now & " -- Specific branch listed, proceeding with direct download."
    LogInformation dictTemplateFile("Log"), logString
    Exit Function
  ElseIf dictConfigData.Exists("latest_release") = True Then
    strLatestVersion = dictConfigData("latest_release")
    logString = Now & " -- Latest release is " & strLatestVersion
    LogInformation dictTemplateFile("Log"), logString
  Else
  ' If no branch OR latest_release, default to master branch
    NeedUpdate = True
    logString = Now & " -- No specific branch or release listed, proceeding " & _
      "with direct download from master branch."
    LogInformation dictTemplateFile("Log"), logString
    Exit Function
  End If

' -------------------- Compare version numbers --------------------------------
' Convert version strings to arrays
  Dim arrLocalVersion As Variant
  Dim arrLatestVersion As Variant
  arrLocalVersion = ParseVersion(strLocalVersion)
  arrLatestVersion = ParseVersion(strLatestVersion)

' Compare each element of the versions to each other
  Dim arrBase As Variant
  Dim arrComp As Variant
  Dim blnBaseLocal As Boolean
  Dim iBase As Long
  Dim iComp As Long
  Dim lngLB As Long
  Dim lngUB As Long
  Dim lngIndexDiff As Long
  Dim lngEqualItems As Long

' Might be different lengths, need to loop through shorter array to avoid errors
  If UBound(arrLocalVersion) < UBound(arrLatestVersion) Then
    arrBase = arrLocalVersion
    arrComp = arrLatestVersion
    blnBaseLocal = True
  Else
    arrBase = arrLatestVersion
    arrComp = arrLocalVersion
    blnBaseLocal = False
  End If
  
  lngLB = LBound(arrBase)
  lngUB = UBound(arrBase)
' Becasue can't guarantee both will be base 0 or base 1
  lngIndexDiff = LBound(arrComp) - lngLB

  For iBase = lngLB To lngUB
    iComp = iBase + lngIndexDiff
    If arrBase(iBase) > arrComp(iComp) Then
      NeedUpdate = Not blnBaseLocal
      Exit For
    ElseIf arrBase(iBase) < arrComp(iComp) Then
      NeedUpdate = blnBaseLocal
      Exit For
    Else
      lngEqualItems = lngEqualItems + 1
    ' Handling for one version having more elements than the other
      If lngEqualItems = lngUB And UBound(arrComp) > UBound(arrBase) Then
        NeedUpdate = blnBaseLocal
      End If
    End If
  Next iBase
  
  If NeedUpdate = False Then
    logString = Now & " -- No update needed."
  Else
    logString = Now & " -- Updating to newer release."
  End If

  LogInformation dictTemplateFile("Log"), logString

End Function

Private Sub OpenDocMac(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True ', Visible:=False      'Mac Word 2011 doesn't allow Visible as an argument :(
End Sub

Private Sub OpenDocPC(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True, visible:=False      'Win Word DOES allow Visible as an argument :)
End Sub

' ===== ParseVersion ==========================================================
' Convert version number string to individual integers for semantic versioning.

' RETURNS
' version segments (e.g., major, minor, patch) as an array

Private Function ParseVersion(VersionStr As String) As Variant
' Check for prefix and remove
  If Left(VersionStr, 1) = "v" Then
    VersionStr = Right(VersionStr, Len(VersionStr) - 1)
  End If
  
' Split string on points
  ParseVersion = Split(VersionStr, ".")
End Function

' ===== FullURL ===============================================================
' Takes file name (with extension) you want to download as an argument, returns
' URL to download that file from, based on config file info.

Private Function FullURL(FileName As String) As String
  Dim dictWorkingData As Dictionary
  Set dictWorkingData = GetWorkingData(FileName)

' Add path elements to a collection, we'll Join later
  Dim collPath As Collection
  Set collPath = New Collection

  With collPath
  ' Start with base URL
    .Add dictWorkingData("base_url")
  
  ' Add next two elements (always same format)
    .Add dictWorkingData("owner")
    .Add dictWorkingData("repo")
  
  ' Add string for direct download
    .Add "raw"
  
  ' Add branch or tag (branch takes precedence if it exists)
    If dictWorkingData.Exists("branch") = True Then
      .Add dictWorkingData("branch")
    ElseIf dictWorkingData.Exists("latest_release") = True Then
      .Add dictWorkingData("latest_release")
    Else
    ' If no branch or latest_release in any config, defaults to master branch
      .Add "master"
    End If
    
  ' Add subfolders if we have any
    If dictWorkingData.Exists("subfolders") = True Then
      Dim collSubfolders As Collection
      Set collSubfolders = dictWorkingData("subfolders")

      Dim varDir As Variant
      For Each varDir In collSubfolders
        .Add varDir
      Next varDir
    End If

  ' Add the file name to the end!
    .Add FileName
  End With
  
' No native Join function for Collections, so convert to an array first
  Dim varPathArray As Variant
  varPathArray = Utils.ToArray(collPath)
  
  FullURL = Join(varPathArray, "/")
End Function

' ===== GetWorkingData ========================================================
' Loop through all config files in order of least important to more important
' and add each value to working dictionary. Items that take precedence will
' overwrite the value from the previous configs.

' only gets data from "files" object for specific file

Private Function GetWorkingData(FileName As String) As Dictionary
  Dim collConfigs As Collection
  Set collConfigs = New Collection
  Dim objDict As Object
  Dim dictWorkingData As Dictionary
  
' global config data handled separately
  If FileName = "global_config.json" Then
    Set GetWorkingData = GetGlobalConfigData
    Exit Function
  Else
  ' Read data from global config file for other files
    collConfigs.Add WT_Settings.GlobalConfig
  End If

' region config data from info in global config and local config files
'  If InStr(FileName, "_region_config") = False Then
'    collConfigs.Add WT_Settings.RegionConfig
'  End If

' Always add local config (though might be empty Dictionary)
  collConfigs.Add WT_Settings.LocalConfig
  
' Add each config in order
  Set dictWorkingData = New Dictionary
  For Each objDict In collConfigs
    AddConfigData DestinationDictionary:=dictWorkingData, file:=FileName, _
      ConfigData:=objDict
  Next objDict
    
  Set GetWorkingData = dictWorkingData
End Function

' ===== GetGlobalConfigData ===================================================
' data to download global_config.json is handled differently than all other files
' since we can't get the data from the file we don't yet have. So those items
' are all stored in the CustomDocumentProperties of ThisDocument.

' Needs to check against local_config.json in case we have a branch overrride.

Private Function GetGlobalConfigData() As Dictionary
  Dim dictConfig As Dictionary
  Set dictConfig = New Dictionary

' Read from CustomDocumentProperties
  With dictConfig
  ' If value is a JSON object or array, we've written the full JSON string to
  ' the value of the CustomDocProp and appended ".json" to the key

    Dim docProps As DocumentProperties
    Set docProps = ThisDocument.CustomDocumentProperties

    Dim varProperty As DocumentProperty
    Dim strShortKey As String
    If docProps.Count > 0 Then
      For Each varProperty In docProps
        If Utils.GetFileExtension(varProperty.Name) = "json" Then
          .Add Key:=Utils.GetFileNameOnly(varProperty.Name), _
            item:=JsonConverter.ParseJson(JsonString:=varProperty.value)
        Else
          .Add Key:=varProperty.Name, item:=varProperty.value
        End If
      Next varProperty
    End If
  End With
  
  ' If local_config exists, add that data as well. Note items with the same key
  ' will overwrite anything already added to the dictionary.
  AddConfigData DestinationDictionary:=dictConfig, file:="global_config.json", _
    ConfigData:=WT_Settings.LocalConfig
  
  Set GetGlobalConfigData = dictConfig
End Function

' ===== AddConfigData =========================================================
' Adds data from a config.json file to an already existant dictionary. If a key
' being added already exists in the dictionary, the value will be overwritten.

' PARAMS:
' DestinationDictionary: BY REF, so it changes the object passed to it (don't
' need to return a new dictionary object.

' File: Name of the file that we're trying to download/get data ABOUT

' ConfigData: the config data we're adding

Private Sub AddConfigData(ByRef DestinationDictionary As Dictionary, file As String, _
  ConfigData As Dictionary)
  
  Dim varKey1 As Variant
  Dim varKey2 As Variant
  Dim dictFileData As Variant

  With DestinationDictionary
    For Each varKey1 In ConfigData.Keys
    ' only get data for the file we're looking for
      If varKey1 = "files" Then
        If ConfigData(varKey1).Exists(file) Then
          Set dictFileData = ConfigData("files")(file)
          For Each varKey2 In dictFileData.Keys
            If IsObject(dictFileData(varKey2)) = True Then
              Dim col As Collection
              Set col = dictFileData(varKey2)
              Set .item(varKey2) = col
            Else
              .item(varKey2) = dictFileData(varKey2)
            End If
          Next varKey2
        End If
      Else
        If IsObject(ConfigData(varKey1)) = True Then
          Set .item(varKey1) = ConfigData(varKey1)
        Else
          .item(varKey1) = ConfigData(varKey1)
        End If
      End If
    Next varKey1
  End With

End Sub

Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function


Private Sub CloseOpenDocs()

  '-------------Check for/close open documents---------------------------------
  Dim strInstallerName As String
  Dim strSaveWarning As String
  Dim objDocument As Document
  Dim b As Long
  Dim Doc As Document
  
  strInstallerName = ThisDocument.Name

  'MsgBox "Installer Name: " & strInstallerName
  'MsgBox "Open docs: " & Documents.Count


  If Documents.Count > 1 Then
    strSaveWarning = "All other Word documents must be closed to run the macro." & vbNewLine & vbNewLine & _
      "Click OK and I will save and close your documents." & vbNewLine & _
      "Click Cancel to exit without running the macro and close the documents yourself."
    If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
      activeDoc.Close
      Exit Sub
    Else
      For Each Doc In Documents
        On Error Resume Next        'To skip error if user is prompted to save new doc and clicks Cancel
          'DebugPrint doc.Name
          If Doc.Name <> strInstallerName Then       'But don't close THIS document
            Doc.Save   'separate step to trigger Save As prompt for previously unsaved docs
            Doc.Close
          End If
        On Error GoTo 0
      Next Doc
    End If
  End If
    
End Sub



' ===== FileInfo ==============================================================
' Returns dictionary with paths to final file location and other things for
' downloading.

' PARAMS
' FileName: File name with extension that you want to download.

Public Function FileInfo(FileName As String) As Dictionary
  Dim dictFileInfo As Dictionary
  Set dictFileInfo = New Dictionary
  Dim strStyleDir As String
  Dim strTmpDir As String
  Dim strBaseName As String
  Dim strFinalPath As String
  Dim strLogPath As String
  Dim thisVersion
  Dim strFilePath As String

  Dim strSep As String
  strSep = Application.PathSeparator

' Full path to file in Tmp dir
  strTmpDir = WT_Settings.TmpDir & strSep & FileName

' Get root Style directory
  strStyleDir = WT_Settings.StyleDir

' Create full path to the log file for this file
  strBaseName = Utils.GetFileNameOnly(FileName)
  strLogPath = strStyleDir & strSep & "log" & strSep & strBaseName & "_updates.log"

' Create full path to file in final location
' Config files go in their own subfolder
' Style templates go in their own subfolder
  If FileName Like "*_config*" Then
  ' Verify that directory exists, create if it doesn't
    strFilePath = strStyleDir & strSep & "config"
    If Utils.IsItThere(strFilePath) = False Then
      MkDir strFilePath
    End If
  ' add to path
    strFinalPath = strFilePath
  ElseIf FileName Like "template_switcher*" Then
    strFinalPath = Application.StartupPath
  ElseIf FileName Like "AutoMacros*" Then
    strFinalPath = strStyleDir & strSep & "modules"
  ElseIf FileName Like "RSuite_Word-template*" Then
    strFinalPath = strStyleDir
  ElseIf FileName Like "RSuite*" Then
    ' Verify that directory exists, create if it doesn't
    strFilePath = strStyleDir & strSep & "StyleTemplate_auto-generate"
    If Utils.IsItThere(strFilePath) = False Then
      MkDir strFilePath
    End If
  ' add to path
    strFinalPath = strFilePath
  Else
    strFinalPath = strStyleDir
  End If

  strFinalPath = strFinalPath & strSep & FileName
  
  With dictFileInfo
    .Add "Final", strFinalPath
    .Add "Tmp", strTmpDir
    .Add "Log", strLogPath
  End With
  
  Set FileInfo = dictFileInfo

End Function


Function DownloadFile()

Dim myURL As Variant
Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", "https://www.google.com"
WinHttpReq.send

myURL = WinHttpReq.responseBody

MsgBox WinHttpReq.Status

'If WinHttpReq.Status = 200 Then
'    Set oStream = CreateObject("ADODB.Stream")
'    oStream.Open
'    oStream.Type = 1
'    oStream.Write WinHttpReq.responseBody
'    oStream.SaveToFile saveToPath, 2 ' 1 = no overwrite, 2 = overwrite
'    oStream.Close
'End If

End Function


