Attribute VB_Name = "DownloadMac2016"
Option Explicit

' execShell() function courtesy of Robert Knight via StackOverflow
' http://stackoverflow.com/questions/6136798/vba-shell-function-in-office-2011-for-mac

#If Mac Then
 #If MAC_OFFICE_VERSION >= 15 Then
  #If VBA7 Then ' 64-bit Word 2016 for Mac
    Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As LongPtr
    Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
    Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
    Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
  #Else ' 32-bit Word 2016 for Mac
    Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As Long
    Private Declare Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
    Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
    Private Declare Function feof Lib "libc.dylib" (ByVal file As LongPtr) As Long
  #End If
 #Else
  #If VBA7 Then ' does not exist, but why take a chance
    Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As LongPtr
    Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
    Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As LongPtr
    Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
  #Else ' 32-bit Word 2011 for Mac
    Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As Long
    Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
    Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal items As Long, ByVal stream As Long) As Long
    Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long
  #End If
 #End If
#Else ' Word for Windows
' #If VBA7 Then ' Word 2010 or later for Windows
'    Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As Long
'    Private Declare Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
'    Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
'    Private Declare Function feof Lib "libc.dylib" (ByVal file As LongPtr) As Long
' #End If
    Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As LongPtr
    Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
    Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As LongPtr
    Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
#End If

Sub testdl()
    DownloadFileMac ("https://github.com/johnwangel/HCML/blob/master/HCML.dotm?raw=true")
End Sub

Function DownloadFileMac(Path As String) As String


    On Error GoTo Handler

    Dim User, UserPath, dlFile, base_dest, final_dest, FName As String
    Dim Timer As Boolean
    
    Timer = False

    Dim timeout As Variant
    timeout = Now + TimeValue("00:00:10")
    
    User = Environ("USER")
    UserPath = "/Users/" & User & "/Downloads/"
    
    Dim splitString As Variant
    splitString = Split(Path, "/")
    FName = splitString(UBound(splitString))
    splitString = Split(FName, "?")
    FName = splitString(LBound(splitString))
    dlFile = UserPath & FName
    base_dest = "/Users/" & User & "/Library/Containers/com.microsoft.Word/Data/Documents/RSuiteStyleTemplate/"
    
    If InStr(1, Path, "json") Or InStr(1, Path, "txt") Then
        Dim sCmd As String
        Dim sResult As String
        Dim lExitCode As Long
        final_dest = base_dest & "config/" & FName
        sCmd = "curl -L " & Path
        sResult = execShell(sCmd, lExitCode)
        Dim textDoc As Variant
        
        Application.DisplayAlerts = wdAlertsNone
        Documents.Add visible:=False
        ActiveDocument.Content.InsertAfter Text:=sResult
        ActiveDocument.SaveAs FileName:=dlFile, FileFormat:=wdFormatText
        Application.DisplayAlerts = wdAlertsAll
        ActiveDocument.Close
    Else:
'        final_dest = base_dest & FName
        ActiveDocument.FollowHyperlink Address:=Path
    End If

'    If Dir(final_dest) <> "" Then
'        Kill final_dest
'    End If
'
    Timer = True

ReTry:

    If Dir(dlFile) <> "" Then
'        FileCopy dlFile, final_dest
'        Kill dlFile
'        ActiveDocument.Activate
        DownloadFileMac = dlFile
        Exit Function
    End If

Handler:
    If Err.Number = 53 And Timer = False Then
        Resume Next
     ElseIf Err.Number = 53 And Timer = True Then
        If Now > timeout Then
            MsgBox "There is a problem downloading the file. Please check your internet connection."
            End
        Else
            Resume ReTry
        End If
    Else
        MsgBox Err.Number & vbNewLine & Err.Description
        End
    End If
    

End Function



Sub downloadParams()

    Dim sQuery As String
    Dim myURL As String
    myURL = "https://github.com/johnwangel/HCML/blob/master/HCML.dotm?raw=true"
    
    Dim destFile As String
    destFile = "/Users/johnatkins/Downloads/HCML.dotm"
    
    'sQuery = "rm -f " & destFile & " ; curl -L -o " & destFile & " " & myURL
    'sQuery = "curl -L -s -o /dev/null -w '%{http_code}' " & myURL
    sQuery = "curl -L -o /Users/johnatkins/Downloads/HCML.dotm https://github.com/johnwangel/HCML/blob/master/HCML.dotm\?raw\=true"
    
    Dim sResult As String
    Dim lExitCode As Long
    
    'sResult = ShellAndWaitMac2016(sQuery, lExitCode)

End Sub

Function execShell(Command As String, Optional ByRef exitCode As Long) As String
    Dim file As LongPtr
    file = popen(Command, "r")

    If file = 0 Then
        Exit Function
    End If
    
    While feof(file) = 0
        Dim chunk As String
        Dim read As Long
        chunk = Space(50)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        If read > 0 Then
            chunk = Left$(chunk, read)
            execShell = execShell & chunk
        End If
    Wend

    exitCode = pclose(file)
End Function
