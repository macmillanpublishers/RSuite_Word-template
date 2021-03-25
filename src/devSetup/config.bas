Attribute VB_Name = "config"
Public rswt As New clsDotm 'RSuiteWordTemplate.dotm
Public tsw As New clsDotm 'Template_switcher.dotm
Public dev As New clsDotm 'devSetup.dotm
Public st As New clsDotm  'RSuite.dotx
Public stnc As New clsDotm 'RSuite-NoColor.dotx
Public dotms As Collection

Function defineVBAProjectParams()
Dim currentUser As String
Dim gitBasePath As String
Dim fso As Object

    ' setup
    Set dotms = New Collection
    currentUser = Environ("Username")
    gitBasePath = ThisDocument.Path
    Set fso = CreateObject("scripting.filesystemobject")
    
    'define file 1
    rswt.filename = "RSuite_Word-template.dotm"
    rswt.installedPath = "C:\Users\" & currentUser & "\AppData\Roaming\RSuiteStyleTemplate\" & rswt.filename
    rswt.dotmRepoPath = gitBasePath & "\" & rswt.filename
    rswt.modulesRepoPath = gitBasePath & "\src\" & fso.GetBaseName(rswt.filename)
    ' add it to file collection
    dotms.Add rswt

    ' define file 2
    tsw.filename = "template_switcher.dotm"
    tsw.installedPath = "C:\Users\" & currentUser & "\AppData\Roaming\Microsoft\Word\STARTUP\" & tsw.filename
    tsw.dotmRepoPath = gitBasePath & "\" & tsw.filename
    tsw.modulesRepoPath = gitBasePath & "\src\" & fso.GetBaseName(tsw.filename)
    ' add it to file collection
    dotms.Add tsw
    
    ' define file 3
    dev.filename = "devSetup.docm"
    dev.dotmRepoPath = gitBasePath & "\" & dev.filename
    dev.installedPath = dev.dotmRepoPath
    dev.modulesRepoPath = gitBasePath & "\src\" & fso.GetBaseName(dev.filename)
    ' add it to file collection
    dotms.Add dev
    
    ' define styletemplate 1 (Color, std)
    st.filename = "RSuite.dotx"
    st.dotmRepoPath = gitBasePath & "\StyleTemplate_auto-generate\" & st.filename
    st.installedPath = st.dotmRepoPath
   
    ' define styletemplate 2 (Color, std)
    stnc.filename = "RSuite_NoColor.dotx"
    stnc.dotmRepoPath = gitBasePath & "\StyleTemplate_auto-generate\" & stnc.filename
    stnc.installedPath = stnc.dotmRepoPath
    
End Function
Public Function GetGitBasepath() As String
GetGitBasepath = ThisDocument.Path

End Function

Sub Copy_Installed_Binary_Back_to_Repo(fname As String)
Dim dotm As Variant
Dim curr_dotm As clsDotm
Dim matchcheck As Boolean
Dim fso As Object
    
    ' setup
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    matchcheck = False
    Call defineVBAProjectParams
    
    ' get matching template object
    For Each dotm In dotms
        Debug.Print "fname: " & fname
        Debug.Print "other: " & dotm.filename
        If dotm.filename = fname Then
            Set curr_dotm = dotm
            matchcheck = True
            Exit For
        End If
    Next dotm
    
    ' alert on unknown file
    If matchcheck = False Then
        MsgBox ("Could not cp installed binary to repo: dotm not defined in config, repo path unknown")
        Exit Sub
    End If
    
    ' skip files that already live in the repo all the time (like devSetup)
    If curr_dotm.installedPath = curr_dotm.dotmRepoPath Then
        Exit Sub
    End If
    
SaveFile:
    ' Try to save, if error, try to open the file
    On Error GoTo Opendotms
        Word.Documents(curr_dotm.filename).Save
    On Error GoTo 0
    
    ' attempt to save & copy file from current location to repopath
    Call fso.CopyFile(curr_dotm.installedPath, curr_dotm.dotmRepoPath, True)
    MsgBox "successfully copied " & curr_dotm.filename & " to local cloned Repo-path"
    Exit Sub
    
Opendotms:
    On Error GoTo 0
    'Call Open_All_Defined_VBA_Projects
    Documents.Open filename:=dotm.installedPath
    'Word.Documents(dotm.filename).Save
    Resume SaveFile
End Sub

Function configPathCheck(curr_filename As String)
    Dim dotm As Variant
    Dim filenamecheck As String
    Call defineVBAProjectParams
    filenamecheck = curr_filename
    
    For Each dotm In dotms
        If dotm.filename = curr_filename Then
            filenamecheck = dotm.modulesRepoPath
            Exit For
        End If
    Next dotm
    configPathCheck = filenamecheck
    
End Function

Function IsFileOpen(filename As String)

Dim fileNum As Integer
Dim errNum As Integer

'Allow all errors to happen
On Error Resume Next
fileNum = FreeFile()

'Try to open and close the file for input.
'Errors mean the file is already open
Open filename For Input Lock Read As #fileNum
Close fileNum

'Get the error number
errNum = Err

'Do not allow errors to happen
On Error GoTo 0

'Check the Error Number
Select Case errNum

    'errNum = 0 means no errors, therefore file closed
    Case 0
    IsFileOpen = False
 
    'errNum = 70 means the file is already open
    Case 70
    IsFileOpen = True

    'Something else went wrong
    Case Else
    IsFileOpen = errNum

End Select

End Function
