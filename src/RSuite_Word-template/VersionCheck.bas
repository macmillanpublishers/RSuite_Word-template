Attribute VB_Name = "VersionCheck"
Option Explicit
Sub CheckMacmillanGT()
' can't change name to Word-template because "CheckMacmillanGT" is in customUI.xml
' and I don't want to muck with it right now.
'----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim templateFile As String
    Dim strMacDocs As String
    Dim strTemplatePath As String
    
    templateFile = "RSuite_Word-template.dotm"  'the template file you are checking
    strTemplatePath = WT_Settings.StyleDir(FileType:="tools")

    Call VersionCheck(strTemplatePath, templateFile)

End Sub
Sub CheckMacmillan()

    '----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim templateFile As String
    Dim strTemplatePath As String
    
    templateFile = "RSuite.dotx"  'the template file you are checking
    strTemplatePath = WT_Settings.StyleDir(FileType:="styles")
    
    Call VersionCheck(strTemplatePath, templateFile)

End Sub
Private Sub VersionCheck(fullPath As String, FileName As String)

    '------------------------------
    'created by Erica Warren 2014-04-08         erica.warren@macmillan.com
    'Alerts user to the version number of the template file
    Dim installedVersion As String
    
    installedVersion = "v" + GetVersion(fullPath, FileName)

    'Now we tell the user what version they have
    If installedVersion <> "none" Then
        MsgBox "You currently have version " & installedVersion & " of the file " & FileName & " installed."
    Else
        MsgBox "You do not have " & FileName & " installed on your computer."
    End If

End Sub

Public Function GetVersion(ByVal fullPath As String, ByVal FileName As String)
    Dim installedVersion As String
    Dim fullFilePath As String
    'DebugPrint fullPath
    
    fullFilePath = fullPath & Application.PathSeparator & FileName
    
    If IsItThere(fullFilePath) = False Then            ' the template file is not installed, or is not in the correct place
        installedVersion = "none"
    Else                                                                'the template file is installed in the correct place
        Documents.Open FileName:=fullFilePath, ReadOnly:=True                   ' Note can't set Visible:=False because that's not an argument in Word Mac VBA :(
        installedVersion = Documents(fullFilePath).CustomDocumentProperties("version")
        Documents(fullFilePath).Close
    End If
    
    If Left(installedVersion, 1) = "v" Then
        GetVersion = Right(installedVersion, Len(installedVersion) - 1)
    Else
        GetVersion = installedVersion
    End If

End Function

Private Sub SetVersion()
    Dim d As Variant
    Dim Prop As Variant
    
    ActiveDocument.CustomDocumentProperties.Add Name:="repo", LinkToContent:=False, value:="RSuite_Word-template", Type:=msoPropertyTypeString
    ActiveDocument.Save
    
'    For Each d In Documents
'        If d.Name = "RSuite_Word-template.dotm" Then
'            d.CustomDocumentProperties("version").Value = "v1.0"
'            d.CustomDocumentProperties.Add Name:="repo", LinkToContent:=False, Value:="RSuite_Word-template"
'            MsgBox d.CustomDocumentProperties("version")
'            MsgBox d.CustomDocumentProperties("repo")
'        End If
'
'    Next

End Sub

Private Sub seeVersion()
    MsgBox ActiveDocument.CustomDocumentProperties("repo").value
    MsgBox ActiveDocument.CustomDocumentProperties("version").value
End Sub
