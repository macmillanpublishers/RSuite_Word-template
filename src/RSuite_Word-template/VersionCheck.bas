Attribute VB_Name = "VersionCheck"
Option Private Module
Option Explicit
Sub AttachedVersion()
    ' declare valid style template files
    Dim templateFile As String, templateFileNoColor As String
    Dim templateNameStr As String
    Dim strTemplatePath As String
    Dim installVersion As String, cdpVersion As String, attachVersion As String
    Dim msgstr As String
    
    templateFile = "RSuite.dotx"  'the template file you are checking
    templateFileNoColor = "RSuite_NoColor.dotx"
    strTemplatePath = WT_Settings.StyleDir(FileType:="styles")
    
    ' see if valid template files attached, if so get version & report.
    If ActiveDocument.AttachedTemplate = templateFile Then
        attachVersion = GetVersion(strTemplatePath, templateFile)
        templateNameStr = templateFile
        GoTo ShowMsgbox
    ElseIf ActiveDocument.AttachedTemplate = templateFileNoColor Then
        attachVersion = GetVersion(strTemplatePath, templateFileNoColor)
        templateNameStr = templateFileNoColor
        GoTo ShowMsgbox
    End If
    
    ' if valid template files not attached, check version from DocProps
    If attachVersion = "" Then
        cdpVersion = customDocPropValue("Version")
        templateNameStr = customDocPropValue("TemplateName")
        ' nothing currently attached, nothing ever attached:
        If cdpVersion = "" Then
            msgstr = "Unable to determine this document's RSuite-styles version:" + vbCr + vbCr + _
                "Please click 'Activate Template' in the RSuite Tools toolbar and check again."
            GoTo ShowMsgbox
        ' let's see if installed version matches
        Else
            If templateNameStr = "" Then templateNameStr = templateFile
            installVersion = GetVersion(strTemplatePath, templateNameStr)
            ' if no installversion, versions match, or comparison fails, use cdpversion:
            If installVersion = "none" Or _
                versionCompare(cdpVersion, installVersion) = "same" Or _
                versionCompare(cdpVersion, installVersion) = "unable to compare" Then
                attachVersion = cdpVersion
                GoTo ShowMsgbox
            ' warn if local style-template is older
            ElseIf versionCompare(cdpVersion, installVersion) = ">" Then
                msgstr = "The version of RSuite styles in this document (v" & cdpVersion & _
                    ") is newer than the version of RSuite styles in your installed RSuite style-template (v" & installVersion & _
                    ")" & vbCr & vbCr & _
                    "Please contact workflows@macmillan.com for assistance in getting your installed template updated!"
                GoTo ShowMsgbox
            ' warn if docstyles are older
            ElseIf versionCompare(cdpVersion, installVersion) = "<" And cdpVersion >= 6 Then
                msgstr = "The version of RSuite styles in this document (v" & cdpVersion & _
                    ") is older than the version of RSuite styles in your installed RSuite style-template (v" & installVersion & _
                    ")" & vbCr & vbCr & _
                    "Please click 'Activate Template' in the RSuite Tools toolbar to update this document to the latest RSuite styles!"
                GoTo ShowMsgbox
            ' warn if pre-rsuite styles as per older version cdp declaration
            ElseIf cdpVersion < 6 Then
                 msgstr = "This document may be styled with Macmillan's legacy, pre-RSuite style-set." & vbCr & vbCr & _
                    "You may want to update/edit styles using the old Macmillan template, or you can click 'Activate Template'" & _
                    " in the RSuite Tools toolbar to add RSuite styles." & vbCr & vbCr & _
                    "If you're not sure what to do, please reach out to workflows@macmillan.com for assistance!"
                GoTo ShowMsgbox
            End If
        End If
    End If
    Exit Sub
    
ShowMsgbox:
    ' if we haven't already set a problem msgstr, then we set up a success-one
    If msgstr = "" Then
        msgstr = "This document is using the RSuite style-set, version v" + attachVersion + "."
        If templateNameStr <> "" Then
            msgstr = msgstr + vbCr + vbCr + "(from template file: '" + templateNameStr + "')"
        End If
    End If
    MsgBox prompt:=msgstr, Title:="Document Styles Version Check"
    
End Sub
Sub installedVersion()
    
    Dim templateFile As String
    Dim strMacDocs As String
    Dim strTemplatePath As String
    Dim installedVersion As String
    
    templateFile = "RSuite_Word-template.dotm"  'the template file you are checking
    strTemplatePath = WT_Settings.StyleDir(FileType:="RS_wt")

    'Call VersionCheck(strTemplatePath, templateFile)
    installedVersion = GetVersion(strTemplatePath, templateFile)
    
    If installedVersion <> "none" Then
        MsgBox Title:="Installed Template Version", _
            prompt:="You currently have version v" & installedVersion & " of the Macmillan template & tools installed."
    Else
        MsgBox Title:="Warning", _
            prompt:="You do not have " & templateFile & " installed on your computer."
    End If


End Sub


Public Function GetVersion(ByVal fullPath As String, ByVal fileName As String)
    Dim installedVersion As String
    Dim fullFilePath As String
    
    fullFilePath = fullPath & Application.PathSeparator & fileName
    
    If IsItThere(fullFilePath) = False Then            ' the template file is not installed, or is not in the correct place
        installedVersion = "none"
    Else                                                                'the template file is installed in the correct place
        #If Mac Then
            Documents.Open fileName:=fullFilePath, ReadOnly:=True  ' Note can't set Visible:=False because it behaves inconsistently on Mac
        #Else
            Documents.Open fileName:=fullFilePath, visible:=False
        #End If
        installedVersion = Documents(fullFilePath).CustomDocumentProperties("version")
        Documents(fullFilePath).Close
    End If
    
    GetVersion = numericVersionStr(installedVersion)

End Function
Function numericVersionStr(vStr) As String
    If Left(vStr, 1) = "v" Then
        numericVersionStr = Right(vStr, Len(vStr) - 1)
    Else
        numericVersionStr = vStr
    End If
End Function

Function cleanVersionStr(ByVal versionStr As String) As String
    If Len(versionStr) - Len(Replace(versionStr, ".", "")) > 1 Then
        Dim splitArray() As String, versionStrArray(1) As String
        splitArray() = Split(versionStr, ".")
        versionStrArray(0) = splitArray(0)
        versionStrArray(1) = splitArray(1)
        versionStr = Join(versionStrArray, ".")
    End If
    cleanVersionStr = versionStr
End Function

Function versionCompare(ByVal versionA, ByVal versionB) As String
    Dim vA As Double, vB As Double
    Dim countA As Long, countB As Long
    ' get rid of last 'x' in 'x.x.x'
    versionA = cleanVersionStr(versionA)
    versionB = cleanVersionStr(versionB)
    ' try convert to double
    On Error GoTo returnErr
    vA = VBA.CDbl(versionA)
    vB = VBA.CDbl(versionB)
    On Error GoTo 0
    If vA > vB Then
        versionCompare = ">"
    ElseIf vA < vB Then
        versionCompare = "<"
    ElseIf vA = vB Then
        versionCompare = "same"
    Else
        versionCompare = "unable to compare"
    End If
    Exit Function
returnErr:
    versionCompare = "unable to compare"
End Function

Function customDocPropValue(cdpName As String) As String
    If Utils.DocPropExists(ActiveDocument, cdpName) Then
        customDocPropValue = numericVersionStr(ActiveDocument.CustomDocumentProperties(cdpName).value)
    Else
        customDocPropValue = ""
    End If
End Function


