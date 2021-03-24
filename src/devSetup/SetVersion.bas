Attribute VB_Name = "SetVersion"

Const versionFilename As String = "version.txt"
Const docPropNameStr As String = "Version"
' Const versionTxtDir ' currently same path as this file
Private versionTxtDir As String

Sub zz_updateVersionsForRepoTemplates_silent()
  Dim versionStr As String
  Dim versionTxtPath As String
  ' get version str
  versionTxtDir = ThisDocument.Path
  versionTxtPath = versionTxtDir & "\" & versionFilename
  versionStr = localReadTextFile(versionTxtPath)
  ' get doc paths
  Call config.defineVBAProjectParams
  ' set versions:
  Call setTemplateVersionNumber(rswt.dotmRepoPath, versionStr, True)
  Call setTemplateVersionNumber(st.installedPath, versionStr, True)
  Call setTemplateVersionNumber(stnc.installedPath, versionStr, True)
End Sub

Sub updateVersionsForRepoTemplates()
  Dim versionStr As String
  Dim versionTxtPath As String
  ' get version str
  versionTxtDir = ThisDocument.Path
  versionTxtPath = versionTxtDir & "\" & versionFilename
  versionStr = localReadTextFile(versionTxtPath)
  ' get doc paths
  Call config.defineVBAProjectParams
  ' set versions:
  Call setTemplateVersionNumber(rswt.dotmRepoPath, versionStr)
  Call setTemplateVersionNumber(st.installedPath, versionStr)
  Call setTemplateVersionNumber(stnc.installedPath, versionStr)
End Sub

Sub setTemplateVersionNumber(templatePath As String, versionStr As String, Optional silentBool As Boolean = False)
  Dim myDoc As Document
  
  ' Open template, set docprop, close document
  Set myDoc = Documents.Open(filename:=templatePath, Visible:=False)
  Call SetTemplateDocProp(myDoc, versionStr)
  myDoc.Saved = False   'Without this, just a change to custom docprops is not consistently saved!
  myDoc.Close SaveChanges:=True
  
  ' notify
  If silentBool = False Then
    MsgBox "'" & docPropNameStr & "' custom Doc Property set, to: '" & _
      versionStr & "', for file:" & vbCr & vbCr & templatePath
  End If
  
End Sub


Private Sub SetTemplateDocProp(myDoc As Document, dpValStr As String)

    If DocPropExists(myDoc, docPropNameStr) Then
        myDoc.CustomDocumentProperties(docPropNameStr).value = dpValStr
    Else
        myDoc.CustomDocumentProperties.Add Name:=docPropNameStr, LinkToContent:=False, _
            Type:=msoPropertyTypeString, value:=dpValStr
    End If

End Sub

'copied / co-opted from Utils.bas for local use
Public Function DocPropExists(ByRef objDoc As Document, ByVal PropName As String) As Boolean
  DocPropExists = False

' Note DocumentProperties returns a Collection
  Dim docProps As DocumentProperties
  Set docProps = objDoc.CustomDocumentProperties

  Dim varProp As Variant

  If docProps.Count > 0 Then
      For Each varProp In docProps
          If varProp.Name = PropName Then
              DocPropExists = True
              Exit Function
          End If
      Next varProp
  Else
      DocPropExists = False
  End If
End Function
' ===== ReadTextFile ==========================================================
'copied / co-opted from Utils.bas for local use
Public Function localReadTextFile(Path As String, Optional FirstLineOnly As Boolean _
  = True) As String

' load string from text file

    Dim fnum As Long
    Dim strTextWeWant As String
    
    fnum = FreeFile()
    Open Path For Input As fnum
    
    If FirstLineOnly = False Then
        strTextWeWant = Input$(LOF(fnum), fnum)
    Else
        Line Input #fnum, strTextWeWant
    End If
    
    Close fnum
    
    localReadTextFile = strTextWeWant

End Function


