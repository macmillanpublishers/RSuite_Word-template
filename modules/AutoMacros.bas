Attribute VB_Name = "AutoMacros"
Public Sub AutoNew()
    Call template_switcher
End Sub

Public Sub AutoOpen()
    Call template_switcher
End Sub

Private Function template_switcher()

    If isTemplate(ActiveDocument) Then Exit Function

    If DocPropExists(objDoc:=ActiveDocument, PropName:="Template") Then
     Dim this_val As String
     this_val = ActiveDocument.CustomDocumentProperties("Template").Value
    Else
     this_val = "0"
    End If
     
    If this_val = "1" Or whichTemplate = 1 Then
        setAddins Disable:="Word-template.dotm", Enable:="RSuite_Word-template.dotm"
        If this_val = "0" Then setProperty ("1")
        SaveProfileSetting myToken:="LastTemplate", mySetting:="1"
    ElseIf this_val = "2" Or whichTemplate = 2 Then
        setAddins Disable:="RSuite_Word-template.dotm", Enable:="Word-template.dotm"
        If this_val = "0" Then setProperty ("2")
        SaveProfileSetting myToken:="LastTemplate", mySetting:="2"
    Else:
       MsgBox "No style template attached. Would you like to enable a Macmillan template?"
    End If

End Function

Private Function DocPropExists(ByRef objDoc As Document, ByVal PropName As String) As Boolean
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

Private Function isTemplate(ByVal doc As Document) As Boolean
  Select Case doc.SaveFormat
    Case wdFormatTemplate, wdFormatDocument97, _
         wdFormatXMLTemplate, wdFormatXMLTemplateMacroEnabled, _
         wdFormatFlatXMLTemplate, wdFormatFlatXMLTemplateMacroEnabled
      isTemplate = True
    Case Else
      isTemplate = False
  End Select
End Function

Private Function whichTemplate() As Integer

    If ActiveDocument.AttachedTemplate Like "RSuite*" Then
        whichTemplate = 2
    ElseIf ActiveDocument.AttachedTemplate Like "macmillan*" Then
        whichTemplate = 1
    Else
     whichTemplate = 0
    End If
    
End Function

Private Function setAddins(ByRef Disable As String, ByRef Enable As String)

 On Error Resume Next
    AddIns(TemplateName).Installed = False
    AddIns(Enable).Installed = True
 On Error GoTo 0

End Function

Private Function setProperty(ByRef VersionNumber As String)
    ActiveDocument.CustomDocumentProperties.Add Name:="Template", LinkToContent:=False, _
        Type:=msoPropertyTypeString, Value:=VersionNumber
End Function

Private Function SaveProfileSetting(ByVal myToken, ByVal mySetting)
    If System.OperatingSystem = "Macintosh" Then
         Call SaveSetting("Word", "Macmillan", myToken, mySetting)
    Else
         System.ProfileString("Macmillan", myToken) = mySetting
    End If
End Function

Private Function GetProfileSetting(ByVal myToken) As String
    On Error GoTo ErrorHandler
    
    If System.OperatingSystem = "Macintosh" Then
        GetProfileSetting = GetSetting("Word", "Macmillan", myToken, "false")
    Else
        GetProfileSetting = System.ProfileString("Macmillan", myToken)
    End If
    
    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        Case 5843
            System.ProfileString("Word", myToken) = ""
            GetProfileSetting = ""
    End Select
    
End Function



