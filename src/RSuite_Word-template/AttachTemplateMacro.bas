Attribute VB_Name = "AttachTemplateMacro"
Option Explicit
'created by Erica Warren - erica.warren@macmillan.com
' ======== PURPOSE =================
' Attaches assorted templates with custom Macmillan styles to the current document

' ======== DEPENDENCIES ============
' 1. Requires MacroHelpers module be installed in the same template
' 2. Requires the macmillan style templates be saved in the correct directories
'    that were nstalled on user's computer with Installer file or updated from Word-template.dotm

''' CHECK IT OUT

Sub zz_AttachStyleTemplate()
    Call AttachMe("RSuite.dotx")
End Sub

Sub zz_AttachBoundMSTemplate()
    Call AttachMe("RSuite_NoColor.dotx")
End Sub

Sub zz_AttachCoverTemplate()
    Call AttachMe("RSuite_CoverCopy.dotm")
End Sub

Sub AttachMe(TemplateName As String)
'Attaches a style template from the RSuiteStyleTemplate directory

' Get path to actual template
  Dim dictTemplateInfo As Dictionary
  Set dictTemplateInfo = SharedFileInstaller.FileInfo(TemplateName)

  Dim strTemplatePath As String
  strTemplatePath = dictTemplateInfo("Final")
  
  ' Can't attach template to another template, so
  If IsTemplate(ActiveDocument) = False Then
    'Check that file exists
    If IsItThere(strTemplatePath) = True Then
    
      'Apply template with Styles
      With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = strTemplatePath
      End With
      
      Dim templatePath, strVersionNumber As String
      templatePath = Left(strTemplatePath, (Len(strTemplatePath) - (Len(TemplateName) + 1)))
      strVersionNumber = VersionCheck.GetVersion(templatePath, TemplateName)
      SetStyleVersion VersionNumber:=strVersionNumber, templateNameStr:=TemplateName
      
    Else
      MsgBox "That style template doesn't seem to exist." & vbNewLine & vbNewLine & _
        "Install the Macmillan Style Template and try again, or contact workflows@macmillan.com for assistance.", _
        vbCritical, "Oh no!"
    End If
  End If

End Sub

Private Sub SetStyleVersion(ByRef VersionNumber As String, ByRef templateNameStr As String)
  Dim strProps(), Prop As Variant
  strProps = Array(Array("Version", VersionNumber), Array("TemplateName", templateNameStr))
  
  For Each Prop In strProps
    
    
    If Utils.DocPropExists(objDoc:=ActiveDocument, PropName:=Prop(0)) Then
        ActiveDocument.CustomDocumentProperties(Prop(0)).value = Prop(1)
    Else
        ActiveDocument.CustomDocumentProperties.Add Name:=Prop(0), LinkToContent:=False, _
            Type:=msoPropertyTypeString, value:=Prop(1)
    End If
  Next

End Sub

Private Sub listProps()
    Dim Prop As Variant

    For Each Prop In ActiveDocument.CustomDocumentProperties
        MsgBox Prop.Name & " " & Prop.value
    Next
    
End Sub

Private Function IsTemplate(ByVal objDoc As Document) As Boolean
  Select Case objDoc.saveFormat
    Case wdFormatTemplate, _
         wdFormatXMLTemplate, wdFormatXMLTemplateMacroEnabled, _
         wdFormatFlatXMLTemplate, wdFormatFlatXMLTemplateMacroEnabled
      IsTemplate = True
    Case Else
      IsTemplate = False
  End Select
End Function

