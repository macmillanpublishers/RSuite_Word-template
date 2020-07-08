Attribute VB_Name = "Switch"
Option Explicit
Public switchRibbon As IRibbonUI

Public Function Onload(ribbon As IRibbonUI)
  'Creates a ribbon instance for use in this project
  Set switchRibbon = ribbon
End Function

Public Function EnableTemplate(r As IRibbonControl)
    Dim myVariables As Variant
    
    If r.Tag = "RSuite" Then
        myVariables = SetAddins("2")
    ElseIf r.Tag = "Macmillan" Then
        myVariables = SetAddins("1")
    End If
    AttachTemplate (myVariables)
    
End Function

Sub test_enable()
Dim myVariables As Variant
 myVariables = SetAddins("2")
 AttachTemplate (myVariables)
End Sub

Public Function Template_Inspector(control As IRibbonControl)
    Template_Inspector_Macro
End Function

Private Function Template_Inspector_Macro()
    Dim AttTemp, AddIn, tempName As String
    AttTemp = getAttachedTemplate
    AddIn = getAddIn
    
    tempName = "RSuite"
    If AttTemp = 1 Then tempName = "Macmillan"
    
    If AttTemp = 0 Then
        #If Mac Then
            MsgBox Prompt:="This document does not have RSuite or Macmillan styles attached." & vbNewLine & vbNewLine & _
                "Please click on the button in the Home tab to activate the RSuite template and tools." & vbNewLine & vbNewLine & _
                "If you wish to attach the Macmillan template, please quit and reopen the file in Word 2011.", _
                Title:="NO TEMPLATE ATTACHED"
        #Else
            MsgBox Prompt:="This document does not have RSuite or Macmillan styles attached." & vbNewLine & vbNewLine & _
                "Please click on the button in the Home tab to activate the desired template and tools.", _
                Title:="NO TEMPLATE ATTACHED"
        #End If
    ElseIf AddIn = 3 Then
        MsgBox Prompt:="WARNING: You have both Macmillan and RSuite tools enabled." & vbNewLine & vbNewLine & _
                "This document has the " & tempName & "template attached." & vbNewLine & vbNewLine & _
                "Please click on the " & tempName & " button in the Home tab to load ONLY the " & tempName & " tools to your Ribbon.", _
                Title:="TWO TOOLBARS LOADED"
    
    ElseIf AttTemp = 2 And AddIn = 1 Then
            MsgBox Prompt:="WARNING: This document has the RSuite style-template attached, but you have the Macmillan tools loaded!" & vbNewLine & vbNewLine & _
                "Please click on the RSuite button in the Home tab to load the RSuite tools to your Ribbon.", _
                Title:="WRONG TEMPLATE ATTACHED"

    ElseIf AttTemp = 1 And AddIn = 2 Then
        #If Mac Then
            MsgBox Prompt:="WARNING: This document has the Macmillan style-template attached!" & vbNewLine & vbNewLine & _
                "If you wish to style using the Macmillan template, please quit and reopen the file in Word 2011." & vbNewLine & vbNewLine & _
                "If you wish to apply the RSuite template and load the RSuite tools, please click on the RSuite button in the Home tab.", _
                Title:="WRONG TEMPLATE ATTACHED"
        #Else
            MsgBox Prompt:="WARNING: This document has the Macmillan style-template attached, but you have the RSuite tools loaded!" & vbNewLine & vbNewLine & _
                "Please click on the Macmillan button in the Home tab to load the Macmillan tools to your Ribbon.", _
                Title:="WRONG TEMPLATE ATTACHED"
        #End If
    ElseIf AttTemp = 3 Then
        MsgBox Prompt:="WARNING: This appears to be an RSuite document but the RSuite style-template is not attached!" & vbNewLine & vbNewLine & _
                "Please click on the RSuite button in the Home tab to load the RSuite tools to your Ribbon.", _
                Title:="NO TEMPLATE ATTACHED"
    ElseIf AttTemp = 4 Then
        #If Mac Then
                MsgBox Prompt:="WARNING: This document has the Macmillan style-template attached!" & vbNewLine & vbNewLine & _
                "If you wish to style using the Macmillan template, please quit and reopen the file in Word 2011." & vbNewLine & vbNewLine & _
                "If you wish to apply the RSuite template and load the RSuite tools, please click on the RSuite button in the Home tab.", _
                Title:="WRONG TEMPLATE ATTACHED"
        #Else
            MsgBox Prompt:="WARNING: This appears to be a Macmillan document but the Macmillan style-template is not attached!" & vbNewLine & vbNewLine & _
                "Please click on the Macmillan button in the Home tab to load the Macmillan tools to your Ribbon.", _
                Title:="NO TEMPLATE ATTACHED"
        #End If
    Else
        'If the right Ribbon/toolset is loaded for the attached template (switch the names to Macmillan as needed):
        MsgBox Prompt:="This document has the " & tempName & " style-template attached and you have the " & tempName & " tools loaded in the Ribbon." & vbNewLine & vbNewLine & _
            "Go ahead, get styling!", _
            Title:="ALL GOOD"
    End If
    
End Function

Private Function AttachTemplate(ByVal myVariables As Variant)
    
    Dim templatePath, strVersionNumber As String
    templatePath = WT_Settings.StyleDir(myVariables(1), , "styles") & Application.PathSeparator & myVariables(0)
    
    If IsItThere(Path:=templatePath) Then
        With ActiveDocument
            .UpdateStylesOnOpen = True
            .attachedTemplate = templatePath
        End With
    End If
    
    If myVariables(1) = "3" Then
        strVersionNumber = "0"
    Else
        strVersionNumber = Utils.GetVersion(templatePath)
    End If
    
    Utils.SetStyleVersion VersionNumber:=strVersionNumber, templateNo:=myVariables(1)

End Function

Public Function template_switcher()

    Dim Attached_Template
    Dim lastTemplate, this_val As String
    Dim stored_template_value As String
    Dim lastTemplateName, enabledTemplateName, disabledTemplateName, enabledTemplateNumber, disabledTemplateNumber, templatePath, t1, t2, t3 As String
    lastTemplate = "2"
    lastTemplate = GetProfileSetting("LastTemplate")

    t1 = "Macmillan Word template"
    t2 = "RSuite Word Template"
    
    If lastTemplate = "1" Then
        enabledTemplateName = t1
        enabledTemplateNumber = "1"
        disabledTemplateName = t2
        disabledTemplateNumber = "2"
    ElseIf lastTemplate = "2" Then
        enabledTemplateName = t2
        enabledTemplateNumber = "2"
        disabledTemplateName = t1
        disabledTemplateNumber = "1"
    End If
    
    If isTemplate(ActiveDocument) Then Exit Function

    If DocPropExists(objDoc:=ActiveDocument, PropName:="Template") Then
     Attached_Template = ActiveDocument.CustomDocumentProperties("Template").Value
    Else
     this_val = "0"
    End If
            
    Dim AttTemp, AddIn As String
    AttTemp = getAttachedTemplate
    AddIn = getAddIn
            
    If stored_template_value = "1" Or AttTemp = 1 Then
        SetAddins Enable:="1"
    ElseIf stored_template_value = "2" Or AttTemp = 2 Then
        SetAddins Enable:="2"
    Else:
        Dim Msg, Style, Title, Response
        Msg = "No style template attached." & vbNewLine & vbNewLine & _
            "The currently enabled style toolbar is " & enabledTemplateName & vbNewLine & vbNewLine & _
            "Would you like to enable " & disabledTemplateName & " instead?"
        Style = vbYesNo
        Title = "Select a Template"
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            SetAddins Enable:=disabledTemplateNumber
        Else
            SetAddins Enable:=enabledTemplateNumber
        End If
    End If

End Function

Private Function SetAddins(ByVal Enable As String) As String()

    On Error GoTo Handler

    Dim returnVal(2) As String
    Dim t1, t2, e, d, curr, templatePath, setting, attachedTemplate As String
    Dim inst, isGT As Boolean
    isGT = False

    t1 = WT_Settings.StyleDir(TemplateNumber:="1") & Application.PathSeparator & "Word-template.dotm"
    t2 = WT_Settings.StyleDir(TemplateNumber:="2") & Application.PathSeparator & "RSuite_Word-template.dotm"
    setting = Enable
    
    #If Mac Then
'        If Dir(t1, MacID("TEXT")) = "" Then
'            isGT = True
'            t1 = WT_Settings.StyleDir(TemplateNumber:="1") & Application.PathSeparator & "MacmillanGT.dotm"
'        End If
    #Else
        If Dir(t1) = "" Then
            isGT = True
            t1 = WT_Settings.StyleDir(TemplateNumber:="1") & Application.PathSeparator & "MacmillanGT.dotm"
        End If
    #End If

    If Enable = "1" Then
        e = t1
        d = t2
        If isGT Then
            returnVal(0) = "macmillan.dotm"
            returnVal(1) = "3"
        Else
            returnVal(0) = "macmillan.dotx"
            returnVal(1) = "1"
        End If
    Else
        returnVal(0) = "RSuite.dotx"
        returnVal(1) = "2"
        e = t2
        d = t1
    End If
    
    curr = e
    inst = True
    AddIns(e).Installed = inst
    
    #If Mac Then
    #Else
        curr = d
        inst = False
        AddIns(d).Installed = inst
    #End If
    
    SetAddins = returnVal
    
    Exit Function
    
Handler:

    Dim s_err
    s_err = Err.Number
    
    If s_err = "5941" Then
        Dim mydir
        mydir = Dir(curr)
        If mydir <> "" Then
            AddIns.Add FileName:=curr, Install:=inst
            Resume
        Else
            MsgBox "Cannot file the file " & curr & "."
            Exit Function
        End If
'    ElseIf s_err = "53" Then
'        Resume Next
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

    
End Function

Private Function GetDocProp(ByVal PropName)

   GetDocProp = ActiveDocument.CustomDocumentProperties(PropName)

End Function

Private Function isTemplate(ByVal Doc As Document) As Boolean
  Select Case Doc.SaveFormat
    Case wdFormatTemplate, wdFormatDocument97, _
         wdFormatXMLTemplate, wdFormatXMLTemplateMacroEnabled, _
         wdFormatFlatXMLTemplate, wdFormatFlatXMLTemplateMacroEnabled
      isTemplate = True
    Case Else
      isTemplate = False
  End Select
End Function

Private Function getAddIn() As Integer

    Dim a As AddIn
    Dim r, m As Boolean

    For Each a In Application.AddIns
        If a Like "RSuite*" And a.Installed = True Then
            r = True
        ElseIf (a Like "Macmillan*" Or a Like "Word-template*") And a.Installed = True Then
            m = True
        End If
    Next
    
    If r And m Then
        getAddIn = 3
    ElseIf r Then
        getAddIn = 2
    ElseIf m Then
        getAddIn = 1
    Else
        getAddIn = 0
    End If
    
End Function

Private Function StyleTest(ByVal strStyleName As String) As Boolean

    Dim sCheck As String
    Dim thisDoc As Document

    StyleTest = False
    Set thisDoc = ActiveDocument

    sCheck = "x" & strStyleName ' makes the two strings different
    On Error Resume Next
        sCheck = thisDoc.Styles(strStyleName).NameLocal
    On Error GoTo 0

    If sCheck = strStyleName Then
        If thisDoc.Styles(strStyleName).InUse Then
            StyleTest = True
        End If
    End If

End Function

Private Function getAttachedTemplate() As Integer

    Dim this_version As String
    Dim is_rs As Boolean
    Dim is_mc As Boolean
    is_rs = False
    is_mc = False
    
    If DocPropExists(ActiveDocument, "Template") Then
         this_version = GetDocProp("Template")
         If this_version = "2" Then is_rs = True
         If this_version = "1" Or "3" Then is_mc = True
    Else
        If StyleTest("Text - Standard (tx)") Then is_mc = True
    End If
    
    If ActiveDocument.attachedTemplate Like "RSuite*" Then
        getAttachedTemplate = 2
    ElseIf ActiveDocument.attachedTemplate Like "macmillan*" Then
        getAttachedTemplate = 1
    Else
        If is_rs Then
            getAttachedTemplate = 3
        ElseIf is_mc Then
            getAttachedTemplate = 4
        Else
            getAttachedTemplate = 0
        End If
    End If
    
End Function

Private Function setProperty(ByRef VersionNumber As String)
    ActiveDocument.CustomDocumentProperties.Add Name:="Template", LinkToContent:=False, _
        Type:=msoPropertyTypeString, Value:=VersionNumber
End Function

Public Function SaveProfileSetting(ByVal myToken As String, ByVal mySetting As String)
    If System.OperatingSystem = "Macintosh" Then
         Call SaveSetting("Word", "Macmillan", myToken, mySetting)
    Else
         System.ProfileString("Macmillan", myToken) = mySetting
    End If
End Function

Public Function GetProfileSetting(ByVal myToken) As String
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
