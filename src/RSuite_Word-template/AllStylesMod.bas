Attribute VB_Name = "AllStylesMod"
Option Explicit
Public FullList

Sub GetVisible(control As IRibbonControl, ByRef visible)
    
    visible = False
    #If Mac Then
        If control.Tag = "MacStylesGroup" Then visible = True
    #Else
        If control.Tag = "PCStylesGroup" Then visible = True
    #End If

End Sub

Public Function OnGetItemLabel(ctl As IRibbonControl, index As Integer, ByRef Label)
    Label = FullList(index)
End Function

Public Function OnGetItemCount(ctl As IRibbonControl, ByRef Count)
    Call getAllStyles
    Count = UBound(FullList)
End Function


Public Function ApplyTheStyle(ctl As IRibbonControl, Text As String)

    On Error GoTo ErrorTrap

    If ActiveDocument.Styles(Text).Type = 1 And ActiveDocument.Styles(Text).Type = 2 Then
        ApplyParaStyFromCombo (Text)
    ElseIf ActiveDocument.Styles(Text).Type = 1 Then
        ApplyParaStyFromCombo (Text)
    ElseIf ActiveDocument.Styles(Text).Type = 2 Then
        ApplyCharStyFromCombo (Text)
    Else
        ApplyParaStyFromCombo (Text)
    End If
    
    Selection.Collapse direction:=wdCollapseEnd
    myRibbon.InvalidateControl ("cboApplyStyles")
    
    Exit Function
    
ErrorTrap:
    If Err.Number = 5941 Then
        MsgBox "That style does not exist in the template. Please check spelling and try again."
    Else
        MsgBox "Error " + Err.Number + ": " + Err.Description + " Exiting routine."
    End If
End Function

Public Function GetCurrentStyle(ctl As IRibbonControl, ByRef Text)
    On Error GoTo ErrorHandler

    Text = Selection.Style
    Exit Function
    
ErrorHandler:
    Text = "Multiple Styles"
End Function

Sub ApplyParaStyFromCombo(Text)
    On Error GoTo ErrorHandler
    
    Dim MyTag As String
    MyTag = Text
    ApplyStyle.ApplyParaStyleB (MyTag)
    Exit Sub
    
ErrorHandler:
    Selection.Style = Text
End Sub

Sub ApplyCharStyFromCombo(Text)
    On Error GoTo ErrorHandler
    
    Dim MyTag As String
    MyTag = Text
    ApplyStyle.ApplyCharStyleB (MyTag)
    Exit Sub
    
ErrorHandler:
    Selection.Style = Text
End Sub

Private Function getAllStylesXXX()

    Dim sty As Style
    Dim allStyles() As Variant
    Dim i As Integer
    i = 0
    
    For Each sty In ActiveDocument.Styles
        If Not sty.BuiltIn Then
            ReDim Preserve allStyles(i)
            allStyles(i) = sty.NameLocal
            i = i + 1
        End If
    Next
    
    If IsEmpty(allStyles) = False Then
        ReDim Preserve allStyles(i)
        allStyles(i) = "Normal"
    End If
    
    FullList = allStyles
    
End Function

Public Function getAllStyles()

    Dim FileNum As Integer
    Dim DataLine As String
    Dim StylePath As String
    
    Dim allStyles() As Variant
    Dim i As Integer
    i = 0
    
    StylePath = WT_Settings.StyleDir(FileType:="styles") & Application.PathSeparator & "RSuite_styles.txt"
    
    If IsItThere(StylePath) = True Then
        FileNum = FreeFile()
        Open StylePath For Input As #FileNum
        
        While Not EOF(FileNum)
            Line Input #FileNum, DataLine
            ReDim Preserve allStyles(i)
            allStyles(i) = DataLine
            i = i + 1
        Wend
        
        Close FileNum
        
    Else
        MessageBox Title:="Style List Not Found", Msg:="Cannot locate the RSuite Styles file."
    End If
    
    FullList = allStyles
    
    getAllStyles = FullList
    
End Function
