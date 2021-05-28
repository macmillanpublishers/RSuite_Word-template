Attribute VB_Name = "ApplyStyle"
Option Explicit
Option Base 1

Public ApplyType As String
Public MyTag As String

Sub ApplyParaStyle(myControl As IRibbonControl)
    Dim MyTag As String
    MyTag = myControl.Tag
    ApplyParaSty (MyTag)
End Sub
Sub ApplyParaStyleB(MyTag As String)
    ApplyParaSty (MyTag)
End Sub
Sub ApplyCharStyle(myControl As IRibbonControl)
    Dim MyTag As String
    MyTag = myControl.Tag
    ApplyCharSty (MyTag)
End Sub
Sub ApplyCharStyleB(MyTag As String)
    ApplyCharSty (MyTag)
End Sub
Sub ApplyNumberList(myControl As IRibbonControl)
    Dim MyTag As String
    MyTag = myControl.Tag
    ApplyNumList (MyTag)
End Sub

Function ApplyParaSty(myStyle As String)
    On Error GoTo ErrorHandler

    Selection.Expand Unit:=wdParagraph
    Selection.Style = myStyle
    Selection.Collapse Direction:=wdCollapseEnd
    ActiveWindow.SmallScroll Up:=3, down:=2
    Selection.GoTo What:=wdGoToBookmark, Name:="\Sel"
    Application.ScreenRefresh
    
    Exit Function
    
ErrorHandler:

     If Err.Number = 5834 Then
        Clean_helpers.MessageBox _
            buttonType:=vbOKOnly, _
            Title:="STYLE NOT FOUND", _
            Msg:="Style not found." & vbNewLine & vbNewLine & "Make sure you have the Macmillan Template attached to the document."
        Exit Function
     Else
        Clean_helpers.MessageBox _
                buttonType:=vbOKOnly, _
                Title:="UNEXPECTED ERROR", _
                Msg:="Error number " & Err.Number & " occurred." & vbCr & "Description: " & Err.Description
     End If

End Function

Function ApplyCharSty(myStyle As String)

    On Error GoTo ErrorHandler
    
    If Selection.Style Is Nothing Then
        Selection.Style = myStyle
    ElseIf Selection.Style = myStyle Then
        Selection.ClearFormatting
    Else
        Selection.Style = myStyle
    End If
    Application.ScreenRefresh
    
    Exit Function
    
ErrorHandler:

     If Err.Number = 5834 Then
        Clean_helpers.MessageBox _
            buttonType:=vbOKOnly, _
            Title:="STYLE NOT FOUND", _
            Msg:="Style not found." & vbNewLine & vbNewLine & "Make sure you have the Macmillan Template attached to the document."
        Exit Function
     Else
        Clean_helpers.MessageBox _
                buttonType:=vbOKOnly, _
                Title:="UNEXPECTED ERROR", _
                Msg:="Error number " & Err.Number & " occurred." & vbCr & "Description: " & Err.Description
     End If
     
End Function

Function ApplyNumList(myStyle As String)

On Error GoTo ErrorHandler
    
    Dim StartAt1 As Boolean
    Application.DisplayAlerts = wdAlertsNone
    
    Selection.Expand Unit:=wdParagraph
    ActiveDocument.Bookmarks.Add ("templine")
    Selection.Collapse Direction:=wdCollapseStart
    Selection.Move Unit:=wdParagraph, Count:=-1
    If Selection.Style <> myStyle Then StartAt1 = True
    ActiveDocument.Bookmarks("templine").Select
    Selection.Style = myStyle
    If StartAt1 = True Then Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(1), _
                ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList
    Selection.Collapse Direction:=wdCollapseEnd
    Application.ScreenRefresh
    
        With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = "nl"
    End With
    ActiveDocument.styles("nl").LinkToListTemplate ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(1), ListLevelNumber:=1
    
    Application.DisplayAlerts = wdAlertsAll
    
        Exit Function
    
ErrorHandler:

     If Err.Number = 5834 Then
        Clean_helpers.MessageBox _
            buttonType:=vbOKOnly, _
            Title:="STYLE NOT FOUND", _
            Msg:="Style not found." & vbNewLine & vbNewLine & "Make sure you have the Macmillan Template attached to the document."
        Exit Function
     Else
        Clean_helpers.MessageBox _
                buttonType:=vbOKOnly, _
                Title:="UNEXPECTED ERROR", _
                Msg:="Error number " & Err.Number & " occurred." & vbCr & "Description: " & Err.Description
     End If

End Function

Sub testlist()

    Dim myList
    myList = getList("sections")
    
    MsgBox "done"

End Sub

Private Function getList(listName As String)

    Dim FileNum As Integer
    Dim DataLine As String
    Dim StylePath As String
    
    Dim all() As Variant
    Dim i As Integer
    i = 1
    
    StylePath = WT_Settings.StyleDir(FileType:="styles") & Application.PathSeparator & listName & ".txt"
    
    If IsItThere(StylePath) = True Then
        FileNum = FreeFile()
        Open StylePath For Input As #FileNum
        
        While Not EOF(FileNum)
            Line Input #FileNum, DataLine
            Dim result() As String
            result() = Split(DataLine, ",")
            If Right(result(0), 1) = vbLf Or Right(result(0), 1) = vbCr Then result(0) = Left(result(0), Len(result(0)) - 1)
            If Right(result(1), 1) = vbLf Or Right(result(1), 1) = vbCr Then result(1) = Left(result(1), Len(result(1)) - 1)
            result(0) = CleanString(RTrim(Trim(result(0))))
            result(1) = CleanString(RTrim(Trim(result(1))))
            ReDim Preserve all(i)
            all(i) = result
            i = i + 1
        Wend
        
        Close FileNum
        
    Else
        MessageBox Title:="Style List Not Found", Msg:="Cannot locate the RSuite Styles file."
    End If
    
    getList = all
    
End Function






Sub InsertSection(control As IRibbonControl)
    Dim myList As Variant
    Dim i As Integer
    
    Call Clean_helpers.CheckTemplate
    
'    MyList = Array(Array("Section-About-Author (SAA)", "About the Author"), Array("Section-Acknowledgments (SAK)", "Acknowledgments"), Array("Section-Ad-Card (SAC)", "Ad Card"), _
'        Array("Section-Afterword (SAW)", "Afterword"), Array("Section-Appendix (SAP)", "Appendix"), Array("Section-Back-Ad (SBA)", "Back Ad"), Array("Section-Back-Matter-General (SBM)", _
'        "Back Matter General"), Array("Section-Bibliography (SBI)", "Bibliography"), Array("Section-Book (BOOK)", "Book"), Array("Section-Chapter (SCP)", "Chapter"), Array("Section-Conclusion (SCL)", _
'        "Conclusion"), Array("Section-Contents (STC)", "Contents"), Array("Section-Copyright (SCR)", "Copyright"), Array("Section-Dedication (SDE)", "Dedication"), Array("Section-Ebook-Copyright (SECR)", "Ebook Copyright"), _
'        Array("Section-Epigraph (SEP)", "Epigraph"), Array("Section-Excerpt-Chapter (SEC)", "Excerpt Chapter"), Array("Section-Excerpt-Opener (SEO)", "Excerpt Opener"), Array("Section-Foreword (SFW)", "Foreword"), Array("Section-Front-Matter-General (SFM)", _
'        "Front Matter General"), Array("Section-Front-Sales (SFS)", "Front Sales"), Array("Section-Glossary (SGL)", "Glossary"), Array("Section-Halftitle (SHT)", "Halftitle Page"), _
'        Array("Section-Index (SIN)", "Index"), Array("Section-Interlude (SIN)", "Interlude"), Array("Section-Introduction (SIC)", "Introduction"), Array("Section-Notes (SNT)", "Notes"), _
'        Array("Section-Part (SPT)", "Part"), Array("Section-Preface (SPF)", "Preface"), Array("Section-Recipe (SREC)", "Recipe"), Array("Section-Series-Page (SSP)", "Series Page"), Array("Section-Titlepage (STI)", "Titlepage"))
    myList = getList("sections")
    For i = 1 To UBound(myList)
        With frmApply.cbList
            .AddItem myList(i)(0)
            .List(.ListCount - 1, 1) = myList(i)(1)
        End With
    Next
    
    frmApply.cbList.Text = "Select section type..."
    frmApply.Caption = "Insert Section"
    frmApply.Tag = "section"
    frmApply.Show
End Sub

Sub InsertContainer(control As IRibbonControl)
    InsertContainerMacro
End Sub

Sub InsertContainerMacro()

    On Error GoTo ErrorHandler

    Dim myList As Variant
    Dim i As Integer
    
    Call Clean_helpers.CheckTemplate
    
    'myList = Array(Array("EXTRACT-A (EXT-A)", "EXTRACT-A"), Array("EXTRACT-B (EXT-B)", "EXTRACT-B"), Array("EXTRACT-C (EXT-C)", "EXTRACT-C"), Array("EXTRACT-D (EXT-D)", "EXTRACT-D"), Array("VERSE-A (VRS-A)", "VERSE-A"), Array("VERSE-B (VRS-B)", "VERSE-B"), Array("BOX-A (BOX-A)", "BOX-A"), Array("BOX-B (BOX-B)", "BOX-B"), Array("LETTER-A (LTR-A)", "LETTER-A"), Array("LETTER-B (LTR-B)", "LETTER-B"), Array("LETTER-C (LTR-C)", "LETTER-C"), Array("LETTER-D (LTR-D)", "LETTER-D"), Array("COMPUTER-A (COM-A)", "COMPUTER-A"), Array("COMPUTER-B (COM-B)", "COMPUTER-B"), Array("SIDEBAR-A (SIDE-A)", "SIDEBAR-A"), Array("SIDEBAR-B (SIDE-B)", "SIDEBAR-B"), Array("PULL-QUOTE (PQ)", "PULL-QUOTE"), Array("IMAGE (IMG)", "IMAGE"), Array("TABLE (TBL)", "TABLE"), Array("RECIPE (REC)", "RECIPE"))
    myList = getList("containers")
    For i = 1 To UBound(myList)
        With frmApply.cbList
            .AddItem myList(i)(0)
            .List(.ListCount - 1, 1) = myList(i)(1)
        End With
    Next
    
    frmApply.cbList.Text = "Select container type..."
    frmApply.Caption = "Insert Container"
    frmApply.Tag = "container"
    frmApply.Show
    
    Exit Sub
    
ErrorHandler:

    Clean_helpers.MessageBox _
                buttonType:=vbOKOnly, _
                Title:="UNEXPECTED ERROR", _
                Msg:="Error number " & Err.Number & " occurred." & vbCr & "Description: " & Err.Description
End Sub

Sub InsertBreak(control As IRibbonControl)

    On Error GoTo ErrorHandler
    
    Dim myList() As Variant
    Dim i As Integer
    
    Call Clean_helpers.CheckTemplate
    
    'myList = Array(Array("Blank-Space-Break (Bsbrk)", "[blank]"), Array("Ornamental-Space-Break (Osbrk)", "***"), Array("Separator (Sep)", "[separator]"))
    myList = getList("breaks")
    For i = 1 To UBound(myList)
        With frmApply.cbList
            .AddItem myList(i)(0)
            .List(.ListCount - 1, 1) = myList(i)(1)
        End With
    Next
    
    frmApply.cbList.Text = "Select break type..."
    frmApply.Caption = "Insert Break"
    frmApply.Tag = "break"
    frmApply.Show
    
    Exit Sub
    
ErrorHandler:

    Clean_helpers.MessageBox _
                buttonType:=vbOKOnly, _
                Title:="UNEXPECTED ERROR", _
                Msg:="Error number " & Err.Number & " occurred." & vbCr & "Description: " & Err.Description
    
End Sub
