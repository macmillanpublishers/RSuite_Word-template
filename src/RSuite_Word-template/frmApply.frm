VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmApply 
   Caption         =   "UserForm1"
   ClientHeight    =   1890
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   5130
   OleObjectBlob   =   "frmApply.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cbList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    HookListBoxScroll Me, Me.cbList
'End Sub
'
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' UnhookListBoxScroll
'End Sub

Private Sub cbCancel_Click()
    Unload frmApply
End Sub

Private Sub cbOK_Click()
    Select Case Me.Tag
        Case "section"
            Section sStyle:=Me.cbList.Column(0), sText:=Me.cbList.Column(1)
        Case "container"
            Container sStyle:=Me.cbList.Column(0), sText:=Me.cbList.Column(1)
        Case "break"
            Break sStyle:=Me.cbList.Column(0), sText:=Me.cbList.Column(1)
    End Select
    Unload frmApply
End Sub

Private Function ClearFormat()
    Selection.Expand unit:=wdParagraph
    Selection.ClearFormatting
    Selection.Collapse wdCollapseStart
End Function

Function Section(sStyle As String, sText As String)

    Selection.Paragraphs(1).Range.Select
    Selection.Collapse wdCollapseStart
    Selection.TypeText sText
    Selection.TypeParagraph
    Selection.MoveUp
    ClearFormat
    Selection.Style = sStyle
    
End Function

Function Container(sStyle As String, sText As String)

    If Selection.Characters.Count = 1 Then
    
        Selection.Paragraphs(1).Range.Select
        Selection.Collapse wdCollapseStart
        Selection.TypeText sText
        Selection.TypeParagraph
        Selection.MoveUp
        ClearFormat
        Selection.Style = sStyle
        Selection.MoveDown
        Selection.Expand unit:=wdParagraph
        Selection.Collapse wdCollapseStart
        Selection.TypeText "END " & sText
        Selection.TypeParagraph
        Selection.MoveUp
        ClearFormat
        Selection.Style = "END (END)"
        
    Else
    
        Dim paraCt As Integer
        Dim aRange As Range
        Dim sRange, eRange As Integer
    
        paraCt = Selection.Paragraphs.Count
        sRange = ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
        eRange = ActiveDocument.Range(0, Selection.Paragraphs(paraCt).Range.End).Paragraphs.Count
        
        Set aRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Paragraphs(sRange).Range.Start, _
            End:=ActiveDocument.Paragraphs(eRange).Range.End)
        aRange.Select
        Selection.Collapse wdCollapseStart
        Selection.TypeText sText
        Selection.TypeParagraph
        Selection.MoveUp
        ClearFormat
        Selection.Style = sStyle
        aRange.Select
        Selection.Collapse wdCollapseEnd
        If Clean_helpers.EndOfDocumentReached Then
            Selection.TypeParagraph
            Selection.TypeText "END " & sText
        Else
            Selection.TypeText "END " & sText
            Selection.TypeParagraph
            Selection.MoveUp
        End If
        ClearFormat
        Selection.Style = "END (END)"
        
    End If
    
End Function

Function Break(sStyle As String, sText As String)

    Selection.Paragraphs(1).Range.Select
    Selection.Collapse wdCollapseStart
    Selection.TypeText sText
    Selection.TypeParagraph
    Selection.MoveUp
    ClearFormat
    Selection.Style = sStyle
    
End Function
