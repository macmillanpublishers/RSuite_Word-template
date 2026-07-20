Attribute VB_Name = "IndexStyler"
Option Explicit

' These must match the style names in RSuite_Word-template.dotm exactly.
Private Const STYLE_BOOK As String = "Section-Book (BOOK)"
Private Const STYLE_SIN  As String = "Section-Index (SIN)"
Private Const STYLE_IDX1 As String = "Index-Entry (Idx1)"
Private Const STYLE_IDX2 As String = "Index-Sub-Entry (Idx2)"
Private Const STYLE_IDX3 As String = "Index-Sub-Sub-Entry (Idx3)"
Private Const STYLE_SEP  As String = "Separator (Sep)"

Private Const SEP_TEXT   As String = "Separator (Sep)"

Private Const SIN_TEXT   As String = "Index"

Public Sub LaunchIndexStyler()

    Call Clean_helpers.CheckTemplate

    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
      If WT_Settings.InstallType = "user" Then
        If MacroHelpers.StartupSettings(AcceptAll:=False) = True Then
          Call MacroHelpers.Cleanup
          Exit Sub
        End If
      Else
        If MacroHelpers.StartupSettings(AcceptAll:=True) = True Then
          Call MacroHelpers.Cleanup
          Exit Sub
        End If
      End If

    ApplyIndexStyles

End Sub

Public Sub ApplyIndexStyles()

    Dim doc As Document
    Set doc = ActiveDocument

    '  Confirm before making changes
    Dim msg As String
    msg = "Apply RSuite index styles to:" & vbCrLf & vbCrLf & _
          "    " & doc.Name & vbCrLf & vbCrLf & _
          "This will replace all existing paragraph styles and remove " & _
          "direct formatting overrides." & vbCrLf & vbCrLf & _
          "Proceed?"

    If MsgBox(msg, vbYesNo + vbQuestion, "RSuite Index Styler") <> vbYes Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim t0 As Single
    t0 = Timer

    On Error GoTo ErrHandler

    
    Debug.Print "stripping headers"
    '  Step 2: Strip page headers
    Call StripHeaders(doc)
    
    ' Clear blank end of Doc paras or they will get Separators
    Call trimDocEnd(doc)
    
    Debug.Print "applying para styles"
    '  Step 3: Apply paragraph styles (Idx1 / Idx2 / Idx3 / Sep)
    Dim nIdx1 As Long, nIdx2 As Long, nIdx3 As Long, nSep As Long
    Call ApplyParagraphStyles(doc, nIdx1, nIdx2, nIdx3, nSep)
    
    Debug.Print "applying char stylee"
    '  Step 4: Apply character styles to formatted runs
    Call Clean.CheckAppliedCharStyles(1)
    Call Clean.FixAppliedCharStyles(1)
    Call Clean.CheckSpecialCharactersPC(1)
    Call Clean.LocalFormatting(1)

    '  Step 6: Insert Section-Book and Section-Index at document start
    Call InsertBookAndSin(doc)

    ActiveDocument.UndoClear
    
    Application.ScreenUpdating = True

    '  Report results
    MsgBox "RSuite styles applied successfully!" & vbCrLf & _
           "(" & Format(Timer - t0, "0.0") & " seconds)" & vbCrLf & vbCrLf & _
           "Idx1 entries:         " & nIdx1 & vbCrLf & _
           "Idx2 sub-entries:     " & nIdx2 & vbCrLf & _
           "Idx3 sub-sub-entries:  " & nIdx3 & vbCrLf & _
           "Separators:           " & nSep & vbCrLf

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "If the error mentions a style name, confirm that the RSuite " & _
           "template is attached and its styles match the constants at the " & _
           "top of RSuiteIndexStyler.bas.", _
           vbCritical, "RSuite Index Styler"

End Sub

Private Sub StripHeaders(doc As Document)
    Dim sect As section
    Dim hdr  As HeaderFooter

    For Each sect In doc.sections
        For Each hdr In sect.Headers
            hdr.LinkToPrevious = False
            hdr.Range.Delete
        Next hdr
    Next sect
End Sub

' =============================================================================
'  STEP 3 — Apply paragraph styles based on leading-tab depth
'
'   0 tabs  → Index-Entry (Idx1)
'   1 tab   → Index-Sub-Entry (Idx2)
'   2+ tabs → Index-Sub-Sub-Entry (Idx3)
'   blank   → Separator (Sep)  with text "Separator (Sep)"
' =============================================================================

Private Sub ApplyParagraphStyles(doc As Document, _
                                  ByRef nIdx1 As Long, _
                                  ByRef nIdx2 As Long, _
                                  ByRef nIdx3 As Long, _
                                  ByRef nSep As Long)
    nIdx1 = 0: nIdx2 = 0: nIdx3 = 0: nSep = 0

    Dim para As Paragraph
    Dim txt  As String
    Dim tabs As Integer
    Dim i    As Integer

    For Each para In doc.Paragraphs

        ' Paragraph text without the trailing paragraph-mark character
        txt = para.Range.Text
        If Len(txt) > 0 Then txt = Left(txt, Len(txt) - 1)

        If Trim(txt) = "" Then
            '  Blank / whitespace-only → Separator
            If (para.style <> STYLE_SEP) Then
                para.style = doc.styles(STYLE_SEP)
                nSep = nSep + 1
            End If
            
            ' Replace any existing content with the required separator text
            Dim sepRng As Range
            Set sepRng = para.Range
            sepRng.MoveEnd wdCharacter, -1   ' Exclude the paragraph mark
            If (sepRng.Text <> SEP_TEXT) Then
                sepRng.Text = SEP_TEXT
            End If
        Else
            '  Count leading tab characters
            tabs = 0
            For i = 1 To Len(txt)
                If Mid(txt, i, 1) = Chr(9) Then
                    tabs = tabs + 1
                Else
                    Exit For
                End If
            Next i

            Select Case tabs
                Case 0
                    If (para.style <> doc.styles(STYLE_IDX1) And para.style <> STYLE_SEP And para.style <> STYLE_SIN And para.style <> STYLE_BOOK) Then
                        para.style = doc.styles(STYLE_IDX1)
                        nIdx1 = nIdx1 + 1
                    End If
                Case 1
                    If (para.style <> doc.styles(STYLE_IDX2)) Then
                        para.style = doc.styles(STYLE_IDX2)
                        nIdx2 = nIdx2 + 1
                    End If
                Case Else
                    If (para.style <> doc.styles(STYLE_IDX3)) Then
                        para.style = doc.styles(STYLE_IDX3)
                        nIdx3 = nIdx3 + 1
                    End If
            End Select
        End If

    Next para
End Sub


' =============================================================================
'  STEP 6 — Insert Section-Book and Section-Index at document start
' =============================================================================

Private Sub InsertBookAndSin(doc As Document)
'Dim doc As Document
'Set doc = ActiveDocument

    ' Idempotency guard: if BOOK + SIN are already in place, just refresh the text
    If doc.Paragraphs.Count >= 2 Then
        On Error Resume Next
        Dim p1Style As style
        Dim p2Style As style
        Set p1Style = doc.Paragraphs(1).style
        Set p2Style = doc.Paragraphs(2).style
        Dim rng As Range
        
       On Error GoTo 0

        If Not p1Style Is Nothing And Not p2Style Is Nothing Then
            If p1Style.NameLocal = STYLE_BOOK And p2Style.NameLocal = STYLE_SIN Then
                Call updateParaTexts(doc.Paragraphs(1), SIN_TEXT)
                Call updateParaTexts(doc.Paragraphs(2), SIN_TEXT)
                Exit Sub
            End If
        End If
    
            Dim bRng As Range
        Set bRng = doc.Paragraphs(1).Range
        bRng.MoveEnd wdCharacter, -1
        Dim iRng As Range
        Set iRng = doc.Paragraphs(2).Range
        iRng.MoveEnd wdCharacter, -1
    

'    ' if 1st para (and 2nd para) has the right text but wrong style, apply
'    Debug.Print bRng.Text
    If (bRng.Text = SIN_TEXT) Then
        If (iRng.Text = SIN_TEXT) Then
            doc.Paragraphs(1).style = doc.styles(STYLE_BOOK)
            doc.Paragraphs(2).style = doc.styles(STYLE_SIN)
            Debug.Print "A"
            Exit Sub
        Else
            doc.Paragraphs(1).style = doc.styles(STYLE_SIN)
            ' … then insert BOOK before it
            Set rng = doc.Range(0, 0)
            rng.InsertBefore SIN_TEXT & vbCr
            doc.Paragraphs(1).style = doc.styles(STYLE_BOOK)
            Debug.Print "B"
            Exit Sub
        End If
    Else
    Debug.Print "d"
        ' Insert SIN paragraph first at position 0 …
        Set rng = doc.Range(0, 0)
        rng.InsertBefore SIN_TEXT & vbCr
        doc.Paragraphs(1).style = doc.styles(STYLE_SIN)
        ' … then insert BOOK before it
        Set rng = doc.Range(0, 0)
        rng.InsertBefore SIN_TEXT & vbCr
        doc.Paragraphs(1).style = doc.styles(STYLE_BOOK)
    End If

End If
End Sub


Sub updateParaTexts(para As Paragraph, newText As String)
Dim rng As Range
Set rng = para.Range
rng.MoveEnd wdCharacter, -1
rng.Text = newText
End Sub


Sub trimDocEnd(doc)
'Dim doc As Document
'Set doc = ActiveDocument
While (Trim(doc.Paragraphs.Last.Range.Text) = vbCr)
    doc.Paragraphs.Last.Range.Delete
Wend
End Sub

