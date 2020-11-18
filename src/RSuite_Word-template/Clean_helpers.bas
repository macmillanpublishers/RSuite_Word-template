Attribute VB_Name = "Clean_helpers"
Public Function CheckTemplate()

    Dim templateName As String
    templateName = ActiveDocument.AttachedTemplate
    
    If InStr(templateName, "RSuite") < 1 And InStr(templateName, "Macmillan") < 1 Then
        MsgBox "You do not have a style template applied. This will cause errors. Please attach a style template and try again."
        End
    End If
    
End Function


Public Function FindReplaceSimple(ByVal sFind As String, ByVal sReplace As String, Optional storyNumber As Variant = 1)
    ActiveDocument.StoryRanges(storyNumber).Select
    Selection.Collapse Direction:=wdCollapseStart
    Call ClearSearch
    
    With Selection.Find
        .Text = sFind
        .Replacement.Text = sReplace
        .Execute Replace:=wdReplaceAll, Forward:=True
        Err.Clear
        On Error GoTo 0
      End With

End Function

Public Function FindReplaceComplex(ByVal sFind As String, _
                                    ByVal sReplace As String, _
                                    Optional bMatchCase As Boolean = False, _
                                    Optional bUseWildcards As Boolean = False, _
                                    Optional bSmallCaps As Boolean = False, _
                                    Optional bIncludeFormat As Boolean = False, _
                                    Optional storyNumber As Variant = 1)

    Call ClearSearch

    ActiveDocument.StoryRanges(storyNumber).Select
    Selection.Collapse Direction:=wdCollapseStart
    
    With Selection.Find
        .Forward = True
        .Text = sFind
        .Wrap = wdFindContinue
        .MatchWildcards = bUseWildcards
        .MatchSoundsLike = False
        .MatchCase = bMatchCase
        .MatchWholeWord = False
        .MatchAllWordForms = False
        If bIncludeFormat = True Then
            .Format = True
        Else: .Format = False
        End If
        .Font.SmallCaps = bSmallCaps
        With .Replacement
          .ClearFormatting
          .Text = sReplace
          .Font.SmallCaps = False
        End With
        .Execute Replace:=wdReplaceAll
        Err.Clear
        On Error GoTo 0
      End With
      
End Function

Function ClearSearch()

    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .MatchAllWordForms = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With
    ActiveDocument.UndoClear
    
End Function

Public Function EndOfDocumentReached() As Boolean
    Select Case ActiveDocument.Content.End
        Case Selection.End, Selection.End + 1
           EndOfDocumentReached = True
        Case Else
           EndOfDocumentReached = False
    End Select
End Function



Public Function EndOfStoryReached(storyNumber As Variant) As Boolean
    Select Case ActiveDocument.StoryRanges(storyNumber).End
        Case Selection.End, Selection.End + 1
           EndOfStoryReached = True
        Case Else
           EndOfStoryReached = False
    End Select
End Function


Public Function AtStartOfDocument() As Boolean
    Select Case ActiveDocument.Content.Start
        Case Selection.Start
           AtStartOfDocument = True
        Case Else
           AtStartOfDocument = False
    End Select
End Function

Sub TitleCase()
            
    Application.ScreenUpdating = False
    
    Dim HeadingSoFar As String, Q As String, NumWords As Integer
    Dim LowerCaseWords(), Acronyms() As Variant
    Dim MyWords As Variant
    Dim CaseHandled, AllCaps As Boolean

'CHICAGO MANUAL OF STYLE RULES:
'Capitalize the first and last words in titles and subtitles (but see rule 7), and capitalize all other major words (nouns, pronouns, verbs, adjectives, adverbs, and some conjunctions_but see rule 4).
'Lowercase the articles the, a, and an.
'Lowercase prepositions, regardless of length, except when they are used adverbially or adjectivally (up in Look Up, down in Turn Down, on in The On Button, to in Come To, etc.) or when they compose part of a Latin expression used adjectivally or adverbially (De Facto, In Vitro, etc.).
'Lowercase the conjunctions and, but, for, or, and nor.
'Lowercase to not only as a preposition (rule 3) but also as part of an infinitive (to Run, to Hide, etc.), and lowercase as in any grammatical function.
    
    
    LowerCaseWords = Array("amid ", "as ", "at ", "by ", "down ", "for ", "from ", "in ", "into ", "like ", "minus ", _
       "near ", "of ", "off ", "on ", "onto ", "over ", "past ", "per ", "plus ", "than ", "to ", "up ", "upon ", "via ", "with ", _
       "a ", "an ", "the ", _
       "and ", "but ", "or ", "nor ")
       
    Acronyms = Array("AA", "AAA", "AARP", "ABC", "ADA", "ADHD", "AFL", "AMA", "APA", "ASAP", "AWOL", "CBS", "CDC", "CIA", "CSI", _
                "DIY", "DMV", "DNC", "ESPN", "FAQ", "FBI", "GIF", "HBO", "HTML", "HIV", "I", "II", "III", "IV", "IX", "HR", "MBA", "MD", "MIA", "MLA", _
                "NAFTA", "NASA", "NASDAQ", "NBA", "NBC", "NFL", "NHL", "PBS", "PGA", "POTUS", "RADAR", "RNC", _
                "SCOTUS", "SONAR", "SPCA", "SUV", "SWAT", "UFO", "V", "VI", "VII", "VIII", "WWE", "XML", "X")
    
    HeadingSoFar = ""
    If Selection.Type <> wdSelectionIP Then Selection.Collapse
    Selection.Paragraphs(1).Range.Select
    NumWords = Selection.Words.Count
    
    For i = 1 To NumWords
        Q = LCase(Selection.Words(i))
        R = UCase(Trim(Selection.Words(i)))
        s = Trim(Selection.Words(i))
        
        For Each MyWord In LowerCaseWords
            If Q = MyWord Then
                If i <> 1 And i <> NumWords Then
                    Selection.Words(i).Case = wdLowerCase
                    CaseHandled = True
                    Exit For
                ElseIf i = 1 Or i = NumWords Then
                    Selection.Words(i).Case = wdTitleWord
                    CaseHandled = True
                    Exit For
                End If
            End If
        Next
        
        AllCaps = False
        If CaseHandled = False Then
            'All caps handled here
            For Each Acro In Acronyms
                If R = Acro Then
                    Selection.Words(i).Case = wdUpperCase
                    AllCaps = True
                End If
            Next
            If AllCaps = False Then Selection.Words(i).Case = wdTitleWord
        End If
        
        CaseHandled = False
    Next i
    
    Selection.Collapse wdEnd

End Sub

Function MessageBox(Title As String, Msg As String, Optional ByVal buttonType As Variant = vbYesNo)
    MessageBox = MsgBox(Msg, buttonType, Title)
End Function

Public Function ConvertLocalFormatting(MyStoryNo, Optional ByVal ItalTF As Boolean = False, _
                                        Optional ByVal BoldTF As Boolean = False, _
                                        Optional ByVal CapsTF As Boolean = False, _
                                        Optional ByVal SmallCapsTF As Boolean = False, _
                                        Optional ByVal UnderlineTF As Boolean = False, _
                                        Optional ByVal StrikeTF As Boolean = False, _
                                        Optional ByVal superTF As Boolean = False, _
                                        Optional ByVal subTF As Boolean = False, _
                                        Optional ByVal NewStyle As String = "")
                     
        Application.ScreenUpdating = False
        
        Dim oStyle As Style
        Dim oRng As Range
        Dim tRng As Range
        Dim currentPage, CurrentCol, CurrentLine, PrevPage, PrevCol, PrevLine
        Dim CurrSel As String
        
        If MyStoryNo < 1 Then MyStoryNo = 1
        
        Clean_helpers.ClearSearch
        
        ActiveDocument.StoryRanges(MyStoryNo).Select
        Selection.Collapse Direction:=wdCollapseStart
        
            With Selection.Find
                .Text = ""
                .Format = True
                .Font.Italic = ItalTF
                .Font.Bold = BoldTF
                .Font.AllCaps = CapsTF
                .Font.SmallCaps = SmallCapsTF
                .Font.Underline = UnderlineTF
                .Font.StrikeThrough = StrikeTF
                .Font.Superscript = superTF
                .Font.Subscript = subTF
            End With
            
            Selection.Find.Execute
            
            Do While Selection.Find.Found
                CurrSel = Selection.Text
                Set oStyle = Selection.Style
                If CurrSel = PrevSel Then
                    If oStyle = "Endnote Reference" Then
                        GoTo NextOne
                    End If
                End If
                
                If CurrSel = vbCr Or CurrSel = vbLf Or CurrSel = vbCrLf Or CurrSel = vbNewLine Or CurrSel = "" Then
                    Selection.Font.Reset
                    GoTo NextOne
                End If
                
                If InStr(CurrSel, vbCr) Or InStr(currse1, vbLf) Or InStr(CurrSel, vbCrLf) Or InStr(CurrSel, vbNewLine) Then
                    Selection.MoveEnd Unit:=wdCharacter, Count:=-1
                End If
                
                Select Case NewStyle
                    Case "bold-ital (bi)"
                        If Not oStyle.Font.Italic And Not oStyle.Font.Bold Then
                            Selection.Style = NewStyle
                        ElseIf oStyle.Font.Italic And Not oStyle.Font.Bold Then
                            Selection.Style = "bold (b)"
                        ElseIf Not oStyle.Font.Italic And oStyle.Font.Bold Then
                            Selection.Style = "ital (i)"
                        End If
                    Case "smallcaps-ital (sci)"
                        If Not oStyle.Font.Italic And Not oStyle.Font.SmallCaps Then
                            Selection.Style = NewStyle
                        ElseIf oStyle.Font.Italic And Not oStyle.Font.SmallCaps Then
                            Selection.Style = "smallcaps (sc)"
                        ElseIf Not oStyle.Font.Italic And oStyle.Font.SmallCaps Then
                            Selection.Style = "ital (i)"
                        End If
                    Case "smallcaps-bold (scb)"
                        If Not oStyle.Font.Bold And Not oStyle.Font.SmallCaps Then
                            Selection.Style = NewStyle
                        ElseIf oStyle.Font.Bold And Not oStyle.Font.SmallCaps Then
                            Selection.Style = "smallcaps (sc)"
                        ElseIf Not oStyle.Font.Bold And oStyle.Font.SmallCaps Then
                            Selection.Style = "bold (b)"
                        End If
                    Case "smallcaps-bold-ital (scbi)"
                        If Not oStyle.Font.Bold And Not oStyle.Font.SmallCaps And Not oStyle.Font.Italic Then
                            Selection.Style = NewStyle
                        ElseIf oStyle.Font.Bold And Not oStyle.Font.SmallCaps And Not oStyle.Font.Italic Then
                            Selection.Style = "smallcaps-ital (sci)"
                        ElseIf Not oStyle.Font.Bold And Not oStyle.Font.SmallCaps And oStyle.Font.Italic Then
                            Selection.Style = "smallcaps-bold (scb)"
                        ElseIf Not oStyle.Font.Bold And oStyle.Font.SmallCaps And oStyle.Font.Italic Then
                            Selection.Style = "bold (b)"
                        ElseIf oStyle.Font.Bold And oStyle.Font.SmallCaps And Not oStyle.Font.Italic Then
                            Selection.Style = "ital (i)"
                        ElseIf oStyle.Font.Bold And Not oStyle.Font.SmallCaps And oStyle.Font.Italic Then
                            Selection.Style = "smallcaps (sc)"
                        End If
                    Case "super-ital (supi)"
                        If Not oStyle.Font.Superscript And Not oStyle.Font.Italic Then
                            Selection.Style = NewStyle
                        ElseIf Not oStyle.Font.Superscript And oStyle.Font.Italic Then
                            Selection.Style = "super (sup)"
                        ElseIf oStyle.Font.Superscript And Not oStyle.Font.Italic Then
                            Selection.Style = "ital (i)"
                        End If
                    Case "ital (i)"
                        If Not oStyle.Font.Italic Then
                            Selection.Style = NewStyle
                        End If
                    Case "bold (b)"
                        If Not oStyle.Font.Bold Then
                            Selection.Style = NewStyle
                        End If
                    Case "smallcaps (sc)"
                        If Not oStyle.Font.SmallCaps Then
                            Selection.Style = NewStyle
                        End If
                    Case "underline (u)"
                        If Not oStyle.Font.Underline Then
                            Selection.Style = NewStyle
                        End If
                    Case "super (sup)"
                        If Not oStyle.Font.Superscript Then
                            Selection.Style = NewStyle
                        End If
                    Case "sub (sub)"
                        If Not oStyle.Font.Subscript Then
                            Selection.Style = NewStyle
                        End If
                    Case "strike (str)"
                        Selection.Style = NewStyle
                End Select
NextOne:
                PrevSel = Selection.Text
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                If Clean_helpers.EndOfDocumentReached Then Exit Do
                Selection.Find.Execute
            Loop

End Function

Function updateStatus(ByVal update As String)
    pBar.Status.Caption = completeStatus & vbNewLine & update
    pBar.Repaint
    Application.ScreenRefresh
End Function
