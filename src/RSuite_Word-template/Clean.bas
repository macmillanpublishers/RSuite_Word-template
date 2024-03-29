Attribute VB_Name = "Clean"
Const excludeStyle As String = "cs-cleanup-exclude (cex)"

Sub Ellipses(MyStoryNo)

        Application.ScreenUpdating = False
                
        thisstatus = "Fixing ellipses "
        If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

        'replace anythign that's already fixed, in case it's run again
        Clean_helpers.FindReplaceSimple ELLIPSIS, "<doneellipsis>", MyStoryNo
        'this makes sure all ellipses are consistent
        'ellipsis.dot = ellipsis
        Clean_helpers.FindReplaceSimple "." & ELLIPSIS_SYM, "." & TEMP_ELL, MyStoryNo
        
        'ellipsis = ellipsis
        Clean_helpers.FindReplaceSimple ELLIPSIS_SYM, TEMP_ELL, MyStoryNo
        'dot.dot.dot.dot=dot.ellipsis
        Clean_helpers.FindReplaceSimple "....", "." & TEMP_ELL, MyStoryNo
        
        'dot.space.dot.space.dot.space.dot=dot.ellipsis
        Clean_helpers.FindReplaceSimple ". . . .", "." & TEMP_ELL, MyStoryNo
        'dot.dot.dot=ellipsis
        Clean_helpers.FindReplaceSimple "...", TEMP_ELL, MyStoryNo
        
        'dot.space.dot.space.dot=ellipsis
        Clean_helpers.FindReplaceSimple ". . .", TEMP_ELL, MyStoryNo
        'dot.space.dot.space.dot=ellipsis
        Clean_helpers.FindReplaceSimple TEMP_ELL & "." & aSPACE, "." & TEMP_ELL, MyStoryNo
        
        'space.dot.tempell=tempell
        Clean_helpers.FindReplaceSimple aSPACE & "." & TEMP_ELL, "." & TEMP_ELL, MyStoryNo
         'dot.space.dot.space.dot=ellipsis
        Clean_helpers.FindReplaceSimple TEMP_ELL & aSPACE, TEMP_ELL, MyStoryNo
        
        'fix all double spaces before and after ellipses
        Clean_helpers.FindReplaceComplex aSPACE & "{1,}" & TEMP_ELL, TEMP_ELL, False, True, , , MyStoryNo
        Clean_helpers.FindReplaceComplex TEMP_ELL & aSPACE & "{2,}", TEMP_ELL & aSPACE, False, True, , , MyStoryNo
    
        ActiveDocument.StoryRanges(MyStoryNo).Select
        With Selection.Find
            .MatchWildcards = False
            .ClearFormatting
            .Execute FindText:=TEMP_ELL
        End With
        While Selection.Find.Found
            
            ' moveLeft 2 looks at the character preceding found selection, resolve by case
            ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Select Case Selection.Text
                Case RTN, vbCr
                    'do nothing
                Case DP, DOQ, SOQ
                    ' add an nbsp trailing the ellipse
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.TypeText NBSPchar
                Case EMDASH
                    ' add a space following emdash, preceding ellipse
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.TypeText aSPACE
                    ' and add an nbsp trailing the ellipse
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.TypeText NBSPchar
                Case Else
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Select Case Selection.Text
                        ' if emdash _trails_ ellipses
                        Case EMDASH
                            ' add preceding _and_ trailing nbsps to ellipse
                            Selection.TypeText NBSPchar
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.MoveLeft Unit:=wdCharacter, Count:=1
                            Selection.TypeText NBSPchar
                        ' for all other leading chars,
                        Case Else
                            ' add nbsp preceding ellipse
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.MoveLeft Unit:=wdCharacter, Count:=1
                            Selection.TypeText NBSPchar
                        End Select
            End Select

            ActiveDocument.Bookmarks("temp").Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Select Case Selection.Text
                Case DP, DCQ, SCQ, RTN, NBSPchar, aSPACE, vbCr
                    'do nothing
                Case ";", ";", ":", ";", ",", "?", "!", ")"
                    Selection.TypeText NBSPchar
                Case Else
                    Selection.TypeText aSPACE
            End Select
            ActiveDocument.Bookmarks("temp").Delete
            Selection.Find.Execute
        Wend
        
    'dot.space.dot.space.dot=space.ellipsis.space
    Clean_helpers.FindReplaceSimple TEMP_ELL, ELLIPSIS, MyStoryNo
    
    'replace anything that's already fixed, in case it's run again
    Clean_helpers.FindReplaceComplex "<doneellipsis>", ELLIPSIS, True, False, , , MyStoryNo
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
        
End Sub

Sub Spaces(MyStoryNo)

    thisstatus = "Fixing spaces "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'temporarily change finished ellipses and delete nonbreaking spaces
    Clean_helpers.FindReplaceSimple EMDASH_ELLIPSIS, "<doneemdashellipsis>", MyStoryNo
    Clean_helpers.FindReplaceSimple PERIOD_ELLIPSIS, "<doneperiodellipsis>", MyStoryNo
    Clean_helpers.FindReplaceSimple NBS_ELLIPSIS, "<donenbsellipsis>", MyStoryNo
    Clean_helpers.FindReplaceSimple QUOTE_ELLIPSIS, "<donequoteellipsis>", MyStoryNo
    Clean_helpers.FindReplaceSimple ELLIPSIS, "<doneellipsis>", MyStoryNo
    'nonbreaking space to regular space
    Clean_helpers.FindReplaceComplex ChrW(202), " ", False, True, , , MyStoryNo
    Clean_helpers.FindReplaceComplex ChrW(160), " ", False, True, , , MyStoryNo
    'change ellipses back
    Clean_helpers.FindReplaceSimple "<doneellipsis>", ELLIPSIS, MyStoryNo
    Clean_helpers.FindReplaceSimple "<donequoteellipsis>", QUOTE_ELLIPSIS, MyStoryNo
    Clean_helpers.FindReplaceSimple "<donenbsellipsis>", NBS_ELLIPSIS, MyStoryNo
    Clean_helpers.FindReplaceSimple "<doneperiodellipsis>", PERIOD_ELLIPSIS, MyStoryNo
    Clean_helpers.FindReplaceSimple "<doneemdashellipsis>", EMDASH_ELLIPSIS, MyStoryNo
    'multiple tabs to regular space
    Clean_helpers.FindReplaceComplex_WithExclude "^9{1,}", " ", False, True, , , MyStoryNo
    'multiple spaces to one space
    Clean_helpers.FindReplaceComplex_WithExclude " {2,}", " ", False, True, , , MyStoryNo
    'soft returns to hard returns
    '  note: replacing with ^p using WithExclude function requires 'vbnewline' instead
    Clean_helpers.FindReplaceSimple_WithExclude "^l", vbNewLine, MyStoryNo

   ' these 2 modified f-and-r's (along with TrimTrailingSpace below (from wdv-387)) help with
   '    extra para problem surfaced in wdv-354, but are more efficient with notes as per wdv-395
    Clean_helpers.TrimLeadingSpace " " + ChrW(13), "^p", MyStoryNo
    Clean_helpers.TrimLeadingSpace " ^p", "^p", MyStoryNo

    ' these 2 f-and-r's get special attention because they span 2 paras.
    Clean_helpers.TrimTrailingSpace_WithExclude ChrW(13) + " ", vbNewLine, MyStoryNo
    Clean_helpers.TrimTrailingSpace_WithExclude "^p ", vbNewLine, MyStoryNo

    'space before/after brackets to no space
    Clean_helpers.FindReplaceSimple_WithExclude "( ", "(", MyStoryNo
    Clean_helpers.FindReplaceSimple_WithExclude "[ ", "[", MyStoryNo
    Clean_helpers.FindReplaceSimple_WithExclude "{ ", "{", MyStoryNo
    Clean_helpers.FindReplaceSimple_WithExclude " )", ")", MyStoryNo
    Clean_helpers.FindReplaceSimple_WithExclude " ]", "]", MyStoryNo
    Clean_helpers.FindReplaceSimple_WithExclude " }", "}", MyStoryNo
    'space after dollar sign to no space
    Clean_helpers.FindReplaceSimple_WithExclude "$ ", "$", MyStoryNo

    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
    
End Sub

Sub Punctuation(MyStoryNo)

    thisstatus = "Fixing punctuation "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'multiple periods to single period
    Clean_helpers.FindReplaceComplex ".{2,}", ".", False, True, , , MyStoryNo
    'multiple commas to single comma
    Clean_helpers.FindReplaceComplex ",{2,}", ",", False, True, , , MyStoryNo
    'optional hyphen to nothing
    Clean_helpers.FindReplaceSimple OPTHYPH, "", MyStoryNo
    Clean_helpers.FindReplaceSimple OPTHYPH2, "", MyStoryNo
    'non-breaking hyphen to regular hyphen
    Clean_helpers.FindReplaceSimple NBHYPH, "-", MyStoryNo
    Clean_helpers.FindReplaceSimple NBHYPH2, "-", MyStoryNo
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
End Sub

Sub DoubleQuotes(MyStoryNo)
            
    Application.ScreenUpdating = False
    ActiveDocument.StoryRanges(MyStoryNo).Select
    
    Dim totalPages, currentPage, nextPercentage As Integer
    Dim currPercentage, newPercentage As Integer
    ActiveDocument.Repaginate
    totalPages = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
    currPercentage = 0
    
    thisstatus = "Fixing double quotes"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    ' Combine double single-primes into Double-prime, also double-backticks
    FindReplaceSimple "``", DP, MyStoryNo
    FindReplaceSimple SP & SP, DP, MyStoryNo
    
    ActiveDocument.StoryRanges(MyStoryNo).Select
    Selection.Find.Execute FindText:=DP
    Do While Selection.Find.Found
        ' Find / Replace tool includes DOQ and DCQ as results in a search for DP
        '   for some reason (Windows/Office2013)
        '   we can filter them out here with the next line:
        If Selection.Text = DP Then

            newPercentage = Selection.Range.Information(wdActiveEndPageNumber) / totalPages * 100
            If newPercentage > currPercentage Then
                thisstatus = "Fixing double quotes: " & CStr(newPercentage) & "%"
                If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
                currPercentage = newPercentage
            End If
            
            ' test preceding char for following case statement
            ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Select Case Selection.Text
                Case EMDASH
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Select Case Selection.Text
                        ' preceding emdash and trailing whitespace = DCQ
                        Case " ", vbCr
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DCQ
                        ' otherwise preceding emdash = DOQ
                        Case Else
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DOQ
                    End Select
                Case " "
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.Expand Unit:=wdCharacter
                    Select Case Selection.Text
                        ' preceding space and trailing whitespace = DCQ
                        Case " ", vbCr
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DCQ
                        ' preceding space and trailing SP = DOQ-SOQ (replacing SP)
                        Case SP
                            Selection.TypeText SOQ
                            Selection.Expand Unit:=wdCharacter
                            Select Case Selection.Text
                                Case DP
                                    Selection.TypeText DOQ
                            End Select
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DOQ
                        ' preceding space other = DOQ
                        Case Else
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DOQ
                    End Select
                Case vbCr, vbTab, "("
                    ' if preceding return, tab or open paren type DOQ
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.TypeText DOQ
                    Selection.Expand Unit:=wdCharacter
                    Select Case Selection.Text
                        ' if trailing char is SP replace it with SOQ
                        Case SP
                            Selection.TypeText SOQ
                            Selection.Expand Unit:=wdCharacter
                            ' if trailing that SP is DP, replace it with another DOQ
                            Select Case Selection.Text
                                Case DP
                                    Selection.TypeText DOQ
                            End Select
                    End Select
                Case Else
                    ' if we are the first char or doc, make DP a DOQ
                    If Clean_helpers.AtStartOfDocument Then
                        ActiveDocument.Bookmarks("temp").Select
                        Selection.TypeText DOQ
                    ' otherwise make DP a DCQ
                    Else
                        ActiveDocument.Bookmarks("temp").Select
                        Selection.TypeText DCQ
                    End If
                    Selection.MoveLeft Unit:=wdCharacter, Count:=2
                    Select Case Selection.Text
                        'if preceding char is SP, change to SCQ
                        Case SP
                            Selection.Delete
                            Selection.TypeText SCQ
                            Selection.MoveLeft Unit:=wdCharacter, Count:=2
                            Select Case Selection.Text
                                ' if char preceding SP is DP, make it a DCQ
                                Case DP
                                    Selection.Delete
                                    Selection.TypeText DCQ
                            End Select
                    End Select
                End Select
            ' get out of that selection region for next find!
            Selection.MoveRight Unit:=wdCharacter, Count:=3
       End If
            If Clean_helpers.EndOfStoryReached(MyStoryNo) Then Exit Do
            Selection.Find.Execute

    Loop

    completeStatus = completeStatus + vbNewLine + "Fixing double quotes: 100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub

Sub SingleQuotes(MyStoryNo)

    Application.ScreenUpdating = False
    
    Dim nextPercentage As Integer
    nextPercentage = 30
    
    thisstatus = "Fixing single quotes "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    Dim ChangeQ As Boolean
    ChangeQ = False
    
    ' check backtick chars
    ActiveDocument.StoryRanges(MyStoryNo).Select
    Selection.Find.ClearFormatting
    Selection.Find.Execute FindText:="`"
    While Selection.Find.Found
        ' get preceding character
        ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
        Selection.MoveLeft Unit:=wdCharacter, Count:=2
        ' if preceding char is space, return, or open paren, replace selection with SOQ
        If Selection.Text = " " Or Selection.Text = vbCr Or Selection.Text = "(" Then
            ActiveDocument.Bookmarks("temp").Select
            Selection.TypeText SOQ
        ' else replace with SXQ
        Else:
            ActiveDocument.Bookmarks("temp").Select
            Selection.TypeText SCQ
        End If
        Selection.Find.Execute
    Wend
    
    Dim StringFound, OpenQuo As Boolean
    Dim SearchString(), QuoStr
    
    ' WDV-281: 7-14-20
    '   "educating" already 'smart' single quotes results in some user-intended use-cases to be overridden
    '   leaving the capability to search SOQ/SCQ via this array setup in case we end up reversing course
    SearchString = Array(SP) ', SOQ, SCQ)
           
    ActiveDocument.StoryRanges(MyStoryNo).Select
    For Each QuoStr In SearchString
        
        thisstatus = "Fixing single quotes: " & CStr(nextPercentage) & "%"
        If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
        nextPercentage = nextPercentage + 30
    
        Selection.Find.Execute FindText:=QuoStr
        
        While Selection.Find.Found
            ' Find / Replace tool includes SOQ and SCQ as results in a search for SP
            '   for some reason (Windows/Office2013)
            '   we can filter them out here
            If Selection.Text = SP Then
            
                ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
                Selection.MoveLeft Unit:=wdCharacter, Count:=2
                
                ' if preceding char is open quote or double-prime, default is SOQ
                '   else default is SCQ
                Select Case Selection.Text
                        Case DP, DOQ, SOQ
                            OpenQuo = True
                End Select
                
                ' if default is SOQ, for any of the following lookaheads we would flip to SCQ
                Select Case Selection.Text
                        Case " ", vbCr, vbTab, vbNewLine, "(", DP, DOQ, SOQ
                            Selection.MoveRight Unit:=wdCharacter, Count:=2
                            Selection.ExtendMode = True
                            
                            '1 character
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case DOQ, "K", "k"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '2 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If IsYear(Selection.Text) = True Then
                                ChangeQ = True
                                GoTo SkipToHere
                            ElseIf LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "em", "Em", "er", "Er", "Im", "im", "n" & SCQ, "N" & SCQ
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '3 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Cuz", "cuz", "Net", "net", "Sup", "sup", "Tar", "tar", "Til", "til", "Tis", "tis"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '4 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Bout", "bout", "Cept", "cept", "Fore", "fore", "Nuff", "nuff", "Post", "post", "Tall", "tall", "Twas", "twas"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '5 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Cause", "cause", "Fraid", "fraid", "Night", "night", "Round", "round", "Scuse", "scuse", "Sides", "sides", "Spect", "spect", "Tever", "tever"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
        
                            '6 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Course", "course", "Gainst", "gainst", "Nother", "nother", "Splain", "splain", "Tain" & SCQ & "t", "tain" & SCQ & "t", "Tisn" & SCQ & "t", "tisn" & SCQ & "t"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '7 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Chother", "chother", "Druther", "druther", "Salmost", "salmost", "Snothin", "snothin", "Twasn" & SCQ & "t", "twasn" & SCQ & "t"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '8 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Druthers", "druthers", "Tweren" & SCQ & "t", "tweren" & SCQ & "t"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '9 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Specially", "specially", "Spossible", "spossible"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '10 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            'If LookAhead() = True Then
                                'Select Case Selection.Text
                                    'Case
                                        'ChangeQ = True
                                'End Select
                                'GoTo SkipToHere
                            'End If
                                
                            '11 characters
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Neverything", "neverything"
                                        ChangeQ = True
                                End Select
                            End If
                            
SkipToHere:
     
                            Selection.ExtendMode = False
                            ActiveDocument.Bookmarks("temp").Select
                            Select Case ChangeQ
                                Case True
                                    Selection.TypeText SCQ
                                Case False
                                    Selection.TypeText SOQ
                            End Select
                        
                    Case Else
                        If Not (OpenQuo = True) Then
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText SCQ
                        End If
    
                End Select
            End If
        
        ChangeQ = False
        OpenQuo = False
        
        Selection.Find.Execute
    
    Wend
Next

' Remove spaces between certain quote combinations.
'   7/14/20 -- Commenting some of these space removals as per WDV-281
'   Also, though they were previously set before DP conversion,
'   SP & DP replacements captured relative OQ and CQ too;
'   So moving them to the end of quote cleanup, where Primes have already become quotes
FindReplaceSimple SOQ & aSPACE & DOQ, SOQ & DOQ, MyStoryNo
FindReplaceSimple DOQ & aSPACE & SOQ, DOQ & SOQ, MyStoryNo
FindReplaceSimple SCQ & aSPACE & DCQ, SCQ & DCQ, MyStoryNo
FindReplaceSimple DOQ & aSPACE & SCQ, DOQ & SCQ, MyStoryNo
FindReplaceSimple DCQ & aSPACE & SCQ, DCQ & SCQ, MyStoryNo
'FindReplaceSimple SCQ & aSPACE & DOQ, SCQ & DOQ
'FindReplaceSimple DCQ & aSPACE & SOQ, DCQ & SOQ
'FindReplaceSimple SOQ & aSPACE & DCQ, SOQ & DCQ

completeStatus = completeStatus + vbNewLine + "Fixing Single Quotes: 100%"
If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub

Function LookAhead() As Boolean

    ActiveDocument.Bookmarks.Add ("myTemp")
    Selection.ExtendMode = False
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Select Case Selection.Text
        Case " ", ".", ",", "?", "!", EMDASH, ")"
            LookAhead = True
        Case Else
            LookAhead = False
    End Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.ExtendMode = True
    ActiveDocument.Bookmarks("myTemp").Select
    ActiveDocument.Bookmarks("myTemp").Delete

End Function

Function IsYear(theNumber) As Boolean

    If theNumber Like "[0-9][0-9]" Then
    
            ActiveDocument.Bookmarks.Add ("myTemp")
            Selection.ExtendMode = False
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Select Case Selection.Text
                Case " ", ".", ",", "?", "!", EMDASH, ")", "s"
                    IsYear = True
                Case Else
                    IsYear = False
            End Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.ExtendMode = True
            ActiveDocument.Bookmarks("myTemp").Select
            ActiveDocument.Bookmarks("myTemp").Delete
            
    End If

End Function

Sub Dashes(MyStoryNo)
   
    thisstatus = "Fixing dashes "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    Application.ScreenUpdating = False
    
    ' Word wildcards isn't as flexible as regex, cannot use ? or | in the normal ways
    '   So we have to run multiple queries
    Call HighlightNumber("1-[0-9]{3}-[0-9]{3}-[0-9]{4}", MyStoryNo)
    Call HighlightNumber("\([0-9]{3}\) [0-9]{3}-[0-9]{4}", MyStoryNo)
    Call HighlightNumber("[0-9]{3}-[0-9]{3}-[0-9]{4}", MyStoryNo)
    
    ' FOLLOWING CAN BE USED TO FIND ISBN PATTERN AND FLAG FOR NO CHANGE
    Call HighlightNumber("97[89]-[0-9]{10,14}", MyStoryNo)
    Call HighlightNumber("97[89]-[0-9]-[0-9]{3}-[0-9]{5}-[0-9]", MyStoryNo)
     
    thisstatus = "Fixing dashes: 10%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    Dim activeRng As Range
    Dim h As Range
    For i = 0 To 9
        For j = 0 To 9
        
            Set activeRng = ActiveDocument.StoryRanges(MyStoryNo)
    
            With activeRng.Find
                .ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
                .Text = LTrim(i) & ("-") & LTrim(j)
                .MatchWildcards = False
                While .Execute
                    ' rst-1249 -- exempting superscript hyphens too
                    If Not (activeRng.FormattedText.HighlightColorIndex = wdPink) And _
                    Not (activeRng.Characters(2).Font.Superscript = True) And _
                    Not (activeRng.style = "Hyperlink") And _
                    Not (activeRng.style = excludeStyle) Then
                        activeRng.Text = LTrim(i) & ENDASH & LTrim(j)
                    End If
                    activeRng.Collapse wdCollapseEnd
                Wend
             End With
        Next
    Next

    thisstatus = "Fixing dashes: 20%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'weird-character = emdash
    FindReplaceSimple ChrW(-3906), EMDASH, MyStoryNo
    'bar character = emdash
    FindReplaceSimple ChrW(8213), EMDASH, MyStoryNo

    thisstatus = "Fixing dashes: 30%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'figure dash=endash
    FindReplaceSimple ChrW(8210), ENDASH, MyStoryNo
    'hyphen.hyphen.hyphen=endash
    FindReplaceSimple_WithExcludeOrHyperlink "---", EMDASH, MyStoryNo
    'space.hyphen.space=emdash
    FindReplaceSimple_WithExclude " - ", "-", MyStoryNo

    thisstatus = "Fixing dashes: 40%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'space.hyphen.hyphen.space=emdash
    FindReplaceSimple_WithExclude " -- ", EMDASH, MyStoryNo
    'hyphen.hyphen=emdash
    FindReplaceSimple_WithExcludeOrHyperlink "--", EMDASH, MyStoryNo

    thisstatus = "Fixing dashes: 50%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

   'dash.space=dash
    FindReplaceSimple_WithExclude "-" & aSPACE, "-", MyStoryNo
    'space.dash=dash
    FindReplaceSimple_WithExclude aSPACE & "-", "-", MyStoryNo

    thisstatus = "Fixing dashes: 60%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'space.endash=emdash
    FindReplaceSimple_WithExclude aSPACE & ENDASH, EMDASH, MyStoryNo
    'endash.space=emdash
    FindReplaceSimple_WithExclude ENDASH & aSPACE, ENDASH, MyStoryNo

    thisstatus = "Fixing dashes: 70%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'emdash.space=emdash
    FindReplaceSimple_WithExclude EMDASH & aSPACE, EMDASH, MyStoryNo
    'space.emdash=emdash
    FindReplaceSimple_WithExclude aSPACE & EMDASH, EMDASH, MyStoryNo

    thisstatus = "Fixing dashes: 80%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
    
    Call removeHighlight(MyStoryNo)
    
    thisstatus = "Fixing dashes: 90%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
    
    completeStatus = completeStatus + vbNewLine + "Fixing Dashes: 100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
End Sub

Function HighlightNumber(myPattern, Optional storyNumber As Variant = 1)
    
    ActiveDocument.StoryRanges(storyNumber).Select
    Selection.Collapse Direction:=wdCollapseStart
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = myPattern
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute
    End With
    
    Do While Selection.Find.Found
        Selection.Range.HighlightColorIndex = wdPink
        Selection.MoveRight
        If Clean_helpers.EndOfStoryReached(storyNumber) Then Exit Do
        Selection.Find.Execute
    Loop

End Function

Function removeHighlight(Optional storyNumber As Variant = 1)
    ActiveDocument.StoryRanges(storyNumber).Select
    Selection.Collapse Direction:=wdCollapseStart
    
    Options.DefaultHighlightColorIndex = wdPink

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Highlight = True
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Replacement.Text = ""
        .Replacement.Highlight = False
        .Execute Replace:=wdReplaceAll
    End With


End Function

Function MakeTitleCase(MyStoryNo)

    thisstatus = "Converting headings to title case "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    If MyStoryNo = 0 Then MyStoryNo = 1
    
    Dim tcStyles() As Variant
    tcStyles = Array("Title (Ttl)", "Number (Num)", "Main-Head (MHead)")
    
    For Each TC In tcStyles
        Clean_helpers.ClearSearch
        
        ActiveDocument.StoryRanges(MyStoryNo).Select
        Selection.Collapse Direction:=wdCollapseStart
    
        With Selection.Find
            .Wrap = wdFindStop
            .style = TC
            .Execute
        End With
        
        Do While Selection.Find.Found
            Clean_helpers.TitleCase
            Selection.MoveRight
            If Clean_helpers.EndOfStoryReached(MyStoryNo) Then Exit Do
            Selection.Find.Execute
        Loop
    Next
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Function

Function CleanBreaks(MyStoryNo)

    thisstatus = "Cleaning breaks "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    FindReplaceSimple_WithExclude "^l", vbNewLine, MyStoryNo
    ' ^ replacing with ^p with WithExclude function must be done with vbnewline instead
    FindReplaceSimple "^m", "^p", MyStoryNo
    FindReplaceSimple "^b", "^p", MyStoryNo

    Call Clean_helpers.CleanConsecutiveBreaks(MyStoryNo)
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
End Function

Function RemoveTrackChanges()

    thisstatus = "Removing Track Changes "
   ' Clean_helpers.updateStatus (thisStatus)
    
    Dim StoryNo As Variant
    Dim tc_sum As Long
    Dim tc_subtotal As Integer
    Dim rm_revisions As Boolean
    tc_sum = 0
    accept_revisions = False
    
    'get a count of track changes for the whole document
    'determine stories in document
    For Each StoryNo In ActiveDocument.StoryRanges
        'run on main (1) endnotes (2), and footnotes (3)
        If StoryNo.StoryType < 4 Then
            tc_sum = tc_sum + ActiveDocument.StoryRanges(StoryNo.StoryType).Revisions.Count
        End If
    Next
    
    'see if the user wants to accept changes
    If tc_sum > 0 Then
        If Clean_helpers.MessageBox("ACCEPT TRACK CHANGES", "Your document contains unacccepted Track Changes, which must be removed before the file is transformed in RSuite." & vbNewLine & vbNewLine & _
          "Select YES to accept all changes in the document." & vbNewLine & vbNewLine & _
          "Select NO to retain Track Changes.") = vbYes Then
                accept_revisions = True
        End If
    End If
    
    ' if YES, delete revisions per story, where present
    If accept_revisions = True Then
        'determine stories in document
        For Each StoryNo In ActiveDocument.StoryRanges
            'run on main (1) endnotes (2), and footnotes (3)
            If StoryNo.StoryType < 4 Then
                If ActiveDocument.StoryRanges(StoryNo.StoryType).Revisions.Count > 0 Then
                    ActiveDocument.StoryRanges(StoryNo.StoryType).Revisions.AcceptAll
                End If
            End If
        Next
    End If
    
End Function


Function RemoveComments()

    thisstatus = "Removing Comments "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
    
    Dim c As Comment
    If ActiveDocument.Comments.Count > 0 Then
        If Clean_helpers.MessageBox("DELETE COMMENTS", "Your document contains Comments, which must be removed before the file is transformed in RSuite." & vbNewLine & vbNewLine & _
            "Select YES to remove all comments in the document." & vbNewLine & vbNewLine & _
            "Select NO to retain comments.") = vbYes Then
                ActiveDocument.DeleteAllComments
        End If
    End If
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
    
End Function

Function DeleteBookmarks()

    thisstatus = "Deleting Bookmarks "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
    
    Dim B As Bookmark
    For Each B In ActiveDocument.Bookmarks
        B.Delete
    Next
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
End Function


Function DeleteObjects(MyStoryNo)

    thisstatus = "Deleting Objects "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    Dim s As Shape
    Dim i As InlineShape
    Dim F As Frame
    Dim R As Range
    Dim G As Integer
    Dim TB As TextFrame
    Dim shape_count As Long
    
    'Like hyperlinks, cycling through via 'For each'  for some reason doesn't make it through all shapes
    '  using same 'do while' concstruct, + exit if count is static through a cycle, to prevent loop on 'undeleteable'
    Do While ActiveDocument.Shapes.Count > 0 And shape_count <> ActiveDocument.Shapes.Count
        shape_count = ActiveDocument.Shapes.Count
        For Each s In ActiveDocument.Shapes
            If s.Type = msoTextBox Then
                s.Anchor.Select
                Selection.MoveLeft Unit:=wdCharacter
                Selection.MoveDown Unit:=wdParagraph
                Selection.TypeText s.TextFrame.TextRange.Text
                s.Delete
            ElseIf s.Type = msoGroup Then
                For G = 1 To s.GroupItems.Count
                    If s.GroupItems(G).Type = 17 Then
                        Set TB = s.GroupItems(G).TextFrame
                        s.Anchor.Select
                        Selection.MoveLeft Unit:=wdCharacter
                        Selection.MoveDown Unit:=wdParagraph
                        Selection.TypeText TB.TextRange.Text
                    End If
                Next G
                s.Delete
            Else
                s.Delete
            End If
        Next
    Loop
    
    For Each i In ActiveDocument.StoryRanges(MyStoryNo).InlineShapes
        i.Delete
    Next
    
    For Each F In ActiveDocument.StoryRanges(MyStoryNo).Frames
        F.Delete
    Next
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
End Function

Function RemoveHyperlinks(Optional MyStoryNo As Variant = 1)

    thisstatus = "Removing hyperlinks "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
    
    Dim link_count As Integer
    Dim h As hyperlink
    link_count = 0

    ' since in testing, one pass of the for loop did not catch all hyperlinks:
    '   while hyperlink count is greater than 0 we run through again. If link_count ever
    '   matches hyperlink count, we know we're passing through and not able to delete remaining links
    '   for whatever reason, so exit to avoid a crash.
    Do While ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count > 0 And link_count <> ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
        link_count = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
        For Each h In ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks
            h.Range.style = "Hyperlink"
            h.Delete
        Next
    Loop
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Function
Sub LocalFormatting(MyStoryNo)

    thisstatus = "Replacing Local Formatting with Character Styles "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    'Application.ScreenUpdating = False '< should already be off unless we are running standalone
    
    ' fix for wdv-479
    Call TableEndCharFontReset
    
    'small caps bold italic
    Call ConvertLocalFormatting(MyStoryNo, SmallCapsTF:=True, ItalTF:=True, BoldTF:=True, NewStyle:="smallcaps-bold-ital (scbi)")
    
    'bold ital
    Call ConvertLocalFormatting(MyStoryNo, ItalTF:=True, BoldTF:=True, NewStyle:="bold-ital (bi)")
    
    'small caps bold
    Call ConvertLocalFormatting(MyStoryNo, SmallCapsTF:=True, BoldTF:=True, NewStyle:="smallcaps-bold (scb)")
    
    'small caps ital
    Call ConvertLocalFormatting(MyStoryNo, SmallCapsTF:=True, ItalTF:=True, NewStyle:="smallcaps-ital (sci)")
    
    'strikethrough
    Call ConvertLocalFormatting(MyStoryNo, StrikeTF:=True, NewStyle:="strike (str)")
    
    'superscript italic
    Call ConvertLocalFormatting(MyStoryNo, superTF:=True, ItalTF:=True, NewStyle:="super-ital (supi)")
    
    'superscript
    Call ConvertLocalFormatting(MyStoryNo, superTF:=True, NewStyle:="super (sup)")
    
    'subscript
    Call ConvertLocalFormatting(MyStoryNo, subTF:=True, NewStyle:="sub (sub)")
    
    'ital
    Call ConvertLocalFormatting(MyStoryNo, ItalTF:=True, NewStyle:="ital (i)")
    
    'bold
    Call ConvertLocalFormatting(MyStoryNo, BoldTF:=True, NewStyle:="bold (b)")

    'small caps
    Call ConvertLocalFormatting(MyStoryNo, SmallCapsTF:=True, NewStyle:="smallcaps (sc)")
        
    'underline
    Call ConvertLocalFormatting(MyStoryNo, UnderlineTF:=True, NewStyle:="underline (u)")
    
    'as per RST-1231, we take some LF-combos with no RS equiv. and revert them to ones that do:
    'bold + strikethrough = strikethrough
    Call ConvertLocalFormatting(MyStoryNo, BoldTF:=True, StrikeTF:=True, NewStyle:="strike (str)")
    'ital + strikethrough = strikethrough
    Call ConvertLocalFormatting(MyStoryNo, ItalTF:=True, StrikeTF:=True, NewStyle:="strike (str)")
    'bold + underline = underline
    Call ConvertLocalFormatting(MyStoryNo, BoldTF:=True, UnderlineTF:=True, NewStyle:="underline (u)")
    'ital + underline = underline
    Call ConvertLocalFormatting(MyStoryNo, ItalTF:=True, UnderlineTF:=True, NewStyle:="underline (u)")
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
    'Application.ScreenUpdating = True '<for debug only

End Sub



Function getFormatCharStyles() As Variant

getFormatCharStyles = Array("bold (b)", "ital (i)", "smallcaps (sc)", "underline (u)", "super (sup)", "sub (sub)", _
                          "bold-ital (bi)", "smallcaps-ital (sci)", "smallcaps-bold (scb)", _
                          "smallcaps-bold-ital (scbi)", "super-ital (supi)", "strike (str)")

End Function
Sub AppliedCharStylesHelper(selectedChar)
    Dim defaultStyle As Variant
    defaultStyle = WdBuiltinStyle.wdStyleDefaultParagraphFont
    With selectedChar.Font
        If .Italic Then
            If .Bold Then
                If .SmallCaps Then
                    selectedChar.style = "smallcaps-bold-ital (scbi)"
                Else
                    selectedChar.style = "bold-ital (bi)"
                End If
            ElseIf .Superscript Then
                selectedChar.style = "super-ital (supi)"
            ElseIf .SmallCaps Then
                selectedChar.style = "smallcaps-ital (sci)"
            Else
                selectedChar.style = "ital (i)"
            End If
        ElseIf .Bold Then
            If .SmallCaps Then
                selectedChar.style = "smallcaps-bold (scb)"
            Else
                selectedChar.style = "bold (b)"
            End If
        ElseIf .SmallCaps Then
            selectedChar.style = "smallcaps (sc)"
        ElseIf .Superscript Then
            selectedChar.style = "super (sup)"
        ElseIf .Underline Then
            selectedChar.style = "underline (u)"
        ElseIf .StrikeThrough Then
            selectedChar.style = "strike (str)"
        ElseIf .Subscript Then
            selectedChar.style = "sub (sub)"
        Else
            selectedChar.style = defaultStyle
        End If
    End With
End Sub


Sub CheckAppliedCharStyles(MyStoryNo)
' this version removes extraneous direct formatting, to match the applied style
    thisstatus = "Checking Applied Character-Format Styles "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
        
      '  Dim t As Single
      '  t = Timer
        
        Application.ScreenUpdating = False

        Dim styleList() As Variant
        Dim B, i, sc, subs, sup, strk, u As Boolean
        Dim styleApplied As Boolean
        Dim numChars As Integer
        Dim selectedChar As Range

        If MyStoryNo < 1 Then MyStoryNo = 1

        styleList = getFormatCharStyles
        Clean_helpers.ClearSearch

        For Each myStyle In styleList
            ActiveDocument.StoryRanges(MyStoryNo).Select
            Selection.Collapse Direction:=wdCollapseStart
            
            On Error Resume Next
            With Selection.Find
                .style = ActiveDocument.styles(myStyle)
                .Execute
            End With
            On Error GoTo 0

            Do While Selection.Find.Found
                numChars = Selection.Characters.Count
                'styleApplied = False
                ' cycle through characters of selected range
                For k = 1 To numChars
                    Set selectedChar = Selection.Characters(k)
                    B = selectedChar.Font.Bold
                    i = selectedChar.Font.Italic
                    sc = selectedChar.Font.SmallCaps
                    subs = selectedChar.Font.Subscript
                    sup = selectedChar.Font.Superscript
                    strk = selectedChar.Font.StrikeThrough
                    u = selectedChar.Font.Underline
                    
                    Select Case myStyle
                        Case "bold (b)"
                            If Not B Or i Or sc Then Call AppliedCharStylesHelper(selectedChar)
                        Case "ital (i)"
                           If Not i Or B Or sc Or sup Then Call AppliedCharStylesHelper(selectedChar)
                        Case "smallcaps (sc)"
                            If Not sc Or B Or i Then Call AppliedCharStylesHelper(selectedChar)
                        Case "underline (u)"
                            If Not u Then Call AppliedCharStylesHelper(selectedChar)
                        Case "super (sup)"
                            If Not sup Or i Then Call AppliedCharStylesHelper(selectedChar)
                        Case "sub (sub)"
                            If Not subs Then Call AppliedCharStylesHelper(selectedChar)
                        Case "bold-ital (bi)"
                            If sc Or Not B Or Not i Then Call AppliedCharStylesHelper(selectedChar)
                        Case "smallcaps-ital (sci)"
                            If B Or Not sc Or Not i Then Call AppliedCharStylesHelper(selectedChar)
                        Case "smallcaps-bold (scb)"
                            If i Or Not sc Or Not B Then Call AppliedCharStylesHelper(selectedChar)
                        Case "smallcaps-bold-ital (scbi)"
                            If Not sc Or Not B Or Not i Then Call AppliedCharStylesHelper(selectedChar)
                        Case "super-ital (supi)"
                            If Not sup Or Not i Then Call AppliedCharStylesHelper(selectedChar)
                        Case "strike (str)"
                            If Not strk Then Call AppliedCharStylesHelper(selectedChar)
                    End Select
                Next k
                
                ' \/ strip out any direct formatting beyond our determined cstyle
                Selection.ClearCharacterDirectFormatting
                ' /\ Adding this in per RST-1231
                
                ' \/ this Collapse assures that the selection keeps moving forward & all items are found
                Selection.Collapse Direction:=wdCollapseEnd
                If Clean_helpers.EndOfStoryReached(MyStoryNo) Then Exit Do
                ' this prevents getting stuck in a table cell (wdv-359)
                If Selection.Tables.Count <> 0 Then
                    If Clean_helpers.EndofTableCellReached Then
                        Selection.MoveRight Unit:=wdCharacter, Count:=1
                    End If
                End If
                Selection.Find.Execute
            Loop
        Next

   ' Debug.Print Timer - t
    ActiveDocument.UndoClear
    
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub
'As per RST-1231: reverting direct formatting for already char-styled content
' exempting "format" charstyles, to allow 2nd & 3rd subs (localformatting & checkapplied) to do their jobs
' running this before local formatting cleanup so that can remain cs-agnostic
Sub FixAppliedCharStyles(MyStoryNo)

    thisstatus = "Checking Other Applied Character Styles "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

        Application.ScreenUpdating = False

        Dim styleList As Collection
        Dim myStyle As Variant

        If MyStoryNo < 1 Then MyStoryNo = 1

        Set styleList = GetNonFormatCharStyles
        Clean_helpers.ClearSearch

        For Each myStyle In styleList
            ActiveDocument.StoryRanges(MyStoryNo).Select
            Selection.Collapse Direction:=wdCollapseStart

            With Selection.Find
                .style = ActiveDocument.styles(myStyle)
                .Execute
                
            End With

            Do While Selection.Find.Found
            
                Selection.ClearCharacterDirectFormatting

                ' \/ this Collapse assures that the selection keeps moving forward & all items are found
                Selection.Collapse Direction:=wdCollapseEnd
                If Clean_helpers.EndOfStoryReached(MyStoryNo) Then Exit Do
                ' this prevents getting stuck in a table cell (wdv-359)
                If Selection.Tables.Count <> 0 Then
                    If Clean_helpers.EndofTableCellReached Then
                        Selection.MoveRight Unit:=wdCharacter, Count:=1
                    End If
                End If
                Selection.Find.Execute
            Loop
            ActiveDocument.UndoClear
        Next

    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub



Sub CheckSpecialCharactersPC(MyStoryNo)

    thisstatus = "Checking for Special Characters "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

        Dim MyUpdate, FoundSomething As Boolean
        Dim myValue As Integer
        Dim R As Range
        Dim B() As Byte, i As Long, A As Long
        
        Application.ScreenRefresh
        MyUpdate = Application.ScreenUpdating
        
        If MyStoryNo < 1 Then MyStoryNo = 1
        
        Clean_helpers.ClearSearch
        
        For Each R In ActiveDocument.StoryRanges(MyStoryNo).Characters
            B = R.Text ' converts the string to byte array (2 or 4 bytes per character)
            For i = 1 To UBound(B) Step 2            ' 2 bytes per Unicode codepoint
                If B(i) > 0 Then                     ' if AscW > 255
                    A = B(i): A = A * 256 + B(i - 1) ' AscW
                    Select Case A
                        Case &H1FFE To &H2022, &H120 To &H17D, &H2BD To &H2C0, &H2DA: ' Curly Quotes, Dashes, Apostrophes
                            'do nothing
                        Case Else:
                            If R.Italic Then
                                R.style = "symbols-ital (symi)"
                            Else
                                R.style = "symbols (sym)"
                            End If
                    End Select
                End If
            Next
        Next
        
        ' for rst-1283
        CheckSpecialCharacters_supplemental (MyStoryNo)
        
        ActiveDocument.UndoClear
        
        completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
        If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
        Application.ScreenUpdating = MyUpdate
        Selection.HomeKey Unit:=wdStory
End Sub
Sub CheckSpecialCharacters_supplemental(MyStoryNo)
        
    If MyStoryNo < 1 Then MyStoryNo = 1
    
    Clean_helpers.ClearSearch
    Dim activeRng As Range
    Dim supplemental_syms
    
    supplemental_syms = Array(322, 380)
    ' ^ decimal unicode encodings for polish l, z
    
    For Each sup_sym_code In supplemental_syms
        Set activeRng = ActiveDocument.StoryRanges(MyStoryNo)
        With activeRng.Find
            .Text = ChrW(sup_sym_code)
            .MatchWildcards = False
            While .Execute
                If activeRng.Italic Then
                    activeRng.style = "symbols-ital (symi)"
                    activeRng.Collapse wdCollapseEnd
                Else
                    activeRng.style = "symbols (sym)"
                    activeRng.Collapse wdCollapseEnd
                End If
            Wend
        End With
    Next
        
End Sub


Sub NextElement(control As IRibbonControl)
    Call NextElementRoutine
End Sub

Sub NextElementRoutine()

    Application.ScreenUpdating = False
    
    Selection.Move Unit:=wdParagraph, Count:=1
    If Clean_helpers.EndOfDocumentReached = True Then
        MsgBox "End of document reached."
        Exit Sub
    End If
    
    While Selection.style = "Body-Text (Tx)"
        Selection.Move Unit:=wdParagraph, Count:=1
            If Clean_helpers.EndOfDocumentReached = True Then
                MsgBox "End of document reached."
                Selection.GoTo What:=wdGoToBookmark, Name:="\Sel"
                Exit Sub
            End If
    Wend
    
    Selection.GoTo What:=wdGoToBookmark, Name:="\Sel"
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
End Sub



Sub ValidateCharStyles()

    Dim docActive As Document
    Dim strMessage As String
    Dim styleLoop As style
    Dim allStyles() As Variant
    Dim charStyles() As Variant
    Dim badStyles() As Variant
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    allStyles = AllStylesMod.getAllStyles
    
    ReDim Preserve charStyles(4)
    charStyles(0) = "Footnote Reference"
    charStyles(1) = "Endnote Reference"
    charStyles(2) = "Default Paragraph Font"
    charStyles(3) = "Hyperlink"
    
    i = UBound(charStyles)
    For Each s In allStyles
        If Left(s, 1) = LCase(Left(s, 1)) Then
            ReDim Preserve charStyles(i)
            If System.OperatingSystem = "Macintosh" And Right(s, 1) = vbCr Then
                s = Left$(s, Len(s) - 1)
            End If
            charStyles(i) = s
            i = i + 1
        End If
    Next
    
    For Each styleLoop In ActiveDocument.styles
        If styleLoop.InUse = True And styleLoop.Type = wdStyleTypeCharacter Then
            If GetIndex(styleLoop.NameLocal, charStyles, False) = -1 Then
                Dim rng As Integer
                For rng = 1 To 3
                  Dim myRange As Range
                  Set myRange = ActiveDocument.StoryRanges(rng)
                  myRange.Select
                  Selection.Collapse Direction:=wdCollapseStart
                  With Selection.Find
                      .ClearFormatting
                      .Text = ""
                      .style = styleLoop.NameLocal
                      .Wrap = wdFindStop
                      .Execute Format:=True
                   End With
                  
                   Do While Selection.Find.Found = True
                        If GetIndex(styleLoop.NameLocal, removeStyles, False) >= 0 Then
                           Selection.style = wdStyleDefaultParagraphFont
                        ElseIf GetIndex(styleLoop.NameLocal, replaceStyles, True) >= 0 Then
                            Dim y, z
                            y = GetIndex(styleLoop.NameLocal, replaceStyles, True)
                            z = replaceStyles(y)(1)
                            Selection.style = z
                        ElseIf GetIndex(styleLoop.NameLocal, skipStyles, False) >= 0 Then
                            'continue to next
                        Else
                            #If Mac Then
                            #Else
                                 ActiveDocument.ActiveWindow.ScrollIntoView Selection.Range, True
                            #End If
  
                            Application.ScreenUpdating = True
                            For i = 4 To UBound(charStyles)
                                frmReplaceCharSty.cbList.AddItem charStyles(i)
                            Next
                            frmReplaceCharSty.Tag = styleLoop.NameLocal
                            frmReplaceCharSty.Caption = "Invalid Character Style: " + styleLoop.NameLocal
                            frmReplaceCharSty.frRemove.Caption = "Remove " + Chr(34) + styleLoop.NameLocal + Chr(34)
                            frmReplaceCharSty.frReplace.Caption = "Replace " + Chr(34) + styleLoop.NameLocal + Chr(34) + " with"
                            frmReplaceCharSty.cbList.Text = "Select a replacement style..."
                            frmReplaceCharSty.Show
                        End If
                        If endCharCheck Then GoTo EndEarly
                        Application.ScreenUpdating = False
                        Selection.MoveRight
                        If Clean_helpers.EndOfDocumentReached Then Exit Do
                        Selection.Find.Execute
                    Loop
NextIteration:
                Next rng
            End If
        End If
    Next styleLoop
    
EndEarly:
    
    Application.ScreenUpdating = True
    
    Erase replaceStyles
    Erase removeStyles
    Erase skipStyles
    
    If endCharCheck = True Then
        MessageBox Title:="Ending Before Completion", Msg:="Character Style Check has been terminated.", buttonType:=vbOKOnly
        endCharCheck = False
    Else
        MessageBox Title:="Done", Msg:="Character Style Check is complete.", buttonType:=vbOKOnly
    End If
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 5941 Then Resume NextIteration
    
End Sub

Function GetIndex(value, iaList, multiDim As Boolean) As Long
    Dim item As String
    Dim i As Integer
    
    On Error GoTo Handler
     GetIndex = -1
     For i = 0 To UBound(iaList)
      If multiDim Then item = iaList(i)(0) Else item = iaList(i)
      If value = item Then
       GetIndex = i
       Exit For
      End If
     Next i
     
     Exit Function
     
Handler:
    If Err.Number = 9 Then GetIndex = -1
End Function

Sub fixCustomFootnotes()
' borrowed basic framework from: https://answers.microsoft.com/en-us/msoffice/forum
'   /msoffice_word-mso_win10-mso_2019/converting-custom-mark-footnotes-to-automatic
'   /5c5b089a-85ec-42f4-8dc4-49eb199ad1dd

thisstatus = "Fixing Custom Footnote marks "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

Dim i As Long
Dim rng As Range
Dim newNote As Footnote
Set DataObj = New MSForms.DataObject
With ActiveDocument
    For i = 1 To .Footnotes.Count
        ' skip notes already set to auto-increment
        If .Footnotes(i).Reference.Text <> Chr(2) Then
            ' special handling for notes with no text (under 'Else'),
            '   otherwise the range doesn't exist to .Copy
            If .Footnotes(i).Range <> "" Then
                .Footnotes(i).Range.Copy
                DataObj.GetFromClipboard
                Set rng = .Footnotes(i).Reference
                rng.Collapse wdCollapseStart
                .Footnotes(i).Reference.Delete
                ' we add a blank new note so we can paste in
                '   Range with styles instead of just text
                Set newNote = .Footnotes.Add(rng)
                ' has the side effect of changing parastyle to default Note style
                '   but since all notes should be styled like this, that's ok
                newNote.Range.PasteAndFormat wdFormatOriginalFormatting
            Else
                Set rng = .Footnotes(i).Reference
                rng.Collapse wdCollapseStart
                .Footnotes(i).Reference.Delete
                .Footnotes.Add rng, , " "
            End If
        End If
    Next i
End With

completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub
Sub fixCustomEndnotes()
' borrowed basic framework from: https://answers.microsoft.com/en-us/msoffice/forum
'   /msoffice_word-mso_win10-mso_2019/converting-custom-mark-Endnotes-to-automatic
'   /5c5b089a-85ec-42f4-8dc4-49eb199ad1dd

thisstatus = "Fixing Custom Endnote marks "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

Dim i As Long
Dim rng As Range
Dim newNote As Endnote
Set DataObj = New MSForms.DataObject

With ActiveDocument
    For i = 1 To .Endnotes.Count
        ' skip notes already set to auto-increment
        If .Endnotes(i).Reference.Text <> Chr(2) Then
            ' special handling for notes with no text (under 'Else'),
            '   otherwise the range doesn't exist to .Copy
            If .Endnotes(i).Range <> "" Then
                .Endnotes(i).Range.Copy
                DataObj.GetFromClipboard
                Set rng = .Endnotes(i).Reference
                rng.Collapse wdCollapseStart
                .Endnotes(i).Reference.Delete
                ' we add a blank new note so we can paste in
                '   Range with styles instead of just text
                Set newNote = .Endnotes.Add(rng)
                newNote.Range.PasteAndFormat wdFormatOriginalFormatting
            Else
                Set rng = .Endnotes(i).Reference
                rng.Collapse wdCollapseStart
                .Endnotes(i).Reference.Delete
                .Endnotes.Add rng, , " "
            End If
        End If
    Next i
End With

completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub
