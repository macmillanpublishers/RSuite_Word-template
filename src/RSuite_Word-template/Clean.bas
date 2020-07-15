Attribute VB_Name = "Clean"

Sub Ellipses()

        Application.ScreenUpdating = False
                
        thisStatus = "Fixing ellipses "
        Clean_helpers.updateStatus (thisStatus)

        'replace anythign that's already fixed, in case it's run again
        Clean_helpers.FindReplaceSimple ELLIPSIS, "<doneellipsis>"
        'this makes sure all ellipses are consistent
        'ellipsis.dot = ellipsis
        Clean_helpers.FindReplaceSimple "." & ELLIPSIS_SYM, "." & TEMP_ELL
        
        'ellipsis = ellipsis
        Clean_helpers.FindReplaceSimple ELLIPSIS_SYM, TEMP_ELL
        'dot.dot.dot.dot=dot.ellipsis
        Clean_helpers.FindReplaceSimple "....", "." & TEMP_ELL
        
        'dot.space.dot.space.dot.space.dot=dot.ellipsis
        Clean_helpers.FindReplaceSimple ". . . .", "." & TEMP_ELL
        'dot.dot.dot=ellipsis
        Clean_helpers.FindReplaceSimple "...", TEMP_ELL
        
        'dot.space.dot.space.dot=ellipsis
        Clean_helpers.FindReplaceSimple ". . .", TEMP_ELL
        'dot.space.dot.space.dot=ellipsis
        Clean_helpers.FindReplaceSimple TEMP_ELL & "." & aSPACE, "." & TEMP_ELL
        
        'space.dot.tempell=tempell
        Clean_helpers.FindReplaceSimple aSPACE & "." & TEMP_ELL, "." & TEMP_ELL
         'dot.space.dot.space.dot=ellipsis
        Clean_helpers.FindReplaceSimple TEMP_ELL & aSPACE, TEMP_ELL
        
        'fix all double spaces before and after ellipses
        Clean_helpers.FindReplaceComplex aSPACE & "{1,}" & TEMP_ELL, TEMP_ELL, False, True
        Clean_helpers.FindReplaceComplex TEMP_ELL & aSPACE & "{2,}", TEMP_ELL & aSPACE, False, True
    
        ActiveDocument.StoryRanges(MyStoryNo).Select
        With Selection.Find
            .MatchWildcards = False
            .ClearFormatting
            .Execute findText:=TEMP_ELL
        End With
        While Selection.Find.Found
        
            ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
            Selection.MoveLeft unit:=wdCharacter, Count:=2
            Select Case Selection.Text
                Case RTN, vbCr
                    'do nothing
                Case DP, DOQ, SOQ
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight unit:=wdCharacter, Count:=1
                    Selection.TypeText NBSPchar
                Case EMDASH
                    Selection.MoveRight unit:=wdCharacter, Count:=1
                    Selection.TypeText aSPACE
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight unit:=wdCharacter, Count:=1
                    Selection.TypeText NBSPchar
                Case Else
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight unit:=wdCharacter, Count:=1
                    Select Case Selection.Text
                        Case EMDASH
                            Selection.TypeText NBSPchar
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.MoveLeft unit:=wdCharacter, Count:=1
                            Selection.TypeText NBSPchar
                        Case Else
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.MoveLeft unit:=wdCharacter, Count:=1
                            Selection.TypeText NBSPchar
                        End Select
            End Select

            ActiveDocument.Bookmarks("temp").Select
            Selection.MoveRight unit:=wdCharacter, Count:=1
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
    Clean_helpers.FindReplaceSimple TEMP_ELL, ELLIPSIS
    
    'replace anything that's already fixed, in case it's run again
    Clean_helpers.FindReplaceComplex "<doneellipsis>", ELLIPSIS, True, False
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
        
End Sub

Sub Spaces()

    thisStatus = "Fixing spaces "
    Clean_helpers.updateStatus (thisStatus)

    'temporarily chanage finished ellipses and delete nonbreaking spaces
    Clean_helpers.FindReplaceSimple PERIOD_ELLIPSIS, "<doneperiodellipsis>"
    'temporarily chanage finished ellipses and delete nonbreaking spaces
    Clean_helpers.FindReplaceSimple NBS_ELLIPSIS, "<doneellipsis>"
    'nonbreaking space to regular space
    Clean_helpers.FindReplaceComplex ChrW(202), " ", False, True
    Clean_helpers.FindReplaceComplex ChrW(160), " ", False, True
    Clean_helpers.FindReplaceSimple "<doneellipsis>", NBS_ELLIPSIS
    Clean_helpers.FindReplaceSimple "<doneperiodellipsis>", PERIOD_ELLIPSIS
    'multiple tabs to regular space
    Clean_helpers.FindReplaceComplex "^9{1,}", " ", False, True
    'multiple spaces to one space
    Clean_helpers.FindReplaceComplex " {2,}", " ", False, True
    'soft returns to hard returns
    Clean_helpers.FindReplaceSimple "^l", "^p"
    'spaces before/after line breaks
    Clean_helpers.FindReplaceSimple ChrW(13) + " ", "^p"
    Clean_helpers.FindReplaceSimple " " + ChrW(13), "^p"
    Clean_helpers.FindReplaceSimple "^p ", "^p"
    Clean_helpers.FindReplaceSimple " ^p", "^p"
    'space before/after brackets to no space
    Clean_helpers.FindReplaceSimple "( ", "("
    Clean_helpers.FindReplaceSimple "[ ", "["
    Clean_helpers.FindReplaceSimple "{ ", "{"
    Clean_helpers.FindReplaceSimple " )", ")"
    Clean_helpers.FindReplaceSimple " ]", "]"
    Clean_helpers.FindReplaceSimple " }", "}"
    'space after dollar sign to no space
    Clean_helpers.FindReplaceSimple "$ ", "$"
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
    
End Sub

Sub Punctuation()

    thisStatus = "Fixing punctuation "
    Clean_helpers.updateStatus (thisStatus)

    'multiple periods to single period
    Clean_helpers.FindReplaceComplex ".{2,}", ".", False, True
    'multiple commas to single comma
    Clean_helpers.FindReplaceComplex ",{2,}", ",", False, True
    'optional hyphen to nothing
    Clean_helpers.FindReplaceSimple OPTHYPH, ""
    Clean_helpers.FindReplaceSimple OPTHYPH2, ""
    'non-breaking hyphen to regular hyphen
    Clean_helpers.FindReplaceSimple NBHYPH, "-"
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
End Sub

Sub DoubleQuotes()
            
    Application.ScreenUpdating = False
    ActiveDocument.StoryRanges(MyStoryNo).Select
    
    Dim totalPages, currentPage, nextPercentage As Integer
    Dim currPercentage, newPercentage As Integer
    ActiveDocument.Repaginate
    totalPages = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
    currPercentage = 0
    
    thisStatus = "Fixing double quotes"
    Clean_helpers.updateStatus (thisStatus)

    ' Combine double single-primes into Double-prime, also double-backticks
    FindReplaceSimple SP & SP, DP
    FindReplaceSimple "``", DP
    
    ActiveDocument.StoryRanges(MyStoryNo).Select
    Selection.Find.Execute findText:=DP
    Do While Selection.Find.Found
        ' Find / Replace tool includes DOQ and DCQ as results in a search for DP
        '   for some reason (Windows/Office2013)
        '   we can filter them out here
        If Selection.Text = DP Then
            
            newPercentage = Selection.Range.Information(wdActiveEndPageNumber) / totalPages * 100
            If newPercentage > currPercentage Then
                thisStatus = "Fixing double quotes: " & CStr(newPercentage) & "%"
                Clean_helpers.updateStatus (thisStatus)
                currPercentage = newPercentage
            End If
             
            ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
            Selection.MoveLeft unit:=wdCharacter, Count:=2
            Select Case Selection.Text
                Case EMDASH
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight unit:=wdCharacter, Count:=1
                    Select Case Selection.Text
                        Case " ", vbCr
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DCQ
                        Case Else
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DOQ
                    End Select
                Case " "
                    Debug.Print Selection.Text + "space case"
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.MoveRight unit:=wdCharacter, Count:=1
                    Selection.Expand unit:=wdCharacter
                    Select Case Selection.Text
                        Case " ", vbCr
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DCQ
                        Case SP
                            Selection.TypeText SOQ
                            Selection.Expand unit:=wdCharacter
                            Select Case Selection.Text
                                Case DP
                                    Selection.TypeText DOQ
                            End Select
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DOQ
                        Case Else
                            ActiveDocument.Bookmarks("temp").Select
                            Selection.TypeText DOQ
                    End Select
                Case vbCr, vbTab, "("
                    Debug.Print Selection.Text + "tab"
                    ActiveDocument.Bookmarks("temp").Select
                    Selection.TypeText DOQ
                    Selection.Expand unit:=wdCharacter
                    Select Case Selection.Text
                        Case SP
                            Selection.TypeText SOQ
                            Selection.Expand unit:=wdCharacter
                            Select Case Selection.Text
                                Case DP
                                    Selection.TypeText DOQ
                            End Select
                    End Select
                Case Else
                    If Clean_helpers.AtStartOfDocument Then
                        ActiveDocument.Bookmarks("temp").Select
                        Selection.TypeText DOQ
                    Else
                        ActiveDocument.Bookmarks("temp").Select
                        Selection.TypeText DCQ
                    End If
                    Selection.MoveLeft unit:=wdCharacter, Count:=2
                    Select Case Selection.Text
                        Case SP
                            Selection.Delete
                            Selection.TypeText SCQ
                            Selection.MoveLeft unit:=wdCharacter, Count:=2
                            Select Case Selection.Text
                                Case DP
                                    Selection.Delete
                                    Selection.TypeText DCQ
                            End Select
                    End Select
                End Select
            Selection.MoveRight unit:=wdCharacter, Count:=3
       End If
            If Clean_helpers.EndOfDocumentReached Then Exit Do
            Selection.Find.Execute
        
    Loop
    
    completeStatus = completeStatus + vbNewLine + "Fixing double quotes: 100%"
    Clean_helpers.updateStatus ("")

End Sub

Sub SingleQuotes()

    Application.ScreenUpdating = False
    
    Dim nextPercentage As Integer
    nextPercentage = 30
    
    thisStatus = "Fixing single quotes "
    Clean_helpers.updateStatus (thisStatus)

    Dim ChangeQ As Boolean
    ChangeQ = False
     
    ActiveDocument.StoryRanges(MyStoryNo).Select
    Selection.Find.ClearFormatting
    Selection.Find.Execute findText:="`"
    While Selection.Find.Found
        ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
        Selection.MoveLeft unit:=wdCharacter, Count:=2
        If Selection.Text = " " Or Selection.Text = vbCr Or Selection.Text = "(" Then
            ActiveDocument.Bookmarks("temp").Select
            Selection.TypeText SOQ
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
    '   leaving the capability to search SOQ/SCQ via this array setup in csae we end up reversing/
    SearchString = Array(SP) ', SOQ, SCQ)
           
    ActiveDocument.StoryRanges(MyStoryNo).Select
    For Each QuoStr In SearchString
        
        thisStatus = "Fixing single quotes: " & CStr(nextPercentage) & "%"
        Clean_helpers.updateStatus (thisStatus)
        nextPercentage = nextPercentage + 30
    
        Selection.Find.Execute findText:=QuoStr
        
        While Selection.Find.Found
            ' Find / Replace tool includes DOQ and DCQ as results in a search for DP
            '   for some reason (Windows/Office2013)
            '   we can filter them out here
            If Selection.Text = SP Then
            
                ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
                Selection.MoveLeft unit:=wdCharacter, Count:=2
                
                Select Case Selection.Text
                        Case DP, DOQ, SOQ
                            OpenQuo = True
                End Select
                
                Select Case Selection.Text
                        Case " ", vbCr, vbTab, vbNewLine, "(", DP, DOQ, SOQ
                            Selection.MoveRight unit:=wdCharacter, Count:=2
                            Selection.ExtendMode = True
                            
                            '1 character
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case DOQ, "K", "k"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '2 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
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
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Cuz", "cuz", "Net", "net", "Sup", "sup", "Tar", "tar", "Til", "til", "Tis", "tis"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '4 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Bout", "bout", "Cept", "cept", "Fore", "fore", "Nuff", "nuff", "Post", "post", "Tall", "tall", "Twas", "twas"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '5 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Cause", "cause", "Fraid", "fraid", "Night", "night", "Round", "round", "Scuse", "scuse", "Sides", "sides", "Spect", "spect", "Tever", "tever"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
        
                            '6 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Course", "course", "Gainst", "gainst", "Nother", "nother", "Splain", "splain", "Tain" & SCQ & "t", "tain" & SCQ & "t", "Tisn" & SCQ & "t", "tisn" & SCQ & "t"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '7 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Chother", "chother", "Druther", "druther", "Salmost", "salmost", "Snothin", "snothin", "Twasn" & SCQ & "t", "twasn" & SCQ & "t"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '8 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Druthers", "druthers", "Tweren" & SCQ & "t", "tweren" & SCQ & "t"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '9 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            If LookAhead() = True Then
                                Select Case Selection.Text
                                    Case "Specially", "specially", "Spossible", "spossible"
                                        ChangeQ = True
                                End Select
                                GoTo SkipToHere
                            End If
                            
                            '10 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
                            'If LookAhead() = True Then
                                'Select Case Selection.Text
                                    'Case
                                        'ChangeQ = True
                                'End Select
                                'GoTo SkipToHere
                            'End If
                                
                            '11 characters
                            Selection.MoveRight unit:=wdCharacter, Count:=1
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
FindReplaceSimple SOQ & aSPACE & DOQ, SOQ & DOQ
FindReplaceSimple DOQ & aSPACE & SOQ, DOQ & SOQ
FindReplaceSimple SCQ & aSPACE & DCQ, SCQ & DCQ
FindReplaceSimple DOQ & aSPACE & SCQ, DOQ & SCQ
FindReplaceSimple DCQ & aSPACE & SCQ, DCQ & SCQ
'FindReplaceSimple SCQ & aSPACE & DOQ, SCQ & DOQ
'FindReplaceSimple DCQ & aSPACE & SOQ, DCQ & SOQ
'FindReplaceSimple SOQ & aSPACE & DCQ, SOQ & DCQ

completeStatus = completeStatus + vbNewLine + "Fixing Single Quotes: 100%"
Clean_helpers.updateStatus ("")

End Sub

Function LookAhead() As Boolean

    ActiveDocument.Bookmarks.Add ("myTemp")
    Selection.ExtendMode = False
    Selection.MoveRight unit:=wdCharacter, Count:=1
    Select Case Selection.Text
        Case " ", ".", ",", "?", "!", EMDASH, ")"
            LookAhead = True
        Case Else
            LookAhead = False
    End Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.ExtendMode = True
    ActiveDocument.Bookmarks("myTemp").Select
    ActiveDocument.Bookmarks("myTemp").Delete

End Function

Function IsYear(theNumber) As Boolean

    If theNumber Like "[0-9][0-9]" Then
    
            ActiveDocument.Bookmarks.Add ("myTemp")
            Selection.ExtendMode = False
            Selection.MoveRight unit:=wdCharacter, Count:=1
            Select Case Selection.Text
                Case " ", ".", ",", "?", "!", EMDASH, ")", "s"
                    IsYear = True
                Case Else
                    IsYear = False
            End Select
            Selection.MoveLeft unit:=wdCharacter, Count:=1
            Selection.ExtendMode = True
            ActiveDocument.Bookmarks("myTemp").Select
            ActiveDocument.Bookmarks("myTemp").Delete
            
    End If

End Function

Sub Dashes()

    thisStatus = "Fixing dashes "
    Clean_helpers.updateStatus (thisStatus)

    Application.ScreenUpdating = False
    
     'phone number pattern
     Call HighlightNumber("[0-9]{3}-[0-9]{3}-[0-9]{4}")
     Call HighlightNumber("\([0-9]{3}\) [0-9]{3}-[0-9]{4}")
    
'    FOLLOWING CAN BE USED TO FIND ISBN PATTERN AND FLAG FOR NO CHANGE
     Call HighlightNumber("97[89]-[0-9]{10,14}")
     Call HighlightNumber("97[89]-[0-9]-[0-9]{3}-[0-9]{5}-[0-9]")
     
    thisStatus = "Fixing dashes: 10%"
    Clean_helpers.updateStatus (thisStatus)
     
    For i = 0 To 9
        For J = 0 To 9
            ActiveDocument.StoryRanges(MyStoryNo).Select
            Selection.Collapse direction:=wdCollapseStart
            
            With Selection.Find
                 .ClearFormatting
                 .Forward = True
                 .Wrap = wdFindStop
                 .Text = LTrim(i) & "-" & LTrim(J)
                 .MatchWildcards = False
                 .Execute
             End With
             
             While Selection.Find.Found
                 If Not (Selection.FormattedText.HighlightColorIndex = wdPink) Then
                     Selection.TypeText LTrim(i) & ENDASH & LTrim(J)
                 End If
                 
                 Selection.MoveRight
                 Selection.Find.Execute
             Wend
        Next
    Next
    
    thisStatus = "Fixing dashes: 20%"
    Clean_helpers.updateStatus (thisStatus)

    'weird-character = emdash
    FindReplaceSimple ChrW(-3906), EMDASH
    'bar character = emdash
    FindReplaceSimple ChrW(8213), EMDASH
    
    thisStatus = "Fixing dashes: 30%"
    Clean_helpers.updateStatus (thisStatus)
    
    'figure dash=endash
    FindReplaceSimple ChrW(8210), ENDASH
    'hyphen.hyphen.hyphen=endash
    FindReplaceSimple "---", EMDASH
    'space.hyphen.space=emdash
    FindReplaceSimple " - ", "-"
    
    thisStatus = "Fixing dashes: 40%"
    Clean_helpers.updateStatus (thisStatus)
    
    'space.hyphen.hyphen.space=emdash
    FindReplaceSimple " -- ", EMDASH
    'hyphen.hyphen=emdash
    FindReplaceSimple "--", EMDASH
    
    thisStatus = "Fixing dashes: 50%"
    Clean_helpers.updateStatus (thisStatus)
    
   'dash.space=dash
    FindReplaceSimple "-" & aSPACE, "-"
    'space.dash=dash
    FindReplaceSimple aSPACE & "-", "-"
    
    thisStatus = "Fixing dashes: 60%"
    Clean_helpers.updateStatus (thisStatus)
    
    'space.endash=emdash
    FindReplaceSimple aSPACE & ENDASH, EMDASH
    'endash.space=emdash
    FindReplaceSimple ENDASH & aSPACE, ENDASH
    
    thisStatus = "Fixing dashes: 70%"
    Clean_helpers.updateStatus (thisStatus)
    
    'emdash.space=emdash
    FindReplaceSimple EMDASH & aSPACE, EMDASH
    'space.emdash=emdash
    FindReplaceSimple aSPACE & EMDASH, EMDASH
    
    thisStatus = "Fixing dashes: 80%"
    Clean_helpers.updateStatus (thisStatus)
    
    Call removeHighlight
    
    thisStatus = "Fixing dashes: 90%"
    Clean_helpers.updateStatus (thisStatus)
    
    completeStatus = completeStatus + vbNewLine + "Fixing Dashes: 100%"
    Clean_helpers.updateStatus ("")
    
End Sub

Function HighlightNumber(myPattern)
    
    Selection.HomeKey unit:=wdStory
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
        If Clean_helpers.EndOfDocumentReached Then Exit Do
        Selection.Find.Execute
    Loop

End Function

Function removeHighlight()

    Options.DefaultHighlightColorIndex = wdPink

    Selection.HomeKey unit:=wdStory
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

Function MakeTitleCase()

    thisStatus = "Converting headings to title case "
    Clean_helpers.updateStatus (thisStatus)

    If MyStoryNo = 0 Then MyStoryNo = 1
    
    Dim tcStyles() As Variant
    tcStyles = Array("Title (Ttl)", "Number (Num)", "Main-Head (MHead)")
    
    For Each TC In tcStyles
        Clean_helpers.ClearSearch
        
        ActiveDocument.StoryRanges(MyStoryNo).Select
        Selection.Collapse direction:=wdCollapseStart
    
        With Selection.Find
            .Wrap = wdFindStop
            .Style = TC
            .Execute
        End With
        
        Do While Selection.Find.Found
            Clean_helpers.TitleCase
            Selection.MoveRight
            If Clean_helpers.EndOfDocumentReached Then Exit Do
            Selection.Find.Execute
        Loop
    Next
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")

End Function

Function CleanBreaks()

    thisStatus = "Cleaning breaks "
    Clean_helpers.updateStatus (thisStatus)

    FindReplaceSimple "^l", "^p"
    FindReplaceSimple "^m", "^p"
    FindReplaceSimple "^b", "^p"
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute
    End With
    
    Do While Selection.Find.Found
        If EndOfDocumentReached = False Then
            FindReplaceSimple "^p^p", "^p"
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "^p^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Execute
            End With
        Else
            Exit Do
        End If
    Loop
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
End Function

Function RemoveTrackChanges()

    thisStatus = "Removing Track Changes "
    Clean_helpers.updateStatus (thisStatus)
    
    If ActiveDocument.Revisions.Count > 0 Then
        If Clean_helpers.MessageBox("ACCEPT TRACK CHANGES", "Your document contains unacccepted Track Changes, which must be removed before the file is transformed in RSuite." & vbNewLine & vbNewLine & _
          "Select YES to accept all changes in the document." & vbNewLine & vbNewLine & _
          "Select NO to retain Track Changes.") = vbYes Then
                ActiveDocument.Revisions.AcceptAll
        End If
    End If
    
End Function

Function RemoveComments()

    thisStatus = "Removing Comments "
    Clean_helpers.updateStatus (thisStatus)
    
    Dim c As Comment
    If ActiveDocument.Comments.Count > 0 Then
        If Clean_helpers.MessageBox("DELETE COMMENTS", "Your document contains Comments, which must be removed before the file is transformed in RSuite." & vbNewLine & vbNewLine & _
            "Select YES to remove all comments in the document." & vbNewLine & vbNewLine & _
            "Select NO to retain comments.") = vbYes Then
                ActiveDocument.DeleteAllComments
        End If
    End If
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
    
End Function

Function DeleteBookmarks()

    thisStatus = "Deleting Bookmarks "
    Clean_helpers.updateStatus (thisStatus)
    
    Dim b As Bookmark
    For Each b In ActiveDocument.Bookmarks
        b.Delete
    Next
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
End Function

Function DeleteObjects()

    thisStatus = "Deleting Objects "
    Clean_helpers.updateStatus (thisStatus)

    Dim s As Shape
    Dim i As InlineShape
    Dim F As Frame
    Dim R As Range
    Dim G As Integer
    Dim TB As TextFrame
    
    For Each s In ActiveDocument.Shapes
        If s.Type = msoTextBox Then
            s.Anchor.Select
            Selection.MoveLeft unit:=wdCharacter
            Selection.MoveDown unit:=wdParagraph
            Selection.TypeText s.TextFrame.TextRange.Text
            s.Delete
        ElseIf s.Type = msoGroup Then
            For G = 1 To s.GroupItems.Count
                If s.GroupItems(G).Type = 17 Then
                    Set TB = s.GroupItems(G).TextFrame
                    s.Anchor.Select
                    Selection.MoveLeft unit:=wdCharacter
                    Selection.MoveDown unit:=wdParagraph
                    Selection.TypeText TB.TextRange.Text
                End If
            Next G
            s.Delete
        Else
            s.Delete
        End If
    Next
    
    For Each i In ActiveDocument.InlineShapes
        i.Delete
    Next
    
    For Each F In ActiveDocument.Frames
        F.Delete
    Next
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
End Function

Function RemoveHyperlinks()

    thisStatus = "Removing hyperlinks "
    Clean_helpers.updateStatus (thisStatus)
    
    Dim H As hyperlink
    For Each H In ActiveDocument.Hyperlinks
        H.Range.Style = "Hyperlink"
        H.Delete
    Next
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")

End Function

Sub LocalFormatting()

    thisStatus = "Replacing Local Formatting with Character Styles "
    Clean_helpers.updateStatus (thisStatus)

    Application.ScreenUpdating = False
    
    'small caps bold italic
    Call ConvertLocalFormatting(SmallCapsTF:=True, ItalTF:=True, BoldTF:=True, NewStyle:="smallcaps-bold-ital (scbi)")
    
    'bold ital
    Call ConvertLocalFormatting(ItalTF:=True, BoldTF:=True, NewStyle:="bold-ital (bi)")
    
    'small caps bold
    Call ConvertLocalFormatting(SmallCapsTF:=True, BoldTF:=True, NewStyle:="smallcaps-bold (scb)")
    
    'small caps ital
    Call ConvertLocalFormatting(SmallCapsTF:=True, ItalTF:=True, NewStyle:="smallcaps-ital (sci)")
    
    'strikethrough
    Call ConvertLocalFormatting(StrikeTF:=True, NewStyle:="strike (str)")
    
    'superscript italic
    Call ConvertLocalFormatting(superTF:=True, ItalTF:=True, NewStyle:="super-ital (supi)")
    
    'superscript
    Call ConvertLocalFormatting(superTF:=True, NewStyle:="super (sup)")
    
    'subscript
    Call ConvertLocalFormatting(subTF:=True, NewStyle:="sub (sub)")
    
    'ital
    Call ConvertLocalFormatting(ItalTF:=True, NewStyle:="ital (i)")
    
    'bold
    Call ConvertLocalFormatting(BoldTF:=True, NewStyle:="bold (b)")

    'small caps
    Call ConvertLocalFormatting(SmallCapsTF:=True, NewStyle:="smallcaps (sc)")
        
    'underline
    Call ConvertLocalFormatting(UnderlineTF:=True, NewStyle:="underline (u)")
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
    
    Application.ScreenUpdating = True

End Sub

Sub CheckAppliedCharStyles()

    thisStatus = "Checking Applied Character Styles "
    Clean_helpers.updateStatus (thisStatus)

        Application.ScreenUpdating = False
        
        Dim styleList() As Variant
        Dim defaultStyle As Variant
        Dim b, i, sc, subs, sup, strk, u As Boolean
        
        defaultStyle = WdBuiltinStyle.wdStyleDefaultParagraphFont
        
        If MyStoryNo < 1 Then MyStoryNo = 1
        
        styleList = Array("bold (b)", "ital (i)", "smallcaps (sc)", "underline (u)", "super (sup)", "sub (sub)", _
                          "bold-ital (bi)", "smallcaps-ital (sci)", "smallcaps-bold (scb)", _
                          "smallcaps-bold-ital (scbi)", "super-ital (supi)", "strike (str)")
        
        Clean_helpers.ClearSearch
        
        For Each MyStyle In styleList
            ActiveDocument.StoryRanges(MyStoryNo).Select
            Selection.Collapse direction:=wdCollapseStart
        
            With Selection.Find
                .Style = ActiveDocument.Styles(MyStyle)
                .Execute
            End With
            
            Do While Selection.Find.Found
                
                b = Selection.Font.Bold
                i = Selection.Font.Italic
                sc = Selection.Font.SmallCaps
                subs = Selection.Font.Subscript
                sup = Selection.Font.Superscript
                strk = Selection.Font.StrikeThrough
                u = Selection.Font.Underline
                
                Select Case MyStyle
                
                    Case "bold (b)"
                        If Not b Then Selection.Style = defaultStyle
                        
                    Case "ital (i)"
                        If Not i Then Selection.Style = defaultStyle
                        
                    Case "smallcaps (sc)"
                        If Not sc Then Selection.Style = defaultStyle
                        
                    Case "underline (u)"
                        If Not u Then Selection.Style = defaultStyle
                        
                    Case "super (sup)"
                        If Not sup Then Selection.Style = defaultStyle
                        
                    Case "sub (sub)"
                        If Not subs Then Selection.Style = defaultStyle
                    
                    Case "bold-ital (bi)"
                        If Not b And Not i Then
                            Selection.Style = defaultStyle
                        ElseIf Not b Then
                            Selection.Range.Style = "ital (i)"
                        ElseIf Not i Then
                            Selection.Range.Style = "bold (b)"
                        End If
                        
                    Case "smallcaps-ital (sci)"
                        If Not sc And Not i Then
                            Selection.Style = defaultStyle
                        ElseIf Not sc Then
                            Selection.Range.Style = "ital (i)"
                        ElseIf Not i Then
                            Selection.Range.Style = "smallcaps (sc)"
                        End If
                    
                    Case "smallcaps-bold (scb)"
                        If Not sc And Not b Then
                            Selection.Style = defaultStyle
                        ElseIf Not sc Then
                            Selection.Range.Style = "bold (b)"
                        ElseIf Not b Then
                            Selection.Range.Style = "smallcaps (sc)"
                        End If
                    
                    Case "smallcaps-bold-ital (scbi)"
                        If Not sc And Not b And Not i Then
                            Selection.Style = defaultStyle
                        ElseIf Not sc And Not i Then
                            Selection.Range.Style = "bold (b)"
                        ElseIf Not sc And Not b Then
                            Selection.Range.Style = "ital (i)"
                        ElseIf Not sc Then
                            Selection.Range.Style = "bold-ital (bi)"
                        ElseIf Not b Then
                            Selection.Range.Style = "smallcaps-ital (sci)"
                        ElseIf Not i Then
                            Selection.Range.Style = "smallcaps-bold (scb)"
                        End If
                    
                    Case "super-ital (supi)"
                        If Not sup And Not i Then
                            Selection.Style = defaultStyle
                        ElseIf Not sup Then
                            Selection.Range.Style = "ital (i)"
                        ElseIf Not i Then
                            Selection.Range.Style = "super (sup)"
                        End If
                    
                    Case "strike (str)"
                        If Not strk Then Selection.Style = defaultStyle
                        
                End Select
                'Selection.ClearCharacterDirectFormatting
                Selection.MoveRight
                If Clean_helpers.EndOfDocumentReached Then Exit Do
                Selection.Find.Execute
            Loop
        Next
        
    completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
    Clean_helpers.updateStatus ("")
        
End Sub


Sub CheckSpecialCharactersPC()

    thisStatus = "Checking for Special Characters "
    Clean_helpers.updateStatus (thisStatus)

        Dim MyUpdate, FoundSomething As Boolean
        Dim myValue As Integer
        Dim R As Range
        Dim b() As Byte, i As Long, a As Long
        
        Application.ScreenRefresh
        MyUpdate = Application.ScreenUpdating
        
        If MyStoryNo < 1 Then MyStoryNo = 1
        
        Clean_helpers.ClearSearch
        
        For Each R In ActiveDocument.StoryRanges(MyStoryNo).Characters
            b = R.Text ' converts the string to byte array (2 or 4 bytes per character)
            For i = 1 To UBound(b) Step 2            ' 2 bytes per Unicode codepoint
                If b(i) > 0 Then                     ' if AscW > 255
                    a = b(i): a = a * 256 + b(i - 1) ' AscW
                    Select Case a
                        Case &H1FFE To &H2022, &H120 To &H17D, &H2BD To &H2C0, &H2DA: ' Curly Quotes, Dashes, Apostrophes
                            'do nothing
                        Case Else:
                            If R.Italic Then
                                R.Style = "symbols-ital (symi)"
                            Else
                                R.Style = "symbols (sym)"
                            End If
                    End Select
                End If
            Next
            
        Next
        
        completeStatus = completeStatus + vbNewLine + thisStatus + "100%"
        Clean_helpers.updateStatus ("")
    
        Application.ScreenUpdating = MyUpdate
        Selection.HomeKey unit:=wdStory
End Sub


Sub NextElement(control As IRibbonControl)
    Call NextElementRoutine
End Sub

Sub NextElementRoutine()

    Application.ScreenUpdating = False
    
    Selection.Move unit:=wdParagraph, Count:=1
    If Clean_helpers.EndOfDocumentReached = True Then
        MsgBox "End of document reached."
        Exit Sub
    End If
    
    While Selection.Style = "Body-Text (Tx)"
        Selection.Move unit:=wdParagraph, Count:=1
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
    Dim styleLoop As Style
    Dim allStyles() As Variant
    Dim charStyles() As Variant
    Dim badStyles() As Variant
    Dim i As Integer
    
    On Error GoTo errHandler
    
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
    
    For Each styleLoop In ActiveDocument.Styles
        If styleLoop.InUse = True And styleLoop.Type = wdStyleTypeCharacter Then
            If GetIndex(styleLoop.NameLocal, charStyles, False) = -1 Then
                Dim rng As Integer
                For rng = 1 To 3
                  Dim myRange As Range
                  Set myRange = ActiveDocument.StoryRanges(rng)
                  myRange.Select
                  Selection.Collapse direction:=wdCollapseStart
                  With Selection.Find
                      .ClearFormatting
                      .Text = ""
                      .Style = styleLoop.NameLocal
                      .Wrap = wdFindStop
                      .Execute Format:=True
                   End With
                  
                   Do While Selection.Find.Found = True
                        If GetIndex(styleLoop.NameLocal, removeStyles, False) >= 0 Then
                           Selection.Style = wdStyleDefaultParagraphFont
                        ElseIf GetIndex(styleLoop.NameLocal, replaceStyles, True) >= 0 Then
                            Dim y, z
                            y = GetIndex(styleLoop.NameLocal, replaceStyles, True)
                            z = replaceStyles(y)(1)
                            Selection.Style = z
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
    
errHandler:
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


