Attribute VB_Name = "TestHelpers"
Function strFromLocation(myDoc As Document, paraNumber As Long, char_offset As Long, returnStr_length As Long) As String
    Dim myRange As Range
    
    Set myRange = myDoc.Paragraphs(paraNumber).Range
    myRange.Collapse Direction:=wdCollapseStart
    With myRange
        .Move Unit:=wdCharacter, Count:=char_offset
        .MoveEnd Unit:=wdCharacter, Count:=returnStr_length
    End With
    
    strFromLocation = myRange.Text
End Function
Function smallCapsCheckFromLocation(myDoc As Document, paraNumber As Long, char_offset As Long, returnStr_length As Long) As Boolean
' for CIP tag insertions, to verify they are not picking up local styling (that could override the tag charCase when converted to txt)
    Dim myRange As Range
    
    Set myRange = myDoc.Paragraphs(paraNumber).Range
    myRange.Collapse Direction:=wdCollapseStart
    With myRange
        .Move Unit:=wdCharacter, Count:=char_offset
        .MoveEnd Unit:=wdCharacter, Count:=returnStr_length
    End With
    
    smallCapsCheckFromLocation = myRange.Font.SmallCaps
End Function
Function stringCountInDoc(myDoc As Document, targetStr As String) As Long
    Dim docTxt As String, newTxt As String
    docTxt = myDoc.Range.Text
    newTxt = Replace(docTxt, targetStr, "")
    stringCountInDoc = ((Len(docTxt) - Len(newTxt)) / Len(targetStr))
End Function
Function lastTablecellText(storyRangeIndex As Long, tableIndex As Long, _
    Optional rowCount As Long = 0, Optional colCount As Long = 0) As String
Dim myTable As Table
Dim cellRange As Range
Set myTable = ActiveDocument.StoryRanges(storyRangeIndex).Tables(tableIndex)
If colCount = 0 Then colCount = myTable.Columns.Count
If rowCount = 0 Then rowCount = myTable.Rows.Count
Set cellRange = myTable.Cell(rowCount, colCount).Range
cellRange.End = cellRange.End - 1
lastTablecellText = cellRange.Text
End Function
Function compareParaStylesInRange(testRange As Range, expectedRange As Range) As String
    Dim i As Long
    Dim resultstr As String
    resultstr = ""
    If testRange.Paragraphs.Count <> expectedRange.Paragraphs.Count Then
        compareParaStylesInRange = "Mismatch: testRange has " & testRange.Paragraphs.Count & _
            " paragraphs and expectedRange has " & expectedRange.Paragraphs.Count
        Exit Function
    End If
    For i = 1 To testRange.Paragraphs.Count
        If testRange.Paragraphs(i).style <> expectedRange.Paragraphs(i).style Then
            If resultstr = "" Then
                resultstr = "Mismatched parastyles found; para number(s): " & i
            Else
                resultstr = resultstr & ", " & i
            End If
        End If
    Next i
    compareParaStylesInRange = resultstr
End Function

Sub testsub()
Debug.Print lastTablecellText(3, 1)
'Debug.Print """" & strFromLocation(ActiveDocument, 25, 0, 4) & """"
End Sub

Function compareRanges(actualRange As Range, expectedRange As Range) As String
Dim returnString As String
Dim i As Integer
returnString = "Same"

If actualRange.Characters.Count <> expectedRange.Characters.Count Then
    returnString = "Compared ranges are different lengths, expected: " + str(expectedRange.Characters.Count) + _
        ", actual: " + str(actualRange.Characters.Count)
    GoTo Theend
ElseIf actualRange.Text <> expectedRange.Text Then
    returnString = "Range text mismatch, expected: '" + actualRange.Text + "', actual: '" + expectedRange.Text + "'"
    GoTo Theend
Else
    For i = 1 To actualRange.Characters.Count
        ' namelocal reports char style where present, otherwise give para style
        If actualRange.Characters(i).style.NameLocal <> expectedRange.Characters(i).style.NameLocal Then
            returnString = "Different styles detected for char #" + str(i) + " ('" + actualRange.Characters(i) + _
                "'), expected: '" + expectedRange.Characters(i).style.NameLocal + "', actual: '" + _
                actualRange.Characters(i).style.NameLocal + "'"
            GoTo Theend
        ' checking local formatting types one by one
        ElseIf actualRange.Characters(i).Font.Bold <> expectedRange.Characters(i).Font.Bold Then
            returnString = "Diff in 'bold' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        ElseIf actualRange.Characters(i).Font.Italic <> expectedRange.Characters(i).Font.Italic Then
            returnString = "Diff in 'italic' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        ElseIf actualRange.Characters(i).Font.SmallCaps <> expectedRange.Characters(i).Font.SmallCaps Then
            returnString = "Diff in 'smallcaps' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        ElseIf actualRange.Characters(i).Font.Subscript <> expectedRange.Characters(i).Font.Subscript Then
            returnString = "Diff in 'Subscript' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        ElseIf actualRange.Characters(i).Font.Superscript <> expectedRange.Characters(i).Font.Superscript Then
            returnString = "Diff in 'Superscript' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        ElseIf actualRange.Characters(i).Font.StrikeThrough <> expectedRange.Characters(i).Font.StrikeThrough Then
            returnString = "Diff in 'strikethrough' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        ElseIf actualRange.Characters(i).Font.Underline <> expectedRange.Characters(i).Font.Underline Then
            returnString = "Diff in 'underline' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo Theend
        End If
    Next i
End If

compareRanges = returnString

Theend:
    compareRanges = returnString

End Function

Function returnTestResultStyle(testProcName, MyStoryNo)
Dim testProcNameStart As String
testProcNameStart = "__" + testProcName + "__^p"
Const testProcNameNext = "^p__"

Dim resultRng As Range, DelStartRange As Range, DelEndRange As Range
Dim FindStartRange As Range, FindEndRange As Range

Set FindStartRange = ActiveDocument.StoryRanges(MyStoryNo)
Set FindEndRange = ActiveDocument.StoryRanges(MyStoryNo)
Set resultRng = ActiveDocument.StoryRanges(MyStoryNo)

With FindStartRange.Find
    .Text = testProcNameStart
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = True
    .MatchWholeWord = True
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        resultRng.Start = FindStartRange.End
        '' \/ helpful for debug
        'Set resultStartRange = FindStartRange
        'resultStartRange.Select
        resultRng.Start = FindStartRange.End
        ' start the next search at this point \/
        FindEndRange.Start = FindStartRange.End
   End If
End With

With FindEndRange.Find
    .Text = testProcNameNext
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        'FindEndRange.Select ' < helpful for debug
        resultRng.End = FindEndRange.Start
    Else
        'if we are at end of story, that is the end of range ...
        resultRng.End = ActiveDocument.StoryRanges(MyStoryNo).End
        ' ... minus the trailing vbcr
        resultRng.MoveEnd Unit:=wdWord, Count:=-1
        'Debug.Print resultRng ' < for debug
    End If
End With

returnTestResultStyle = resultRng.style.NameLocal

End Function

Function returnTestResultString(testProcName, MyStoryNo)

Dim testProcNameStart As String
testProcNameStart = "__" + testProcName + "__^p"
Const testProcNameNext = "^p__"

Dim resultRng As Range, DelStartRange As Range, DelEndRange As Range
Dim FindStartRange As Range, FindEndRange As Range

Set FindStartRange = ActiveDocument.StoryRanges(MyStoryNo)
Set FindEndRange = ActiveDocument.StoryRanges(MyStoryNo)
Set resultRng = ActiveDocument.StoryRanges(MyStoryNo)

With FindStartRange.Find
    .Text = testProcNameStart
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = True
    .MatchWholeWord = True
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        resultRng.Start = FindStartRange.End
        '' \/ helpful for debug
        'Set resultStartRange = FindStartRange
        'resultStartRange.Select
        resultRng.Start = FindStartRange.End
        ' start the next search at this point \/
        FindEndRange.Start = FindStartRange.End
   End If
End With

With FindEndRange.Find
    .Text = testProcNameNext
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        'FindEndRange.Select ' < helpful for debug
        resultRng.End = FindEndRange.Start
    ElseIf MyStoryNo = 3 Then
        resultRng.End = FindEndRange.Endnotes(1).Range.End
    Else
        'if we are at end of story, that is the end of range ...
        resultRng.End = ActiveDocument.StoryRanges(MyStoryNo).End
        ' ... minus the trailing vbcr
        resultRng.MoveEnd Unit:=wdWord, Count:=-1
    End If
    'Debug.Print resultRng ' < for debug
End With

returnTestResultString = resultRng

End Function

Function returnTestResultRange(testProcName, MyStoryNo, testDocument)

Dim testProcNameStart As String
testProcNameStart = "__" + testProcName + "__^p"
Const testProcNameNext = "^p__"

Dim resultRng As Range, DelStartRange As Range, DelEndRange As Range
Dim FindStartRange As Range, FindEndRange As Range

Set FindStartRange = testDocument.StoryRanges(MyStoryNo)
Set FindEndRange = testDocument.StoryRanges(MyStoryNo)
Set resultRng = testDocument.StoryRanges(MyStoryNo)

With FindStartRange.Find
    .Text = testProcNameStart
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = True
    .MatchWholeWord = True
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        resultRng.Start = FindStartRange.End
        '' \/ helpful for debug
        'Set resultStartRange = FindStartRange
        'resultStartRange.Select
        resultRng.Start = FindStartRange.End
        ' start the next search at this point \/
        FindEndRange.Start = FindStartRange.End
   End If
End With

With FindEndRange.Find
    .Text = testProcNameNext
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        'FindEndRange.Select ' < helpful for debug
        resultRng.End = FindEndRange.Start
    Else
        'if we are at end of story, that is the end of range ...
        resultRng.End = testDocument.StoryRanges(MyStoryNo).End
        ' ... minus the trailing vbcr
        resultRng.MoveEnd Unit:=wdWord, Count:=-1
    End If
    'Debug.Print resultRng ' < for debug
End With

Set returnTestResultRange = resultRng

End Function

Function getRepoPath() As String
Dim vbProj As VBIDE.VBProject
Dim strDoc As Variant
Dim i As Long
i = 0
Dim gitrepo_anchorfile As String
Dim repo_pathStr As String

gitrepo_anchorfile = "devSetup"

For Each vbProj In Application.VBE.VBProjects   'Loop through each project
    If Not vbProj.Description = "" Then
        If InStr(vbProj.fileName, gitrepo_anchorfile) Then
            repo_pathStr = Left(vbProj.fileName, InStrRev(vbProj.fileName, "\"))
        End If
    End If
    i = i + 1
Next vbProj

getRepoPath = repo_pathStr  'includes trailing backslash
End Function

Sub copyBodyContentsToEndNotes(Optional firstNote As Boolean = False, Optional lastNote As Boolean = False)
Dim mainContentsRng As Range
Dim mainRngEnd As Range

' added optional params for wdv_396, where we need the second to last note to have the body/test-content

' add trailing 'end' tag to main doc
Set mainContentsRng = ActiveDocument.StoryRanges(1)
If firstNote = False Then
    mainContentsRng.InsertAfter (vbCr + "__END_TESTS__")
End If
' paste contents of main doc into the note (minus trailing note ref and vbcr)
mainContentsRng.MoveEnd Unit:=wdCharacter, Count:=-2

' add an empty note referencing last char of main doc
Set mainRngEnd = ActiveDocument.StoryRanges(1)
mainRngEnd.Collapse Direction:=wdCollapseEnd
ActiveDocument.Endnotes.Add Range:=mainRngEnd

' if last is true it means there was a first note with test content, so we're leaving final note blank
If lastNote = False Then
    mainContentsRng.Copy
    ActiveDocument.Endnotes(1).Range.Paste
End If

End Sub

Sub copyBodyContentsToFootNotes()
Dim mainContentsRng As Range
Dim mainRngEnd As Range

' add trailing 'end' tag to main doc
Set mainContentsRng = ActiveDocument.StoryRanges(1)
mainContentsRng.InsertAfter (vbCr + "__END_TESTS__")

' add an empty note referencing last char of main doc
Set mainRngEnd = ActiveDocument.StoryRanges(1)
mainRngEnd.Collapse Direction:=wdCollapseEnd
ActiveDocument.Footnotes.Add Range:=mainRngEnd

' paste contents of main doc into the note (minus trailing note ref and vbcr)
mainContentsRng.MoveEnd Unit:=wdCharacter, Count:=-2
mainContentsRng.Copy
ActiveDocument.Footnotes(1).Range.Paste

End Sub

