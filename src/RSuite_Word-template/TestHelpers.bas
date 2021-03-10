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

Sub testsub()
Debug.Print """" & strFromLocation(ActiveDocument, 25, 0, 4) & """"
End Sub

Function compareRanges(actualRange As Range, expectedRange As Range) As String
Dim returnString As String
Dim i As Integer
returnString = "Same"

If actualRange.Characters.Count <> expectedRange.Characters.Count Then
    returnString = "Compared ranges are different lengths, expected: " + str(expectedRange.Characters.Count) + _
        ", actual: " + str(actualRange.Characters.Count)
    GoTo TheEnd
ElseIf actualRange.Text <> expectedRange.Text Then
    returnString = "Range text mismatch, expected: '" + actualRange.Text + "', actual: '" + expectedRange.Text + "'"
    GoTo TheEnd
Else
    For i = 1 To actualRange.Characters.Count
        ' namelocal reports char style where present, otherwise give para style
        If actualRange.Characters(i).Style.NameLocal <> expectedRange.Characters(i).Style.NameLocal Then
            returnString = "Different styles detected for char #" + str(i) + " ('" + actualRange.Characters(i) + _
                "'), expected: '" + expectedRange.Characters(i).Style.NameLocal + "', actual: '" + _
                actualRange.Characters(i).Style.NameLocal + "'"
            GoTo TheEnd
        ' checking local formatting types one by one
        ElseIf actualRange.Characters(i).Font.Bold <> expectedRange.Characters(i).Font.Bold Then
            returnString = "Diff in 'bold' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        ElseIf actualRange.Characters(i).Font.Italic <> expectedRange.Characters(i).Font.Italic Then
            returnString = "Diff in 'italic' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        ElseIf actualRange.Characters(i).Font.SmallCaps <> expectedRange.Characters(i).Font.SmallCaps Then
            returnString = "Diff in 'smallcaps' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        ElseIf actualRange.Characters(i).Font.Subscript <> expectedRange.Characters(i).Font.Subscript Then
            returnString = "Diff in 'Subscript' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        ElseIf actualRange.Characters(i).Font.Superscript <> expectedRange.Characters(i).Font.Superscript Then
            returnString = "Diff in 'Superscript' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        ElseIf actualRange.Characters(i).Font.StrikeThrough <> expectedRange.Characters(i).Font.StrikeThrough Then
            returnString = "Diff in 'strikethrough' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        ElseIf actualRange.Characters(i).Font.Underline <> expectedRange.Characters(i).Font.Underline Then
            returnString = "Diff in 'underline' found for char #" + str(i) + " ('" + actualRange.Characters(i) + "')"
            GoTo TheEnd
        End If
    Next i
End If

compareRanges = returnString

TheEnd:
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

returnTestResultStyle = resultRng.Style.NameLocal

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






'
'Sub clearOtherTestContent(testProcName, MyStoryNo)
'Dim testProcNameStart As String
'testProcNameStart = "__" + testProcName + "__"
'Const testProcNameNext = "__"
'
'Dim Pre_DelRange As Range, Post_DelRange As Range
'Dim Post_DelStartRange As Range, Pre_DelEndRange As Range
'Dim Find_PostStartRange As Range, Find_PreEndRange As Range
'
'Set Find_PreEndRange = ActiveDocument.StoryRanges(MyStoryNo)
'Set Find_PostStartRange = ActiveDocument.StoryRanges(MyStoryNo)
'Set Pre_DelRange = ActiveDocument.Range
'Set Post_DelRange = ActiveDocument.Range
'
'Pre_DelRange.Start = ActiveDocument.StoryRanges(MyStoryNo).Start
'Post_DelRange.End = ActiveDocument.StoryRanges(MyStoryNo).End
'
'' RM content preceding stuff for our test
'With Find_PreEndRange.Find
'    .Text = testProcNameStart + "^p"
'    .Replacement.Text = ""
'    .Forward = True
'    .Wrap = wdFindAsk
'    .Format = False
'    .MatchCase = True
'    .MatchWholeWord = True
'    .MatchWildcards = False
'    .MatchSoundsLike = False
'    .MatchAllWordForms = False
'    .Execute
'    If .Found = True Then
'        Set Pre_DelEndRange = Find_PreEndRange
'        Pre_DelEndRange.Select
'    End If
'End With
'
'Pre_DelRange.End = Pre_DelEndRange.End
'Pre_DelRange.Delete
'
'' RM content following stuff for our test
'With Find_PostStartRange.Find
'    .Text = testProcNameNext
'    .Replacement.Text = ""
'    .Forward = True
'    .Wrap = wdFindAsk
'    .Format = False
'    .MatchCase = True
'    .MatchWholeWord = True
'    .MatchWildcards = False
'    .MatchSoundsLike = False
'    .MatchAllWordForms = False
'    .Execute
'    If .Found = True Then
'        Set Post_DelStartRange = Find_PostStartRange
'        Post_DelStartRange.Select
'    End If
'End With
'
'Post_DelRange.Start = Post_DelStartRange.Start
'Post_DelRange.Delete
'
'End Sub

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

Sub copyBodyContentsToEndNotes()
Dim mainContentsRng As Range
Set mainContentsRng = ActiveDocument.StoryRanges(1)

' add trailing 'end' tag to main doc
Set mainContentsRng = ActiveDocument.StoryRanges(1)
mainContentsRng.InsertAfter (vbCr + "__END_TESTS__")

' add an empty note referencing last char of main doc
Set mainRngEnd = ActiveDocument.StoryRanges(1)
mainRngEnd.Collapse Direction:=wdCollapseEnd
ActiveDocument.Endnotes.Add Range:=mainRngEnd

' paste contents of main doc into the note (minus trailing note ref and vbcr)
mainContentsRng.MoveEnd Unit:=wdCharacter, Count:=-2
mainContentsRng.Copy
ActiveDocument.Endnotes(1).Range.Paste
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

