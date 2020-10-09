Attribute VB_Name = "TestHelpers"
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

Sub FindTestFiles()

End Sub

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
        If InStr(vbProj.FileName, gitrepo_anchorfile) Then
            repo_pathStr = Left(vbProj.FileName, InStrRev(vbProj.FileName, "\"))
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

