Attribute VB_Name = "TestHelpers"
Function returnTestResultStyle(testProcName, MyStoryNo)
Dim testProcNameStart As String
testProcNameStart = "__" + testProcName + "__^p"
Const testProcNameNext = "^p__"

Dim resultRng As Range, DelStartRange As Range, DelEndRange As Range
Dim FindStartRange As Range, FindEndRange As Range

Set FindStartRange = ActiveDocument.StoryRanges(MyStoryNo)
Set FindEndRange = ActiveDocument.StoryRanges(MyStoryNo)
Set resultRng = ActiveDocument.Range

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
Set resultRng = ActiveDocument.Range

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

