Attribute VB_Name = "TestModule1"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Dim testDocx As Document
Dim MyStoryNo As Variant
Function returnTestResultString(testProcName, MyStoryNo)
Dim testProcNameStart As String
testProcNameStart = "__" + testProcName + "__^p"
Const testProcNameNext = "^p__"

Dim DelRange As Range, DelStartRange As Range, DelEndRange As Range
Dim FindStartRange As Range, FindEndRange As Range

Set FindStartRange = ActiveDocument.StoryRanges(MyStoryNo)
Set FindEndRange = ActiveDocument.StoryRanges(MyStoryNo)
Set DelRange = ActiveDocument.Range

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
        Set DelStartRange = FindStartRange
        'DelStartRange.Select
        FindEndRange.Start = DelStartRange.End
   End If
End With

With FindEndRange.Find
    .Text = testProcNameNext
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        Set DelEndRange = FindEndRange
        'DelEndRange.Select
    End If
End With

DelRange.Start = DelStartRange.End
DelRange.End = DelEndRange.Start
returnTestResultString = DelRange

End Function

Sub clearOtherTestContent(testProcName, MyStoryNo)
Dim testProcNameStart As String
testProcNameStart = "__" + testProcName + "__"
Const testProcNameNext = "__"

Dim Pre_DelRange As Range, Post_DelRange As Range
Dim Post_DelStartRange As Range, Pre_DelEndRange As Range
Dim Find_PostStartRange As Range, Find_PreEndRange As Range

Set Find_PreEndRange = ActiveDocument.StoryRanges(MyStoryNo)
Set Find_PostStartRange = ActiveDocument.StoryRanges(MyStoryNo)
Set Pre_DelRange = ActiveDocument.Range
Set Post_DelRange = ActiveDocument.Range

Pre_DelRange.Start = ActiveDocument.StoryRanges(MyStoryNo).Start
Post_DelRange.End = ActiveDocument.StoryRanges(MyStoryNo).End

' RM content preceding stuff for our test
With Find_PreEndRange.Find
    .Text = testProcNameStart + "^p"
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
        Set Pre_DelEndRange = Find_PreEndRange
        Pre_DelEndRange.Select
    End If
End With

Pre_DelRange.End = Pre_DelEndRange.End
Pre_DelRange.Delete

' RM content following stuff for our test
With Find_PostStartRange.Find
    .Text = testProcNameNext
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
        Set Post_DelStartRange = Find_PostStartRange
        Post_DelStartRange.Select
    End If
End With

Post_DelRange.Start = Post_DelStartRange.Start
Post_DelRange.Delete

End Sub

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub



'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Get DevSetup filepath.
    Dim testdotx_filepath As String
    testdotx_filepath = devTools.config.GetGitBasepath + "\test_files\testfile1.dotx"
    ' Create new test docx from template
    Set testDocx = Application.Documents.Add(testdotx_filepath)
    MyStoryNo = 1 '1 = Main Body, 2 = Footnotes, 3 = Endnotes. Can override this value per test as needed
End Sub
Function getStoryRange(MyStoryNo)
    getStoryRange = ActiveDocument.StoryRanges(MyStoryNo)
End Function


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Application.Documents(testDocx).Close SaveChanges:=wdDoNotSaveChanges
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDoubleQuotes"  '<-- name of this test procedure
        MyStoryNo = 1 'Main body of docx: use 2 for footnotes, 3 for endnotes
        Set pBar = New Progress_Bar ' < target sub updates progress bar constantly as it executes a find; we create a dummy
        'Call clearOtherTestContent(C_PROC_NAME, MyStoryNo)   ' < can use this to clear content from testdoc unrelated to this test
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        Unload pBar
        'results = ActiveDocument.StoryRanges(MyStoryNo)    ' use this to capture results if you're usign clearOtherTestContent above
        results = returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        'Assert.AreEqual 5, 4, "Test: compare ints"     '< Example
        Assert.Succeed
        Assert.AreEqual "Backtick pairs become doublequotes, ''Two single-primes also''", results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
