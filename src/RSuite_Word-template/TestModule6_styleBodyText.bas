Attribute VB_Name = "TestModule6_styleBodyText"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private testDocx As Document
Private testdotx_filepath As String
Private good_testdoc_filepath As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Get testdot filepath.
    testdotx_filepath = getRepoPath + "test_files\testfile_styleBodyText.dotx"
    good_testdoc_filepath = getRepoPath + "test_files\testfile_styleBodyText-good.dotx"
    ' Load public vars:
    Application.ScreenUpdating = False
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'reset loaded public vars
    Application.ScreenUpdating = True
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Create new test docx from template
    Set testDocx = Nothing
    ' this one has to be visible, for styles.Namelocal to properly apply (in tagtext)
    Set testDocx = Application.Documents.Add(testdotx_filepath)
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Application.Documents(testDocx).Close savechanges:=wdDoNotSaveChanges
    Set testDocx = Nothing
End Sub

'@TestMethod("StyleBodyTxtMacro")
Private Sub TestStyleBodyText() 'TODO Rename test
    Dim testdoc_good As Document
    Dim testRngMain As Range, expectedRngMain As Range, testRngFN As Range, expectedRngFN As Range, _
        testRngEN As Range, expectedRngEN As Range, testRngSecondRun As Range
    Dim resultMain As String, resultFN As String, resultEN As String, resultSecondRun As String
    On Error GoTo TestFail
    'Arrange:
        Fakes.MsgBox.Returns vbOK
        Set testdoc_good = Nothing
    'Act:
        Call TagUnstyledParas.tagText(testDocx)
        Set testRngMain = testDocx.StoryRanges(1)
        Set testRngFN = testDocx.StoryRanges(2)
        Set testRngEN = testDocx.StoryRanges(3)
        
        'get ranges from expected doc
        Set testdoc_good = Application.Documents.Add(good_testdoc_filepath, visible:=False)
        ' open file visibly, for debug
        'Set testdoc_good = Application.Documents.Add(good_testdoc_filepath)
        Set expectedRngMain = testdoc_good.StoryRanges(1)
        Set expectedRngFN = testdoc_good.StoryRanges(2)
        Set expectedRngEN = testdoc_good.StoryRanges(3)
    
        ' compare ranges from expected doc, testdoc
        resultMain = TestHelpers.compareParaStylesInRange(testRngMain, expectedRngMain)
        resultFN = TestHelpers.compareParaStylesInRange(testRngFN, expectedRngFN)
        resultEN = TestHelpers.compareParaStylesInRange(testRngEN, expectedRngEN)
        
        ' test second run
        Call TagUnstyledParas.tagText(testDocx)
        Set testRngSecondRun = testDocx.StoryRanges(1)
        resultSecondRun = TestHelpers.compareParaStylesInRange(testRngSecondRun, expectedRngMain)
    'Assert:
        Assert.Succeed
        Assert.areequal "", resultMain
        Assert.areequal "", resultFN
        Assert.areequal "", resultEN
        Assert.areequal "", resultSecondRun
    'Cleanup:
        testdoc_good.Close savechanges:=wdDoNotSaveChanges
        Set testdoc_good = Nothing
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

