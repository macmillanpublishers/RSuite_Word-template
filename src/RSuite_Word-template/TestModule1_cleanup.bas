Attribute VB_Name = "TestModule1_cleanup"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private DQ_simplefinds_expected, DQ_emdash_expected, DQ_spaces_expected, DQ_special_expected, DQ_other_expected As String
Private testDocx As Document
Private MyStoryNo As Variant
Private Function SetResultStrings()

DQ_simplefinds_expected = DOQ + "Backtick pairs become doublequotes" + DCQ + ", " + DOQ + "Two single-primes also" + DCQ
DQ_emdash_expected = "Testing emdashes pt 1" + EMDASH + DCQ + " Should be DCQ" + vbCr _
            + "Testing emdashes pt 2" + EMDASH + DOQ + "Should be DOQ"
DQ_spaces_expected = "Testing spaces A " + DCQ + " DCQ" + vbCr _
            + "Testing spaces B " + DCQ + vbCr + "DCQ" + vbCr _
            + "Testing spaces C " + DOQ + SOQ + "DoqSoq" + vbCr _
            + "Testing spaces C2 " + DOQ + SOQ + DOQ + "DoqSoqDoq" + vbCr _
            + "Testing spaces D " + DOQ + "DOQ"
DQ_special_expected = "Testing vbcr" + vbCr + DOQ + "DOQ" + vbCr _
            + "Testing tab" + vbTab + DOQ + "DOQ" + vbCr _
            + "Testing oParen (" + DOQ + "DOQ" + vbCr _
            + "Testing special quote combo (" + DOQ + SOQ + "DoqSoq" + vbCr _
            + "Testing special quote combo2 (" + DOQ + SOQ + DOQ + "DoqSoqDoq"
DQ_other_expected = "Testing leading text" + DCQ + " DCQ" + vbCr _
            + "Testing leading and trailing text" + DCQ + "DCQ" + vbCr _
            + "Testing leading text and quote" + SCQ + DCQ + "ScqDcq" + vbCr _
            + "Testing leading text and quote2" + DCQ + SCQ + DCQ + "DcqScqDcq"

End Function

Private Function DestroyResultStrings()

DQ_simplefinds_expected = vbNullString
DQ_emdash_expected = vbNullString
DQ_spaces_expected = vbNullString
DQ_special_expected = vbNullString
DQ_other_expected = vbNullString

End Function


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Load public vars:
    SetCharacters
    SetResultStrings
    Application.ScreenUpdating = False
    Set pBar = New Progress_Bar
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'reset loaded public vars
    'Unload pBar
    DestroyCharacters
    DestroyResultStrings
    Application.ScreenUpdating = True
    MsgBox ("Cleanup Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Get DevSetup filepath.
    Dim testdotx_filepath As String
    testdotx_filepath = devTools.config.GetGitBasepath + "\test_files\testfile1_cleanup.dotx"
    ' Create new test docx from template
    Set testDocx = Application.Documents.Add(testdotx_filepath)
    MyStoryNo = 1 '1 = Main Body, 2 = Footnotes, 3 = Endnotes. Can override this value per test as needed
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Unload pBar
    Application.Documents(testDocx).Close SaveChanges:=wdDoNotSaveChanges
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_simplefinds() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDoubleQuotes_simplefinds"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        'Call clearOtherTestContent(C_PROC_NAME, MyStoryNo)   ' < can use this to clear content from testdoc unrelated to this test
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        'results = ActiveDocument.StoryRanges(MyStoryNo)    ' use this to capture results if you're usign clearOtherTestContent above
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        'Assert.AreEqual 5, 4, "Test: compare ints"     '< Example
        Assert.Succeed
        Assert.AreEqual DQ_simplefinds_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_emdash() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDoubleQuotes_emdash"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.AreEqual DQ_emdash_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_spaces() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDoubleQuotes_spaces"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.AreEqual DQ_spaces_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_special() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDoubleQuotes_special"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.AreEqual DQ_special_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_secondrun() 'TODO Rename test
    Dim results_emdash, results_other, results_simplefinds, results_spaces, results_special As String
    On Error GoTo TestFail
    'Arrange:
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        Call Clean.DoubleQuotes(MyStoryNo)
        results_emdash = TestHelpers.returnTestResultString("TestDoubleQuotes_emdash", MyStoryNo)
        results_other = TestHelpers.returnTestResultString("TestDoubleQuotes_other", MyStoryNo)
        results_simplefinds = TestHelpers.returnTestResultString("TestDoubleQuotes_simplefinds", MyStoryNo)
        results_spaces = TestHelpers.returnTestResultString("TestDoubleQuotes_spaces", MyStoryNo)
        results_special = TestHelpers.returnTestResultString("TestDoubleQuotes_special", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.AreEqual DQ_emdash_expected, results_emdash
        Assert.AreEqual DQ_other_expected, results_other
        Assert.AreEqual DQ_simplefinds_expected, results_simplefinds
        Assert.AreEqual DQ_spaces_expected, results_spaces
        Assert.AreEqual DQ_special_expected, results_special
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_other() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDoubleQuotes_other"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.AreEqual DQ_other_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
