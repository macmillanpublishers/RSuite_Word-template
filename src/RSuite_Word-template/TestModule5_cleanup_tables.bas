Attribute VB_Name = "TestModule5_cleanup_tables"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private testDocx As Document
Private testdotx_filepath As String
Private testdotx As String
Private MyStoryNo As Variant


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Get testdot filepath.
    testdotx_filepath = getRepoPath + "test_files\testfile_cleanup_tables.dotx"
    ' Load public vars:
    SetCharacters
    Application.ScreenUpdating = False
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'reset loaded public vars
    DestroyCharacters
    Application.ScreenUpdating = True
    'MsgBox ("Cleanup Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Create new test docx from template
    '   \/ had set visible:=False but smote tests failed intermittently.
    '       if we want to set invisible, move this to individ. test level and set per test.
    Set testDocx = Application.Documents.Add(testdotx_filepath, visible:=True)
    testDocx.Activate
    ' for debug, make the doc visible:
    'Set testDocx = Application.Documents.Add(testdotx_filepath)
    MyStoryNo = 1 '1 = Main Body, 2 = Footnotes, 3 = Endnotes. Can override this value per test as needed
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Application.Documents(testDocx).Close savechanges:=wdDoNotSaveChanges
End Sub

'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_simplefinds() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
  '  On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_simplefinds"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = DOQ + "Backticks to doublequotes" + DCQ
        testTableIndex = 1
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.DoubleQuotes(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index
        Call Clean.DoubleQuotes(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.DoubleQuotes(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.DoubleQuotes(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_complexfinds() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_complexfinds"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Trailing nbsp "
        testTableIndex = 2
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index
        Call Clean.Spaces(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Spaces(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Spaces(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_simplefind_excludeOff() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_simplefind_excludeOff"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Just a return" + vbCr
        testTableIndex = 3
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index
        Call Clean.Spaces(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Spaces(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Spaces(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_complexfind_excludeOff() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_complexfind_excludeOff"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Trailing space "
        testTableIndex = 4
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.Spaces(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Spaces(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Spaces(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_simplefind_excludeOn() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_simplefind_excludeOn"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Just a return" + vbCr + " "
        testTableIndex = 5
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index
        Call Clean.Spaces(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Spaces(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Spaces(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_complexfind_excludeOn() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_complexfind_excludeOn"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Trailing space        "
        testTableIndex = 6
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.Spaces(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Spaces(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Spaces(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_dashesHighlight() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_dashesHighlight"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "1-800-456-7890"
        testTableIndex = 7
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Dashes(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.Dashes(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Dashes(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Dashes(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_dashesSelectFind() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_dashesSelectFind"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "55" + ENDASH + "678"
        testTableIndex = 8
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Dashes(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.Dashes(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Dashes(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Dashes(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_cleanBreaksFindloop() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_cleanBreaksFindloop"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Just one vbcr!" + vbCr
        testTableIndex = 9
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.CleanBreaks(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.CleanBreaks(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.CleanBreaks(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.CleanBreaks(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_ellipsisLookahead() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_ellipsisLookahead"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "More to come" + NBS_ELLIPSIS + " "
        testTableIndex = 10
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Ellipses(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.Ellipses(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.Ellipses(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Ellipses(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_dquoteLookahead() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_dquoteLookahead"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "That was the end" + DCQ
        testTableIndex = 11
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.DoubleQuotes(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.DoubleQuotes(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.DoubleQuotes(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.DoubleQuotes(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_squoteLookahead() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_squoteLookahead"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "The end again. " + SOQ
        testTableIndex = 12
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.SingleQuotes(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.SingleQuotes(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.SingleQuotes(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.SingleQuotes(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_titleCase() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_titleCase"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "Make This Title Case"
        testTableIndex = 13
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.MakeTitleCase(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.MakeTitleCase(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.MakeTitleCase(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.MakeTitleCase(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_hyperlinks() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_hyperlinks"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "google.com"
        testTableIndex = 14
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.RemoveHyperlinks(1)
        results = TestHelpers.lastTablecellText(1, testTableIndex) ' params: storyrange, table_index, row, column
        Call Clean.RemoveHyperlinks(1)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
        Call Clean.RemoveHyperlinks(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.RemoveHyperlinks(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, results
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_tables")
Private Sub TestTables_trimNoteSpaces() 'TODO Rename test
    Dim expected_str As String, results As String, second_results As String, fnotes_results As String, _
        enotes_results As String
    Dim testTableIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTables_trimNoteSpaces"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        expected_str = "End of the Note "
        testTableIndex = 15
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(2)
        fnotes_results = TestHelpers.lastTablecellText(2, testTableIndex)
        Call Clean.Spaces(3)
        enotes_results = TestHelpers.lastTablecellText(3, testTableIndex)
        Call Clean.Spaces(3)
        second_results = TestHelpers.lastTablecellText(1, testTableIndex)
    'Assert:
        Assert.Succeed
        Assert.areequal expected_str, second_results
        Assert.areequal expected_str, fnotes_results
        Assert.areequal expected_str, enotes_results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

