Attribute VB_Name = "TestModule2_charstyles"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private Sym_symbol_expected, Sym_italsym_expected As String, Sym_validsym_expected As String
Private testDocx As Document
Private testdotx_filepath As String
Private testresults_dotx_filepath As String
Private MyStoryNo As Variant
Private Function SetResultStrings()

Sym_symbol_expected = "symbols (sym)"
Sym_italsym_expected = "symbols-ital (symi)"
Sym_validsym_expected = "Body-Text (Tx)"

End Function

Private Function DestroyResultStrings()

Sym_symbol_expected = vbNullString
Sym_italsym_expected = vbNullString
Sym_validsym_expected = vbNullString

End Function

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Get testdot filepath.
    testdotx_filepath = getRepoPath + "test_files\testfile_charstyles.dotx"
    ' Get results docx filepath
    testresults_dotx_filepath = getRepoPath + "test_files\testfile_charstyle_results.dotx"
    ' Load public vars:
    SetCharacters
    SetResultStrings
    Application.ScreenUpdating = False
    'Set pBar = New Progress_Bar
    'pBarCounter = 0
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'reset loaded public vars
    DestroyCharacters
    DestroyResultStrings
    Application.ScreenUpdating = True
    'MsgBox ("Charstyle Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
   ' Create new test docx from template
   Set testDocx = Application.Documents.Add(testdotx_filepath, visible:=False)
   ' or create doc visibly, for debug:
   'Set testDocx = Application.Documents.Add(testdotx_filepath)
   testDocx.Activate
   MyStoryNo = 1 '1 = Main Body, 2 = Footnotes, 3 = Endnotes. Can override this value per test as needed
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    'Unload pBar
    Application.Documents(testDocx).Close savechanges:=wdDoNotSaveChanges
    Set testDocx = Nothing
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestPCSpecialCharacters_symbol() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestPCSpecialCharacters_symbol"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        results = TestHelpers.returnTestResultStyle(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal Sym_symbol_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestPCSpecialCharacters_italsymbol() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestPCSpecialCharacters_italsymbol"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        results = TestHelpers.returnTestResultStyle(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal Sym_italsym_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestPCSpecialCharacters_validsymbol() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestPCSpecialCharacters_validsymbol"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        results = TestHelpers.returnTestResultStyle(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal Sym_validsym_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestPCSpecialCharacters_secondrun() 'TODO Rename test
    Dim results_symbol, results_italsymbol, results_validsymbol As String
    On Error GoTo TestFail
    'Arrange:
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        results_symbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_symbol", MyStoryNo)
        results_italsymbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_italsymbol", MyStoryNo)
        results_validsymbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_validsymbol", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal Sym_symbol_expected, results_symbol
        Assert.areequal Sym_italsym_expected, results_italsymbol
        Assert.areequal Sym_validsym_expected, results_validsymbol
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestPCSpecialCharacters_footnotes() 'TODO Rename test
    Dim results_symbol, results_italsymbol, results_validsymbol As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        results_symbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_symbol", MyStoryNo)
        results_italsymbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_italsymbol", MyStoryNo)
        results_validsymbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_validsymbol", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal Sym_symbol_expected, results_symbol
        Assert.areequal Sym_italsym_expected, results_italsymbol
        Assert.areequal Sym_validsym_expected, results_validsymbol
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestPCSpecialCharacters_endnotes() 'TODO Rename test
    Dim results_symbol, results_italsymbol, results_validsymbol As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.CheckSpecialCharactersPC(MyStoryNo)
        results_symbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_symbol", MyStoryNo)
        results_italsymbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_italsymbol", MyStoryNo)
        results_validsymbol = TestHelpers.returnTestResultStyle("TestPCSpecialCharacters_validsymbol", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal Sym_symbol_expected, results_symbol
        Assert.areequal Sym_italsym_expected, results_italsymbol
        Assert.areequal Sym_validsym_expected, results_validsymbol
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestCheckAppliedStyles_basic() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestCheckAppliedStyles_basic"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestCheckAppliedStyles_multistyle() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestCheckAppliedStyles_multistyle"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestCheckAppliedStyles_allstyles() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestCheckAppliedStyles_allstyles"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CharStylesMacro")
' testing LocalFormatting in conjunction with Check Applied Styles to verify we are not mis-removing styles again
Private Sub TestWdv321_basic() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        ' for the other 2 wdv-321 tests we are using the CheckAppliedStyles content; bu tthis one has a brk that is handled
        '   differently by the LocalFormatting macro, resulting in slightly different expected output.
        Const C_PROC_NAME = "TestWdv321_basic"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
' testing LocalFormatting in conjunction with Check Applied Styles to verify we are not mis-removing styles again
Private Sub TestWdv321_multistyle() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestCheckAppliedStyles_multistyle"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
' testing LocalFormatting in conjunction with Check Applied Styles to verify we are not mis-removing styles again
Private Sub TestWdv321_allstyles() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestCheckAppliedStyles_allstyles"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestCheckAppliedStyles_footnotes() 'TODO Rename test
    Dim results_basic As Range, basic_expected As Range
    Dim results_multistyle As Range, multi_expected As Range
    Dim results_allstyle As Range, allstyle_expected As Range
    Dim testResultsDocx As Document
    Dim compareStr_basic As String, compareStr_multi As String, compareStr_all As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_basic = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_basic", MyStoryNo, testDocx)
        Set results_multistyle = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_multistyle", MyStoryNo, testDocx)
        Set results_allstyle = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_allstyles", MyStoryNo, testDocx)
        ''' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        ' Get known good results
        Set basic_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_basic", 1, testResultsDocx)
        Set multi_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_multistyle", 1, testResultsDocx)
        Set allstyle_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_allstyles", 1, testResultsDocx)
        ' Compare known good output and output from just now
        compareStr_basic = TestHelpers.compareRanges(results_basic, basic_expected)
        compareStr_multi = TestHelpers.compareRanges(results_multistyle, multi_expected)
        compareStr_all = TestHelpers.compareRanges(results_allstyle, allstyle_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", compareStr_basic
        Assert.areequal "Same", compareStr_multi
        Assert.areequal "Same", compareStr_all
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestCheckAppliedStyles_endnotes() 'TODO Rename test
    Dim results_basic As Range, basic_expected As Range
    Dim results_multistyle As Range, multi_expected As Range
    Dim results_allstyle As Range, allstyle_expected As Range
    Dim testResultsDocx As Document
    Dim compareStr_basic As String, compareStr_multi As String, compareStr_all As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_basic = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_basic", MyStoryNo, testDocx)
        Set results_multistyle = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_multistyle", MyStoryNo, testDocx)
        Set results_allstyle = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_allstyles", MyStoryNo, testDocx)
        ''' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        ' Get known good results
        Set basic_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_basic", 1, testResultsDocx)
        Set multi_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_multistyle", 1, testResultsDocx)
        Set allstyle_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_allstyles", 1, testResultsDocx)
        ' Compare known good output and output from just now
        compareStr_basic = TestHelpers.compareRanges(results_basic, basic_expected)
        compareStr_multi = TestHelpers.compareRanges(results_multistyle, multi_expected)
        compareStr_all = TestHelpers.compareRanges(results_allstyle, allstyle_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", compareStr_basic
        Assert.areequal "Same", compareStr_multi
        Assert.areequal "Same", compareStr_all
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestCheckAppliedStyles_secondrun() 'TODO Rename test
    Dim results_basic As Range, basic_expected As Range
    Dim results_multistyle As Range, multi_expected As Range
    Dim results_allstyle As Range, allstyle_expected As Range
    Dim testResultsDocx As Document
    Dim compareStr_basic As String, compareStr_multi As String, compareStr_all As String
    On Error GoTo TestFail
    'Arrange:
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Call Clean.CheckAppliedCharStyles(MyStoryNo)
        Set results_basic = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_basic", MyStoryNo, testDocx)
        Set results_multistyle = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_multistyle", MyStoryNo, testDocx)
        Set results_allstyle = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_allstyles", MyStoryNo, testDocx)
        ''' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        ' Get known good results
        Set basic_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_basic", 1, testResultsDocx)
        Set multi_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_multistyle", 1, testResultsDocx)
        Set allstyle_expected = TestHelpers.returnTestResultRange("TestCheckAppliedStyles_allstyles", 1, testResultsDocx)
        ' Compare known good output and output from just now
        compareStr_basic = TestHelpers.compareRanges(results_basic, basic_expected)
        compareStr_multi = TestHelpers.compareRanges(results_multistyle, multi_expected)
        compareStr_all = TestHelpers.compareRanges(results_allstyle, allstyle_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", compareStr_basic
        Assert.areequal "Same", compareStr_multi
        Assert.areequal "Same", compareStr_all
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestLocalFormatting() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestLocalFormatting"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestLocalFormatting_tables() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestLocalFormatting_tables"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        Set results_actual = TestHelpers.returnTestResultRange(C_PROC_NAME, MyStoryNo, testDocx)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
        Set results_expected = TestHelpers.returnTestResultRange(C_PROC_NAME, 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
        ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestLocalFormatting_footnotes() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range, results_tables_actual As Range, results_tables_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String, results_tables_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
    'Compare function 1
        Set results_actual = TestHelpers.returnTestResultRange("TestLocalFormatting", MyStoryNo, testDocx)
        Set results_expected = TestHelpers.returnTestResultRange("TestLocalFormatting", 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
    'Compare function 2
        Set results_tables_actual = TestHelpers.returnTestResultRange("TestLocalFormatting_tables", MyStoryNo, testDocx)
        Set results_tables_expected = TestHelpers.returnTestResultRange("TestLocalFormatting_tables", 1, testResultsDocx)
        ' Compare known good output and output from just now
        results_tables_compareStr = TestHelpers.compareRanges(results_tables_actual, results_tables_expected)
    ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
        Assert.areequal "Same", results_tables_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestLocalFormatting_endnotes() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range, results_tables_actual As Range, results_tables_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String, results_tables_compareStr As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
    'Compare function 1
        Set results_actual = TestHelpers.returnTestResultRange("TestLocalFormatting", MyStoryNo, testDocx)
        Set results_expected = TestHelpers.returnTestResultRange("TestLocalFormatting", 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
    'Compare function 2
        Set results_tables_actual = TestHelpers.returnTestResultRange("TestLocalFormatting_tables", MyStoryNo, testDocx)
        Set results_tables_expected = TestHelpers.returnTestResultRange("TestLocalFormatting_tables", 1, testResultsDocx)
        ' Compare known good output and output from just now
        results_tables_compareStr = TestHelpers.compareRanges(results_tables_actual, results_tables_expected)
    ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
        Assert.areequal "Same", results_tables_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CharStylesMacro")
Private Sub TestLocalFormatting_secondrun() 'TODO Rename test
    Dim results_actual As Range, results_expected As Range, results_tables_actual As Range, results_tables_expected As Range
    Dim testResultsDocx As Document
    Dim result_compareStr As String, results_tables_compareStr As String
    On Error GoTo TestFail
    'Act:
        Call Clean.LocalFormatting(MyStoryNo)
        Call Clean.LocalFormatting(MyStoryNo)
        ' Create new results docx from template
        Set testResultsDocx = Application.Documents.Add(testresults_dotx_filepath, visible:=False)
    'Compare function 1
        Set results_actual = TestHelpers.returnTestResultRange("TestLocalFormatting", MyStoryNo, testDocx)
        Set results_expected = TestHelpers.returnTestResultRange("TestLocalFormatting", 1, testResultsDocx)
        ' Compare known good output and output from just now
        result_compareStr = TestHelpers.compareRanges(results_actual, results_expected)
    'Compare function 2
        Set results_tables_actual = TestHelpers.returnTestResultRange("TestLocalFormatting_tables", MyStoryNo, testDocx)
        Set results_tables_expected = TestHelpers.returnTestResultRange("TestLocalFormatting_tables", 1, testResultsDocx)
        ' Compare known good output and output from just now
        results_tables_compareStr = TestHelpers.compareRanges(results_tables_actual, results_tables_expected)
    ' Close results doc
        Application.Documents(testResultsDocx).Close savechanges:=wdDoNotSaveChanges
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
        Assert.areequal "Same", results_tables_compareStr
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



