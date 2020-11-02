Attribute VB_Name = "TestModule2_charstyles"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private Sym_symbol_expected, Sym_italsym_expected, Sym_validsym_expected As String
Private testDocx As Document
Private testdotx_filepath As String
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
    MsgBox ("Charstyle Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
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
        Assert.AreEqual Sym_symbol_expected, results
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
        Assert.AreEqual Sym_italsym_expected, results
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
        Assert.AreEqual Sym_validsym_expected, results
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
        Assert.AreEqual Sym_symbol_expected, results_symbol
        Assert.AreEqual Sym_italsym_expected, results_italsymbol
        Assert.AreEqual Sym_validsym_expected, results_validsymbol
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
        Assert.AreEqual Sym_symbol_expected, results_symbol
        Assert.AreEqual Sym_italsym_expected, results_italsymbol
        Assert.AreEqual Sym_validsym_expected, results_validsymbol
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
        Assert.AreEqual Sym_symbol_expected, results_symbol
        Assert.AreEqual Sym_italsym_expected, results_italsymbol
        Assert.AreEqual Sym_validsym_expected, results_validsymbol
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



