Attribute VB_Name = "TestModule1_cleanup"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private DQ_simplefinds_expected As String, DQ_emdash_expected As String, DQ_spaces_expected As String, DQ_special_expected As String, _
    DQ_other_expected As String, SQ_backtick_expected As String, SQ_setasides_expected As String, SQ_openquo_expected As String, _
    SQ_fandr_expected As String, EL_frbasic_expected As String, EL_fr4dots_expected As String, EL_spaces_expected As String, _
    EL_emdashes_expected As String, SP_nbsps_expected As String, SP_brackets_expected As String, SP_breaks_expected As String, _
    SP_exclude_expected As String, PN_expected As String, DSH_numbers_expected As String, DSH_expected As String, _
    DSH_exclude_expected As String, MTC_expected As String, BR_expected As String
Private HL_expected As Integer

Private testDocx As Document
Private testdotx_filepath As String
Private testdotx As String
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
SQ_backtick_expected = "Testing " + SOQ + "backtick" + SCQ + ":SoqScq" + vbCr _
            + SOQ + "Backtick two. " + SOQ + " :SoqSoq and (" + SOQ + "backtick 3" + SCQ + ")" + SCQ + " :SoqScqScq"
SQ_setasides_expected = "Test preceding space + good char: " + SCQ + "K :SCQ" + vbCr _
            + "Test preceding space + bad char: " + SOQ + "L :SOQ" + vbCr _
            + "Test preceding space + good word: " + SCQ + "spossible :SCQ" + vbCr _
            + "Test preceding space + bad word: " + SOQ + "npossible :SOQ" + vbCr _
            + "Test space + good year: " + SCQ + "30 :SCQ" + vbCr _
            + "Test space + bad year: " + SOQ + "30( :SOQ"
SQ_openquo_expected = "Openquo true 1 " + SOQ + SOQ + " and 2 " + DOQ + SOQ + " SoqSoq" + vbCr _
            + "Openquo true 3 " + SOQ + SOQ + "SOQ and 4 " + DOQ + SOQ + "SOQ" + vbCr _
            + "Openquo false 1 " + DCQ + SCQ + " and 2 X" + SCQ + ". ScqScq"
SQ_fandr_expected = "Tighten " + SOQ + DOQ + "soq doq, tighten " + DOQ + SOQ + "doq soq" + vbCr _
            + "Tighten scq dcq" + SCQ + DCQ + ", tighten dcq scq" + DCQ + SCQ
EL_frbasic_expected = "Finished ellipse " + ELLIPSIS + " finished nbsp_el" + NBS_ELLIPSIS + " ready ellipse" + NBS_ELLIPSIS _
            + " ellipse char" + NBS_ELLIPSIS + " 3 dots" + NBS_ELLIPSIS + " done"
EL_fr4dots_expected = "Ell char." + NBS_ELLIPSIS + " 4 periods." + NBS_ELLIPSIS + " toolong ellipse." + NBS_ELLIPSIS _
            + " dot after ellipse." + NBS_ELLIPSIS + " done"
EL_spaces_expected = "Too much space" + NBS_ELLIPSIS + " leading quotes1, " + DOQ + ELLIPSIS + NBSPchar + "leadquotes2, " _
            + SOQ + ELLIPSIS + NBSPchar + "new pre-nbsp" + NBS_ELLIPSIS + " nl" + vbCr + ELLIPSIS + " done"
EL_emdashes_expected = "Emdash case1 take1" + EMDASH + " " + ELLIPSIS + NBSPchar + "case1 take2" + EMDASH + " " + ELLIPSIS + NBSPchar _
            + "case2 take1" + NBS_ELLIPSIS + NBSPchar + EMDASH + " case2 take2" + NBS_ELLIPSIS + NBSPchar + EMDASH + " done"
SP_nbsps_expected = "Rm nbsps and preserve done ellipses:" + vbCr + ELLIPSIS + " <basic" + NBS_ELLIPSIS + " <nbsp" + PERIOD_ELLIPSIS _
            + " <period " + DOQ + ELLIPSIS + NBSPchar + "<quote" + NBS_ELLIPSIS + NBSPchar + EMDASH + " <emdash"
SP_brackets_expected = "Tabs (parens) [brackets] {braces} $dollars spaces again"
SP_breaks_expected = "Soft" + vbCr + "break; space preceding break" + vbCr + "and space following"
SP_exclude_expected = "Tabs" + vbTab + vbTab + vbTab + "( parens)  [   brackets   ] {braces } $ dollars         spaces again Soft" _
            + vbVerticalTab + "break;" + vbCr + " and space following break"
PN_expected = "Multicomma, multiperiods. Non-breaking-hyphens, optionalhyphens."
DSH_expected = "Bar" + EMDASH + "character, figure" + ENDASH + "dash, triple" + EMDASH + "dash, double" + EMDASH _
            + "dash, double" + EMDASH + "andspaces" + vbCr _
            + "Space-space, dash-space, space-dash" + vbCr _
            + "Space" + EMDASH + "endash, endash" + ENDASH + "space, emdash" + EMDASH + "space, space" + EMDASH + "emdash"
DSH_exclude_expected = "triple---dash, double--dash, double -- andspaces" + vbCr + "Space - space, dash- space, space -dash" + vbCr _
            + "Space " + ENDASH + "endash, endash" + ENDASH + " space, emdash" + EMDASH + " space, space " + EMDASH + "emdash" + vbCr _
            + "7" + ENDASH + "8, from 94" + ENDASH + "112, space emdash" + EMDASH + "space"
DSH_numbers_expected = "Leave alone phone: 703-536-4247, (987) 654-3211, 1-800-123-4567" + vbCr _
            + "Leave alone Isbns: 978-5-426-01234-8, 979-022-2323212" + vbCr _
            + "Now endash: 7" + ENDASH + "8, from 94" + ENDASH + "112, also 15" + ENDASH + "34" + ENDASH + "41, 6.0" + ENDASH + "6.125"
MTC_expected = "Testing All of the Lowercase" + vbCr _
            + "Testing If They Use All Caps" + vbCr _
            + "Tragically the Caps Are Bad" + vbCr _
            + "keywords:" + vbCr _
            + "The the HBO HTML V past the Down"
HL_expected = 3
BR_expected = "Line" + vbCr + "break. Now excluded Line" + vbVerticalTab + "break. Now page" + vbCr + "break. Now double new" + vbCr + "paras. Now five new" + vbCr + "paras."

End Function

Private Function DestroyResultStrings()

DQ_simplefinds_expected = vbNullString
DQ_emdash_expected = vbNullString
DQ_spaces_expected = vbNullString
DQ_special_expected = vbNullString
DQ_other_expected = vbNullString
SQ_backtick_expected = vbNullString
SQ_setasides_expected = vbNullString
SQ_openquo_expected = vbNullString
SQ_fandr_expected = vbNullString
EL_frbasic_expected = vbNullString
EL_fr4dots_expected = vbNullString
EL_spaces_expected = vbNullString
EL_emdashes_expected = vbNullString
SP_nbsps_expected = vbNullString
SP_brackets_expected = vbNullString
SP_breaks_expected = vbNullString
SP_exclude_expected = vbNullString
PN_expected = vbNullString
DSH_expected = vbNullString
DSH_exclude_expected = vbNullString
DSH_numbers_expected = vbNullString
MTC_expected = vbNullString
HL_expected = 0
BR_expected = vbNullString

End Function

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Get testdot filepath.
    testdotx_filepath = getRepoPath + "test_files\testfile_cleanup.dotx"
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
    'MsgBox ("Cleanup Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Create new test docx from template
    Set testDocx = Application.Documents.Add(testdotx_filepath, visible:=False)
    ' for debug, make the doc visible:
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
        Assert.areequal DQ_simplefinds_expected, results
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
        Assert.areequal DQ_emdash_expected, results
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
        Assert.areequal DQ_spaces_expected, results
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
        Assert.areequal DQ_special_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_secondrun() 'TODO Rename test
    Dim results_emdash As String, results_other As String, results_simplefinds As String, results_spaces As String, results_special As String
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
        Assert.areequal DQ_emdash_expected, results_emdash
        Assert.areequal DQ_other_expected, results_other
        Assert.areequal DQ_simplefinds_expected, results_simplefinds
        Assert.areequal DQ_spaces_expected, results_spaces
        Assert.areequal DQ_special_expected, results_special
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_footnotes() 'TODO Rename test
    Dim results_emdash As String, results_other As String, results_simplefinds As String, results_spaces As String, results_special As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        results_emdash = TestHelpers.returnTestResultString("TestDoubleQuotes_emdash", MyStoryNo)
        results_other = TestHelpers.returnTestResultString("TestDoubleQuotes_other", MyStoryNo)
        results_simplefinds = TestHelpers.returnTestResultString("TestDoubleQuotes_simplefinds", MyStoryNo)
        results_spaces = TestHelpers.returnTestResultString("TestDoubleQuotes_spaces", MyStoryNo)
        results_special = TestHelpers.returnTestResultString("TestDoubleQuotes_special", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal DQ_emdash_expected, results_emdash
        Assert.areequal DQ_other_expected, results_other
        Assert.areequal DQ_simplefinds_expected, results_simplefinds
        Assert.areequal DQ_spaces_expected, results_spaces
        Assert.areequal DQ_special_expected, results_special
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDoubleQuotes_endnotes() 'TODO Rename test
    Dim results_emdash As String, results_other As String, results_simplefinds As String, results_spaces As String, results_special As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.DoubleQuotes(MyStoryNo)
        results_emdash = TestHelpers.returnTestResultString("TestDoubleQuotes_emdash", MyStoryNo)
        results_other = TestHelpers.returnTestResultString("TestDoubleQuotes_other", MyStoryNo)
        results_simplefinds = TestHelpers.returnTestResultString("TestDoubleQuotes_simplefinds", MyStoryNo)
        results_spaces = TestHelpers.returnTestResultString("TestDoubleQuotes_spaces", MyStoryNo)
        results_special = TestHelpers.returnTestResultString("TestDoubleQuotes_special", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal DQ_emdash_expected, results_emdash
        Assert.areequal DQ_other_expected, results_other
        Assert.areequal DQ_simplefinds_expected, results_simplefinds
        Assert.areequal DQ_spaces_expected, results_spaces
        Assert.areequal DQ_special_expected, results_special
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
        Assert.areequal DQ_other_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_backtick() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSingleQuotes_backtick"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_backtick_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_word_setasides() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSingleQuotes_word_setasides"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_setasides_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_openquo() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSingleQuotes_openquo"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_openquo_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_fandr() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSingleQuotes_fandr"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_fandr_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_secondrun() 'TODO Rename test
    Dim results_fandr As String, results_openquo As String, results_setasides As String, results_backtick As String
    On Error GoTo TestFail
    'Arrange:
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        Call Clean.SingleQuotes(MyStoryNo)
        results_fandr = TestHelpers.returnTestResultString("TestSingleQuotes_fandr", MyStoryNo)
        results_openquo = TestHelpers.returnTestResultString("TestSingleQuotes_openquo", MyStoryNo)
        results_setasides = TestHelpers.returnTestResultString("TestSingleQuotes_word_setasides", MyStoryNo)
        results_backtick = TestHelpers.returnTestResultString("TestSingleQuotes_backtick", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_fandr_expected, results_fandr
        Assert.areequal SQ_openquo_expected, results_openquo
        Assert.areequal SQ_setasides_expected, results_setasides
        Assert.areequal SQ_backtick_expected, results_backtick
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_endnotes() 'TODO Rename test
    Dim results_fandr As String, results_openquo As String, results_setasides As String, results_backtick As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        results_fandr = TestHelpers.returnTestResultString("TestSingleQuotes_fandr", MyStoryNo)
        results_openquo = TestHelpers.returnTestResultString("TestSingleQuotes_openquo", MyStoryNo)
        results_setasides = TestHelpers.returnTestResultString("TestSingleQuotes_word_setasides", MyStoryNo)
        results_backtick = TestHelpers.returnTestResultString("TestSingleQuotes_backtick", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_fandr_expected, results_fandr
        Assert.areequal SQ_openquo_expected, results_openquo
        Assert.areequal SQ_setasides_expected, results_setasides
        Assert.areequal SQ_backtick_expected, results_backtick
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSingleQuotes_footnotes() 'TODO Rename test
    Dim results_fandr As String, results_openquo As String, results_setasides As String, results_backtick As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.SingleQuotes(MyStoryNo)
        results_fandr = TestHelpers.returnTestResultString("TestSingleQuotes_fandr", MyStoryNo)
        results_openquo = TestHelpers.returnTestResultString("TestSingleQuotes_openquo", MyStoryNo)
        results_setasides = TestHelpers.returnTestResultString("TestSingleQuotes_word_setasides", MyStoryNo)
        results_backtick = TestHelpers.returnTestResultString("TestSingleQuotes_backtick", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SQ_fandr_expected, results_fandr
        Assert.areequal SQ_openquo_expected, results_openquo
        Assert.areequal SQ_setasides_expected, results_setasides
        Assert.areequal SQ_backtick_expected, results_backtick
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestEllipses_frbasic() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestEllipses_frbasic"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_frbasic_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestEllipses_fr4dots() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestEllipses_fr4dots"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_fr4dots_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestEllipses_spaces() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestEllipses_spaces"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_spaces_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestEllipses_emdashes() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestEllipses_emdashes"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_emdashes_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestEllipses_secondrun() 'TODO Rename test
    Dim results_frbasic As String, results_fr4dots As String, results_spaces As String, results_emdashes As String
    On Error GoTo TestFail
    'Arrange:
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        Call Clean.Ellipses(MyStoryNo)
        results_frbasic = TestHelpers.returnTestResultString("TestEllipses_frbasic", MyStoryNo)
        results_fr4dots = TestHelpers.returnTestResultString("TestEllipses_fr4dots", MyStoryNo)
        results_spaces = TestHelpers.returnTestResultString("TestEllipses_spaces", MyStoryNo)
        results_emdashes = TestHelpers.returnTestResultString("TestEllipses_emdashes", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_frbasic_expected, results_frbasic
        Assert.areequal EL_fr4dots_expected, results_fr4dots
        Assert.areequal EL_spaces_expected, results_spaces
        Assert.areequal EL_emdashes_expected, results_emdashes
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestEllipses_footnotes() 'TODO Rename test
    Dim results_frbasic As String, results_fr4dots As String, results_spaces As String, results_emdashes As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        results_frbasic = TestHelpers.returnTestResultString("TestEllipses_frbasic", MyStoryNo)
        results_fr4dots = TestHelpers.returnTestResultString("TestEllipses_fr4dots", MyStoryNo)
        results_spaces = TestHelpers.returnTestResultString("TestEllipses_spaces", MyStoryNo)
        results_emdashes = TestHelpers.returnTestResultString("TestEllipses_emdashes", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_frbasic_expected, results_frbasic
        Assert.areequal EL_fr4dots_expected, results_fr4dots
        Assert.areequal EL_spaces_expected, results_spaces
        Assert.areequal EL_emdashes_expected, results_emdashes
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestEllipses_endnotes() 'TODO Rename test
    Dim results_frbasic As String, results_fr4dots As String, results_spaces As String, results_emdashes As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Ellipses(MyStoryNo)
        results_frbasic = TestHelpers.returnTestResultString("TestEllipses_frbasic", MyStoryNo)
        results_fr4dots = TestHelpers.returnTestResultString("TestEllipses_fr4dots", MyStoryNo)
        results_spaces = TestHelpers.returnTestResultString("TestEllipses_spaces", MyStoryNo)
        results_emdashes = TestHelpers.returnTestResultString("TestEllipses_emdashes", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal EL_frbasic_expected, results_frbasic
        Assert.areequal EL_fr4dots_expected, results_fr4dots
        Assert.areequal EL_spaces_expected, results_spaces
        Assert.areequal EL_emdashes_expected, results_emdashes
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestSpaces_nbsps() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSpaces_nbsps"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_nbsps_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSpaces_brackets_and_whitespace() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSpaces_brackets_and_whitespace"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_brackets_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestSpaces_breaks() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSpaces_breaks"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_breaks_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestSpaces_exclude() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestSpaces_exclude"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_exclude_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestSpaces_secondrun() 'TODO Rename test
    Dim results_nbsps As String, results_brackets As String, results_breaks As String, results_exclude As String
    On Error GoTo TestFail
    'Arrange:
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        Call Clean.Spaces(MyStoryNo)
        results_nbsps = TestHelpers.returnTestResultString("TestSpaces_nbsps", MyStoryNo)
        results_brackets = TestHelpers.returnTestResultString("TestSpaces_brackets_and_whitespace", MyStoryNo)
        results_breaks = TestHelpers.returnTestResultString("TestSpaces_breaks", MyStoryNo)
        results_exclude = TestHelpers.returnTestResultString("TestSpaces_exclude", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_nbsps_expected, results_nbsps
        Assert.areequal SP_brackets_expected, results_brackets
        Assert.areequal SP_breaks_expected, results_breaks
        Assert.areequal SP_exclude_expected, results_exclude
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestSpaces_footnotes() 'TODO Rename test
    Dim results_nbsps As String, results_brackets As String, results_breaks As String, results_exclude As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results_nbsps = TestHelpers.returnTestResultString("TestSpaces_nbsps", MyStoryNo)
        results_brackets = TestHelpers.returnTestResultString("TestSpaces_brackets_and_whitespace", MyStoryNo)
        results_breaks = TestHelpers.returnTestResultString("TestSpaces_breaks", MyStoryNo)
        results_exclude = TestHelpers.returnTestResultString("TestSpaces_exclude", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_nbsps_expected, results_nbsps
        Assert.areequal SP_brackets_expected, results_brackets
        Assert.areequal SP_breaks_expected, results_breaks
        Assert.areequal SP_exclude_expected, results_exclude
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestSpaces_endnotes() 'TODO Rename test
    Dim results_nbsps As String, results_brackets As String, results_breaks As String, results_exclude As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results_nbsps = TestHelpers.returnTestResultString("TestSpaces_nbsps", MyStoryNo)
        results_brackets = TestHelpers.returnTestResultString("TestSpaces_brackets_and_whitespace", MyStoryNo)
        results_breaks = TestHelpers.returnTestResultString("TestSpaces_breaks", MyStoryNo)
        results_exclude = TestHelpers.returnTestResultString("TestSpaces_exclude", MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal SP_nbsps_expected, results_nbsps
        Assert.areequal SP_brackets_expected, results_brackets
        Assert.areequal SP_breaks_expected, results_breaks
        Assert.areequal SP_exclude_expected, results_exclude
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestPunctuation() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestPunctuation"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Punctuation(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal PN_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestPunctuation_notes_and_secondrun() 'TODO Rename test
    Dim results_fnotes As String, results_enotes As String, results_secondrun As String
    On Error GoTo TestFail
    'Act:
        Call Clean.Punctuation(MyStoryNo)
        Call Clean.Punctuation(MyStoryNo)
        results_secondrun = TestHelpers.returnTestResultString("TestPunctuation", MyStoryNo)
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.Punctuation(MyStoryNo)
        results_fnotes = TestHelpers.returnTestResultString("TestPunctuation", MyStoryNo)
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Punctuation(MyStoryNo)
        results_enotes = TestHelpers.returnTestResultString("TestPunctuation", MyStoryNo)
     'Assert:
        Assert.Succeed
        Assert.areequal PN_expected, results_fnotes
        Assert.areequal PN_expected, results_enotes
        Assert.areequal PN_expected, results_secondrun
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDashes() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDashes"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Dashes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal DSH_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDashes_exclude() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDashes_exclude"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Dashes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal DSH_exclude_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDashes_numbers() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestDashes_numbers"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Dashes(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal DSH_numbers_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDashes_secondrun() 'TODO Rename test
    Dim results_dashes As String, results_numbers As String, results_exclude As String
    On Error GoTo TestFail
    'Act:
        Call Clean.Dashes(MyStoryNo)
        Call Clean.Dashes(MyStoryNo)
        results_dashes = TestHelpers.returnTestResultString("TestDashes", MyStoryNo)
        results_numbers = TestHelpers.returnTestResultString("TestDashes_numbers", MyStoryNo)
        results_exclude = TestHelpers.returnTestResultString("TestDashes_exclude", MyStoryNo)
     'Assert:
        Assert.Succeed
        Assert.areequal DSH_expected, results_dashes
        Assert.areequal DSH_numbers_expected, results_numbers
        Assert.areequal DSH_exclude_expected, results_exclude
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestDashes_notes() 'TODO Rename test
    Dim results_fnotes As String, results_enotes As String, results_fnotes_numbers As String, results_enotes_numbers As String, _
        results_fnotes_exclude As String, results_enotes_exclude As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.Dashes(MyStoryNo)
        results_fnotes = TestHelpers.returnTestResultString("TestDashes", MyStoryNo)
        results_fnotes_numbers = TestHelpers.returnTestResultString("TestDashes_numbers", MyStoryNo)
        results_fnotes_exclude = TestHelpers.returnTestResultString("TestDashes_exclude", MyStoryNo)
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.Dashes(MyStoryNo)
        results_enotes = TestHelpers.returnTestResultString("TestDashes", MyStoryNo)
        results_enotes_numbers = TestHelpers.returnTestResultString("TestDashes_numbers", MyStoryNo)
        results_enotes_exclude = TestHelpers.returnTestResultString("TestDashes_exclude", MyStoryNo)
     'Assert:
        Assert.Succeed
        Assert.areequal DSH_expected, results_fnotes
        Assert.areequal DSH_numbers_expected, results_fnotes_numbers
        Assert.areequal DSH_exclude_expected, results_fnotes_exclude
        Assert.areequal DSH_expected, results_enotes
        Assert.areequal DSH_numbers_expected, results_enotes_numbers
        Assert.areequal DSH_exclude_expected, results_enotes_exclude
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestBreaks() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestBreaks"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CleanBreaks(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal BR_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestBreaks_notes_and_secondrun() 'TODO Rename test
    Dim results_fnotes As String, results_enotes As String, results_secondrun As String
    Dim testDocx_local As Document
    On Error GoTo TestFail
    'Arrange:
        ' this test requires document be opened with visibility, else it loses track of activedoc.
        '   so we open our own just for this test
        Set testDocx_local = Application.Documents.Add(testdotx_filepath)
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.CleanBreaks(MyStoryNo)
        results_fnotes = TestHelpers.returnTestResultString("TestBreaks", MyStoryNo)
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.CleanBreaks(MyStoryNo)
        results_enotes = TestHelpers.returnTestResultString("TestBreaks", MyStoryNo)
     'Act:
        Call Clean.CleanBreaks(MyStoryNo)
        Call Clean.CleanBreaks(MyStoryNo)
        results_secondrun = TestHelpers.returnTestResultString("TestBreaks", MyStoryNo)
     'Assert:
        Assert.Succeed
        Assert.areequal BR_expected, results_fnotes
        Assert.areequal BR_expected, results_enotes
        Assert.areequal BR_expected, results_secondrun
     'Cleanup:
        testDocx_local.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestTitlecase() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestTitlecase"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.MakeTitleCase(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal MTC_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro")
Private Sub TestTitlecase_notes_and_secondrun() 'TODO Rename test
    Dim results_fnotes As String, results_enotes As String, results_secondrun As String
    On Error GoTo TestFail
    'Arrange (setup for footnotes):
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToFootNotes
    'Act:
        Call Clean.MakeTitleCase(MyStoryNo)
        results_fnotes = TestHelpers.returnTestResultString("TestTitlecase", MyStoryNo)
    'Arrange (setup for endnotes):
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        copyBodyContentsToEndNotes
    'Act:
        Call Clean.MakeTitleCase(MyStoryNo)
        results_enotes = TestHelpers.returnTestResultString("TestTitlecase", MyStoryNo)
    'Act (main, second run):
        Call Clean.MakeTitleCase(MyStoryNo)
        Call Clean.MakeTitleCase(MyStoryNo)
        results_secondrun = TestHelpers.returnTestResultString("TestTitlecase", MyStoryNo)
     
     'Assert:
        Assert.Succeed
        Assert.areequal MTC_expected, results_fnotes
        Assert.areequal MTC_expected, results_enotes
        Assert.areequal MTC_expected, results_secondrun
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro")
Private Sub TestHyperlinks() 'TODO Rename test
    Dim init_link_count As Integer, final_link_count As Integer, second_run_count As Integer, init_link_count_fn As Integer, _
        final_link_count_fn As Integer, init_link_count_en As Integer, final_link_count_en As Integer
    On Error GoTo TestFail
    'Arrange:
        copyBodyContentsToFootNotes
        copyBodyContentsToEndNotes
    ''' test main story and second run
    'Act:
        init_link_count = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
        Call Clean.RemoveHyperlinks(MyStoryNo)
        final_link_count = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
        Call Clean.RemoveHyperlinks(MyStoryNo)
        second_run_count = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
    ''' test footnotes
    'Arrange:
        MyStoryNo = 2
    'Act:
        init_link_count_fn = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
        Call Clean.RemoveHyperlinks(MyStoryNo)
        final_link_count_fn = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
    ''' test endnotes
    'Arrange:
        MyStoryNo = 3
    'Act:
        init_link_count_en = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
        Call Clean.RemoveHyperlinks(MyStoryNo)
        final_link_count_en = ActiveDocument.StoryRanges(MyStoryNo).Hyperlinks.Count
       
    'Assert:
        Assert.Succeed
        Assert.areequal init_link_count, HL_expected
        Assert.areequal final_link_count, 0
        Assert.areequal second_run_count, 0
        Assert.areequal init_link_count_fn, HL_expected
        Assert.areequal final_link_count_fn, 0
        Assert.areequal init_link_count_en, HL_expected
        Assert.areequal final_link_count_en, 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
