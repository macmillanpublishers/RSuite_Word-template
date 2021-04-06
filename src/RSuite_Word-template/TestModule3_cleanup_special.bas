Attribute VB_Name = "TestModule3_cleanup_special"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private BR_sbreaks_expected As String
Private BM_expected As Integer, CM_expected As Integer, TC_main_expected As Integer, TC_en_expected As Integer, _
    TC_fn_expected As Integer, OB_shapes_expected As Integer, OB_ishapes_expected As Integer, OB_frames_expected As Integer, _
    OB_en_ishapes_expected As Integer, OB_fn_ishapes_expected As Integer, SP_fn_control_expected As String, _
    SP_fn_problem_expected As String, SP_en_control_expected As String, SP_en_problem_expected As String, _
    SP_fn_num_c As Long, SP_fn_num_p As Long, SP_en_num_c As Long, SP_en_num_p As Long, SP_en_num_b As Long, _
    SP_fn_num_b As Long, SP_fn_blank_expected As String, SP_en_blank_expected As String, CFn_num_c As Long, _
    CFn_num_b As Long, CEn_num_c As Long, CEn_num_b As Long
Private testDocx As Document
Private testdotx_filepath As String
Private testdotx As String
Private MyStoryNo As Variant

Private Function SetResultStrings()

BM_expected = 9
CM_expected = 6
TC_main_expected = 14
TC_en_expected = 5
TC_fn_expected = 7
BR_sbreaks_expected = "Continuous section" + vbCr _
    + "break. Now nextpage section" + vbCr _
    + "break. Now evenpage section" + vbCr _
    + "break."
OB_shapes_expected = 3
OB_ishapes_expected = 3
OB_frames_expected = 5
OB_en_ishapes_expected = 2
OB_fn_ishapes_expected = 3
SP_fn_num_c = 6
SP_fn_control_expected = "Control footnote. A little extra white space, but no trimmable ws." + vbCr + "Same deal, 2nd paragraph."
SP_fn_num_p = 7
SP_fn_problem_expected = "Bad footnote." + vbCr + "Extra spaces everywhere."
SP_fn_num_b = 8
SP_fn_blank_expected = ""
SP_en_num_c = 7
SP_en_control_expected = "Control Endnote. Some random extra white space." + vbCr + "Second para. More of same."
SP_en_num_p = 8
SP_en_problem_expected = "Bad Endnote. Leading and trailing spaces." + vbCr + "For whole note! Not ideal?"
SP_en_num_b = 9
SP_en_blank_expected = ""
CFn_num_b = 10
CFn_num_c = 9
CEn_num_b = 10
CEn_num_c = 11

End Function

Private Function DestroyResultStrings()

'EX_example_expected = vbNullString
BM_expected = 0
CM_expected = 0
TC_main_expected = 0
TC_en_expected = 0
TC_fn_expected = 0
BR_sbreaks_expected = vbNullString
OB_shapes_expected = 0
OB_ishapes_expected = 0
OB_frames_expected = 0
OB_en_ishapes_expected = 0
OB_fn_ishapes_expected = 0
SP_fn_num_c = 0
SP_fn_control_expected = vbNullString
SP_fn_num_p = 0
SP_fn_problem_expected = vbNullString
SP_en_num_b = 0
SP_fn_blank_expected = vbNullString
SP_en_num_c = 0
SP_en_control_expected = vbNullString
SP_en_num_p = 0
SP_en_problem_expected = vbNullString
SP_en_num_b = 0
SP_en_blank_expected = vbNullString
CFn_num_b = 0
CFn_num_c = 0
CEn_num_b = 0
CEn_num_c = 0

End Function

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Get testdot filepath.
    testdotx_filepath = getRepoPath + "test_files\testfile_cleanup_special.dotx"
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
    'Unload pBar
    DestroyCharacters
    DestroyResultStrings
    Application.ScreenUpdating = True
    'MsgBox ("Cleanup Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Create new test docx from template
    ' (visible version for debug):
    'Set testDocx = Application.Documents.Add(testdotx_filepath)
    Set testDocx = Application.Documents.Add(testdotx_filepath, visible:=False)
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
'@TestMethod("CleanupMacro_special")
Private Sub TestBookmarks() 'TODO Rename test
    Dim init_bm_count As Integer, final_bm_count As Integer, second_run_count As Integer
    On Error GoTo TestFail
    'Act:
        init_bm_count = testDocx.Bookmarks.Count
        Call Clean.DeleteBookmarks
        final_bm_count = testDocx.Bookmarks.Count
        Call Clean.DeleteBookmarks
        second_run_count = testDocx.Bookmarks.Count
    'Assert:
        Assert.Succeed
        Assert.areequal init_bm_count, BM_expected
        Assert.areequal final_bm_count, 0
        Assert.areequal second_run_count, 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestTrackchanges_msgbox_y_and_secondrun() 'TODO Rename test
    Dim init_tc_count As Integer, final_tc_count As Integer, second_run_count As Integer
    On Error GoTo TestFail
    'Arrange
        Fakes.MsgBox.Returns vbYes
    'Act:
        init_tc_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        With Fakes.MsgBox.verify
            .Parameter "title", "ACCEPT TRACK CHANGES"
            .Parameter "buttons", vbYesNo
        End With
        final_tc_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        second_run_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
    'Assert:
        Assert.Succeed
        Assert.areequal init_tc_count, TC_main_expected + TC_fn_expected + TC_en_expected
        Assert.areequal final_tc_count, 0
        Assert.areequal second_run_count, 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestTrackchanges_none() 'TODO Rename test
    Dim init_tc_count As Integer, final_tc_count As Integer
    On Error GoTo TestFail
    'Arrange
        testDocx.StoryRanges(1).Revisions.AcceptAll
        testDocx.StoryRanges(2).Revisions.AcceptAll
        testDocx.StoryRanges(3).Revisions.AcceptAll
    'Act:
        init_tc_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        final_tc_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
      'Assert:
        Assert.Succeed
        Assert.areequal init_tc_count, 0
        Assert.areequal final_tc_count, 0
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestTrackchanges_msgbox_n() 'TODO Rename test
    Dim init_tc_count As Integer, final_tc_count As Integer
    On Error GoTo TestFail
    'Arrange
        Fakes.MsgBox.Returns vbNo
    'Act:
        init_tc_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        With Fakes.MsgBox.verify
            .Parameter "title", "ACCEPT TRACK CHANGES"
            .Parameter "buttons", vbYesNo
        End With
        final_tc_count = testDocx.StoryRanges(1).Revisions.Count + _
            testDocx.StoryRanges(2).Revisions.Count + _
            testDocx.StoryRanges(3).Revisions.Count
      'Assert:
        Assert.Succeed
        Assert.areequal init_tc_count, TC_main_expected + TC_fn_expected + TC_en_expected
        Assert.areequal final_tc_count, init_tc_count
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestComments_msgbox_no() 'TODO Rename test
    Dim init_c_count As Integer, final_c_count As Integer
    On Error GoTo TestFail
    'Arrange
        Fakes.MsgBox.Returns vbNo
    'Act:
        init_c_count = testDocx.Comments.Count
        Call Clean.RemoveComments
        With Fakes.MsgBox.verify
            .Parameter "title", "DELETE COMMENTS"
            .Parameter "buttons", vbYesNo
        End With
        final_c_count = testDocx.Comments.Count
    'Assert:
        Assert.Succeed
        Assert.areequal init_c_count, CM_expected
        Assert.areequal final_c_count, CM_expected
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestComments_none() 'TODO Rename test
    Dim init_c_count As Integer, final_c_count As Integer
    On Error GoTo TestFail
    'Arrange
        testDocx.DeleteAllComments
    'Act:
        init_c_count = testDocx.Comments.Count
        Call Clean.RemoveComments
        final_c_count = testDocx.Comments.Count
    'Assert:
        Assert.Succeed
        Assert.areequal init_c_count, 0
        Assert.areequal final_c_count, 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestComments_msgbox_y_and_secondrun() 'TODO Rename test
    Dim init_c_count As Integer, final_c_count As Integer, second_run_count As Integer
    On Error GoTo TestFail
    'Arrange
        Fakes.MsgBox.Returns vbYes
    'Act
        init_c_count = testDocx.Comments.Count
        Call Clean.RemoveComments
        With Fakes.MsgBox.verify
            .Parameter "title", "DELETE COMMENTS"
            .Parameter "buttons", vbYesNo
        End With
        final_c_count = testDocx.Comments.Count
        Call Clean.RemoveComments
        second_run_count = testDocx.Comments.Count
    'Assert:
        Assert.Succeed
        Assert.areequal init_c_count, CM_expected
        Assert.areequal final_c_count, 0
        Assert.areequal second_run_count, 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestBreaks_sectionbreaks() 'TODO Rename test
    Dim results As String
    On Error GoTo TestFail
    'Arrange:
        Const C_PROC_NAME = "TestBreaks_sectionbreaks"  '<-- name of this test procedure
        'MyStoryNo = 1 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.CleanBreaks(MyStoryNo)
        results = TestHelpers.returnTestResultString(C_PROC_NAME, MyStoryNo)
    'Assert:
        Assert.Succeed
        Assert.areequal BR_sbreaks_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestSpaces_note_trim() 'TODO Rename test
    ' fn_c = control note, fn_p = problem, fn_b = blank
    Dim results_fn_c As String, results_fn_p As String, results_en_c As String, results_en_p As String, _
        results_fn_b As String, results_en_b As String
    Dim testDocx_local As Document
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
        ' this test requires document be opened with visibility, else it loses track of activedoc.
        '   so we open our own just for this test
        Set testDocx_local = Application.Documents.Add(testdotx_filepath)
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results_fn_c = testDocx_local.Footnotes(SP_fn_num_c).Range
        results_fn_p = testDocx_local.Footnotes(SP_fn_num_p).Range
        results_fn_b = testDocx_local.Footnotes(SP_fn_num_b).Range
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results_en_c = testDocx_local.Endnotes(SP_en_num_c).Range
        results_en_p = testDocx_local.Endnotes(SP_en_num_p).Range
        results_en_b = testDocx_local.Endnotes(SP_en_num_b).Range
    'Assert:
        Assert.Succeed
        Assert.areequal SP_fn_control_expected, results_fn_c
        Assert.areequal SP_fn_problem_expected, results_fn_p
        Assert.areequal SP_fn_blank_expected, results_fn_b
        Assert.areequal SP_en_control_expected, results_en_c
        Assert.areequal SP_en_problem_expected, results_en_p
        Assert.areequal SP_en_blank_expected, results_en_b
    'Cleanup:
        testDocx_local.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestCustomFootnoteMark_fix() 'TODO Rename test
    ' fn_c = note w. contents, fn_b = blank
    Dim result_fn_c As Range, init_ctrlRng As Range, final_ctrlRng As Range, init_fn_c As Range
    Dim result_fn_b As String, init_ctrl_refStr As String, result_fn_refStr As String, _
        final_ctrl_refStr As String, result_fn_refStr_b As String, result_compareStr As String, _
        result_compareCtrlStr As String
    Dim testDocx_local As Document, origDocx_local As Document
    On Error GoTo TestFail
    'Arrange:
        ' have to open two documents; otherwise the range no longer exists after 'fix' function is called
        Set origDocx_local = Application.Documents.Add(testdotx_filepath, visible:=False)
        Set init_fn_c = origDocx_local.Footnotes(CFn_num_c).Range
        ' using reg. footnote from another test as a control.
        Set init_ctrlRng = origDocx_local.Footnotes(SP_fn_num_c).Range
        init_ctrl_refStr = origDocx_local.Footnotes(SP_fn_num_c).Reference.Text
        Set testDocx_local = Application.Documents.Add(testdotx_filepath)
    'Act:
        Call Clean.fixCustomFootnotes
        Set result_fn_c = testDocx_local.Footnotes(CFn_num_c).Range
        result_fn_refStr = testDocx_local.Footnotes(CFn_num_c).Reference.Text
        result_fn_b = testDocx_local.Footnotes(CFn_num_b).Range
        result_fn_refStr_b = testDocx_local.Footnotes(CFn_num_b).Reference.Text
        Set final_ctrlRng = testDocx_local.Footnotes(SP_fn_num_c).Range
        final_ctrl_refStr = testDocx_local.Footnotes(SP_fn_num_c).Reference.Text
        ' Compare after and before ranges, including charstyles
        result_compareStr = TestHelpers.compareRanges(result_fn_c, init_fn_c)
        result_compareCtrlStr = TestHelpers.compareRanges(final_ctrlRng, init_ctrlRng)
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
        Assert.areequal " ", result_fn_b         ' blank notes should get a single space inserted for content on recreation.
        Assert.areequal result_fn_refStr, Chr(2) ' Chr(2) is the sys default for auto-increment note-mark
        Assert.areequal result_fn_refStr_b, Chr(2)
        ' checking control note
        Assert.areequal "Same", result_compareCtrlStr
        Assert.areequal init_ctrl_refStr, Chr(2)
        Assert.areequal final_ctrl_refStr, init_ctrl_refStr
    'Cleanup:
        testDocx_local.Close savechanges:=wdDoNotSaveChanges
        origDocx_local.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestCustomEndnoteMark_fix() 'TODO Rename test
    ' en_c = note w. contents, en_b = blank
    Dim result_en_c As Range, init_ctrlRng As Range, final_ctrlRng As Range, init_en_c As Range
    Dim result_en_b As String, init_ctrl_refStr As String, result_en_refStr As String, _
        result_en_refStr_b As String, final_ctrl_refStr As String, result_compareStr As String, _
        result_compareCtrlStr As String
    Dim testDocx_local As Document, origDocx_local As Document
    On Error GoTo TestFail
    'Arrange:
        ' have to open two documents; otherwise the range no longer exists after 'fix' function is called
        Set origDocx_local = Application.Documents.Add(testdotx_filepath, visible:=False)
        Set init_en_c = origDocx_local.Endnotes(CEn_num_c).Range
        ' using reg. endnote from another test as a control.
        Set init_ctrlRng = origDocx_local.Endnotes(SP_en_num_p).Range
        init_ctrl_refStr = origDocx_local.Endnotes(SP_en_num_p).Reference.Text
        Set testDocx_local = Application.Documents.Add(testdotx_filepath)
    'Act:
        Call Clean.fixCustomEndnotes
        Set result_en_c = testDocx_local.Endnotes(CEn_num_c).Range
        result_en_refStr = testDocx_local.Endnotes(CEn_num_c).Reference.Text
        result_en_b = testDocx_local.Endnotes(CEn_num_b).Range
        result_en_refStr_b = testDocx_local.Endnotes(CEn_num_b).Reference.Text
        Set final_ctrlRng = testDocx_local.Endnotes(SP_en_num_p).Range
        final_ctrl_refStr = testDocx_local.Endnotes(SP_en_num_p).Reference.Text
        ' Compare after and before ranges, including charstyles
        result_compareStr = TestHelpers.compareRanges(result_en_c, init_en_c)
        result_compareCtrlStr = TestHelpers.compareRanges(final_ctrlRng, init_ctrlRng)
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", result_compareStr
        Assert.areequal " ", result_en_b         ' blank notes should get a single space inserted for content on recreation.
        Assert.areequal result_en_refStr, Chr(2) ' Chr(2) is the sys default for auto-increment note-mark
        Assert.areequal result_en_refStr_b, Chr(2)
        ' checking control note
        Assert.areequal "Same", result_compareCtrlStr
        Assert.areequal init_ctrl_refStr, Chr(2)
        Assert.areequal final_ctrl_refStr, init_ctrl_refStr
        ' check second run reults
    'Cleanup:
        testDocx_local.Close savechanges:=wdDoNotSaveChanges
        origDocx_local.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CleanupMacro_special")
Private Sub TestCustomNoteMark_fix_secondRun() 'TODO Rename test
    ' fn_c = note w. contents, fn_b = blank
    Dim result_en_c As Range, init_en_c As Range, result_fn_c As Range, init_fn_c As Range
    Dim result_fn_b As String, result_en_b As String, result_fn_refStr As String, _
        result_fn_refStr_b As String, result_en_refStr As String, result_en_refStr_b As String, _
        result_compareStr_en As String, result_compareStr_fn As String
    Dim testDocx_local As Document, origDocx_local As Document
    On Error GoTo TestFail
    'Arrange:
        ' have to open two documents; otherwise the range no longer exists after 'fix' function is called
        Set origDocx_local = Application.Documents.Add(testdotx_filepath, visible:=False)
        Set init_fn_c = origDocx_local.Footnotes(CFn_num_c).Range
        Set init_en_c = origDocx_local.Endnotes(CEn_num_c).Range
        Set testDocx_local = Application.Documents.Add(testdotx_filepath)
    'Act: (footnotes)
        Call Clean.fixCustomFootnotes
        Call Clean.fixCustomFootnotes
        Set result_fn_c = testDocx_local.Footnotes(CFn_num_c).Range
        result_fn_refStr = testDocx_local.Footnotes(CFn_num_c).Reference.Text
        result_fn_b = testDocx_local.Footnotes(CFn_num_b).Range
        result_fn_refStr_b = testDocx_local.Footnotes(CFn_num_b).Reference.Text
        ' Compare after and before ranges, including charstyles
        result_compareStr_fn = TestHelpers.compareRanges(result_fn_c, init_fn_c)
    'Act: (endnotes)
        Call Clean.fixCustomEndnotes
        Call Clean.fixCustomEndnotes
        Set result_en_c = testDocx_local.Endnotes(CEn_num_c).Range
        result_en_refStr = testDocx_local.Endnotes(CEn_num_c).Reference.Text
        result_en_b = testDocx_local.Endnotes(CEn_num_b).Range
        result_en_refStr_b = testDocx_local.Endnotes(CEn_num_b).Reference.Text
        ' Compare after and before ranges, including charstyles
        result_compareStr_en = TestHelpers.compareRanges(result_en_c, init_en_c)
    'Assert:
        Assert.Succeed
        ' checking footnotes
        Assert.areequal "Same", result_compareStr_fn
        Assert.areequal " ", result_fn_b         ' blank notes should get a single space inserted for content on recreation.
        Assert.areequal result_fn_refStr, Chr(2) ' Chr(2) is the sys default for auto-increment note-mark
        Assert.areequal result_fn_refStr_b, Chr(2)
        ' checking endnotes
        Assert.areequal "Same", result_compareStr_en
        Assert.areequal " ", result_en_b         ' blank notes should get a single space inserted for content on recreation.
        Assert.areequal result_en_refStr, Chr(2) ' Chr(2) is the sys default for auto-increment note-mark
        Assert.areequal result_en_refStr_b, Chr(2)
    'Cleanup:
      testDocx_local.Close savechanges:=wdDoNotSaveChanges
      origDocx_local.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestObjects_secondrun() 'TODO Rename test
    Dim init_shape_count As Integer, init_frame_count As Integer, init_ishape_count As Integer, _
        final_shape_count As Integer, final_frame_count As Integer, final_ishape_count As Integer
    On Error GoTo TestFail
    'Arrange:
    'Act:
        init_shape_count = testDocx.Shapes.Count
        init_frame_count = testDocx.StoryRanges(1).Frames.Count
        init_ishape_count = testDocx.StoryRanges(1).InlineShapes.Count
        Call Clean.DeleteObjects(1)
        Call Clean.DeleteObjects(1)
        final_shape_count = testDocx.Shapes.Count
        final_frame_count = testDocx.StoryRanges(1).Frames.Count
        final_ishape_count = testDocx.StoryRanges(1).InlineShapes.Count
    'Assert:
        Assert.Succeed
        Assert.areequal OB_shapes_expected, init_shape_count
        Assert.areequal OB_frames_expected, init_frame_count
        Assert.areequal OB_ishapes_expected, init_ishape_count
        Assert.areequal 0, final_shape_count
        Assert.areequal 0, final_frame_count
        Assert.areequal 0, final_ishape_count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestObjects_notes() 'TODO Rename test
    Dim init_shape_count As Integer, init_frame_count As Integer, init_ishape_count As Integer, _
        final_shape_count As Integer, final_frame_count As Integer, final_ishape_count As Integer
    On Error GoTo TestFail
    'Arrange:
    'Act:
        init_shape_count = testDocx.Shapes.Count
        init_frame_count = testDocx.StoryRanges(2).Frames.Count _
            + testDocx.StoryRanges(3).Frames.Count
        init_ishape_count = testDocx.StoryRanges(2).InlineShapes.Count _
            + testDocx.StoryRanges(3).InlineShapes.Count
        Call Clean.DeleteObjects(2)
        Call Clean.DeleteObjects(3)
        final_shape_count = testDocx.Shapes.Count
        final_frame_count = testDocx.StoryRanges(2).Frames.Count _
            + testDocx.StoryRanges(3).Frames.Count
        final_ishape_count = testDocx.StoryRanges(2).InlineShapes.Count _
            + testDocx.StoryRanges(3).InlineShapes.Count
    'Assert:
        Assert.Succeed
        Assert.areequal OB_shapes_expected, init_shape_count
        Assert.areequal 0, init_frame_count 'Was unable to create Frame in notes.
        Assert.areequal OB_en_ishapes_expected + OB_fn_ishapes_expected, init_ishape_count
        Assert.areequal 0, final_shape_count
        Assert.areequal 0, final_frame_count
        Assert.areequal 0, final_ishape_count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestObjects() 'TODO Rename test
    Dim init_shape_count As Integer, init_frame_count As Integer, init_ishape_count As Integer, _
        final_shape_count As Integer, final_frame_count As Integer, final_ishape_count As Integer
    On Error GoTo TestFail
    'Arrange:
    'Act:
        init_shape_count = testDocx.Shapes.Count
        init_frame_count = testDocx.StoryRanges(1).Frames.Count
        init_ishape_count = testDocx.StoryRanges(1).InlineShapes.Count
        Call Clean.DeleteObjects(1)
        final_shape_count = testDocx.Shapes.Count
        final_frame_count = testDocx.StoryRanges(1).Frames.Count
        final_ishape_count = testDocx.StoryRanges(1).InlineShapes.Count
    'Assert:
        Assert.Succeed
        Assert.areequal OB_shapes_expected, init_shape_count
        Assert.areequal OB_frames_expected, init_frame_count
        Assert.areequal OB_ishapes_expected, init_ishape_count
        Assert.areequal 0, final_shape_count
        Assert.areequal 0, final_frame_count
        Assert.areequal 0, final_ishape_count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


