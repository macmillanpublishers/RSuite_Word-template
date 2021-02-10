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
    SP_fn_num_c As Long, SP_fn_num_p As Long, SP_en_num_c As Long, SP_en_num_p As Long
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
SP_en_num_c = 7
SP_en_control_expected = "Control Endnote. Some random extra white space." + vbCr + "Second para. More of same."
SP_en_num_p = 8
SP_en_problem_expected = "Bad Endnote. Leading and trailing spaces." + vbCr + "For whole note! Not ideal?"

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
SP_en_num_c = 0
SP_en_control_expected = vbNullString
SP_en_num_p = 0
SP_en_problem_expected = vbNullString

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
    Set pBar = New Progress_Bar
    pBarCounter = 0
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
'@TestMethod("CleanupMacro_special")
Private Sub TestBookmarks() 'TODO Rename test
    Dim init_bm_count As Integer, final_bm_count As Integer, second_run_count As Integer
    On Error GoTo TestFail
    'Act:
        init_bm_count = ActiveDocument.Bookmarks.Count
        Call Clean.DeleteBookmarks
        final_bm_count = ActiveDocument.Bookmarks.Count
        Call Clean.DeleteBookmarks
        second_run_count = ActiveDocument.Bookmarks.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual init_bm_count, BM_expected
        Assert.AreEqual final_bm_count, 0
        Assert.AreEqual second_run_count, 0
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
        init_tc_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        With Fakes.MsgBox.Verify
            .Parameter "title", "ACCEPT TRACK CHANGES"
            .Parameter "buttons", vbYesNo
        End With
        final_tc_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        second_run_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual init_tc_count, TC_main_expected + TC_fn_expected + TC_en_expected
        Assert.AreEqual final_tc_count, 0
        Assert.AreEqual second_run_count, 0
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
        ActiveDocument.StoryRanges(1).Revisions.AcceptAll
        ActiveDocument.StoryRanges(2).Revisions.AcceptAll
        ActiveDocument.StoryRanges(3).Revisions.AcceptAll
    'Act:
        init_tc_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        final_tc_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
      'Assert:
        Assert.Succeed
        Assert.AreEqual init_tc_count, 0
        Assert.AreEqual final_tc_count, 0
        
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
        init_tc_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
        Call Clean.RemoveTrackChanges
        With Fakes.MsgBox.Verify
            .Parameter "title", "ACCEPT TRACK CHANGES"
            .Parameter "buttons", vbYesNo
        End With
        final_tc_count = ActiveDocument.StoryRanges(1).Revisions.Count + _
            ActiveDocument.StoryRanges(2).Revisions.Count + _
            ActiveDocument.StoryRanges(3).Revisions.Count
      'Assert:
        Assert.Succeed
        Assert.AreEqual init_tc_count, TC_main_expected + TC_fn_expected + TC_en_expected
        Assert.AreEqual final_tc_count, init_tc_count
        
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
        init_c_count = ActiveDocument.Comments.Count
        Call Clean.RemoveComments
        With Fakes.MsgBox.Verify
            .Parameter "title", "DELETE COMMENTS"
            .Parameter "buttons", vbYesNo
        End With
        final_c_count = ActiveDocument.Comments.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual init_c_count, CM_expected
        Assert.AreEqual final_c_count, CM_expected
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
        ActiveDocument.DeleteAllComments
    'Act:
        init_c_count = ActiveDocument.Comments.Count
        Call Clean.RemoveComments
        final_c_count = ActiveDocument.Comments.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual init_c_count, 0
        Assert.AreEqual final_c_count, 0
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
        init_c_count = ActiveDocument.Comments.Count
        Call Clean.RemoveComments
        With Fakes.MsgBox.Verify
            .Parameter "title", "DELETE COMMENTS"
            .Parameter "buttons", vbYesNo
        End With
        final_c_count = ActiveDocument.Comments.Count
        Call Clean.RemoveComments
        second_run_count = ActiveDocument.Comments.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual init_c_count, CM_expected
        Assert.AreEqual final_c_count, 0
        Assert.AreEqual second_run_count, 0
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
        Assert.AreEqual BR_sbreaks_expected, results
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CleanupMacro_special")
Private Sub TestSpaces_note_trim() 'TODO Rename test
    Dim results_fn_c As String, results_fn_p As String, results_en_c As String, results_en_p As String
    On Error GoTo TestFail
    'Arrange:
        MyStoryNo = 2 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results_fn_c = ActiveDocument.Footnotes(SP_fn_num_c).Range
        results_fn_p = ActiveDocument.Footnotes(SP_fn_num_p).Range
    'Arrange:
        MyStoryNo = 3 '<< override test_init here as needed: use 1 for Main body of docx: use 2 for footnotes, 3 for endnotes
    'Act:
        Call Clean.Spaces(MyStoryNo)
        results_en_c = ActiveDocument.Endnotes(SP_en_num_c).Range
        results_en_p = ActiveDocument.Endnotes(SP_en_num_p).Range
    'Assert:
        Assert.Succeed
        Assert.AreEqual SP_fn_control_expected, results_fn_c
        Assert.AreEqual SP_fn_problem_expected, results_fn_p
        Assert.AreEqual SP_en_control_expected, results_en_c
        Assert.AreEqual SP_en_problem_expected, results_en_p
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
        init_shape_count = ActiveDocument.Shapes.Count
        init_frame_count = ActiveDocument.StoryRanges(1).Frames.Count
        init_ishape_count = ActiveDocument.StoryRanges(1).InlineShapes.Count
        Call Clean.DeleteObjects(1)
        Call Clean.DeleteObjects(1)
        final_shape_count = ActiveDocument.Shapes.Count
        final_frame_count = ActiveDocument.StoryRanges(1).Frames.Count
        final_ishape_count = ActiveDocument.StoryRanges(1).InlineShapes.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual OB_shapes_expected, init_shape_count
        Assert.AreEqual OB_frames_expected, init_frame_count
        Assert.AreEqual OB_ishapes_expected, init_ishape_count
        Assert.AreEqual 0, final_shape_count
        Assert.AreEqual 0, final_frame_count
        Assert.AreEqual 0, final_ishape_count
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
        init_shape_count = ActiveDocument.Shapes.Count
        init_frame_count = ActiveDocument.StoryRanges(2).Frames.Count _
            + ActiveDocument.StoryRanges(3).Frames.Count
        init_ishape_count = ActiveDocument.StoryRanges(2).InlineShapes.Count _
            + ActiveDocument.StoryRanges(3).InlineShapes.Count
        Call Clean.DeleteObjects(2)
        Call Clean.DeleteObjects(3)
        final_shape_count = ActiveDocument.Shapes.Count
        final_frame_count = ActiveDocument.StoryRanges(2).Frames.Count _
            + ActiveDocument.StoryRanges(3).Frames.Count
        final_ishape_count = ActiveDocument.StoryRanges(2).InlineShapes.Count _
            + ActiveDocument.StoryRanges(3).InlineShapes.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual OB_shapes_expected, init_shape_count
        Assert.AreEqual 0, init_frame_count 'Was unable to create Frame in notes.
        Assert.AreEqual OB_en_ishapes_expected + OB_fn_ishapes_expected, init_ishape_count
        Assert.AreEqual 0, final_shape_count
        Assert.AreEqual 0, final_frame_count
        Assert.AreEqual 0, final_ishape_count
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
        init_shape_count = ActiveDocument.Shapes.Count
        init_frame_count = ActiveDocument.StoryRanges(1).Frames.Count
        init_ishape_count = ActiveDocument.StoryRanges(1).InlineShapes.Count
        Call Clean.DeleteObjects(1)
        final_shape_count = ActiveDocument.Shapes.Count
        final_frame_count = ActiveDocument.StoryRanges(1).Frames.Count
        final_ishape_count = ActiveDocument.StoryRanges(1).InlineShapes.Count
    'Assert:
        Assert.Succeed
        Assert.AreEqual OB_shapes_expected, init_shape_count
        Assert.AreEqual OB_frames_expected, init_frame_count
        Assert.AreEqual OB_ishapes_expected, init_ishape_count
        Assert.AreEqual 0, final_shape_count
        Assert.AreEqual 0, final_frame_count
        Assert.AreEqual 0, final_ishape_count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


