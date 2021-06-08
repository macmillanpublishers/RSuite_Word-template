Attribute VB_Name = "TestModule4_CIP"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private testDocx As Document
Private goodTestFilePath As String, nostyleTestFilePath As String

Const sectionFileBasename As String = "sections.txt"
Const breakFileBasename As String = "breaks.txt"
Private tpName As String, cpName As String, spName As String, tocName As String, chName As String
Private tpTag As String, cpTag As String, spTag As String, tocTag As String, chTag As String
Private tpDisplayName As String, cpDisplayName As String, spDisplayName As String, tocDisplayName As String
Private tpRequired As Boolean, cpRequired As Boolean, spRequired As Boolean, tocRequired As Boolean
Private tagArray(), tagNameArray(), tagDisplayNameArray(), tagRequiredArray()
Private sectionArray, bmStyleArray, chStyleArray, breakArray, chNamesArray, bmSectionArray

Private Function SetVariables()
    ' these Name descriptors match names in "sectionFile".
    ' tags as laid out by Library of Congress
    tpName = "Titlepage"    ' tp
    tpTag = "tp"
    tpDisplayName = tpName
    tpRequired = True
    cpName = "Copyright"    ' cp
    cpTag = "cp"
    cpDisplayName = cpName & " Page"
    cpRequired = True
    spName = "Series Page"  ' sp
    spTag = "sp"
    spDisplayName = spName
    spRequired = False
    tocName = "Contents"    ' toc
    tocTag = "toc"
    tocDisplayName = "Table of " & tocName
    tocRequired = False
    ' using "parallel" arrays here instead of multidimensional, since there are finite
    '   items, unlikely to change: just need to make sure the below 3 arrays line up
    tagArray = Array(tpTag, cpTag, spTag, tocTag)    'not including chTag
    tagNameArray = Array(tpName, cpName, spName, tocName)
    tagDisplayNameArray = Array(tpDisplayName, cpDisplayName, spDisplayName, tocDisplayName)
    tagRequiredArray = Array(tpRequired, cpRequired, spRequired, tocRequired)

    chNamesArray = Array("Chapter", "Alt Chapter")
    chTag = "ch"
    ' these backmatter strings match names in "sectionFile".
    bmSectionArray = Array("About the Author", _
        "Acknowledgments", _
        "Afterword", _
        "Appendix", _
        "Back Ad", _
        "Back Matter General", _
        "Bibliography", _
        "Conclusion", _
        "Excerpt Chapter", _
        "Excerpt Opener", _
        "Illustration Credits", _
        "Permissions")
End Function

Private Function DestroyResultStrings()

'EX_example_expected = vbNullString


End Function

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Get test file paths
    goodTestFilePath = getRepoPath + "test_files\testfile_CIP_good.dotx"
    nostyleTestFilePath = getRepoPath + "test_files\testfile_CIP_nostyles.dotx"
    
    Call SetVariables
    ' get style arrays
    chStyleArray = CIPmacro.getMultiSectionStyleNames(chNamesArray)
    sectionArray = CIPmacro.getStyleArrayfromFile(sectionFileBasename)
    bmStyleArray = getMultiSectionStyleNames(bmSectionArray)
    breakArray = getStyleArrayfromFile(breakFileBasename)
    
    Application.ScreenUpdating = False
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'reset loaded public vars
    'Unload pBar
    'DestroyCharacters
    DestroyResultStrings
    Application.ScreenUpdating = True
    'MsgBox ("CIP Macro tests complete")
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' Create new test docx from template
    Set testDocx = Nothing
   ' Set activedoc = Nothing
    'MyStoryNo = 1 '1 = Main Body, 2 = Footnotes, 3 = Endnotes. Can override this value per test as needed
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    'Application.Documents(testDocx).Close savechanges:=wdDoNotSaveChanges
    Set testDocx = Nothing
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagFM() 'TODO Rename test
    Dim styleName As String, openTagCheck As String, closeTagCheck As String
    Dim openCapsCheck As Boolean, closeCapsCheck As Boolean
    On Error GoTo TestFail
    'Arrange:
        Set testDocx = Nothing
        'Set testDocx = Application.Documents.Add(Template:=goodTestFilePath, visible:=False)
        ' for debug / regview:
        Set testDocx = Application.Documents.Add(Template:=goodTestFilePath)
        styleName = CIPmacro.getSectionStyleName("Titlepage")
    'Act:
        Call CIPmacro.tagFMSection(styleName, sectionArray, 50, "<" & tpTag & ">", "</" & tpTag & ">", _
            "Titlepage", testDocx)
        openTagCheck = TestHelpers.strFromLocation(testDocx, 10, 0, 4)
        closeTagCheck = TestHelpers.strFromLocation(testDocx, 11, -6, 5)
        'verify tags are not picking up local styling (that could override charCase when converted to txt)
        openCapsCheck = TestHelpers.smallCapsCheckFromLocation(testDocx, 10, 0, 4)
        closeCapsCheck = TestHelpers.smallCapsCheckFromLocation(testDocx, 11, -6, 5)
    'Assert:
        Assert.Succeed
        Assert.areequal "<" & tpTag & ">", openTagCheck
        Assert.areequal "</" & tpTag & ">", closeTagCheck
        Assert.areequal False, openCapsCheck
        Assert.areequal False, closeCapsCheck
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagFM_badFM() 'TODO Rename test
    Dim spStyleName As String, cpStyleName As String
    Dim spOpenTagCount As Long, spCloseTagCount As Long, cpOpenTagCount As Long, _
        cpCloseTagCount As Long, spOpenCountExpected As Long, spCloseCountExpected As Long, _
        cpOpenCountExpected As Long, cpCloseCountExpected As Long
    Dim testdoc_filepath As String
    On Error GoTo TestFail
    'Arrange:
        spOpenCountExpected = 1
        spCloseCountExpected = 0
        cpOpenCountExpected = 1
        cpCloseCountExpected = 0
        ' setup test doc
        testdoc_filepath = getRepoPath + "test_files\testfile_CIP_badFM.dotx"
        Set testDocx = Nothing
        Set testDocx = Application.Documents.Add(Template:=testdoc_filepath)
        spStyleName = CIPmacro.getSectionStyleName("Series Page")
        cpStyleName = CIPmacro.getSectionStyleName("Copyright")
    'Act:
        ' sp exceeds max length, cp ends at end of document
        Call CIPmacro.tagFMSection(spStyleName, sectionArray, 5, "<" & spTag & ">", "</" & spTag & ">", _
            "Series Page", testDocx)
        spOpenTagCount = TestHelpers.stringCountInDoc(testDocx, "<" & spTag & ">")
        spCloseTagCount = TestHelpers.stringCountInDoc(testDocx, "</" & spTag & ">")
        Call CIPmacro.tagFMSection(cpStyleName, sectionArray, 5, "<" & cpTag & ">", "</" & cpTag & ">", _
            "Copyright", testDocx)
        cpOpenTagCount = TestHelpers.stringCountInDoc(testDocx, "<" & cpTag & ">")
        cpCloseTagCount = TestHelpers.stringCountInDoc(testDocx, "</" & cpTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal spOpenCountExpected, spOpenTagCount
        Assert.areequal spCloseCountExpected, spCloseTagCount
        Assert.areequal cpOpenCountExpected, cpOpenTagCount
        Assert.areequal cpCloseCountExpected, cpCloseTagCount
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CIPtests")
Private Sub TestTagFM_nostyle() 'TODO Rename test
    Dim styleName As String, openTagCount As Long, closeTagCount As Long
    Dim expectedCount As Long
    On Error GoTo TestFail
    'Arrange:
        Set testDocx = Nothing
        'Set testDocx = Application.Documents.Add(Template:=nostyleTestFilePath, visible:=False)
        ' for debug / review:
        Set testDocx = Application.Documents.Add(Template:=nostyleTestFilePath)
        expectedCount = 0
        styleName = CIPmacro.getSectionStyleName("Titlepage")
    'Act:
        Call CIPmacro.tagFMSection(styleName, sectionArray, 50, "<" & tpTag & ">", "</" & tpTag & ">", _
            "Titlepage", testDocx)
        openTagCount = TestHelpers.stringCountInDoc(testDocx, "<" & tpTag & ">")
        closeTagCount = TestHelpers.stringCountInDoc(testDocx, "</" & tpTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal expectedCount, openTagCount
        Assert.areequal expectedCount, closeTagCount
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagChStarts() 'TODO Rename test
    Dim styleName As String, tagCheck1 As String, tagCheck2 As String, tagCheck3 As String
    Dim tagTotal As Long, tagExpected As Long
    Dim tagCapsCheck As Boolean
    tagExpected = 3
    On Error GoTo TestFail
    'Arrange:
        Set testDocx = Application.Documents.Add(Template:=goodTestFilePath, visible:=False)
        ' for debug:
        ' Set testDocx = Application.Documents.Add(Template:=goodTestFilePath)
    'Act:
        Call CIPmacro.tagChapterStarts(chStyleArray, chTag, testDocx)
        tagCheck1 = TestHelpers.strFromLocation(testDocx, 25, 0, 4)
        tagCheck2 = TestHelpers.strFromLocation(testDocx, 32, 0, 4)
        tagCapsCheck = TestHelpers.smallCapsCheckFromLocation(testDocx, 32, 0, 4)
        tagCheck3 = TestHelpers.strFromLocation(testDocx, 39, 0, 4)
        tagTotal = TestHelpers.stringCountInDoc(testDocx, "<" & chTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal "<" & chTag & ">", tagCheck1
        Assert.areequal "<" & chTag & ">", tagCheck2
        Assert.areequal False, tagCapsCheck
        Assert.areequal "<" & chTag & ">", tagCheck3
        Assert.areequal tagExpected, tagTotal
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CIPtests")
Private Sub TestTagChNumbers() 'TODO Rename test
    Dim styleName As String, tagCheck1 As String, tagCheck2 As String, tagCheck3 As String
    Dim tagTotal As Long, tagExpected As Long, lastChParaIndex As Long
    Dim tagCapsCheck As Boolean
    tagExpected = 0
    On Error GoTo TestFail
    'Arrange:
        Set testDocx = Application.Documents.Add(Template:=goodTestFilePath, visible:=False)
        ' for debug:
        ' Set testDocx = Application.Documents.Add(Template:=goodTestFilePath)
    'Act:
        Call CIPmacro.tagChapterStarts(chStyleArray, chTag, testDocx)
        lastChParaIndex = CIPmacro.numberChapterTags(chTag, testDocx)
        tagCheck1 = TestHelpers.strFromLocation(testDocx, 25, 0, 5)
        tagCheck2 = TestHelpers.strFromLocation(testDocx, 32, 0, 5)
        tagCapsCheck = TestHelpers.smallCapsCheckFromLocation(testDocx, 32, 0, 5)
        tagCheck3 = TestHelpers.strFromLocation(testDocx, 39, 0, 5)
        tagTotal = TestHelpers.stringCountInDoc(testDocx, "<" & chTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal "<" & chTag & "1>", tagCheck1
        Assert.areequal "<" & chTag & "2>", tagCheck2
        Assert.areequal False, tagCapsCheck
        Assert.areequal "<" & chTag & "3>", tagCheck3
        Assert.areequal tagExpected, tagTotal
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagChEnd() 'TODO Rename test
    Dim styleName As String, tagCheck As String
    Dim tagTotal As Long, tagExpected As Long, lastChParaIndex As Long
    Dim tagCapsCheck As Boolean
    tagExpected = 1
    On Error GoTo TestFail
    'Arrange:
        Set testDocx = Application.Documents.Add(Template:=goodTestFilePath, visible:=False)
        ' for debug:
        'Set testDocx = Application.Documents.Add(Template:=goodTestFilePath)
    'Act:
        Call CIPmacro.tagChapterStarts(chStyleArray, chTag, testDocx)
        lastChParaIndex = CIPmacro.numberChapterTags(chTag, testDocx)
        Call CIPmacro.tagChaptersEnd(lastChParaIndex, bmStyleArray, chTag, testDocx)
        tagCheck = TestHelpers.strFromLocation(testDocx, 45, -6, 5)
        tagCapsCheck = TestHelpers.smallCapsCheckFromLocation(testDocx, 45, -6, 5)
        tagTotal = TestHelpers.stringCountInDoc(testDocx, "</" & chTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal "</" & chTag & ">", tagCheck
        Assert.areequal False, tagCapsCheck
        Assert.areequal tagExpected, tagTotal
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("CIPtests")
Private Sub TestTagCh_noBM() 'TODO Rename test
    Dim styleName As String, tagCheck As String
    Dim tagTotal As Long, tagExpected As Long, lastChParaIndex As Long
    Dim testDocPath As String
    tagExpected = 1
    On Error GoTo TestFail
    'Arrange:
        testDocPath = getRepoPath + "test_files\testfile_CIP_good_noBM.dotx"
        Set testDocx = Application.Documents.Add(Template:=testDocPath, visible:=False)
        ' for debug:
        'Set testDocx = Application.Documents.Add(Template:=testDocPath)
    'Act:
        Call CIPmacro.tagChapterStarts(chStyleArray, chTag, testDocx)
        lastChParaIndex = CIPmacro.numberChapterTags(chTag, testDocx)
        Call CIPmacro.tagChaptersEnd(lastChParaIndex, bmStyleArray, chTag, testDocx)
        tagCheck = TestHelpers.strFromLocation(testDocx, 45, 4, 5)
        tagTotal = TestHelpers.stringCountInDoc(testDocx, "</" & chTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal "'</" & chTag & ">'", "'" & tagCheck & "'"
        Assert.areequal tagExpected, tagTotal
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagCh_nostyle() 'TODO Rename test
    Dim styleName As String, tagCheck As String
    Dim openTagExp As Long, closeTagExp As Long, openTagCount As Long, closeTagCount As Long, _
        lastChParaIndex As Long
    On Error GoTo TestFail
    'Arrange:
        Set testDocx = Nothing
        openTagExp = 0
        closeTagExp = 0
        Set testDocx = Application.Documents.Add(Template:=nostyleTestFilePath, visible:=False)
        ' for debug:
        'Set testDocx = Application.Documents.Add(Template:=nostyleTestFilePath)
        
    'Act:
        Call CIPmacro.tagChapterStarts(chStyleArray, chTag, testDocx)
        lastChParaIndex = CIPmacro.numberChapterTags(chTag, testDocx)
        Call CIPmacro.tagChaptersEnd(lastChParaIndex, bmStyleArray, chTag, testDocx)
        openTagCount = TestHelpers.stringCountInDoc(testDocx, "<" & chTag)
        closeTagCount = TestHelpers.stringCountInDoc(testDocx, "</" & chTag & ">")
    'Assert:
        Assert.Succeed
        Assert.areequal openTagExp, openTagCount
        Assert.areequal closeTagExp, closeTagCount
    'Cleanup:
        testDocx.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestRmParas() 'TODO Rename test
    Dim tstStyleName As String, resultstr
    Dim expectedRange As Range, testRange As Range
    Dim beforeDotxPath As String, afterDotxPath As String
    Dim beforeDoc As Document, afterDoc As Document
    On Error GoTo TestFail
    'Arrange:
        tstStyleName = "Section-Titlepage (STI)"
        beforeDotxPath = getRepoPath + "test_files\testfile_CIP_compareBefore.dotx"
        afterDotxPath = getRepoPath + "test_files\testfile_CIP_compareAfter_rm.dotx"
        Set beforeDoc = Application.Documents.Add(Template:=beforeDotxPath, visible:=False)
        ' for debug:
        'Set beforeDoc = Application.Documents.Add(Template:=beforeDotxPath)
    'Act:
        Call CIPmacro.rmParasWithStyle(tstStyleName, beforeDoc)
        Set testRange = beforeDoc.Range
        Set afterDoc = Application.Documents.Add(Template:=afterDotxPath, visible:=False)
        ' for debug:
        'Set afterDoc = Application.Documents.Add(Template:=afterDotxPath)
        Set expectedRange = afterDoc.Range
        resultstr = TestHelpers.compareRanges(testRange, expectedRange)
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", resultstr
    'Cleanup:
        beforeDoc.Close savechanges:=wdDoNotSaveChanges
        afterDoc.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestChangeBreaks() 'TODO Rename test
    Dim tstStyleName As String, resultstr
    Dim expectedRange As Range, testRange As Range
    Dim beforeDotxPath As String, afterDotxPath As String
    Dim beforeDoc As Document, afterDoc As Document
    On Error GoTo TestFail
    'Arrange:
        beforeDotxPath = getRepoPath + "test_files\testfile_CIP_compareBefore.dotx"
        afterDotxPath = getRepoPath + "test_files\testfile_CIP_compareAfter_breaks.dotx"
        Set beforeDoc = Application.Documents.Add(Template:=beforeDotxPath, visible:=False)
        ' for debug:
        'Set beforeDoc = Application.Documents.Add(Template:=beforeDotxPath)
    'Act:
        Call CIPmacro.changeBreakParaContents(breakArray, "^p", beforeDoc)
        Set testRange = beforeDoc.Range
        Set afterDoc = Application.Documents.Add(Template:=afterDotxPath, visible:=False)
        ' for debug:
        'Set afterDoc = Application.Documents.Add(Template:=afterDotxPath)
        Set expectedRange = afterDoc.Range
        resultstr = TestHelpers.compareRanges(testRange, expectedRange)
    'Assert:
        Assert.Succeed
        Assert.areequal "Same", resultstr
    'Cleanup:
        beforeDoc.Close savechanges:=wdDoNotSaveChanges
        afterDoc.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CIPtests")
Private Sub TestTagReport() 'TODO Rename test
    Dim taggedDoc As Document, reportTxt As Document, expReportTxt As Document
    Dim taggedDocxPath As String, expectedReportPath As String, reportPath As String
    Dim taggedBool As Boolean
    Dim reportStr As String, expReportStr As String
    On Error GoTo TestFail
    'Arrange:
        ' set file paths
        taggedDocxPath = getRepoPath + "test_files\testfile_CIP_tagged-good.docx"
        reportPath = getRepoPath + "test_files\testfile_CIP_tagged-good_CIPtagReport.txt"
        expectedReportPath = getRepoPath + "test_files\testfile_CIP_tagged-good_CIPtagReport_expected.txt"
        ' Open tagged doc for report
        Set taggedDoc = Application.Documents.Open(taggedDocxPath)
    'Act:
        ' run report
        taggedBool = CIPmacro.reportOnTags(tagArray, tagDisplayNameArray, tagRequiredArray, chTag, _
            True, taggedDoc, taggedDoc)
        ' get txtfile outputs, current vs expected
        Set reportTxt = Application.Documents.Open(reportPath)
        reportStr = reportTxt.Range.Text
        Set expReportTxt = Application.Documents.Open(expectedReportPath)
        expReportStr = expReportTxt.Range.Text
    'Assert:
        Assert.Succeed
        Assert.areequal True, taggedBool
        Assert.areequal expReportStr, reportStr
    'Cleanup:
        taggedDoc.Close savechanges:=wdDoNotSaveChanges
        reportTxt.Close savechanges:=wdDoNotSaveChanges
        expReportTxt.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagReport_badFM() 'TODO Rename test
    Dim taggedDoc As Document, reportTxt As Document, expReportTxt As Document
    Dim taggedDocxPath As String, expectedReportPath As String, reportPath As String
    Dim taggedBool As Boolean
    Dim reportStr As String, expReportStr As String
    On Error GoTo TestFail
    'Arrange
        ' set file paths
        taggedDocxPath = getRepoPath + "test_files\testfile_CIP_tagged_badFM.docx"
        reportPath = getRepoPath + "test_files\testfile_CIP_tagged_badFM_CIPtagReport.txt"
        expectedReportPath = getRepoPath + "test_files\testfile_CIP_tagged_badFM_CIPtagReport_expected.txt"
        ' Open tagged doc for report
        Set taggedDoc = Application.Documents.Open(taggedDocxPath)
    'Act:
        ' run report
        taggedBool = CIPmacro.reportOnTags(tagArray, tagDisplayNameArray, tagRequiredArray, chTag, _
            True, taggedDoc, taggedDoc)
        ' get txtfile outputs, current vs expected
        Set reportTxt = Application.Documents.Open(reportPath)
        reportStr = reportTxt.Range.Text
        Set expReportTxt = Application.Documents.Open(expectedReportPath)
        expReportStr = expReportTxt.Range.Text
    'Assert:
        Assert.Succeed
        Assert.areequal True, taggedBool
        Assert.areequal expReportStr, reportStr
    'Cleanup:
        taggedDoc.Close savechanges:=wdDoNotSaveChanges
        reportTxt.Close savechanges:=wdDoNotSaveChanges
        expReportTxt.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagReport_badFM_nochapters() 'TODO Rename test
    ' same as tagreport badFM but not requiring chapters.
    Dim taggedDoc As Document, reportTxt As Document, expReportTxt As Document
    Dim taggedDocxPath As String, expectedReportPath As String, reportPath As String
    Dim taggedBool As Boolean
    Dim reportStr As String, expReportStr As String
    On Error GoTo TestFail
    'Arrange
        ' set file paths
        taggedDocxPath = getRepoPath + "test_files\testfile_CIP_tagged_badFM.docx"
        reportPath = getRepoPath + "test_files\testfile_CIP_tagged_badFM_CIPtagReport.txt"
        expectedReportPath = getRepoPath + "test_files\testfile_CIP_tagged_badFM_CIPtagReport_nochapExpected.txt"
        ' Open tagged doc for report
        Set taggedDoc = Application.Documents.Open(taggedDocxPath)
    'Act:
        ' run report
        taggedBool = CIPmacro.reportOnTags(tagArray, tagDisplayNameArray, tagRequiredArray, chTag, _
            False, taggedDoc, taggedDoc)
        ' get txtfile outputs, current vs expected
        Set reportTxt = Application.Documents.Open(reportPath)
        reportStr = reportTxt.Range.Text
        Set expReportTxt = Application.Documents.Open(expectedReportPath)
        expReportStr = expReportTxt.Range.Text
    'Assert:
        Assert.Succeed
        Assert.areequal True, taggedBool
        Assert.areequal expReportStr, reportStr
    'Cleanup:
        taggedDoc.Close savechanges:=wdDoNotSaveChanges
        reportTxt.Close savechanges:=wdDoNotSaveChanges
        expReportTxt.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagReport_nostyle() 'TODO Rename test
    Dim taggedDoc As Document
    Dim taggedDocxPath As String
    Dim taggedBool As Boolean
    On Error GoTo TestFail
    'Arrange:
        Fakes.MsgBox.Returns vbOK
        ' set file paths
        taggedDocxPath = getRepoPath + "test_files\testfile_CIP_nostyles.dotx"
        Set taggedDoc = Application.Documents.Open(taggedDocxPath, visible:=False)
    'Act:
        ' run report
        taggedBool = CIPmacro.reportOnTags(tagArray, tagDisplayNameArray, tagRequiredArray, chTag, _
            True, taggedDoc, taggedDoc)
        With Fakes.MsgBox.verify
            .Parameter "title", "No Styles Found"
        End With
    'Assert:
        Assert.Succeed
        Assert.areequal False, taggedBool
    'Cleanup:
        taggedDoc.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagPreCheck_false() 'TODO Rename test
    Dim testDoc As Document
    Dim testDocxPath As String
    Dim tagsBool As Boolean
    On Error GoTo TestFail
    'Arrange:
        ' set file paths
        testDocxPath = getRepoPath + "test_files\testfile_CIP_nostyles.dotx"
        Set testDoc = Application.Documents.Open(testDocxPath, visible:=False)
    'Act:
        ' run report
        tagsBool = CIPmacro.preCheckTags(tagArray, chTag, testDoc)
    'Assert:
        Assert.Succeed
        Assert.areequal False, tagsBool
    'Cleanup:
        testDoc.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CIPtests")
Private Sub TestTagPreCheck_true() 'TODO Rename test
    Dim testDoc As Document
    Dim testDocxPath As String
    Dim tagsBool As Boolean
    On Error GoTo TestFail
    'Arrange:
        Fakes.MsgBox.Returns vbOK
        ' set file paths
        testDocxPath = getRepoPath + "test_files\testfile_CIP_tagged-good.docx"
        Set testDoc = Application.Documents.Open(testDocxPath, visible:=False)
    'Act:
        ' run report
        tagsBool = CIPmacro.preCheckTags(tagArray, chTag, testDoc)
        With Fakes.MsgBox.verify
            .Parameter "prompt", "Your document: 'testfile_CIP_tagged-good.docx' already contains the following CIP tag(s):" _
                & vbNewLine & vbNewLine & "      <tp>, </tp>, <cp>, </cp>, <sp>, </sp>, <toc>, </toc>, </ch>, as well as " & _
                "one or more chapter heading tags (e.g. <ch1>, <ch2>, ... )" & vbNewLine & vbNewLine & _
                "This macro may have already been run on this document. To run this macro, you MUST find and remove all existing CIP tags first."
        End With
    'Assert:
        Assert.Succeed
        Assert.areequal True, tagsBool
    'Cleanup:
        testDoc.Close savechanges:=wdDoNotSaveChanges
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub









