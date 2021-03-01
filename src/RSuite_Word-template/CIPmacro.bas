Attribute VB_Name = "CIPmacro"
Const sectionFileBasename As String = "sections.txt"
Const containerFileBasename As String = "containers.txt"
Const breakFileBasename As String = "breaks.txt"
Const containerEndStylename As String = "END (END)"
Const bookmakerPiStylename As String = "Bookmaker-Processing-Instruction (Bpi)"
Const designNoteStylename As String = "Design-Note (Dn)"
Public Const BM_IN_MACRO As String = "_UndoBkmrk_"
Dim activeDoc As Document
'Const elcMarker As String = "``````"

Sub Main()

' ========================== Dimming and Setting Variables ======================
Dim tagCheck As Boolean, tagChaptersBool As Boolean, tagsPresent As Boolean
Dim tpName As String, cpName As String, spName As String, tocName As String, chName As String
Dim tpTag As String, cpTag As String, spTag As String, tocTag As String, chTag As String
Dim tpDisplayName As String, cpDisplayName As String, spDisplayName As String, tocDisplayName As String
Dim tpRequired As Boolean, cpRequired As Boolean, spRequired As Boolean, tocRequired As Boolean
Dim bmsectionarray(), tagArray(), chNamesArray(), tagNameArray(), tagDisplayNameArray(), tagRequiredArray()
Dim sectionArray, bmStyleArray, containerArray, breakArray, chStyleArray
Dim maxsectionlength As Long
Dim lastChapParaIndex As Long

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

' chapter tagging handled separately
chNamesArray = Array("Chapter", "Chapter 2")
chTag = "ch"
' these backmatter strings match names in "sectionFile".
bmsectionarray = Array("About the Author", _
    "Acknowledgments", _
    "Afterword", _
    "Appendix", _
    "Back Ad", _
    "Back Matter General", _
    "Bibliography", _
    "Conclusion", _
    "Excerpt Chapter", _
    "Excerpt Opener")

' for tp, copyright, series page, TOC ending tags:
' none of these special sections should be > 5 paras, with possible exception of TOC
maxsectionlength = 1000

' activeDoc is global variable to hold our document, so if user clicks a different
' document during execution, won't switch to that doc.
' ALWAYS set to Nothing first to reset for this macro.
' Then only refer to this object, not ActiveDocument directly.
Set activeDoc = Nothing
Set activeDoc = ActiveDocument

' ========================== Running Checks, UI Setup ==============================
' pre check: is doc saved, protected, track changes etc
' True means a check failed (e.g., doc protection on)
If WT_Settings.InstallType = "user" Then
  If MacroHelpers.StartupSettings(AcceptAll:=False) = True Then
    Call MacroHelpers.Cleanup
    Exit Sub
  End If
Else
  If MacroHelpers.StartupSettings(AcceptAll:=True) = True Then
    Call MacroHelpers.Cleanup
    Exit Sub
  End If
End If

' pre check: volume userform
tagChaptersBool = volumeCheckPrompt
If cipVolumePrompt.Cancelled = True Then
    Unload cipVolumePrompt
    Exit Sub
End If
Unload cipVolumePrompt

' start progress bar
Set pBar = New Progress_Bar
pBarCounter = 0
pBar.Caption = "CIP Application Tagging Macro"
thisstatus = "* Running pre-tagging checks "
completeStatus = "Starting CIP Macro" + vbNewLine + "=========================" + _
    vbNewLine + thisstatus
pBar.Status.Caption = completeStatus
Clean_helpers.updateStatus ("")

' setup
Application.ScreenUpdating = False
Clean_helpers.ClearSearch

' pre-check: are tags already present?
tagCheck = preCheckTags(tagArray, chTag)
If tagCheck = True Then
    Unload pBar
    Call zz_clearFindB
    Application.ScreenUpdating = True
    Exit Sub
End If
'update pre-check progress bar
completeStatus = completeStatus + "100%" & Clean_helpers.updateStatus("")

'run backup
Call backupFile

' add bookmark for Undo later
Call setupUndoBookmark

' ========================== Perform tagging, Tag Report ==============================
' insert tags for FM sections
sectionArray = getStyleArrayfromFile(sectionFileBasename)
Call tagFMSections(sectionArray, maxsectionlength, tagNameArray, tagDisplayNameArray, tagArray)

' tag chapters
chStyleArray = getMultiSectionStyleNames(chNamesArray)
Call tagChapters(tagChaptersBool, chStyleArray, chTag, bmsectionarray)

' Run tag summary / post-checks
tagsPresent = reportOnTags(tagArray, tagDisplayNameArray, tagRequiredArray, chTag, tagChaptersBool)
If tagsPresent = False Then
    Unload pBar
    ' Call undoChanges ' we can exit, no tags were added. Can add undo out of excess caution
    Call zz_clearFindB
    Application.ScreenUpdating = True
    Exit Sub
End If


' ========================== Rm non-content, save file, cleanup ==============================
' Clean document of Sections and container paras, etc
Call stripNonContentParas(sectionArray)

' save out text file
Call SaveAsTextFile

' Undo changes in original doc back to UNDO bookmark:
Call undoChanges

''' also to do:  error handling, unit testing, revising document shuffling (make copy up front, operate on that?, don't save?
'Call MacroHelpers.Cleanup '?
Unload pBar
Application.ScreenUpdating = True

Call Clean_helpers.MessageBox("Done", "CIP Tagging macro complete!", vbOK)

End Sub


Sub setupUndoBookmark()
    ' remove bookmark if its already here
    If activeDoc.Bookmarks.Exists(BM_IN_MACRO) Then
        activeDoc.Bookmarks(BM_IN_MACRO).Delete
    End If
    ' add bookmark anew.
    activeDoc.Range.Bookmarks.Add BM_IN_MACRO
End Sub

Private Sub backupFile(Optional suffixStr As String = "_CIPbackup")
'Dim suffixStr As String
'suffixStr = "_testss"
Dim fileExt As String
Dim saveFormat As WdSaveFormat
Dim tmpDoc As Document

saveFormat = wdFormatDocumentDefault

' begin update progress bar
thisstatus = "* Making backup of original file "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

' for debug tests:
If activeDoc Is Nothing Then Set activeDoc = ActiveDocument

' get fileExtension
fileExt = Right(activeDoc, Len(activeDoc) - InStrRev(activeDoc.Name, "."))

' get target filepath
Dim targetFilePath As String
targetFilePath = getThisFilePathWithNewSuffix(suffixStr & "." & fileExt)

' make a copy of document
Set tmpDoc = Documents.Add(activeDoc.FullName, visible:=False)
' 'the next line saves the copy to your location and name
'ActiveDocument.SaveAs sSaveAsPath
' 'next line closes the copy leaving you with the original document
'ActiveDocument.Close

'If fileExt = "doc" Then
'    saveFormat = wdFormatDocument
'End If

'Turn off alerts for save to avoid compat. notices
'Application.DisplayAlerts = wdAlertsNone

' "save as" backup file
tmpDoc.SaveAs fileName:=targetFilePath ', FileFormat:=saveFormat

'Turn alerts on again
'Application.DisplayAlerts = wdAlertsAll

'close copy
tmpDoc.Close

' tell pbar we finished something
completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")

End Sub

Private Function volumeCheckPrompt()

With cipVolumePrompt
    .Caption = "Notice"
    .text1.Caption = "The Library of Congress has special rules for tagging 'Multivolume Sets'" & vbNewLine & vbNewLine & _
        "If this book is not part of a multivolume set, or you wish to submit this as a single volume, click 'STANDARD CIP TAGS'" & vbNewLine & vbNewLine & _
        "If submitting as a 'single application for multiple volumes': click 'SKIP CHAPTER TAGS' to proceed with auto-tagging, exempting chapter tags (<ch></ch>)" & vbNewLine & vbNewLine & _
        "Chapter tags will unfortunately need to be applied manually in this case." & vbNewLine & vbNewLine & _
        "For more information, please peruse CIP guidleines on Multivolume Sets via the link below."
    .text1.FontSize = 9
    .button1.FontSize = 10
    .button2.FontSize = 10
    .cbCancel.FontSize = 9
    .LOC_hyperlink.FontSize = 9
    .LOC_hyperlink.Caption = "https://www.loc.gov/publish/cip/techinfo/formattingecip.html#multi"
    .Show
End With

volumeCheckPrompt = cipVolumePrompt.tagChapters
End Function

Private Sub stripNonContentParas(sectionArray)

'' begin update progress bar
thisstatus = "* Stripping 'Section' and 'Container' paras " ', standardizing 'Break' para contents "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

' get array of all containerstylenames (Except END), and breaks
containerArray = getStyleArrayfromFile(containerFileBasename)
breakArray = getStyleArrayfromFile(breakFileBasename)

' rm section styles, container styles, break styles, "END" style (can we save as text file first??)
Call rmParasWithStylesArray(sectionArray)
Call rmParasWithStylesArray(containerArray)
Call rmParasWithStyle(containerEndStylename)
Call rmParasWithStyle(designNoteStylename)
Call rmParasWithStyle(bookmakerPiStylename)

' empty contents of break paras
Call changeBreakParaContents(breakArray, "^p")

' update progress bar - done
completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
Clean_helpers.updateStatus ("")

End Sub
'
Private Sub tagChapters(tagChaptersBool, chStyleArray, chTag As String, bmsectionarray)
'Dim chaptercount As Long

' tag chapters if user agreed in volume-prompt
If tagChaptersBool = True Then
    ' begin update progress bar
    thisstatus = "* Adding tags for Chapters "
    If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

    ' insert chapter start tags
    'chaptercount = tagChapterStartsWithCount(chStyleArray, chTag)
    Call tagChapterStarts(chStyleArray, chTag)
    lastChapParaIndex = numberChapterTags(chTag)

    ' insert chapter end tag
    bmStyleArray = getMultiSectionStyleNames(bmsectionarray)
    Call tagChaptersEnd(lastChapParaIndex, bmStyleArray, chTag)

    ' update progress bar
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    Clean_helpers.updateStatus ("")
Else
    completeStatus = completeStatus + vbNewLine + "* Skipping Chapter tags"
    Clean_helpers.updateStatus ("")
End If

End Sub

Private Function preCheckTags(tagArray, chTag) As Boolean
Dim i As Long, j As Long
Dim foundTags()
Dim openTagStr As String, closeTagStr As String, chOpenTagStr As String, chCloseTagStr As String
Dim chFoundStr As String, nonChFoundStr As String, foundStr As String
preCheckTags = False
j = 0
chFoundStr = ""
nonChFoundStr = ""
foundStr = ""

' all non-chapter tags
For i = 0 To UBound(tagArray)
    openTagStr = "<" & tagArray(i) & ">"
    If tagCheck(openTagStr) = True Then
        ReDim Preserve foundTags(j)
        foundTags(j) = openTagStr
        j = j + 1
    End If
    closeTagStr = "</" & tagArray(i) & ">"
    If tagCheck(closeTagStr) = True Then
        ReDim Preserve foundTags(j)
        foundTags(j) = closeTagStr
        j = j + 1
    End If
Next i

' now chapter tags:
chOpenTagStr = "\<" & chTag & "[0-9]{1,}\>" 'backslashes to include angle brackets, since we will search w. wildcards
If tagCheck(chOpenTagStr, True) = True Then
    chFoundStr = "one or more chapter heading tags (e.g. <ch1>, <ch2>, ... )"
End If

chCloseTagStr = "</" & chTag & ">"
If tagCheck(chCloseTagStr) = True Then
    ReDim Preserve foundTags(j)
    foundTags(j) = chCloseTagStr
End If

' Sort out our return string
If j > 0 Then nonChFoundStr = "      " & Join(foundTags, ", ")

If nonChFoundStr <> "" And chFoundStr <> "" Then
    foundStr = nonChFoundStr & ", as well as " & chFoundStr
ElseIf nonChFoundStr <> "" Then
    foundStr = nonChFoundStr
ElseIf chFoundStr <> "" Then
    foundStr = "      " & chFoundStr
End If

If foundStr <> "" Then
    MsgBox "Your document: '" & activeDoc & "' already contains the following CIP tag(s):" & vbNewLine & vbNewLine & foundStr & vbNewLine & vbNewLine & _
    "This macro may have already been run on this document. To run this macro, you MUST find and remove all existing CIP tags first.", , "Alert"
    preCheckTags = True
    Exit Function
End If

End Function
Private Function tagCheck(tagName, Optional wildcardsBool = False) As Boolean
Dim tagFound As Boolean
tagFound = False

With activeDoc.Range.Find
    .ClearFormatting
    .Text = tagName
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = wildcardsBool
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        tagFound = True
    End If
End With

tagCheck = tagFound
End Function


Private Sub tagChaptersEnd(lastChapParaIndex, bmStyleArray, chTag)
Dim endChapsTag As String
Dim activeRng As Range, targetRange As Range
Dim i As Long
Dim tmpTargetParaIndex As Long, targetParaindex As Long
Dim lastChapterIndex As Long

endChapsTag = "</" & chTag & ">"
lastChapterIndex = activeDoc.Paragraphs.Count
targetParaindex = lastChapterIndex
tmpTargetParaIndex = lastChapterIndex
Set activeRng = activeDoc.Range
' skip if this index val is 0, that means hthere are no tagged chaps here,
'   which means we cannot determine if a section like Acknowledgements is fm or bm
If lastChapParaIndex <> 0 Then
    ' cycle through bm sections, comparing index of first of each to lastchaptag index.
    For i = 0 To UBound(bmStyleArray)
        'Debug.Print "here"
        tmpTargetParaIndex = getFirstParaIndexAfterChapterEnd(bmStyleArray(i), lastChapParaIndex)
        If tmpTargetParaIndex < targetParaindex Then
            targetParaindex = tmpTargetParaIndex
        End If
    Next i
    ' no backmatter sections found, insert tag at end of document
    If targetParaindex = lastChapterIndex Then
        activeRng.InsertAfter endChapsTag
    ' else insert right before first bm section
    Else
        Set targetRange = activeDoc.Paragraphs(targetParaindex).Range
        With targetRange
            .MoveEnd Unit:=wdParagraph, Count:=-1
            .MoveEnd Unit:=wdCharacter, Count:=-1
            .InsertAfter endChapsTag
        End With
    End If
Else
    Debug.Print "No tagged chapters, skipping contents of tagChaptersEnd sub"
End If

End Sub

Sub rmParasWithStylesArray(targetStyleArray)
Dim i As Long
For i = 0 To UBound(targetStyleArray)
    rmParasWithStyle (targetStyleArray(i))
Next i
End Sub

Sub rmParasWithStyle(targetStyle)

With activeDoc.Range.Find
    .ClearFormatting
    .Text = ""
    .Format = True
    .Style = targetStyle
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .MatchWildcards = False
    .Execute ReplaceWith:="", Replace:=wdReplaceAll
End With
    
End Sub

Sub changeBreakParaContents(targetStyleArray, replacementStr)
Dim i As Long
For i = 0 To UBound(targetStyleArray)
    With activeDoc.Range.Find
        .ClearFormatting
        .Text = ""
        .Format = True
        .Style = targetStyleArray(i)
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchWildcards = False
        .Execute ReplaceWith:=replacementStr, Replace:=wdReplaceAll
    End With
Next i
End Sub


Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
Private Function countStyleUses(targetStyle As String) As Long
Dim stylecount As Long
stylecount = 0
With activeDoc.Range.Find
    .ClearFormatting
    .Text = ""
    .Wrap = wdFindStop
    .Format = True
    .Style = targetStyle
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .MatchWildcards = False
    Do While .Execute(Forward:=True) = True
        stylecount = stylecount + 1
    Loop
End With
countStyleUses = stylecount
End Function

'
'
Private Function reportOnTags(tagArray, tagDisplayNames, tagRequired, chTag, tagChaptersBool) As Boolean
Dim activeRng As Range
Dim docTxt As String, newTxt As String
Dim thisOpenTag As String, thisCloseTag As String, newTxtA As String, newTxtB As String, chCloseTag As String
Dim fmOpenTagCounts(), fmCloseTagCounts(), chOpenTagCount As Long, chCloseTagCount As Long, tagTotal As Long
Dim i As Long, j As Long, k As Long, L As Long
Dim errStr As String, reportStr As String

' begin update progress bar
thisstatus = "* Generating tag Report "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

'  =============== Count Tags ================

Set activeRng = activeDoc.Range
docTxt = activeDoc.Range.Text
tagTotal = 0

' count fm tags
For i = 0 To UBound(tagArray)
    thisOpenTag = "<" & tagArray(i) & ">"                           ' open tag
    newTxtA = Replace(docTxt, thisOpenTag, "")
    ReDim Preserve fmOpenTagCounts(i)
    fmOpenTagCounts(i) = ((Len(docTxt) - Len(newTxtA)) / Len(thisOpenTag))
    
    thisCloseTag = "</" & tagArray(i) & ">"                         ' close tag
    newTxtB = Replace(docTxt, thisCloseTag, "")
    ReDim Preserve fmCloseTagCounts(i)
    fmCloseTagCounts(i) = ((Len(docTxt) - Len(newTxtB)) / Len(thisCloseTag))
    
    tagTotal = tagTotal + fmOpenTagCounts(i) + fmCloseTagCounts(i)  ' total
Next

' count chapterEnd tags
chCloseTag = "</" & chTag & ">"
newTxt = Replace(docTxt, chCloseTag, "")
chCloseTagCount = ((Len(docTxt) - Len(newTxt)) / Len(chCloseTag))

'Count occurences of Chapter Tags
Call zz_clearFindB
chOpenTagCount = 0
With activeRng.Find
    .ClearFormatting
    .Text = "\<" & chTag & "[0-9]{1,}\>"
    .MatchWildcards = True
    Do While .Execute(Forward:=True) = True
        chOpenTagCount = chOpenTagCount + 1
    Loop
End With

'  =============== Exit Early no NO Tags ================

' if we have NO tags, post msgbox and exit
tagTotal = tagTotal + chCloseTagCount + chOpenTagCount
If tagTotal = 0 Then
    reportStr = "CIP tags cannot be added:" & vbNewLine & "Unable to find Macmillan-styled paragraphs indicating titlepage, " & _
        "copyright page, table of contents, or chapter title pages. Please add the correct styles and try again."
    MsgBox reportStr, vbCritical, "No Styles Found"
    reportOnTags = False
    Exit Function
Else
    reportOnTags = True
End If
    
'  =============== Report/Verify Tags ================

'Prepare error string(s)
errStr = ""
For j = 0 To UBound(tagRequired)
    If tagRequired(j) = True And fmOpenTagCounts(j) + fmCloseTagCounts(j) = 0 Then
        errStr = errStr & "ERROR: No " & tagDisplayNames(j) & " tags found. " & tagDisplayNames(j) & " tags are REQUIRED for LOC submission." & vbNewLine
    End If
Next j
For k = 0 To UBound(tagArray)
    If fmOpenTagCounts(k) + fmCloseTagCounts(k) = 1 Or fmOpenTagCounts(k) + fmCloseTagCounts(k) > 2 Then
        errStr = errStr & "ERROR: Problem with " & tagDisplayNames(k) & " tags: either too many were found or one is missing" & vbNewLine
    End If
Next k
Debug.Print tagChapter & chOpenTagCount & " " & chCloseTagCount
If tagChaptersBool = True And chOpenTagCount = 0 Then
    errStr = errStr & "ERROR: No Chapter Heading tags were found." & vbNewLine
End If
If tagChaptersBool = True And chCloseTagCount = 0 Then
    errStr = errStr & "ERROR: No 'End of Last Chapter' tag was found." & vbNewLine
End If

'Create full message text
reportStr = ""
If errStr = "" Then
    reportStr = reportStr & "Congratulations!" & vbNewLine
    reportStr = reportStr & "LOC Tags look good for " & activeDoc.Name & vbNewLine
    reportStr = reportStr & "See summary below:" & vbNewLine & vbNewLine
Else
    reportStr = reportStr & "BAD NEWS:" & vbNewLine & vbNewLine
    reportStr = reportStr & "Problems were found with LOC tags in your document '" & activeDoc.Name & "':" & vbNewLine
    reportStr = reportStr & vbNewLine & vbNewLine
    reportStr = reportStr & "------------------------- ERRORS -------------------------" & vbNewLine
    reportStr = reportStr & errStr & vbNewLine & vbNewLine
End If
    reportStr = reportStr & "------------------------- Tag Summary -------------------------" & vbNewLine
    For L = 0 To UBound(tagArray)
        reportStr = reportStr & fmOpenTagCounts(L) & "  " & tagDisplayNames(L) & " open tag(s) found <" & tagArray(L) & ">" & vbNewLine
        reportStr = reportStr & fmCloseTagCounts(L) & "  " & tagDisplayNames(L) & " close tag(s) found </" & tagArray(L) & ">" & vbNewLine
    Next L
    reportStr = reportStr & chOpenTagCount & "  Chapter beginning tag(s) found (<" & chTag & "1>, <" & chTag & "2>, etc)" & vbNewLine
    reportStr = reportStr & chCloseTagCount & "  End of last chapter tag(s) found </" & chTag & ">" & vbNewLine
    
' Print to text file
Call MacroHelpers.CreateTextFile(strText:=reportStr, suffix:="CIPtagReport")

' update progress bar
completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
Clean_helpers.updateStatus ("")

End Function


Private Sub tagChapterStarts(targetStyles, chTag As String)
Dim activeRng As Range
Dim chaptercount As Long
Dim i As Long, h As Long, j As Long
chaptercount = 0
Dim targetStyle As String

For h = 0 To UBound(targetStyles)
    targetStyle = targetStyles(h)
    chaptercount = chaptercount + countStyleUses(targetStyle)
Next h

If chaptercount > 0 Then
    ' tag all chapters with unnumbered tag, since we may be handling > 1 chap style
    For i = 0 To UBound(targetStyles)
        Set activeRng = activeDoc.Range
        With activeRng.Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = ""
          .Replacement.Text = "^&<" & chTag & ">"
          .Wrap = wdFindContinue
          .Format = True
          .Style = targetStyles(i)
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute Replace:=wdReplaceAll
        End With
    Next i
End If

End Sub

'Private Function tagChapterStartsWithCount(targetStyles, chTag As String)
'Dim activeRng As Range
'Dim chaptercount As Long
'Dim i As Long, h As Long, j As Long
'chaptercount = 0
'Dim targetStyle As String
'
'For h = 0 To UBound(targetStyles)
'    targetStyle = targetStyles(h)
'    chaptercount = chaptercount + countStyleUses(targetStyle)
'Next h
'
'If chaptercount > 0 Then
'    ' tag all chapters with unnumbered tag, since we may be handling > 1 chap style
'    For i = 0 To UBound(targetStyles)
'        Set activeRng = activeDoc.Range
'        With activeRng.Find
'          .ClearFormatting
'          .Replacement.ClearFormatting
'          .Text = ""
'          .Replacement.Text = "^&<" & chTag & ">"
'          .Wrap = wdFindContinue
'          .Format = True
'          .Style = targetStyles(i)
'          .MatchCase = False
'          .MatchWholeWord = False
'          .MatchWildcards = False
'          .MatchSoundsLike = False
'          .MatchAllWordForms = False
'          .Execute Replace:=wdReplaceAll
'        End With
'    Next i
'End If
'
'tagChapterStarts = chaptercount
'
'End Function
'
'
Function getSelectParaIndex() As Long ' gets para-index of 1st selected para
    getSelectParaIndex = activeDoc.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
End Function
Function getRangeParaIndex(myRange As Range) As Long ' gets para-index of 1st para in Range
    getRangeParaIndex = activeDoc.Range(0, myRange.Paragraphs(1).Range.End).Paragraphs.Count
End Function


Private Function numberChapterTags(chTag As String) As Long
' cycle back through and add increments
Dim activeRng As Range
Dim j As Long, paraIndex As Long
Set activeRng = activeDoc.Range
j = 1
With activeRng.Find
    .ClearFormatting
    .Text = "<" & chTag & ">"
    .MatchWildcards = False
    Do While .Execute(Forward:=True) = True
        With activeRng
            .MoveEnd Unit:=wdCharacter, Count:=-1
            .InsertAfter (j)
            paraIndex = getRangeParaIndex(activeRng)
            .Collapse direction:=wdCollapseEnd
            .Move Unit:=wdCharacter, Count:=1
        End With
        j = j + 1
    Loop
End With
numberChapterTags = paraIndex
End Function

Private Function getFirstParaIndexAfterChapterEnd(myStyle, chapEndIndex) As Long
' cycle back through and add increments
Dim activeRng As Range
Dim paraIndex As Long, tmpParaIndex As Long

Set activeRng = activeDoc.Range
paraIndex = activeDoc.Paragraphs.Count

With activeRng.Find
    .ClearFormatting
    .Text = ""
    .Format = True
    .Style = myStyle
    .Wrap = wdFindStop
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    Do While .Execute(Forward:=True) = True
        With activeRng
            tmpParaIndex = getRangeParaIndex(activeRng)
            If tmpParaIndex > chapEndIndex And tmpParaIndex < paraIndex Then
                paraIndex = tmpParaIndex
            End If
            .Collapse direction:=wdCollapseEnd
        End With
    Loop
End With

getFirstParaIndexAfterChapterEnd = paraIndex

End Function

Private Sub tagFMSection(targetStyle As String, sectionArray, maxsectionlength As Long, tagStr As String, tagEndStr As String, sectionName)

'activeDoc.StoryRanges(1).Select
'Selection.Collapse Direction:=wdCollapseStart
Dim activeRng As Range
Dim parastyle As String
Dim i As Long

' begin update progress bar
thisstatus = "* Adding tags for " & sectionName & " "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

' search and tag
Set activeRng = activeDoc.Range

With activeRng.Find
    .ClearFormatting
    .Text = ""
    .Wrap = wdFindContinue
    .Format = True
    .Style = targetStyle
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
    If .Found = True Then
        ' go to next paragraph and insert opening tag
        activeRng.MoveStart wdParagraph, 1
        activeRng.InsertBefore (tagStr)
        
        ' now we cycle downwards until we find next section, or reach maxlength (or end of document)
        parastyle = activeRng.ParagraphStyle
        i = 0
        While Not IsInArray(parastyle, sectionArray) And Not EndOfDocumentReached And i < maxsectionlength
            activeRng.MoveStart wdParagraph, 1
            parastyle = activeRng.ParagraphStyle

            i = i + 1
        Wend
        
        ' If we reached a new section, tag end of style
        If IsInArray(parastyle, sectionArray) Then
            activeRng.End = activeRng.End - 1
            activeRng.InsertAfter (tagEndStr)
        ' otherwise we quit and prompt the user:
        ElseIf i = maxsectionlength Then
            Debug.Print "FM section over " & maxsectionlength & " paras!"
        ElseIf EndOfDocumentReached Then
            Debug.Print "endless section!"
        End If
    End If
End With
    
' tell pbar we finished something
completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
If Not pBar Is Nothing Then Clean_helpers.updateStatus ("")
    
End Sub

Private Sub tagFMSections(sectionArray, maxsectionlength As Long, tagNameArray, tagDisplayNameArray, tagArray)
Dim i As Long
Dim tagStylename As String, tagName As String, openTag As String, closeTag As String

For i = 0 To UBound(tagNameArray)
    tagName = tagNameArray(i)
    tagStylename = getSectionStyleName(tagName)
    openTag = "<" & tagArray(i) & ">"
    closeTag = "</" & tagArray(i) & ">"
    Call tagFMSection(tagStylename, sectionArray, maxsectionlength, openTag, closeTag, tagDisplayNameArray(i))
Next i

End Sub

Private Function getMultiSectionStyleNames(sectionNames)
    Dim styleHash
    Dim stylenames()
    Dim j As Long
    Dim hashSectionName As String
    j = 0
    styleHash = getList(sectionFileBasename)
    For i = 1 To UBound(styleHash)
        hashSectionName = styleHash(i)(1)
        If IsInArray(hashSectionName, sectionNames) Then
            ReDim Preserve stylenames(j)
            stylenames(j) = styleHash(i)(0)
            j = j + 1
        End If
    Next i
    getMultiSectionStyleNames = stylenames
End Function
Private Function getSectionStyleName(sectionName As String) As String
    Dim styleHash
    Dim styleName As String
    styleHash = getList(sectionFileBasename)
    For i = 1 To UBound(styleHash)
        If styleHash(i)(1) = sectionName Then
            styleName = styleHash(i)(0)
        End If
    Next i
    getSectionStyleName = styleName
End Function

Private Function getStyleArrayfromFile(fileBasename As String) '(fileBasename As String)
    Dim styleHash
    styleHash = getList(fileBasename)
    Dim sectionStyleHash() As String
    ReDim sectionStyleHash(UBound(styleHash))
    For i = 0 To UBound(styleHash)
        sectionStyleHash(i) = styleHash(i)(0)
    Next i
    getStyleArrayfromFile = sectionStyleHash
End Function


Private Function getList(fileName As String)

    Dim FileNum As Integer
    Dim DataLine As String
    Dim StylePath As String
    
    Dim all() As Variant
    Dim i As Integer
    i = 0
    
    StylePath = WT_Settings.StyleDir(FileType:="styles") & Application.PathSeparator & fileName
    
    If IsItThere(StylePath) = True Then
        FileNum = FreeFile()
        Open StylePath For Input As #FileNum
        
        While Not EOF(FileNum)
            Line Input #FileNum, DataLine
            Dim result() As String
            result() = Split(DataLine, ",")
            If Right(result(0), 1) = vbLf Or Right(result(0), 1) = vbCr Then result(0) = Left(result(0), Len(result(0)) - 1)
            If Right(result(1), 1) = vbLf Or Right(result(1), 1) = vbCr Then result(1) = Left(result(1), Len(result(1)) - 1)
            result(0) = CleanString(RTrim(Trim(result(0))))
            result(1) = CleanString(RTrim(Trim(result(1))))
            ReDim Preserve all(i)
            all(i) = result
            i = i + 1
        Wend
        
        Close FileNum
        
    Else
        MessageBox Title:="Style List Not Found", Msg:="Cannot locate the RSuite Styles file."
    End If
    
    getList = all
    
End Function

Private Sub tagEndLastChapter(ELCstylesArray, chtagcode As String)
' ^\/this sub no t in use!!
' reusing this whole mechanism from prior (oldstyle) LOC macro, with minor tweaks
' Is a little painful to parse, but I believe the logic is basically:
' 1) Tag all BM section starts.
' 2) Then go to the first one, and see if it is followed in the doc by any
'    <ch> or <tp> tags (this is necessary b/c some bm section starts are also
'    Frontmatter section starts in some circumstances).
' 3) If no ch or tp tags trail it, then insert </ch> tag preceding it.
' 4) remove all working tags

    Set activeDoc = activeDoc
    Set activeRng = activeDoc.Range
    Dim i As Long
    
'On Error GoTo ErrHandler
    'Debug.Print UBound(ELCstylesArray)
    For i = 0 To UBound(ELCstylesArray)
    Debug.Print ELCstylesArray(i)
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = elcMarker & "^&"
      .Wrap = wdFindContinue
      .Format = True
      .Style = ELCstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = True
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
    Next
'On Error GoTo 0

    Call zz_clearFindB
    
    ' Declare vars related to part 2 (loop etc)
    Dim testvar As Boolean
    Dim testtag As String
    Dim Q As Long
    Dim bookmarkRng As Range
    Dim dontTag As Boolean
    Dim activeRngB As Range
    Set activeRngB = activeDoc.Range
    dontTag = False
    testvar = False
    testtag = "\<" & chtagcode & "[0-9]{1,}\>"
    Q = 0
    
    ''if <ch> not found, testtag= <tp>
    With activeRng.Find
        .Text = testtag
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If activeRng.Find.Execute = False Then
        testtag = "\<tp\>"
        With activeRngB.Find
            .Text = testtag
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        If activeRngB.Find.Execute = False Then
            dontTag = True
        End If
    End If
    
    'start loop
    Do While testvar = False
    Dim activeRngC As Range
    Set activeRngC = activeDoc.Range
    
        With activeRngC.Find
            .Text = elcMarker
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        ''set range with bookmarks, only search after init tag
        If activeRngC.Find.Execute = True Then
            activeDoc.Bookmarks.Add Name:="elcBookmark", Range:=activeRngC
            Set bookmarkRng = activeDoc.Range(Start:=activeDoc.Bookmarks("elcBookmark").Range.Start, End:=activeDoc.Bookmarks("\EndOfDoc").Range.End)
        Else
            Exit Do
        End If
        
        Set activeRng = activeDoc.Range
        
        Call zz_clearFindB
        
        'check for <ch> tags afer potential </ch> tag
        With bookmarkRng.Find
            .ClearFormatting
            .Text = testtag
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
        End With
        
        If bookmarkRng.Find.Execute = True Then
                'Found one. This one's not it.
                ''Remove first tagged paragraph's tag, will loop
                With activeRng.Find
                    .Text = elcMarker
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceOne
                End With
                Q = Q + 1
        Else
                ''This one's good, tag it right, set var to exit loop
                With activeRng.Find
                    .Text = elcMarker
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                If activeRng.Find.Execute = True Then
                    If dontTag = False Then
                        With activeRng.Find
                            .Text = "[!^13^m`]"
                            .Replacement.Text = "^&</" & chtagcode & ">"
                            .Forward = False
                            .Wrap = wdFindContinue
                            .Format = False
                            .MatchCase = False
                            .MatchWholeWord = True
                            .MatchWildcards = True
                            .MatchSoundsLike = False
                            .MatchAllWordForms = False
                            .Execute Replace:=wdReplaceOne
                        End With
                    End If
                End If
                testvar = True
        End If
            
        If activeDoc.Bookmarks.Exists("elcBookmark") = True Then
            activeDoc.Bookmarks("elcBookmark").Delete
        End If
        
        If Q = 20 Then      'prevent endless loops
            testvar = True
            dontTag = True
        End If
    
    Loop
    
    Call zz_clearFindB
    
    'Get rid of rest of ELC tags
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = elcMarker
      .Replacement.Text = ""
      .Wrap = wdFindContinue
      .Format = False
      .Forward = True
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
    
    'If no </ch> tags exist, add </ch> to the end of the doc
    If dontTag = False Then
        With activeRng.Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = "</ch>"
          .Wrap = wdFindContinue
          .Format = False
          .Forward = True
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
        End With
        If activeRng.Find.Execute = False Then
            Set activeRng = activeDoc.Range
            activeRng.InsertAfter "</ch>"
        End If
    End If
    
    Exit Sub
    
'ErrHandler:
'    Debug.Print "Heooo"
'    If Err.Number = 5941 Or Err.Number = 5834 Then
'        Resume NextLoop
'    End If
End Sub
Private Function getThisFilePathWithNewSuffix(suffixStr) 'expects file extension too, ex: "_2.docx"
Dim strdocname As String
'Separate code by OS because activeDoc.Path returns file name too
' on Mac but doesn't for PC
#If Mac Then        'For Mac
    If Val(Application.Version) > 14 Then
        'Find position of extension in filename
        strdocname = activeDoc.Path
        intPos = InStrRev(strdocname, ".")
        
        'Strip off extension and add ".txt" extension
        strdocname = Left(strdocname, intPos - 1)
        strdocname = strdocname & suffixStr
    End If
#Else                           'For Windows
    'Find position of extension in filename
    strdocname = activeDoc.Name
    DocPath = activeDoc.Path
    intPos = InStrRev(strdocname, ".")
    
    'Strip off extension and add ".txt" extension
    strdocname = Left(strdocname, intPos - 1)
    strdocname = DocPath & "\" & strdocname & suffixStr
#End If
getThisFilePathWithNewSuffix = strdocname
End Function

Private Sub SaveAsTextFile()
 
' Saves a copy of the document as a text file in the same path as the parent document
Dim tmpDoc As Document
Dim strdocname As String
Dim DocPath As String
Dim intPos As Integer
Dim encodingFmt As String
Dim lineBreak As Boolean
' for debug tests:
If activeDoc Is Nothing Then Set activeDoc = ActiveDocument
       
' begin update progress bar
thisstatus = "* Saving CIP text file "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)
       
'Application.ScreenUpdating = False

strdocname = getThisFilePathWithNewSuffix("_CIP.txt")
''Separate code by OS because activeDoc.Path returns file name too
'' on Mac but doesn't for PC
'#If Mac Then        'For Mac
'    If Val(Application.Version) > 14 Then
'        'Find position of extension in filename
'        strDocName = activeDoc.Path
'        intPos = InStrRev(strDocName, ".")
'
'        'Strip off extension and add ".txt" extension
'        strDocName = Left(strDocName, intPos - 1)
'        strDocName = strDocName & "_CIP.txt"
'    End If
'#Else                           'For Windows
'    'Find position of extension in filename
'    strDocName = activeDoc.Name
'    DocPath = activeDoc.Path
'    intPos = InStrRev(strDocName, ".")
'
'    'Strip off extension and add ".txt" extension
'    strDocName = Left(strDocName, intPos - 1)
'    strDocName = DocPath & "\" & strDocName & "_CIP.txt"
'#End If
    
'Copy text of active document and paste into a new document
'Because otherwise open document is converted to .txt, and we want it to stay .doc*
' ^ 2/23/21-mr- we can revise this if we end up closing docx anyways
activeDoc.Select
Selection.Copy

'DebugPrint Len(Selection)
'Because if Len = 1, then no text in doc (only a paragraph return) and causes an error
If Len(Selection) > 1 Then
'PasteSpecial because otherwise gives a warning about too many styles being pasted
    Set tmpDoc = Documents.Add(visible:=False)
    tmpDoc.Content.PasteSpecial Datatype:=wdPasteText
Else
    MsgBox "Your document doesn't appear to have any content. " & _
            "This macro needs a styled manuscript to run correctly.", vbCritical, "Oops!"
    Exit Sub
End If
        
' Set different text encoding based on OS
' And Mac can't create file with line breaks
#If Mac Then
    If Val(Application.Version) > 14 Then
        encodingFmt = msoEncodingMacRoman
        lineBreak = False
    End If
#Else               'For Windows
    encodingFmt = msoEncodingUSASCII
    lineBreak = True
#End If
    
'Turn off alerts because PC warns before saving with this encoding
Application.DisplayAlerts = wdAlertsNone

'Save new document as a text file. Encoding/Line Breaks/Substitutions per LOC info
tmpDoc.SaveAs fileName:=strdocname, _
    FileFormat:=wdFormatEncodedText, _
    Encoding:=encodingFmt, _
    InsertLineBreaks:=lineBreak, _
    AllowSubstitutions:=True
           
' wrap up
Application.DisplayAlerts = wdAlertsAll

tmpDoc.Close SaveChanges:=wdDoNotSaveChanges
'activeDoc.Close SaveChanges:=wdDoNotSaveChanges    ' < optionally...
        
'Application.ScreenUpdating = True
    
' update progress bar - done
completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
Clean_helpers.updateStatus ("")
    
End Sub



Sub undoChanges()

' begin update progress bar
thisstatus = "* Undoing changes to original document "
If Not pBar Is Nothing Then Clean_helpers.updateStatus (thisstatus)

' Undo changes in original doc back to UNDO bookmark:
While activeDoc.Bookmarks.Exists(BM_IN_MACRO)
    activeDoc.Undo
Wend

' update progress bar - done
completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
Clean_helpers.updateStatus ("")

End Sub

Private Sub zz_clearFindB()

    Dim clearRng As Range
    Set clearRng = activeDoc.Words.First
    
    With clearRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = ""
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute
    End With
End Sub

