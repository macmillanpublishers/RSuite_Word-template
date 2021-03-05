Attribute VB_Name = "CIPmacro"
Const sectionFileBasename As String = "sections.txt"
Const containerFileBasename As String = "containers.txt"
Const breakFileBasename As String = "breaks.txt"
Const containerEndStylename As String = "END (END)"
Const bookmakerPiStylename As String = "Bookmaker-Processing-Instruction (Bpi)"
Const designNoteStylename As String = "Design-Note (Dn)"
Dim activeDoc As Document
Dim vPrompt As cipVolumePrompt

Sub Main()
    On Error GoTo ErrorHandler
    
    ' ========================== Dimming and Setting Variables ======================
    Dim successBool As Boolean, tagChaptersBool As Boolean, tagsPresent As Boolean
    Dim tpName As String, cpName As String, spName As String, tocName As String, chName As String
    Dim tpTag As String, cpTag As String, spTag As String, tocTag As String, chTag As String
    Dim tpDisplayName As String, cpDisplayName As String, spDisplayName As String, tocDisplayName As String
    Dim tpRequired As Boolean, cpRequired As Boolean, spRequired As Boolean, tocRequired As Boolean
    Dim bmsectionarray(), tagArray(), chNamesArray(), tagNameArray(), tagDisplayNameArray(), tagRequiredArray()
    Dim sectionArray, bmStyleArray, containerArray, breakArray, chStyleArray
    Dim maxsectionlength As Long
    Dim lastChapParaIndex As Long
    Dim originalDoc As Document, tmpDoc As Document
    Dim tmpDocName As String
    
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
    
    ' ========================== Running Checks, UI Setup ==============================
    ' pre check: is doc saved, protected, track changes etc
    ' True means a check failed (e.g., doc protection on)
    
    ' set activeDoc, check if saved, check protection, check track changes.
    If StartupSettings_CIP() = True Then
        GoTo ProcessExit
    End If
    
    ' pre check: volume userform
    tagChaptersBool = volumeCheckPrompt
    If vPrompt.Cancelled = True Then
        GoTo ProcessExit
    End If
    Unload vPrompt
    
    ' start progress bar
    Set pBar = New Progress_Bar
    pBarCounter = 0
    pBar.Caption = "CIP Application Tagging Macro"
    thisstatus = "* Running pre-tagging checks "
    completeStatus = "Starting CIP Macro" + vbNewLine + "=========================" + _
        vbNewLine + thisstatus
    pBar.Status.Caption = completeStatus
    Clean_helpers.updateStatus ("")
    
    Call clearFind
    
    ' pre-check: are tags already present?
    If preCheckTags(tagArray, chTag) = True Then
        GoTo ProcessExit
    End If
    'update pre-check progress bar
    completeStatus = completeStatus + "100%" & Clean_helpers.updateStatus("")
    
    ' create tmpDoc, set as activeDoc
    ' Create var to track original active doc
    Set originalDoc = activeDoc
    Set tmpDoc = Documents.Add(activeDoc.FullName, visible:=False)
    tmpDocName = tmpDoc.Name
    Set activeDoc = tmpDoc
    activeDoc.TrackRevisions = False 'stop tracking on tmpDoc (if we were still tracking).
    ' strip content controls, fieldcodes < this may not be necessary / apropos for CIP tagging
    '   but was part of prior incarnation of the application
    Call rmContentCntrlsAndFieldCodes(activeDoc)
    
    ' ========================== Perform tagging, Tag Report ==============================
    ' insert tags for FM sections
    sectionArray = getStyleArrayfromFile(sectionFileBasename)
    Call tagFMSections(sectionArray, maxsectionlength, tagNameArray, tagDisplayNameArray, tagArray)
    
    ' tag chapters
    chStyleArray = getMultiSectionStyleNames(chNamesArray)
    Call tagChapters(tagChaptersBool, chStyleArray, chTag, bmsectionarray)
    
    ' Run tag summary / post-checks
    tagsPresent = reportOnTags(tagArray, tagDisplayNameArray, tagRequiredArray, chTag, tagChaptersBool, originalDoc)
    If tagsPresent = False Then
        GoTo ProcessExit
    End If
    
    ' ========================== Rm non-content, save file, cleanup ==============================
    ' Clean document of Sections and container paras, etc
    Call stripNonContentParas(sectionArray)
    
    ' save out text file
    If SaveAsTextFile(originalDoc) = False Then
        GoTo ProcessExit
    End If
    
    'flag a successful finish for cleanOnExit
    successBool = True
    
ProcessExit:
    Call cleanOnExit(successBool, originalDoc, tmpDoc, tmpDocName)
    Exit Sub

ErrorHandler:
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & Err.Number & " - " & Err.Description
    Resume ProcessExit
End Sub

Function StartupSettings_CIP() As Boolean
    On Error GoTo StartupSettings_CIPError
    ' records/adjusts/checks settings and stuff before running the rest of the macro
    ' returns TRUE if some check is bad and we can't run the macro
    
    ' activeDoc is global variable to hold our document, so if user clicks a different
    ' document during execution, won't switch to that doc.
    ' ALWAYS set to Nothing first to reset for this macro.
    ' Then only refer to this object, not ActiveDocument directly.
     Set activeDoc = Nothing
     Set activeDoc = ActiveDocument
    
    ' check if file has doc protection on, quit function if it does
    If activeDoc.ProtectionType <> wdNoProtection Then
      'If WT_Settings.InstallType = "server" Then
      '  Err.Raise MacError.err_DocProtectionOn
      'Else
        MsgBox "Uh oh ... protection is enabled on document '" & activeDoc.Name & "'." & vbNewLine & _
          "Please unprotect the document and run the macro again." & vbNewLine & vbNewLine & _
          "TIP: If you don't know the protection password, try pasting contents of this file into " & _
          "a new file, and run the macro on that.", , "Error 2"
        StartupSettings_CIP = True
        Exit Function
      'End If
    End If
    
    ' check if file has been saved (we can assume it was for Validator)
    'If WT_Settings.InstallType = "user" Then
    Dim iReply As Integer
    Dim docSaved As Boolean
    docSaved = activeDoc.Saved
    
    If docSaved = False Then
      iReply = MsgBox("Your document '" & activeDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
          "Click OK to save your document and run the macro." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", _
              vbOKCancel, "WARNING")
      If iReply = vbOK Then
        activeDoc.Save
      Else
        StartupSettings_CIP = True
        Exit Function
      End If
    End If
    'End If

    ' ========== Turn off screen updating ==========
    Application.ScreenUpdating = False
    
    ' ========== Check if changes present and offer to accept all ==========
    ' FixTrackChanges_CIP returns false if changes are present and user cancels cleanup
    
    If FixTrackChanges_CIP = False Then
        StartupSettings_CIP = True
    End If
    
    Exit Function

StartupSettings_CIPError:
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & _
        Err.Number & " - " & Err.Description & " - Sub:StartupSettings"
    StartupSettings_CIP = True
End Function

Sub rmContentCntrlsAndFieldCodes(thisDoc As Document)
' These are borrowed / modified from MacroHelpers.StartupSettings macro
' Running on tmpDoc.
' Leaving out bookmark removal, since they have no effect on txt output.

'  ' ========== Remove content controls ==========
' This can't be run on Mac; breaks the sub. Since this is a true edge case,
'   and may have no affect on content in txtdoc anyways,
'   not worrying about alerting the occasional Mac user
' MacroHelpers has another function specifically for collating content from cookstr templated cc's
'   since this project is defunct for awhile now, not including it.
  #If Not Mac Then
      Call MacroHelpers.ClearContentControls(thisDoc)
  #End If

'' ========== Delete field codes ==========
'' Fields break cleanup and char styles, so we delete them (but retain their
'' result, if any). Furthermore, fields make no sense in a manuscript, so
'' even if they didn't break anything we don't want them.
'' Note, however, that even though linked endnotes and footnotes are
'' types of fields, this loop doesn't affect them.
'' NOTE: Must run AFTER content control cleanup.

  Dim colStoriesUsed As Collection
  Set colStoriesUsed = MacroHelpers.ActiveStories(thisDoc)
  Call MacroHelpers.UpdateUnlinkFieldCodes(colStoriesUsed, thisDoc)

End Sub

Private Function FixTrackChanges_CIP() As Boolean
    ' borrowed/copied from MacroHelpers module. Main diff. is msgbox text.
    ' returns True if changes were fixed or not present, False if changes remain in doc
    Dim n As Long
    Dim oComments As Comments
    Set oComments = activeDoc.Comments
    Dim tcPresentBool As Boolean
    
    FixTrackChanges_CIP = True
    tcPresentBool = False
    
    ' check if TC are present.
    Dim stry As Object
    For Each stry In activeDoc.StoryRanges
        If stry.Revisions.Count >= 1 Then tcPresentBool = True
    Next
    
    'Debug.Print "commentscount " & oComments.Count & " tc " & tcPresentBool
    
    ' If there are changes, ask user if they want macro to accept changes or cancel
    If oComments.Count > 0 Or tcPresentBool = True Then
        If MsgBox("Comments or tracked changes are present in this file. They will result in problems with output CIP.txt." _
          & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the CIP Macro." _
          & vbCr & vbCr & "Click CANCEL to stop the CIP Macro and deal with the tracked changes and comments on your own.", _
          273, "Alert") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
              FixTrackChanges_CIP = False
              Exit Function
        Else 'User clicked OK, so accept all tracked changes and delete all comments
          activeDoc.AcceptAllRevisions
          For n = oComments.Count To 1 Step -1
              oComments(n).Delete
          Next n
          Set oComments = Nothing
        End If
    Else
      FixTrackChanges_CIP = True
    End If
    
End Function

Sub cleanOnExit(successBool, originalDoc, tmpDoc, tmpDocName)
    On Error GoTo cleanOnExit_ErrorHandler:
    
    ' close progress bar if up
    If isFormLoaded("Progress_Bar") = True Then Unload pBar
    ' close voluime prompt if its still loaded
    If isFormLoaded("cipVolumePrompt") = True Then Unload vPrompt
    ' close tmpdoc if it is defined and open
    If Not tmpDoc Is Nothing Then
        Dim i As Long
        For i = 1 To Application.Documents.Count
            If Application.Documents(i).Name = tmpDocName Then
                tmpDoc.Close savechanges:=wdDoNotSaveChanges
                Exit For
            End If
        Next i
    End If
    ' if we've created originalDoc object, make sure that is activeDoc, and activated
    If Not originalDoc Is Nothing Then
        originalDoc.Activate
        Set activeDoc = originalDoc
    End If
    ' clear find
    Call clearFind
    ' force enable alerts, screenupdating
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    ' display pop-up for a complete run
    If successBool = True Then
        Call Clean_helpers.MessageBox("Done", "CIP Tagging macro complete!", vbOK)
    End If
    Exit Sub
    
cleanOnExit_ErrorHandler:
    ' force enable alerts, screenupdating, msgbox
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & _
        Err.Number & " - " & Err.Description & "Sub:cleanOnExit"
    ' \/ Halts ALL execution, resets all variables, unloads all userforms, etc.
    End
End Sub
Private Function isFormLoaded(formname As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formname Then
            isFormLoaded = True
            Exit Function
        End If
    Next frm
    isFormLoaded = False
End Function


Sub debugCheckOpenItems()
    For i = 1 To Application.Documents.Count
        Debug.Print Application.Documents(i).Name
    Next i
    Debug.Print VBA.UserForms.Count
    
    For Each frm In VBA.UserForms
        Debug.Print "fn: " & frm.Name
    Next frm
End Sub

Private Function volumeCheckPrompt()
    Set vPrompt = New cipVolumePrompt
    
    With vPrompt
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
    
    volumeCheckPrompt = vPrompt.tagChapters
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
Private Function reportOnTags(tagArray, tagDisplayNames, tagRequired, chTag, tagChaptersBool, originalDoc As Document) As Boolean
    On Error GoTo reportOnTagsError
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
    Call clearFind
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
        reportStr = reportStr & "LOC Tags look good for " & originalDoc.Name & vbNewLine
        reportStr = reportStr & "See summary below:" & vbNewLine & vbNewLine
    Else
        reportStr = reportStr & "BAD NEWS:" & vbNewLine & vbNewLine
        reportStr = reportStr & "Problems were found with LOC tags in your document '" & originalDoc.Name & "':" & vbNewLine
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
    Dim thisDoc As Document
    Set thisDoc = activeDoc
    Call MacroHelpers.CreateTextFile(strText:=reportStr, suffix:="CIPtagReport", thisDoc:=originalDoc)

    ' update progress bar
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    Clean_helpers.updateStatus ("")
    
    Exit Function
    
reportOnTagsError:
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & _
        Err.Number & " - " & Err.Description & " - Sub:ReportOnTags"
    reportOnTags = False
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
Private Function getFilePathWithNewSuffix(myDocument, suffixStr) 'expects file extension too, ex: "_2.docx"
    Dim strdocname As String
    'Separate code by OS because myDocument.Path returns file name too
    ' on Mac but doesn't for PC
    #If Mac Then        'For Mac
        If Val(Application.Version) > 14 Then
            'Find position of extension in filename
            strdocname = myDocument.Path
            intPos = InStrRev(strdocname, ".")
            
            'Strip off extension and add ".txt" extension
            strdocname = Left(strdocname, intPos - 1)
            strdocname = strdocname & suffixStr
        End If
    #Else                           'For Windows
        'Find position of extension in filename
        strdocname = myDocument.Name
        DocPath = myDocument.Path
        intPos = InStrRev(strdocname, ".")
        
        'Strip off extension and add ".txt" extension
        strdocname = Left(strdocname, intPos - 1)
        strdocname = DocPath & "\" & strdocname & suffixStr
    #End If
    getFilePathWithNewSuffix = strdocname
End Function
Private Function SaveAsTextFile(originalDocument) As Boolean
    On Error GoTo SaveAsTextFileError
    ' Saves a copy of the document as a text file in the same path as the parent document
    Dim txtDoc As Document
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
    
    strdocname = getFilePathWithNewSuffix(originalDocument, "_CIP.txt")
    
    'Copy text of active document and paste into a new document
    'Because otherwise open document is converted to .txt, and we want it to stay .doc*
    ' ^ 2/23/21-mr- we can revise this if we end up closing docx anyways
    activeDoc.Select
    Selection.Copy
    
    'DebugPrint Len(Selection)
    'Because if Len = 1, then no text in doc (only a paragraph return) and causes an error
    If Len(Selection) > 1 Then
    'PasteSpecial because otherwise gives a warning about too many styles being pasted
        Set txtDoc = Documents.Add(visible:=False)
        txtDoc.Content.PasteSpecial Datatype:=wdPasteText
    Else
        MsgBox "Your document doesn't appear to have any content. " & _
                "This macro needs a styled manuscript to run correctly.", vbCritical, "Oops!"
        Exit Function
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
    txtDoc.SaveAs fileName:=strdocname, _
        FileFormat:=wdFormatEncodedText, _
        Encoding:=encodingFmt, _
        InsertLineBreaks:=lineBreak, _
        AllowSubstitutions:=True
               
    ' wrap up
    Application.DisplayAlerts = wdAlertsAll
    
    Debug.Print
    txtDoc.Close savechanges:=wdDoNotSaveChanges
    activeDoc.Close savechanges:=wdDoNotSaveChanges
            
    'Application.ScreenUpdating = True
        
    ' update progress bar - done
    completeStatus = completeStatus + vbNewLine + thisstatus + "100%"
    Clean_helpers.updateStatus ("")
    
    SaveAsTextFile = True
    
    Exit Function
    
SaveAsTextFileError:
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & _
        Err.Number & " - " & Err.Description & " - Sub:SaveAsTextFile"
    SaveAsTextFile = False
End Function
Private Sub clearFind()
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

