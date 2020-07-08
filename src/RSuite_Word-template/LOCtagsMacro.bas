Attribute VB_Name = "LOCtagsMacro"
' ======= PURPOSE ============================
' Produces a text file of the manuscript with tags added for CIP application

' ======= DEPENDENCIES =======================
' 1. Requires ProgressBar userform and MacroHelpers module
' 2. Manuscript must be tagged with Macmillan styles.
      
Option Explicit
Option Base 1
Dim activeRng As Range

Sub LibraryOfCongressTags()

    '''''''''''''''''''''''''''''''''
    '''created by Matt Retzer  - matthew.retzer@macmillan.com
    '''2/25/15
    '''Version 1.6
    '''Updated: 4/14/15: adding progress bar
    '''Updated: 4/13/15: adding content control handling for PC
    '''Updated: 3/4/15 : revised chapter numbering loop for performance, edited ELC styles and added tag for ELC with no end styles
    '''Updated: 3/8/15 : switching to Whole word searches for the 4 items with closing tags
    '''                           : & allowing for ^m Page Break check/fix to get </ch> inline with final chapter text
    '''         3/10/15 : revamped ELC </ch> again to make it inline.
    '''                 : used same method to make cp, tp, toc and sp closing tags inline-- match whole words broke with hyperlinks
    '''         3/24/15 : re-did ELC in case of atax or other styles present early in manuscript ; uses while loop to scan for first backmatter style that
    '''                 is not eventually followed by <ch#> or <tp> tag
    ''''''''''''''''''''''''''''''

    
    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
    If StartupSettings() = True Then
        Call Cleanup
        Exit Sub
    End If
    
    
    '-----------run preliminary error checks------------
    Dim skipChapterTags As Boolean
    
    If zz_errorChecksB <> False Then
        Call zz_clearFindB
        Call Cleanup
        Exit Sub
    End If
    
    skipChapterTags = volumestylecheck()
    
    
'    '--------Progress Bar------------------------------
'    'Percent complete and status for progress bar (PC) and status bar (Mac)
'    'Requires ProgressBar custom UserForm and Class
'    Dim sglPercentComplete As Single
'    Dim strStatus As String
'    Dim strTitle As String
'
'    'First status shown will be randomly pulled from array, for funzies
'    Dim funArray() As String
'    ReDim funArray(1 To 10)      'Declare bounds of array here
'
'    funArray(1) = "* Checking out library books..."
'    funArray(2) = "* Returning overdue library books..."
'    funArray(3) = "* Reshelving books..."
'    funArray(4) = "* Researching a term paper..."
'    funArray(5) = "* Calling the Librarian of Congress..."
'    funArray(6) = "* Adjusting reading glasses..."
'    funArray(7) = "* Sshhh!..."
'    funArray(8) = "* Roaming the stacks..."
'
'    Dim x As Integer
'
'    'Rnd returns random number between (0,0.8], rest of expression is to return an integer (1,8)
'    Randomize           'Sets seed for Rnd below to value of system timer
'    x = Int(UBound(funArray()) * Rnd()) + 1
'
'    'DebugPrint x
'
'    strTitle = "CIP Application Tagging Macro"
'    sglPercentComplete = 0.1
'    strStatus = funArray(x)
'
'    'All Progress Bar statements for PC only because won't run modeless on Mac
'    Dim oProgressCIP As Progress_Bar
'    Set oProgressCIP = New Progress_Bar  ' Triggers Initialize event, which uses Show method for PC
'
'    oProgressCIP.Title = strTitle
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
'
'
'    '=========================the rest of the macro========================
'
'    '-------------------------tagging Title page---------------------------
'    sglPercentComplete = 0.2
'    strStatus = "* Adding tags for Title page..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call tagTitlePage
    Call zz_clearFindB
    
    '-------------------------tagging Copyright page---------------------------
'    sglPercentComplete = 0.3
'    strStatus = "* Adding tags for Copyright page..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call tagCopyrightPage
    Call zz_clearFindB
        
    '-------------------------tagging Series page---------------------------
'    sglPercentComplete = 0.4
'    strStatus = "* Adding tags for Series page..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call tagSeriesPage
    Call zz_clearFindB
    
    '-------------------------tagging Table of Contents---------------------------
'    sglPercentComplete = 0.5
'    strStatus = "* Adding tags for Table of Contents..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call tagTOC
    Call zz_clearFindB
    
    '-------------------------tagging Chapters---------------------------
'    sglPercentComplete = 0.6
'    strStatus = "* Adding tags for chapters..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    If skipChapterTags = False Then
        Call tagChapterHeads
        Call zz_clearFindB
        
        Call tagEndLastChapter
        Call zz_clearFindB
    End If
    
    
    '-------------------------Checking tags--------------------------
'    sglPercentComplete = 0.7
'    strStatus = "* Running tag check & generating report..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    If zz_TagReport = False Then
'        oProgressCIP.Hide
        
        Dim strMessage As String
        strMessage = "CIP tags cannot be added because no paragraphs were tagged with Macmillan styles for the titlepage, " & _
            "copyright page, table of contents, or chapter title pages. Please add the correct styles and try again."
        MsgBox strMessage, vbCritical, "No Styles Found"
        
        GoTo Finish
        
        Exit Sub
        
    End If
    
    '-------------------------Save as text doc---------------------------
'    sglPercentComplete = 0.8
'    strStatus = "* Saving as text document..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call SaveAsTextFile
    

    '-------------------------Delete tags from orig doc---------------------------
'    sglPercentComplete = 0.9
'    strStatus = "* Cleaning up file..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call cleanFile
    Call zz_clearFindB
    
    '-------------------------cleanup---------------------------
'    sglPercentComplete = 0.99
'    strStatus = "* Finishing up..." & vbNewLine & strStatus
'
'    Call UpdateBarAndWait(Bar:=oProgressCIP, Status:=strStatus, Percent:=sglPercentComplete)
   
Finish:
    Call Cleanup
'    Unload oProgressCIP
    
    'If skipChapterTags = True Then
    '    MsgBox "Library of Congress tagging is complete except for Chapter tags." & vbNewLine & vbNewLine & "Chapter tags will need to be manually applied."
    'End If
    
End Sub

Private Sub tagChapterHeads()
    Set activeRng = activeDoc.Range
    Dim CHstylesArray(3) As String                                   ' number of items in array should be declared here
    Dim i As Long
    Dim chTag As Integer
    
    CHstylesArray(1) = "Chap Number (cn)"
    CHstylesArray(2) = "Chap Title (ct)"
    CHstylesArray(3) = "Chap Title Nonprinting (ctnp)"
    
On Error GoTo errHandler
    
    For i = 1 To UBound(CHstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "`CH|^&|CH`"
      .Wrap = wdFindContinue
      .Format = True
      .Style = CHstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = True
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
    Next
    
On Error GoTo 0

    Call zz_clearFindB
    
    Dim CHfauxTags(4) As String         ' number of items in arrays should be declared here
    Dim CHLOCtags(4) As String
    
    CHfauxTags(1) = "`CH||CH`"
    CHfauxTags(2) = "|CH``CH|"
    CHfauxTags(3) = "|CH`"
    CHfauxTags(4) = "`CH|"
                                                       
    CHLOCtags(1) = ""
    CHLOCtags(2) = ""
    CHLOCtags(3) = ""
    CHLOCtags(4) = "<ch>"
    
    For i = 1 To UBound(CHfauxTags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = CHfauxTags(i)
      .Replacement.Text = CHLOCtags(i)
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
    Next
    
    'adapted from fumei's Range find here: http://www.vbaexpress.com/forum/showthread.php?41244-Loop-until-not-found
    chTag = 1
    
    With activeRng.Find
    .Text = "<ch>"
    Do While .Execute(Forward:=True) = True
    With activeRng
    .MoveEnd unit:=wdCharacter, Count:=-1
    .InsertAfter (chTag)
    .Collapse direction:=wdCollapseEnd
    .Move unit:=wdCharacter, Count:=1
    End With
    chTag = chTag + 1
    Loop
    End With
    
    'previous chapter number tag method (too slow)
    'Dim chapNum As Integer
    'Dim chapNumString As String
    'chapNum = 1
    'chapNumString = "<ch" & chapNum & ">"
    '
    ''this is borrowed form here:  http://stackoverflow.com/questions/11234358/word-2007-macro-to-automatically-number-items-in-a-document
    'Do While InStr(activeDoc.Content, "<ch>") > 0
    '    chapNumString = "<ch" & chapNum & ">"
    '    With activeDoc.Content.Find
    '        .ClearFormatting
    '        .Text = "<ch>"
    '        .Execute Replace:=wdReplaceOne, ReplaceWith:=chapNumString, Forward:=True
    '    End With
    '    chapNum = chapNum + 1
    'Loop
    Exit Sub
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If
End Sub

Private Sub tagTitlePage()

    'to update this for a different tag, replace all in procedure for the two char tag, eg: TP->CH ; this will update array variables too
    'update styles array manually, and Dim'd stylesarray length,
    'update the LOC tags to match LOC:  http://www.loc.gov/publish/cip/techinfo/formattingecip.html#tags
    ''' NOTE:  if you are tagging something only at the beginning or end (eg chapter heads), obviously you need to touch up the second loop
    
    Set activeRng = activeDoc.Range
    Dim TPstylesArray(10) As String                                   ' number of items in array should be declared here
    Dim i As Long
    
    TPstylesArray(1) = "Titlepage Author Name (au)"
    TPstylesArray(2) = "Titlepage Book Subtitle (stit)"
    TPstylesArray(3) = "Titlepage Book Title (tit)"
    TPstylesArray(4) = "Titlepage Cities (cit)"
    TPstylesArray(5) = "Titlepage Contributor Name (con)"
    TPstylesArray(6) = "Titlepage Imprint Line (imp)"
    TPstylesArray(7) = "Titlepage Publisher Name (pub)"
    TPstylesArray(8) = "Titlepage Reading Line (rl)"
    TPstylesArray(9) = "Titlepage Series Title (ser)"
    TPstylesArray(10) = "Titlepage Translator Name (tran)"
    
On Error GoTo errHandler

    For i = 1 To UBound(TPstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "`TP|^&|TP`"
      .Wrap = wdFindContinue
      .Format = True
      .Style = TPstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
    Next
On Error GoTo 0
    
    Call zz_clearFindB
    
    Dim TPfauxTags(3) As String
    Dim TPLOCtags(2) As String
    Dim directionBool(2) As Boolean
    
    TPfauxTags(1) = "`TP|"
    TPfauxTags(2) = "|TP`"
    TPfauxTags(3) = "``````"          'this bit is to make sure tagging is inline with last styled paragraph,
                                                        'instead of the tag falling into the following style eblock
    TPLOCtags(1) = "<tp>"
    TPLOCtags(2) = "``````"
    
    directionBool(1) = True
    directionBool(2) = False
    
    For i = 1 To UBound(TPLOCtags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = TPfauxTags(i)
      .Replacement.Text = TPLOCtags(i)
      .Wrap = wdFindContinue
      .Format = False
      .Forward = directionBool(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceOne
    End With
    Next
    
    Call zz_clearFindB
    
    With activeRng.Find
        .Text = "``````"
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
        With activeRng.Find
            .Text = "[!^13^m`]"
            .Replacement.Text = "^&</tp>"
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
    
    Call zz_clearFindB
    
    For i = 1 To UBound(TPfauxTags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = TPfauxTags(i)
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
    Next
    
    Exit Sub
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If
End Sub

Private Sub tagCopyrightPage()

    'to update this for a different tag, replace all in procedure two char code, eg: TP->CP
    'update styles array manually, and Dim'd stylesarray length, & that's it
    
    Set activeRng = activeDoc.Range
    Dim CPstylesArray(2) As String                                   ' number of items in array should be declared here
    Dim i As Long
    
    CPstylesArray(1) = "Copyright Text double space (crtxd)"
    CPstylesArray(2) = "Copyright Text single space (crtx)"

On Error GoTo errHandler

    For i = 1 To UBound(CPstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "`CP|^&|CP`"
      .Wrap = wdFindContinue
      .Format = True
      .Style = CPstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
    Next
On Error GoTo 0
    
    Call zz_clearFindB
    
    Dim CPfauxTags(3) As String
    Dim CPLOCtags(2) As String
    Dim directionBool(2) As Boolean
    
    CPfauxTags(1) = "`CP|"
    CPfauxTags(2) = "|CP`"
    CPfauxTags(3) = "``````"          'this bit is to make sure tagging is inline with last styled paragraph,
                                                        'instead of the tag falling into the following style eblock
    CPLOCtags(1) = "<cp>"
    CPLOCtags(2) = "``````"
    
    directionBool(1) = True
    directionBool(2) = False
    
    For i = 1 To UBound(CPLOCtags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = CPfauxTags(i)
      .Replacement.Text = CPLOCtags(i)
      .Wrap = wdFindContinue
      .Format = False
      .Forward = directionBool(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceOne
    End With
    Next
    
    Call zz_clearFindB
    
    With activeRng.Find
        .Text = "``````"
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
        With activeRng.Find
            .Text = "[!^13^m`]"
            .Replacement.Text = "^&</cp>"
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
    
    Call zz_clearFindB
    
    For i = 1 To UBound(CPfauxTags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = CPfauxTags(i)
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
    Next
    
    Call zz_clearFindB
    
    Exit Sub
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If
End Sub

Private Sub tagTOC()

    'to update this for a different tag, replace all in procedure two char code, eg: TP->TOC
    'update styles array manually, and Dim'd stylesarray length, & that's it
    
    Set activeRng = activeDoc.Range
    Dim TOCstylesArray(10) As String                                   ' number of items in array should be declared here
    Dim i As Long
    
    TOCstylesArray(1) = "TOC Frontmatter Head (cfmh)"
    TOCstylesArray(2) = "TOC Author (cau)"
    TOCstylesArray(3) = "TOC Part Number  (cpn)"
    TOCstylesArray(4) = "TOC Part Title (cpt)"
    TOCstylesArray(5) = "TOC Chapter Number (ccn)"
    TOCstylesArray(6) = "TOC Chapter Title (cct)"
    TOCstylesArray(7) = "TOC Chapter Subtitle (ccst)"
    TOCstylesArray(8) = "TOC Level-1 Chapter Head (ch1)"
    TOCstylesArray(9) = "TOC Backmatter Head (cbmh)"
    TOCstylesArray(10) = "TOC Page Number (cnum)"
    
On Error GoTo errHandler
    For i = 1 To UBound(TOCstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "`TOC|^&|TOC`"
      .Wrap = wdFindContinue
      .Format = True
      .Style = TOCstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
    Next
On Error GoTo 0
    
    Call zz_clearFindB
    
    Dim TOCfauxTags(3) As String
    Dim TOCLOCtags(2) As String
    Dim directionBool(2) As Boolean
    
    TOCfauxTags(1) = "`TOC|"
    TOCfauxTags(2) = "|TOC`"
    TOCfauxTags(3) = "``````"          'this bit is to make sure tagging is inline with last styled paragraph,
                                                        'instead of the tag falling into the following style eblock
    TOCLOCtags(1) = "<toc>"
    TOCLOCtags(2) = "``````"
    
    directionBool(1) = True
    directionBool(2) = False
    
    For i = 1 To UBound(TOCLOCtags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = TOCfauxTags(i)
      .Replacement.Text = TOCLOCtags(i)
      .Wrap = wdFindContinue
      .Format = False
      .Forward = directionBool(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceOne
    End With
    Next
    
    Call zz_clearFindB
    
    With activeRng.Find
        .Text = "``````"
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
        With activeRng.Find
            .Text = "[!^13^m`]"
            .Replacement.Text = "^&</toc>"
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
    
    Call zz_clearFindB
    
    For i = 1 To UBound(TOCfauxTags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = TOCfauxTags(i)
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
    Next
    
    Exit Sub
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If
End Sub

Private Sub tagSeriesPage()

    'to update this for a different tag, replace all in procedure two char code, eg: TP->SP
    'update styles array manually, and Dim'd stylesarray length, & that's it
    
    Set activeRng = activeDoc.Range
    Dim SPstylesArray(8) As String                                   ' number of items in array should be declared here
    Dim i As Long
    
    SPstylesArray(1) = "Series Page Heading (sh)"
    SPstylesArray(2) = "Series Page Text (stx)"
    SPstylesArray(3) = "Series Page Text No-Indent (stx1)"
    SPstylesArray(4) = "Series Page List of Titles (slt)"
    SPstylesArray(5) = "Series Page Author (sau)"
    SPstylesArray(6) = "Series Page Subhead 1 (sh1)"
    SPstylesArray(7) = "Series Page Subhead 2 (sh2)"
    SPstylesArray(8) = "Series Page Subhead 3 (sh3)"
    
On Error GoTo errHandler

    For i = 1 To UBound(SPstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "`SP|^&|SP`"
      .Wrap = wdFindContinue
      .Format = True
      .Style = SPstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
    Next
On Error GoTo 0

    Call zz_clearFindB
    
    Dim SPfauxTags(3) As String
    Dim SPLOCtags(2) As String
    Dim directionBool(2) As Boolean
    
    SPfauxTags(1) = "`SP|"
    SPfauxTags(2) = "|SP`"
    SPfauxTags(3) = "``````"          'this bit is to make sure tagging is inline with last styled paragraph,
                                                        'instead of the tag falling into the following style eblock
    SPLOCtags(1) = "<sp>"
    SPLOCtags(2) = "``````"
    
    directionBool(1) = True
    directionBool(2) = False
    
    For i = 1 To UBound(SPLOCtags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = SPfauxTags(i)
      .Replacement.Text = SPLOCtags(i)
      .Wrap = wdFindContinue
      .Format = False
      .Forward = directionBool(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceOne
    End With
    Next
    
    Call zz_clearFindB
    
    With activeRng.Find
        .Text = "``````"
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
        With activeRng.Find
            .Text = "[!^13^m`]"
            .Replacement.Text = "^&</sp>"
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
    
    Call zz_clearFindB
    
    For i = 1 To UBound(SPfauxTags())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = SPfauxTags(i)
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
    Next
    
    Exit Sub
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If

End Sub

Private Sub tagEndLastChapter()

    Set activeRng = activeDoc.Range
    Dim ELCstylesArray(9) As String                                   ' number of items in array should be declared here
    Dim i As Long
    
    ELCstylesArray(1) = "BM Head (bmh)"
    ELCstylesArray(2) = "BM Title (bmt)"
    ELCstylesArray(3) = "Appendix Head (aph)"
    ELCstylesArray(4) = "Appendix Subhead (apsh)"
    ELCstylesArray(5) = "Note Level-1 Subhead (n1)"
    ELCstylesArray(6) = "Biblio Level-1 Subhead (b1)"
    ELCstylesArray(7) = "About Author Text (atatx)"
    ELCstylesArray(8) = "About Author Text No-Indent (atatx1)"
    ELCstylesArray(9) = "About Author Text Head (atah)"
    
On Error GoTo errHandler
    For i = 1 To UBound(ELCstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "``````^&"
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
On Error GoTo 0

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
    testtag = "\<ch[0-9]{1,}\>"
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
            .Text = "``````"
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
                    .Text = "``````"
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
                    .Text = "``````"
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
                            .Replacement.Text = "^&</ch>"
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
      .Text = "``````"
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
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If
End Sub

Private Sub SaveAsTextFile()
 
 ' Saves a copy of the document as a text file in the same path as the parent document
    Dim strDocName As String
    Dim DocPath As String
    Dim intPos As Integer
    Dim encodingFmt As String
    Dim linebreak As Boolean

    Application.ScreenUpdating = False
    
    'Separate code by OS because activeDoc.Path returns file name too
    ' on Mac but doesn't for PC
    
    #If Mac Then        'For Mac
        If Val(Application.Version) > 14 Then
            
            'Find position of extension in filename
            strDocName = activeDoc.Path
            intPos = InStrRev(strDocName, ".")
            
                'Strip off extension and add ".txt" extension
                strDocName = Left(strDocName, intPos - 1)
                strDocName = strDocName & "_CIP.txt"
            
        End If
        
    #Else                           'For Windows
    
        'Find position of extension in filename
        strDocName = activeDoc.Name
        DocPath = activeDoc.Path
        intPos = InStrRev(strDocName, ".")
        
                'Strip off extension and add ".txt" extension
                strDocName = Left(strDocName, intPos - 1)
                strDocName = DocPath & "\" & strDocName & "_CIP.txt"
            
    #End If
    
        'Copy text of active document and paste into a new document
        'Because otherwise open document is converted to .txt, and we want it to stay .doc*

        activeDoc.Select
        Selection.Copy

        'DebugPrint Len(Selection)
        'Because if Len = 1, then no text in doc (only a paragraph return) and causes an error
        If Len(Selection) > 1 Then
        'PasteSpecial because otherwise gives a warning about too many styles being pasted
            Documents.Add.Content.PasteSpecial DataType:=wdPasteText
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
            linebreak = False
        End If
    #Else               'For Windows
        encodingFmt = msoEncodingUSASCII
        linebreak = True
    #End If
    
    'Turn off alerts because PC warns before saving with this encoding
    Application.DisplayAlerts = wdAlertsNone
    
        'Save new document as a text file. Encoding/Line Breaks/Substitutions per LOC info
        activeDoc.SaveAs FileName:=strDocName, _
            FileFormat:=wdFormatEncodedText, _
            Encoding:=encodingFmt, _
            InsertLineBreaks:=linebreak, _
            AllowSubstitutions:=True
            
    Application.DisplayAlerts = wdAlertsAll
        Documents(strDocName).Close
        
    Application.ScreenUpdating = True
    
End Sub

Private Sub cleanFile()
    Set activeRng = activeDoc.Range
    Dim tagsFind(10) As String         ' number of items in arrays should be declared here
    Dim a As Long
    
    tagsFind(1) = "\<tp\>"
    tagsFind(2) = "\<\/tp\>"
    tagsFind(3) = "\<cp\>"
    tagsFind(4) = "\<\/cp\>"
    tagsFind(5) = "\<sp\>"
    tagsFind(6) = "\<\/sp\>"
    tagsFind(7) = "\<toc\>"
    tagsFind(8) = "\<\/toc\>"
    tagsFind(9) = "\<ch*\>"
    tagsFind(10) = "\<\/ch\>"
    
    For a = 1 To UBound(tagsFind())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = tagsFind(a)
      .Replacement.Text = ""
      .Wrap = wdFindContinue
      .Format = False
      .Forward = True
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = True
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
    Next

End Sub

Private Function volumestylecheck()

    Set activeRng = activeDoc.Range
    volumestylecheck = False
    Dim VOLstylesArray(2) As String                                   ' number of items in array should be declared here
    Dim i As Long
    Dim mainDoc As Document
    Set mainDoc = activeDoc
    Dim iReply As Integer
    
    VOLstylesArray(1) = "Volume Number (voln)"
    VOLstylesArray(2) = "Volume Title (volt)"

On Error GoTo errHandler

    For i = 1 To UBound(VOLstylesArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Wrap = wdFindContinue
      .Format = True
      .Style = VOLstylesArray(i)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      If .Execute Then                          'Returns true if text was found.
         iReply = MsgBox(mainDoc & "' contains a 'Volume' paragraph style." & vbNewLine & vbNewLine & _
            "To continue submitting this for Library of Congress ingestion as a single volume (standard tags), click 'YES'." & vbNewLine & vbNewLine & _
            "If submitting as a 'single application for multiple volumes': click 'NO' to proceed with auto-tagging exempting chapter tags (<ch></ch>)." & vbNewLine & _
            "Chapter tags wil unfortunately need to be manually applied in this case." & vbNewLine & vbNewLine & _
            "For further guidance please email macsupport@macmillanusa.com", vbYesNoCancel, "Alert")
        If iReply = vbYes Then
            Exit Function
        ElseIf iReply = vbNo Then
            volumestylecheck = True
            Exit Function
        Else
            End
        End If
      End If
    End With
NextLoop:
    Next
On Error GoTo 0

    Exit Function
    
errHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        Resume NextLoop
    End If

End Function



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


Private Function zz_errorChecksB()                       'kidnapped this whole function from macmillan.dotx
                                                                'adding tag checking to include LOC stuff
    zz_errorChecksB = False
    Dim mainDoc As Document
    Set mainDoc = activeDoc
    Dim iReply As Integer
    
    '-----test if backtick style tag already exists
    Set activeRng = mainDoc.Range
    
    Dim existingTagArray(7) As String                                   ' number of items in array should be declared here
    Dim b As Long
    Dim foundBad As Boolean
    foundBad = False
    
    existingTagArray(1) = "[`|]CH[|`]"
    existingTagArray(2) = "`ELC|"
    existingTagArray(3) = "[`|]CP[|`]"
    existingTagArray(4) = "[`|]TP[|`]"
    existingTagArray(5) = "[`|]TOC[|`]"
    existingTagArray(6) = "[`|]SP[|`]"
    existingTagArray(7) = "``````"
    
    For b = 1 To UBound(existingTagArray())
    With activeRng.Find
      .ClearFormatting
      .Text = existingTagArray(b)
      .Wrap = wdFindContinue
      .MatchWildcards = True
    End With
    If activeRng.Find.Execute Then foundBad = True: Exit For
    Next

    If foundBad = True Then                'If activeRng.Find.Execute Then
        MsgBox "Something went wrong! The LOC tags Macro cannot be run on Document:" & vbNewLine & "'" & mainDoc & "'" _
        & vbNewLine & vbNewLine & "Please contact Digital Workflow group for support, I am sure they will be happy to help.", , "Error Code: 1"
        zz_errorChecksB = True
        Exit Function
    End If
    
    '-----test if LOC tags already exists
    
    Dim existingLOCArray(9) As String
    Dim c As Long
    Dim foundLOC As Boolean
    Dim foundLOCitem As String
    foundLOC = False
    Dim iReplyB As Integer
    
    existingLOCArray(1) = "<sp>"
    existingLOCArray(2) = "</sp>"
    existingLOCArray(3) = "</ch>"
    existingLOCArray(4) = "<cp>"
    existingLOCArray(5) = "</cp>"
    existingLOCArray(6) = "<toc>"
    existingLOCArray(7) = "</toc>"
    existingLOCArray(8) = "<tp>"
    existingLOCArray(9) = "</tp>"
    'existingLOCArray(10) = "<ch[0-9]{1,}>"
    
    For c = 1 To UBound(existingLOCArray())
    With activeRng.Find
      .ClearFormatting
      .Text = existingLOCArray(c)
      .Wrap = wdFindContinue
      .MatchWildcards = False
    End With
    If activeRng.Find.Execute Then
        foundLOC = True
        foundLOCitem = existingLOCArray(c)
        Exit For
    End If
    Next
    
    'doing it again with wildcards=True, to catch numbered chapters
    With activeRng.Find
      .ClearFormatting
      .Text = "<ch[0-9]{1,}>"
      .Wrap = wdFindContinue
      .MatchWildcards = True
    End With
    If activeRng.Find.Execute Then
        foundLOC = True
        foundLOCitem = "(chapter heading tag, e.g. <ch1>, <ch2>, ... )"
    End If

    If foundLOC = True Then
        MsgBox "Your document: '" & mainDoc & "' already contains at least one Library of Congress tag:" & vbNewLine & vbNewLine & foundLOCitem & vbNewLine & vbNewLine & _
        "This macro may have already been run on this document. To run this macro, you MUST find and remove all existing LOC tags first.", , "Alert"
        zz_errorChecksB = True
        Exit Function
    End If

End Function

Private Function zz_TagReport()
    
    Set activeRng = activeDoc.Range
    
    'count occurences of all but Chapter Heads
    Dim MyDoc As String, txt As String, t As String
    Dim LOCtagArray(9) As String
    Dim LOCtagCount(9) As Integer
    Dim d As Long
    MyDoc = activeDoc.Range.Text
    
    LOCtagArray(1) = "<tp>"
    LOCtagArray(2) = "</tp>"
    LOCtagArray(3) = "<cp>"
    LOCtagArray(4) = "</cp>"
    LOCtagArray(5) = "<sp>"
    LOCtagArray(6) = "</sp>"
    LOCtagArray(7) = "<toc>"
    LOCtagArray(8) = "</toc>"
    LOCtagArray(9) = "</ch>"
    
    For d = 1 To UBound(LOCtagArray())
        txt = LOCtagArray(d)
        t = Replace(MyDoc, txt, "")
        LOCtagCount(d) = ((Len(MyDoc) - Len(t)) / Len(txt))
    Next
    
    Call zz_clearFindB
    
    Dim chTagCount As Long
    
    'Count occurences of Chapter Heads
    With activeRng.Find
      .ClearFormatting
      .Text = "<ch[0-9]{1,}>"
      .MatchWildcards = True
    Do While .Execute(Forward:=True) = True
    chTagCount = chTagCount + 1
    Loop
    End With
    
    Call zz_clearFindB
    
    'Check if there are ANY tags; if not, styles not used so don't continue.
    If LOCtagCount(1) = 0 And _
        LOCtagCount(2) = 0 And _
        LOCtagCount(3) = 0 And _
        LOCtagCount(4) = 0 And _
        LOCtagCount(5) = 0 And _
        LOCtagCount(6) = 0 And _
        LOCtagCount(7) = 0 And _
        LOCtagCount(8) = 0 And _
        LOCtagCount(9) = 0 And _
        chTagCount = 0 Then
            zz_TagReport = False
            Exit Function
    Else
        zz_TagReport = True
    End If
        
    'Prepare error message
    Dim errorList As String
    errorList = ""
    If LOCtagCount(1) = 0 And LOCtagCount(2) = 0 Then errorList = errorList & "ERROR: No Title Page tags found. Title page tags are REQUIRED for LOC submission." & vbNewLine
    If LOCtagCount(3) = 0 And LOCtagCount(4) = 0 Then errorList = errorList & "ERROR: No Copyright Page tags found. Copyright page tags are REQUIRED for LOC submission." & vbNewLine
    If LOCtagCount(1) > 1 Or LOCtagCount(1) <> LOCtagCount(2) Then errorList = errorList & "ERROR: Problem with Title Page tags: either too many were found or one is missing" & vbNewLine
    If LOCtagCount(3) > 1 Or LOCtagCount(3) <> LOCtagCount(4) Then errorList = errorList & "ERROR: Problem with Copyright Page tags: either too many were found or one is missing" & vbNewLine
    If LOCtagCount(5) > 1 Or LOCtagCount(5) <> LOCtagCount(6) Then errorList = errorList & "ERROR: Problem with Series Page tags: either too many were found or one is missing" & vbNewLine
    If LOCtagCount(7) > 1 Or LOCtagCount(7) <> LOCtagCount(8) Then errorList = errorList & "ERROR: Problem with Table of Contents tags: either too many were found or one is missing" & vbNewLine
    If chTagCount = 0 Then errorList = errorList & "WARNING: No Chapter Heading tags were found." & vbNewLine
    If LOCtagCount(9) = 0 Then errorList = errorList & "WARNING: No 'End of Last Chapter' tag was found." & vbNewLine
    
    'Create full message text
    Dim strTagReportText As String

    If errorList = "" Then
        strTagReportText = strTagReportText & "Congratulations!" & vbNewLine
        strTagReportText = strTagReportText & "LOC Tags look good for " & activeDoc.Name & vbNewLine
        strTagReportText = strTagReportText & "See summary below:" & vbNewLine
        strTagReportText = strTagReportText & vbNewLine
    Else
        strTagReportText = strTagReportText & "BAD NEWS:" & vbNewLine
        strTagReportText = strTagReportText & vbNewLine
        strTagReportText = strTagReportText & "Problems were found with LOC tags in your document '" & activeDoc.Name & "':" & vbNewLine
        strTagReportText = strTagReportText & vbNewLine
        strTagReportText = strTagReportText & vbNewLine
        strTagReportText = strTagReportText & "------------------------- ERRORS -------------------------" & vbNewLine
        strTagReportText = strTagReportText & errorList
        strTagReportText = strTagReportText & vbNewLine
        strTagReportText = strTagReportText & vbNewLine
    End If
        strTagReportText = strTagReportText & "------------------------- Tag Summary -------------------------" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(1) & "  Title page open tag(s) found <tp>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(2) & "  Title page close tag(s) found </tp>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(3) & "  Copyright page open tag(s) found <cp>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(4) & "  Copyright page close tag(s) found </cp>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(5) & "  Series page open tag(s) found <sp>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(6) & "  Series page close tag(s) found </sp>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(7) & "  Table of Contents open tag(s) found <toc>" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(8) & "  Table of Contents close tag(s) found </toc>" & vbNewLine
        strTagReportText = strTagReportText & chTagCount & "  Chapter beginning tag(s) found (<ch1>, <ch2>, etc)" & vbNewLine
        strTagReportText = strTagReportText & LOCtagCount(9) & "  End of last chapter tag(s) found </ch>" & vbNewLine
        
    ' Print to text file
    Call MacroHelpers.CreateTextFile(strText:=strTagReportText, suffix:="CIPtagReport")

End Function

