Attribute VB_Name = "CharacterStyles"
'' Created by Erica Warren -- erica.warren@macmillan.com
'' Split off from MacmillanCleanupMacro: https://github.com/macmillanpublishers/Word-template/blob/master/macmillan/CleanupMacro.bas
'
'' ======== PURPOSE ============================
'' Applies Macmillan character styles to direct-styled text in current
'' document
'
'' ======== DEPENDENCIES =======================
'' 1. Requires ProgressBar userform module
'' 2. Requires MacroHelpers module
'
'' Note: have already used all numerals and capital letters for tagging,
'' starting with lowercase letters. through a.
'
'Option Explicit
'Option Base 1
'
'Private Const strCharStyles As String = "CharacterStyles."
'Private activeRng As Range
'
'Public Sub MacmillanCharStylesSub()
'    'Call MacmillanCharStyles
'End Sub
'
'Public Function MacmillanCharStyles() As Dictionary
'
'  On Error GoTo MacmillanCharStylesError
'  Dim dictReturn As Dictionary
'  Set dictReturn = New Dictionary
'  dictReturn.Add "pass", False
'
'  Dim CharacterProgress As Progress_Bar
'  Set CharacterProgress = New Progress_Bar
'
'  CharacterProgress.Title = "Macmillan Character Styles Macro"
'  DebugPrint "Starting Character Styles macro"
'
'' ======= Run startup checks ========
'' True means a check failed (e.g., doc protection on)
'  If WT_Settings.InstallType = "user" Then
'    If MacroHelpers.StartupSettings(AcceptAll:=False) = True Then
'      Call MacroHelpers.Cleanup
'      Exit Function
'    End If
'  Else
'    If MacroHelpers.StartupSettings(AcceptAll:=True) = True Then
'      Call MacroHelpers.Cleanup
'      Exit Function
'    End If
'  End If
'
'' --------Progress Bar---------------------------------------------------------
'' Percent complete and status for progress bar (PC) and status bar (Mac)
'  Dim sglPercentComplete As Single
'  Dim strStatus As String
'
'  'First status shown will be randomly pulled from array, for funzies
'  Dim funArray() As String
'  ReDim funArray(1 To 10)      'Declare bounds of array here
'
'  funArray(1) = "* Mixing metaphors..."
'  funArray(2) = "* Arguing about the serial comma..."
'  funArray(3) = "* Un-mixing metaphors..."
'  funArray(4) = "* Avoiding the passive voice..."
'  funArray(5) = "* Ending sentences in prepositions..."
'  funArray(6) = "* Splitting infinitives..."
'  funArray(7) = "* Ooh, what an interesting manuscript..."
'  funArray(8) = "* Un-dangling modifiers..."
'  funArray(9) = "* Jazzing up author bio..."
'  funArray(10) = "* Filling in plot holes..."
'
'  Dim x As Integer
'
'' Rnd returns random number between (0,1], rest of expression is to return an
'' integer (1,10)
'  Randomize           'Sets seed for Rnd below to value of system timer
'  x = Int(UBound(funArray()) * Rnd()) + 1
'
'' first number is percent of THIS macro completed
'  sglPercentComplete = 0.09
'  strStatus = funArray(x)
'
'' Calls ProgressBar.Increment mathod and waits for it to complete
'  Call ClassHelpers.UpdateBarAndWait(Bar:=CharacterProgress, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  Call CharacterStyles.ActualCharStyles(oProgressChar:= _
'    CharacterProgress, StartPercent:=sglPercentComplete, TotalPercent:=1, Status:=strStatus)
'
'  dictReturn.Item("pass") = True
'  Set MacmillanCharStyles = dictReturn
'  Exit Function
'
'MacmillanCharStylesError:
'  Err.Source = strCharStyles & "MacmillanCharStyles"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Function
'
'
'Sub ActualCharStyles(ByVal oProgressChar As Progress_Bar, StartPercent As Single, TotalPercent As Single, Status As String)
'' Have to pass the ProgressBar so this can be run from within another macro
'' StartPercent is the percentage the progress bar is at when this sub starts
'' TotalPercent is the total percent of the progress bar that this sub will cover
'
'  'On Error GoTo ActualCharStylesError
''------------------Time Start-----------------
''Dim StartTime As Double
''Dim SecondsElapsed As Double
'
''Remember time when macro starts
''StartTime = Timer
'
'' ------------check for endnotes and footnotes---------------------------------
'  Dim colStories As Collection
'  Set colStories = MacroHelpers.ActiveStories
'  Dim varStory As Variant
'  Dim currentStory As WdStoryType
'
'' -----------Delete hidden text ------------------------------------------------
'
'  For Each varStory In colStories
'    currentStory = varStory
'  ' Note, if you don't delete hidden text, this macro turns it into reg. text.
'    If MacroHelpers.HiddenTextSucks(StoryType:=currentStory) = True Then
'    ' Notify user maybe?
'    End If
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'' -------------- Clear formatting from paragraph marks ------------------------
'' can cause errors
'
'  For Each varStory In colStories
'    currentStory = varStory
'    Call MacroHelpers.ClearPilcrowFormat(StoryType:=currentStory)
'  Next
'
'' -------------- Clean up page break characters -------------------------------
'  Call MacroHelpers.PageBreakCleanup
'
'' ===================== Replace Local Styles Start ============================
'
'' -----------------------Tag space break styles--------------------------------
'  Call MacroHelpers.zz_clearFind
'  Dim sglPercentComplete As Single
'  Dim strStatus As String
'  strStatus = Status
'
'
'  sglPercentComplete = (0.18 * TotalPercent) + StartPercent
'  strStatus = "* Preserving styled whitespace..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  For Each varStory In colStories
'    currentStory = varStory
'    Call PreserveWhiteSpaceinBrkStylesA(StoryType:=currentStory)
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'' ----------------------------Fix hyperlinks-----------------------------------
'  sglPercentComplete = (0.28 * TotalPercent) + StartPercent
'  strStatus = "* Applying styles to hyperlinks..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  Call MacroHelpers.StyleAllHyperlinks(StoriesInUse:=colStories)
'
'  Call MacroHelpers.zz_clearFind
'
'' --------------------------Remove unstyled space breaks-----------------------
'  sglPercentComplete = (0.39 * TotalPercent) + StartPercent
'  strStatus = "* Removing unstyled breaks..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  For Each varStory In colStories
'    currentStory = varStory
'    Call RemoveBreaks(StoryType:=currentStory)
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'' --------------------------Tag existing character styles----------------------
'  sglPercentComplete = (0.52 * TotalPercent) + StartPercent
'  strStatus = "* Tagging character styles..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  For Each varStory In colStories
'    currentStory = varStory
'    Call TagExistingCharStyles(StoryType:=currentStory)
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'' -------------------------Tag direct formatting-------------------------------
'  sglPercentComplete = (0.65 * TotalPercent) + StartPercent
'  strStatus = "* Tagging direct formatting..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  ' allBkmkrStyles is a jagged array (array of arrays) to hold in-use Bookmaker styles.
'  ' i.e., one array for each story. Must be Variant.
'  Dim allBkmkrStyles() As Variant
'  Dim S As Long
'  S = 0
'  For Each varStory In colStories
'    currentStory = varStory
'    S = S + 1
'  'tag local styling, reset local styling, remove text highlights
'    Call LocalStyleTag(StoryType:=currentStory)
'
'    ReDim Preserve allBkmkrStyles(1 To S)
'    allBkmkrStyles(S) = TagBkmkrCharStyles(StoryType:=currentStory)
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'' ----------------------------Apply Macmillan character styles to tagged text--
'  sglPercentComplete = (0.81 * TotalPercent) + StartPercent
'  strStatus = "* Applying Macmillan character styles..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  For Each varStory In colStories
'    currentStory = varStory
'    Call LocalStyleReplace(StoryType:=currentStory, BkmkrStyles:=allBkmkrStyles(S))
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'' ---------------------------Remove tags from styled space breaks--------------
'  sglPercentComplete = (0.95 * TotalPercent) + StartPercent
'  strStatus = "* Cleaning up styled whitespace..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'  For Each varStory In colStories
'    currentStory = varStory
'    Call PreserveWhiteSpaceinBrkStylesB(StoryType:=currentStory)
'  Next
'
'  Call MacroHelpers.zz_clearFind
'
'
'
'' ---------------------------Return settings to original-----------------------
'  sglPercentComplete = TotalPercent + StartPercent
'  strStatus = "* Finishing up..." & vbCr & strStatus
'
'  Call ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
'    Status:=strStatus, Percent:=sglPercentComplete)
'
'' If this is the whole macro, close out; otherwise calling macro will close it all down
'  If TotalPercent = 1 Then
'    Call MacroHelpers.Cleanup
'    Unload oProgressChar
''        MsgBox "Macmillan character styles have been applied throughout your manuscript."
'  End If
'
'
'' ----------------------Timer End-------------------------------------------
'' Determine how many seconds code took to run
'' SecondsElapsed = Round(Timer - StartTime, 2)
'
'' Notify user in seconds
''  DebugPrint "This code ran successfully in " & SecondsElapsed & " seconds"
'  Exit Sub
'
'ActualCharStylesError:
'  Err.Source = strCharStyles & "ActualCharStyles"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Sub
'
'
'Private Sub PreserveWhiteSpaceinBrkStylesA(StoryType As WdStoryType)
' On Error GoTo PreserveWhiteSpaceinBrkStylesAError:
'  Set activeRng = activeDoc.StoryRanges(StoryType)
'
'' Find/Replace (which we'll use later on) will not replace a paragraph mark
'' in the first or last paragraph, so add dummy paragraphs here (with tags)
'' that we can remove later on.
'
''  activeRng.InsertBefore "|||" & vbNewLine
''  activeRng.Paragraphs.First.Style = "Normal"
'  activeRng.InsertAfter vbNewLine
'  activeRng.Paragraphs.Last.Style = "Normal"
'
'' tag paragraphs allowed to be blank
'  Dim varStyle As Variant
'  For Each varStyle In WT_StyleConfig.AllowedBlankStyles
'  ' Only search if style is in doc (so can run on non-styled manuscripts)
'  ' If found, add an optional hyphen (^31) before the final newline (^13).
'  ' The ^31 will prevent matching as multiple newlines when we remove those,
'  ' but it's still whitespace so if something goes awry, we don't leave visible
'  ' tags in the doc and freak people out.
'
'    If MacroHelpers.IsStyleInUse(CStr(varStyle)) = True Then
'      MacroHelpers.zz_clearFind
'      Selection.HomeKey unit:=wdStory
'      With Selection.Find
'        .Text = "(*)(^13)"
'        .Replacement.Text = "\1```\2"    ' add optional hyphen before trailing newline
'        .Format = True
'        .Style = CStr(varStyle)
'        .Replacement.Style = CStr(varStyle)
'        .MatchWildcards = True
'        .Execute replace:=wdReplaceAll
'      End With
'    End If
'NextLoop:
'  Next varStyle
'  Exit Sub
'
'PreserveWhiteSpaceinBrkStylesAError:
'  ' skips tagging that style if it's missing from doc; if missing, obv nothing has that style
'  'DebugPrint StylePreserveArray(e)
'  '5834 "Item with specified name does not exist" i.e. style not present in doc
'  '5941 item not available in collection
'  If Err.Number = 5834 Or Err.Number = 5941 Then
'      Resume NextLoop:
'  End If
'
'  Err.Source = strCharStyles & "PreserveWhiteSpaceinBrkStylesA"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Sub
'
'Private Sub RemoveBreaks(StoryType As WdStoryType)
'  On Error GoTo RemoveBreaksError
'  Set activeRng = activeDoc.StoryRanges(StoryType)
'
'  Call MacroHelpers.zz_clearFind
'    With activeRng.Find
'      .Text = "^13{2,}"
'      .Replacement.Text = "^p"
'      .Wrap = wdFindContinue
'      .MatchWildcards = True
'      .Execute replace:=wdReplaceAll
'    End With
'  Exit Sub
'
'RemoveBreaksError:
'  Err.Source = strCharStyles & "RemoveBreaks"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Sub
'
'Private Sub PreserveWhiteSpaceinBrkStylesB(StoryType As WdStoryType)
'  On Error GoTo PreserveWhiteSpaceinBrkStylesBError
'
'  Set activeRng = activeDoc.StoryRanges(StoryType)
'' Remove those optional hyphens we added earlier, also all optional hyphens
'  Call MacroHelpers.zz_clearFind
'  With activeRng.Find
'    .Text = "```"
'    .Replacement.Text = ""
'    .Wrap = wdFindContinue
'    .MatchWildcards = False
'    .Execute replace:=wdReplaceAll
'  End With
'
'' We also want to remove last para if it only contains a blank para, of any style
'' Loop until we find a paragraph with text.
'  Dim rngLast As Range
'  Dim lngCount As Long
'
'  lngCount = 0
'  Do
'    ' counter to prevent runaway loops
'    lngCount = lngCount + 1
'    Set rngLast = activeDoc.Paragraphs.Last.Range
'    If MacroHelpers.IsNewLine(rngLast.Text) = True Then
'      rngLast.Delete
'    Else
'      Exit Do
'    End If
'  Loop Until lngCount = 20
'  Exit Sub
'
'PreserveWhiteSpaceinBrkStylesBError:
'  Err.Source = strCharStyles & "PreserveWhiteSpaceinBrkStylesB"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Sub
'
'Private Sub TagExistingCharStyles(StoryType As WdStoryType)
'  On Error GoTo TagExistingCharStylesError
'  Set activeRng = activeDoc.StoryRanges(StoryType)                        'this whole sub (except last stanza) is basically a v. 3.1 patch.  correspondingly updated sub name, call in main, and replacements go along with bold and common replacements
'
'  Dim tagCharStylesArray(16) As String                                   ' number of items in array should be declared here
'  Dim CharStylePreserveArray(16) As String              ' number of items in array should be declared here
'  Dim d As Long
'
'  CharStylePreserveArray(1) = "span hyperlink (url)"
'  CharStylePreserveArray(2) = "span symbols (sym)"
'  CharStylePreserveArray(3) = "span accent characters (acc)"
'  CharStylePreserveArray(4) = "span cross-reference (xref)"
'  CharStylePreserveArray(5) = "span material to come (tk)"
'  CharStylePreserveArray(6) = "span carry query (cq)"
'  CharStylePreserveArray(7) = "span key phrase (kp)"
'  CharStylePreserveArray(8) = "span preserve characters (pre)"  'added v. 3.2
'  CharStylePreserveArray(9) = "span ISBN (isbn)"  'added v. 3.7
'  CharStylePreserveArray(10) = "span symbols ital (symi)"     'added v. 3.8
'  CharStylePreserveArray(11) = "span symbols bold (symb)"
'  CharStylePreserveArray(12) = "span run-in computer type (comp)"
'  CharStylePreserveArray(13) = "span alt font 1 (span1)"
'  CharStylePreserveArray(14) = "span alt font 2 (span2)"
'  CharStylePreserveArray(15) = "span illustration holder (illi)"
'  CharStylePreserveArray(16) = "span design note (dni)"
'
'
'  tagCharStylesArray(1) = "`H|^&|H`"
'  tagCharStylesArray(2) = "`Z|^&|Z`"
'  tagCharStylesArray(3) = "`Y|^&|Y`"
'  tagCharStylesArray(4) = "`X|^&|X`"
'  tagCharStylesArray(5) = "`W|^&|W`"
'  tagCharStylesArray(6) = "`V|^&|V`"
'  tagCharStylesArray(7) = "`T|^&|T`"
'  tagCharStylesArray(8) = "`F|^&|F`"
'  tagCharStylesArray(9) = "`Q|^&|Q`"
'  tagCharStylesArray(10) = "`E|^&|E`"
'  tagCharStylesArray(11) = "`G|^&|G`"
'  tagCharStylesArray(12) = "`J|^&|J`"
'  tagCharStylesArray(13) = "`CC|^&|CC`"
'  tagCharStylesArray(14) = "`DD|^&|DD`"
'  tagCharStylesArray(15) = "`EE|^&|EE`"
'  tagCharStylesArray(16) = "`FF|^&|FF`"
'
'  Call MacroHelpers.zz_clearFind
'  For d = 1 To UBound(CharStylePreserveArray())
'    With activeRng.Find
'      .Replacement.Text = tagCharStylesArray(d)
'      .Wrap = wdFindContinue
'      .Format = True
'      .Style = CharStylePreserveArray(d)
'      .MatchWildcards = True
'      .Execute replace:=wdReplaceAll
'    End With
'NextLoop:
'  Next
'  Exit Sub
'
'TagExistingCharStylesError:
'' skips tagging that style if it's missing from doc;
'' if missing, obv nothing has that style
'
'' 5834 "Item with specified name does not exist" i.e. style not present in doc
'' 5941 item is not present in collection
'  If Err.Number = 5834 Or Err.Number = 5941 Then
'    Resume NextLoop
'  End If
'
'  Err.Source = strCharStyles & "TagExistingCharStyles"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Sub
'
'Private Sub LocalStyleTag(StoryType As WdStoryType)
'  On Error GoTo LocalStyleTagError
'
'  Set activeRng = activeDoc.StoryRanges(StoryType)
'
'' ------------tag key styles---------------------------------------------------
'  Dim tagStyleFindArray(11) As Boolean
'  Dim tagStyleReplaceArray(11) As String
'  Dim G As Long
'
'  tagStyleFindArray(1) = False        'Bold
'  tagStyleFindArray(2) = False        'Italic
'  tagStyleFindArray(3) = False        'Underline
'  tagStyleFindArray(4) = False        'Smallcaps
'  tagStyleFindArray(5) = False        'Subscript
'  tagStyleFindArray(6) = False        'Superscript
'  tagStyleFindArray(7) = False        'Highlights
'  ' note 8 - 10 are below
'  tagStyleFindArray(11) = False       'Strikethrough
'
'  tagStyleReplaceArray(1) = "`B|^&|B`"
'  tagStyleReplaceArray(2) = "`I|^&|I`"
'  tagStyleReplaceArray(3) = "`U|^&|U`"
'  tagStyleReplaceArray(4) = "`M|^&|M`"
'  tagStyleReplaceArray(5) = "`S|^&|S`"
'  tagStyleReplaceArray(6) = "`P|^&|P`"
'  tagStyleReplaceArray(8) = "`A|^&|A`"
'  tagStyleReplaceArray(9) = "`C|^&|C`"
'  tagStyleReplaceArray(10) = "`D|^&|D`"
'  tagStyleReplaceArray(11) = "`a|^&|a`"
'
'  Call MacroHelpers.zz_clearFind
'  For G = 1 To UBound(tagStyleFindArray())
'
'    tagStyleFindArray(G) = True
'
'    If tagStyleFindArray(8) = True Then
'      tagStyleFindArray(1) = True
'      tagStyleFindArray(2) = True     'bold and italic
'    End If
'
'    If tagStyleFindArray(9) = True Then
'      tagStyleFindArray(1) = True
'      tagStyleFindArray(4) = True
'      tagStyleFindArray(2) = False  'bold and smallcaps
'    End If
'
'    If tagStyleFindArray(10) = True Then
'      tagStyleFindArray(2) = True
'      tagStyleFindArray(4) = True
'      tagStyleFindArray(1) = False 'smallcaps and italic
'    End If
'
'    If tagStyleFindArray(11) = True Then
'      tagStyleFindArray(2) = False
'      tagStyleFindArray(4) = False ' reset tags for strikethrough
'    End If
'    With activeRng.Find
'      .Replacement.Text = tagStyleReplaceArray(G)
'      .Wrap = wdFindContinue
'      .Format = True
'      .Font.Bold = tagStyleFindArray(1)
'      .Font.Italic = tagStyleFindArray(2)
'      .Font.Underline = tagStyleFindArray(3)
'      .Font.SmallCaps = tagStyleFindArray(4)
'      .Font.Subscript = tagStyleFindArray(5)
'      .Font.Superscript = tagStyleFindArray(6)
'      .Highlight = tagStyleFindArray(7)
'      .Font.StrikeThrough = tagStyleFindArray(11)
'      .Replacement.Highlight = False
'      .MatchWildcards = True
'      .Execute replace:=wdReplaceAll
'    End With
'
'    tagStyleFindArray(G) = False
'
'  Next
'  Exit Sub
'
'LocalStyleTagError:
'  Err.Source = strCharStyles & "LocalStyleTag"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Sub
'
'Private Sub LocalStyleReplace(StoryType As WdStoryType, BkmkrStyles As Variant)
'  On Error GoTo LocalStyleReplaceError
'
'  Set activeRng = activeDoc.StoryRanges(StoryType)
'
'  ' Determine if we need to do the bookmaker styles thing
'  ' BkmkrStyles is an array of bookmaker character styles in use. If it's empty,
'  ' there are none in use so we don't have to check
'
'  Dim blnCheckBkmkr As Boolean
'
'  If IsArrayEmpty(BkmkrStyles) = False Then
'      blnCheckBkmkr = True
'  Else
'      blnCheckBkmkr = False
'  End If
'
'  '-------------apply styles to tags
'  'number of items in array should = styles in LocalStyleTag + styles in TagExistingCharStyles
'  Dim tagFindArray(1 To 28) As String              ' number of items in array should be declared here
'  Dim tagReplaceArray(1 To 28) As String         'and here
'  Dim H As Long
'
'  tagFindArray(1) = "`B|(*)|B`"
'  tagFindArray(2) = "`I|(*)|I`"
'  tagFindArray(3) = "`U|(*)|U`"
'  tagFindArray(4) = "`M|(*)|M`"
'  tagFindArray(5) = "`H|(*)|H`"
'  tagFindArray(6) = "`S|(*)|S`"
'  tagFindArray(7) = "`P|(*)|P`"
'  tagFindArray(8) = "`Z|(*)|Z`"
'  tagFindArray(9) = "`Y|(*)|Y`"
'  tagFindArray(10) = "`X|(*)|X`"
'  tagFindArray(11) = "`W|(*)|W`"
'  tagFindArray(12) = "`V|(*)|V`"
'  tagFindArray(13) = "`T|(*)|T`"
'  tagFindArray(14) = "`A|(*)|A`"                'v. 3.1 patch
'  tagFindArray(15) = "`C|(*)|C`"                 'v. 3.1 patch
'  tagFindArray(16) = "`D|(*)|D`"                       'v. 3.1 patch
'  tagFindArray(17) = "`F|(*)|F`"
'  tagFindArray(18) = "`K|(*)|K`"          'v. 3.7 added
'  tagFindArray(19) = "`Q|(*)|Q`"          'v. 3.7 added
'  tagFindArray(20) = "`E|(*)|E`"
'  tagFindArray(21) = "`G|(*)|G`"          'v. 3.8 added
'  tagFindArray(22) = "`J|(*)|J`"
'  tagFindArray(23) = "`AA|(*)|AA`"
'  tagFindArray(24) = "`BB|(*)|BB`"
'  tagFindArray(25) = "`CC|(*)|CC`"
'  tagFindArray(26) = "`DD|(*)|DD`"
'  tagFindArray(27) = "`EE|(*)|EE`"
'  tagFindArray(28) = "`FF|(*)|FF`"
'
'  tagReplaceArray(1) = "span boldface characters (bf)"
'  tagReplaceArray(2) = "span italic characters (ital)"
'  tagReplaceArray(3) = "span underscore characters (us)"
'  tagReplaceArray(4) = "span small caps characters (sc)"
'  tagReplaceArray(5) = "span hyperlink (url)"
'  tagReplaceArray(6) = "span subscript characters (sub)"
'  tagReplaceArray(7) = "span superscript characters (sup)"
'  tagReplaceArray(8) = "span symbols (sym)"
'  tagReplaceArray(9) = "span accent characters (acc)"
'  tagReplaceArray(10) = "span cross-reference (xref)"
'  tagReplaceArray(11) = "span material to come (tk)"
'  tagReplaceArray(12) = "span carry query (cq)"
'  tagReplaceArray(13) = "span key phrase (kp)"
'  tagReplaceArray(14) = "span bold ital (bem)"
'  tagReplaceArray(15) = "span smcap bold (scbold)"
'  tagReplaceArray(16) = "span smcap ital (scital)"
'  tagReplaceArray(17) = "span preserve characters (pre)"
'  tagReplaceArray(18) = "bookmaker keep together (kt)"            'v. 3.7 added
'  tagReplaceArray(19) = "span ISBN (isbn)"                        'v. 3.7 added
'  tagReplaceArray(20) = "span symbols ital (symi)"                ' v. 3.8 added
'  tagReplaceArray(21) = "span symbols bold (symb)"                ' v. 3.8 added
'  tagReplaceArray(22) = "span run-in computer type (comp)"
'  tagReplaceArray(23) = "span strikethrough characters (str)"
'  tagReplaceArray(24) = "span smcap bold ital (scbi)"
'  tagReplaceArray(25) = "span alt font 1 (span1)"
'  tagReplaceArray(26) = "span alt font 2 (span2)"
'  tagReplaceArray(27) = "span illustration holder (illi)"
'  tagReplaceArray(28) = "span design note (dni)"
'
'  For H = LBound(tagFindArray()) To UBound(tagFindArray())
'
'  ' ----------- bookmaker char styles ----------------------
'    ' tag bookmaker line-ending character styles and
'    ' adjust name if have additional styles applied
'    ' because if you append "tighten" or "loosen" to
'    ' regular style name, Bookmaker does that.
'    If blnCheckBkmkr = True Then
'      Dim Q As Long
'      Dim qCount As Long
'      Dim strAction As String
'      Dim strNewName As String
'      Dim strTag As String
'
'      ' deal with bookmaker styles
'      For Q = LBound(BkmkrStyles) To UBound(BkmkrStyles)
'        ' replace bookmaker-tagged text with bookmaker styles
'        strTag = "bk" & Format(Q, "0000")
'        With activeRng.Find
'          .ClearFormatting
'          .Replacement.ClearFormatting
'          .Text = "`" & strTag & "|(*)|" & strTag & "`"
'          .Replacement.Text = "\1"
'          .Wrap = wdFindContinue
'          .Format = True
'          .Replacement.Style = BkmkrStyles(Q)
'          .matchCase = False
'          .MatchWholeWord = False
'          .MatchWildcards = True
'          .MatchSoundsLike = False
'          .MatchAllWordForms = False
'          .Execute replace:=wdReplaceAll
'        End With
'
'        'Move selection to start of document
'        Selection.HomeKey unit:=wdStory
'
'
'        qCount = 0
'        With Selection.Find
'          .ClearFormatting
'          .Replacement.ClearFormatting
'          .Text = tagFindArray(H)
'          .Replacement.Text = "\1"
'          .Wrap = wdFindStop
'          .Forward = True
'          .Style = BkmkrStyles(Q)
'          .Replacement.Style = tagReplaceArray(H)
'          .Format = True
'          .matchCase = False
'          .MatchWholeWord = False
'          .MatchWildcards = True
'          .MatchSoundsLike = False
'          .MatchAllWordForms = False
'
'          Do While .Execute = True And qCount < 200
'            qCount = qCount + 1
'            .Execute replace:=wdReplaceOne
'            ' pull just action to add to style name
'            ' always starts w/ "bookmaker ", but we want to include the space,
'            ' hence start at 10
'            strAction = Mid(BkmkrStyles(Q), 10, InStr(BkmkrStyles(Q), "(") - 11)
'            strNewName = tagReplaceArray(H) & strAction
'
'            ' Note these hybrid styles aren't in std template, so if they
'            ' haven't been created in this doc yet, will error.
'            Selection.Style = strNewName
'          Loop
'        End With
'      Next Q
'    End If
'
'On Error GoTo ErrorHandler
'    ' tag the rest of the character styles
'    With activeRng.Find
'        .ClearFormatting
'        .Replacement.ClearFormatting
'        .Text = tagFindArray(H)
'        .Replacement.Text = "\1"
'        .Wrap = wdFindContinue
'        .Format = True
'        .Replacement.Style = tagReplaceArray(H)
'        .matchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = True
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'        .Execute replace:=wdReplaceAll
'    End With
'
'NextLoop:
'  Next
'  Exit Sub
'
'' TO DO: Move this to a more universal error handler for missing styles
'' in main ErrorChecker function
'
'ErrorHandler:
'
'  Dim MyStyle As Style
'
'  If Err.Number = 5834 Or Err.Number = 5941 Then
'    Select Case tagReplaceArray(H)
'
'      'If style from LocalStyleTag is not present, add style
'      Case "span boldface characters (bf)":
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .Bold = True
'        End With
'        Resume
'
'      Case "span italic characters (ital)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .Italic = True
'        End With
'        Resume
'
'      Case "span underscore characters (us)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .Underline = wdUnderlineSingle
'        End With
'        Resume
'
'      Case "span small caps characters (sc)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .SmallCaps = False
'          .AllCaps = True
'          .size = 9
'        End With
'        Resume
'
'      Case "span subscript characters (sub)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .Subscript = True
'        End With
'        Resume
'
'      Case "span superscript characters (sup)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .Superscript = True
'        End With
'        Resume
'
'      Case "span bold ital (bem)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .Bold = True
'          .Italic = True
'        End With
'        Resume
'
'      Case "span smcap bold (scbold)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .SmallCaps = False
'          .AllCaps = True
'          .size = 9
'          .Bold = True
'        End With
'        Resume
'
'      Case "span smcap ital (scital)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
'          Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .SmallCaps = False
'          .AllCaps = True
'          .size = 9
'          .Italic = True
'        End With
'        Resume
'
'      Case "span strikethrough characters (str)"
'        Set MyStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), Type:=wdStyleTypeCharacter)
'        With MyStyle.Font
'          .Shading.BackgroundPatternColor = wdColorLightTurquoise
'          .StrikeThrough = True
'        End With
'        Resume
'
'    ' Else just skip if not from direct formatting
'      Case Else
'        Resume NextLoop:
'
'    End Select
'  End If
'
'  Exit Sub
'
'LocalStyleReplaceError:
''    DebugPrint Err.Number & ": " & Err.Description
''    DebugPrint "New name: " & strNewName
''    DebugPrint "Old name: " & tagReplaceArray(h)
'
'  Dim myStyle2 As Style
'
'  If Err.Number = 5834 Or Err.Number = 5941 Then
'
'    Set myStyle2 = activeDoc.Styles.Add(Name:=strNewName, _
'        Type:=wdStyleTypeCharacter)
'
'On Error GoTo ErrorHandler
'    ' If the original style did not exist yet, will error here
'    ' but ErrorHandler will add the style
'    myStyle2.BaseStyle = tagReplaceArray(H)
'    ' Then go back to BkmkrError so further errors will route
'    ' correctly
'On Error GoTo LocalStyleReplaceError
'    Resume
'  Else
'    Err.Source = strCharStyles & "LocalStyleReplace"
'    If ErrorChecker(Err) = False Then
'      Resume
'    Else
'      Call MacroHelpers.GlobalCleanup
'    End If
'  End If
'
'End Sub
'
'
'Private Function TagBkmkrCharStyles(StoryType As Variant) As Variant
'  On Error GoTo TagBkmkrCharStylesError
''    Set activeRng = activeDoc.Range
'  Set activeRng = activeDoc.StoryRanges(StoryType)
'
'' Will need to loop through stories as well
'' And be a function that returns an array
'
'  Dim objStyle As Style
'  Dim strBkmkrNames() As String
'  Dim z As Long
'
'' Loop through all styles to get array of bkmkr styles in document
'' NOTE! The .InUse property does NOT mean "in use in the document"; it means
'' "any custom style or any modified built-in style". Ugh. Anyway, now we
'' have to loop through all styles to see if bookmaker styles are present,
'' then search for each of those styles to see if they are in use.
'
'  For Each objStyle In activeDoc.Styles
'    ' If char style with "bookmaker" in name is in use...
'    ' binary compare is default, but adding here to be clear that we are doing
'    ' a CASE SENSITIVE search, because "Bookmaker" is only for Paragraph styles,
'    ' which we don't want to mess with.
'    If InStr(1, objStyle.NameLocal, "bookmaker", vbBinaryCompare) <> 0 And _
'      objStyle.Type = wdStyleTypeCharacter Then
''      DebugPrint StoryType & ": " & objStyle.NameLocal
'      Selection.HomeKey unit:=wdStory
'      ' Now see if it's being used ...
'      With Selection.Find
'        .ClearFormatting
'        .Text = ""
'        .Style = objStyle.NameLocal
'        .Wrap = wdFindContinue
'        .Format = True
'        .Forward = True
'        .Execute
'      End With
'
'      If Selection.Find.Found = True Then
'        '... add it to an array
'        z = z + 1
'        ReDim Preserve strBkmkrNames(1 To z)
'        strBkmkrNames(z) = objStyle.NameLocal
'      End If
'    End If
'  Next objStyle
'
'  If IsArrayEmpty(strBkmkrNames) = True Then
'    TagBkmkrCharStyles = strBkmkrNames
'    Exit Function
'  End If
'
'' Tag in-use bkmkr styles
'' Make sure if text also has formatting,
'' the tags do not have it...
'Dim x As Long
'  Dim strTag As String
'  Dim strAction As String
'  Dim lngCount As Long
'
'  For x = LBound(strBkmkrNames) To UBound(strBkmkrNames)
'    strTag = "bk" & Format(x, "0000")
''        DebugPrint strTag
'
'    With activeRng.Find
'      .ClearFormatting
'      .Replacement.ClearFormatting
'      .Text = ""
'      .Replacement.Text = "`" & strTag & "|^&|" & strTag & "`"
'      .Wrap = wdFindContinue
'      .Format = True
'      .Style = strBkmkrNames(x)
'      .matchCase = False
'      .MatchWholeWord = False
'      .MatchWildcards = True
'      .MatchSoundsLike = False
'      .MatchAllWordForms = False
'      .Execute replace:=wdReplaceAll
'    End With
'  Next
'
'  '-------------Reset everything -- clears all direct formatting!
'  activeRng.Font.Reset
'
'  ' return array of in-use bookmaker styles so we can tag later
'  TagBkmkrCharStyles = strBkmkrNames
'  Exit Function
'
'TagBkmkrCharStylesError:
'  Err.Source = strCharStyles & "TagBkmkrCharStyles"
'  If ErrorChecker(Err) = False Then
'    Resume
'  Else
'    Call MacroHelpers.GlobalCleanup
'  End If
'End Function
'
'
