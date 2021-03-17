Attribute VB_Name = "TagUnstyledParas"
Option Private Module
Option Explicit

Sub TagUnstyledText()
  
'  #If Not Mac Then
'  Dim objTagProgress As Progress_Bar
'
'  Set objTagProgress = New Progress_Bar
'
'  objTagProgress.Title = "Tag Unstyled Text Macro"
'  DebugPrint "Tag Unstyled Paragraphs Macro"
'  #End If

  
' ======= Run startup checks ========
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
  
Call tagText(activeDoc)
'  #If Not Mac Then
' --------Progress Bar---------------------------------------------------------
' Percent complete and status for progress bar (PC) and status bar (Mac)
'  Dim sglPercentComplete As Single
'  Dim strStatus As String
'  Dim sglTotalCharStylesPercent As Single
'  Dim sglStartingPercent As Single
'
'  First status shown will be randomly pulled from array, for funzies
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
' Rnd returns random number between (0,1], rest of expression is to return an
' integer (1,10)
'  Randomize           'Sets seed for Rnd below to value of system timer
'  x = Int(UBound(funArray()) * Rnd()) + 1
'
' first number is percent of THIS macro completed
'  sglStartingPercent = 0.03
'  strStatus = funArray(x)
'  sglTotalCharStylesPercent = 0.47
'
' Calls ProgressBar.Increment mathod and waits for it to complete
'  Call ClassHelpers.UpdateBarAndWait(Bar:=objTagProgress, _
'    Status:=strStatus, Percent:=sglStartingPercent)
'
'  Call CharacterStyles.ActualCharStyles(oProgressChar:= _
'    objTagProgress, StartPercent:=sglStartingPercent, TotalPercent:=sglTotalCharStylesPercent, _
'    Status:=strStatus)
'
' #End If
End Sub
Sub tagText(myDoc As Document)
' Make sure we're always working with the right document
  Dim thisDoc As Document
  Set thisDoc = myDoc

  ' Rename built-in style that has parens
  thisDoc.styles("Normal (Web)").NameLocal = "_"

  Dim lngParaCount As Long
  Dim A As Long
  Dim strCurrentStyle As String
  Dim strTX As String
  Dim strTX1 As String
  Dim strNewStyle As String
  Dim strParaStatus As String

  Dim sglTotalPercent As Single
  Dim strNextStyle As String
  Dim strNextNextStyle As String
  Dim strCOTX1 As String


' Making these variables so we don't get any input errors with the style names t/o
  strTX = "Body-Text (Tx)"
  'strTX = "Body-Text (Tx)"
  strTX1 = "Text - Std No-Indent (tx1)"
  strCOTX1 = "Chap Opening Text No-Indent (cotx1)"

  lngParaCount = thisDoc.Paragraphs.Count

  Dim myStyle As Style ' For error handlers
'  Dim sglTotalPercentSoFar As Single
'  sglTotalPercentSoFar = sglStartingPercent + sglTotalCharStylesPercent
  
'  Dim sglTotalPercentRemaining As Single
'  sglTotalPercentRemaining = 1 - sglTotalPercentSoFar

' Loop through all paras, tag any w/o close parens as TX or TX1
' (or COTX1 if following chap opener)
  For A = 1 To lngParaCount
'    If a Mod 10 = 0 Then
'      ' Increment progress bar
'      sglPercentComplete = (((a / lngParaCount) * sglTotalPercentRemaining) + sglTotalPercentSoFar)
'      strParaStatus = "* Tagging paragraphs with Macmillan styles: " & a & " of " & lngParaCount & vbNewLine & strStatus
'      Call ClassHelpers.UpdateBarAndWait(Bar:=objTagProgress, Status:=strParaStatus, _
'          Percent:=sglPercentComplete)
'    End If
    
    strCurrentStyle = thisDoc.Paragraphs(A).Style
    
'    DebugPrint strCurrentStyle

  ' tag all non-Macmillan-style paragraphs with standard Macmillan styles
  ' Macmillan styles all end in close parens
    If Right(strCurrentStyle, 1) <> ")" Then                    ' indicates non-Macmillan style
        If thisDoc.Paragraphs(A).Range.Tables.Count = 0 Then    ' skipping table paragraphs
            ' If flush left, make No-Indent
            'If thisDoc.Paragraphs(A).FirstLineIndent = 0 Then
                'strNewStyle = strTX1
            'Else
                strNewStyle = strTX
            'End If
  
    ' Change the style of the paragraph in question
    ' This is where it will error if no style present
        thisDoc.Paragraphs(A).Style = strNewStyle
'    ElseIf a < lngParaCount Then ' it is already a Macmillan style
'    ' but can't check next para if it's the last para
'
'    ' is it a chap head?
'      If InStr(strCurrentStyle, "(cn)") > 0 Or _
'        InStr(strCurrentStyle, "(ct)") > 0 Or _
'        InStr(strCurrentStyle, "(ctnp)") > 0 Then
'
'
'            strNextStyle = thisDoc.Paragraphs(a + 1).Style
'
'
'      ' is the next para non-Macmillan (and thus should be COTX1)
'        If Right(strNextStyle, 1) <> ")" Then     ' it's not a Macmillan style
'        ' so it should be COTX1
'        ' Will error if this style not present in doc
'          strNewStyle = strCOTX1
'          thisDoc.Paragraphs(a + 1).Style = strNewStyle
'        Else ' it IS a Macmillan style too
'        ' it IT a chap opener? (can have CN followed by CT)
'          If InStr(strNextStyle, "(cn)") > 0 Or _
'            InStr(strNextStyle, "(ct)") > 0 Or _
'            InStr(strNextStyle, "(ctnp)") > 0 Then
'
'            strNextNextStyle = thisDoc.Paragraphs(a + 2).Style
'
'            If Right(strNextNextStyle, 1) <> ")" Then ' it's not Macmillan
'            ' so it should be COTX1
'              strNewStyle = strCOTX1
'              thisDoc.Paragraphs(a + 2).Style = strNewStyle
'            End If
'          End If
'        End If
      End If
    End If
  Next A

  ' Change Normal (Web) back
  thisDoc.styles("Normal (Web),_").NameLocal = "Normal (Web)"
  
'    Call MacroHelpers.Cleanup
'    Unload objTagProgress
    MsgBox "RSuite styles have been applied throughout your manuscript."
End Sub

