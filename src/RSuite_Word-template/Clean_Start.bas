Attribute VB_Name = "Clean_Start"
Option Explicit

Sub LaunchCleanup()

    Call Clean_helpers.CheckTemplate

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

    CleanupOptions.Show
End Sub

Sub LaunchTagCharacterStyles()

    On Error GoTo ErrorHandler
    
    Dim StoryNo As Range
    Dim StoryName As Variant
    
    
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
    
    ' init progress bar
    Set pBar = New Progress_Bar
    pBarCounter = 0
    pBar.Caption = "RSuite Character Style"
    completeStatus = "Starting Character Style Replacement"
    pBar.Status.Caption = completeStatus
        
    
    For Each StoryNo In ActiveDocument.StoryRanges
        
        If StoryNo.StoryType < 4 Then
        
            MyStoryNo = StoryNo.StoryType
            
            Select Case MyStoryNo
                Case 1
                    StoryName = "Main Body"
                Case 2
                    StoryName = "Footnotes"
                Case 3
                    StoryName = "Endnotes"
            End Select
            
            completeStatus = completeStatus + vbNewLine + _
                            "=========================" + vbNewLine + _
                            "Cleaning " & StoryName + vbNewLine + _
                            "========================="
            Clean_helpers.updateStatus ("")
               
            ' Clean up characters!
            Call Clean.FixAppliedCharStyles(MyStoryNo)
            Call Clean.LocalFormatting(MyStoryNo)
            Call Clean.CheckSpecialCharactersPC(MyStoryNo)
            Call Clean.CheckAppliedCharStyles(MyStoryNo)
            
        End If
    Next
    
    Unload pBar
    
    Call Clean_helpers.MessageBox("Done", "Character Styles is complete!", vbOK)
    
    Exit Sub
    
ErrorHandler:
    
    If Err.Number = 5834 Then
        Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="NO TEMPLATE ATTACHED", Msg:="Macmillan RSuite styles not found." & vbNewLine & vbNewLine & "Please ensure you have a style template attached to this document."
    Else
        Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & Err.Number & " - " & Err.Description
    End If
    
End Sub

Sub StartCleanup(opts As tpOptions)
    On Error GoTo ErrorHandler
    
    Call PublicVariables.SetCharacters
    
    Dim StoryNo, StoryName As Variant
    ' we only want to run trackchanges / comments once,
    '   b/c these functions cycles through the whole doc at once,
    '   in order to only prompt the user once
    Dim TCrun_bool As Boolean
    TCrun_bool = False

    ' setup progress bar
    Set pBar = New Progress_Bar
    pBarCounter = 0
    pBar.Caption = "RSuite Cleanup Macros"
    completeStatus = "Starting Cleanup"
    pBar.Status.Caption = completeStatus

    'determine stories in document
    For Each StoryNo In ActiveDocument.StoryRanges
        'run on main (1), and endnotes (2), and footnotes (3) if selected
        If StoryNo.StoryType = 1 Or _
            (StoryNo.StoryType = 2 And opts.IncludeNotes = True) Or _
            (StoryNo.StoryType = 3 And opts.IncludeNotes = True) Then
            
            MyStoryNo = StoryNo.StoryType
            
            Select Case MyStoryNo
                Case 1
                    StoryName = "Main Body"
                Case 2
                    StoryName = "Footnotes"
                Case 3
                    StoryName = "Endnotes"
            End Select
            
            completeStatus = completeStatus + vbNewLine + _
                            "=========================" + vbNewLine + _
                            "Cleaning " & StoryName + vbNewLine + _
                            "========================="
            Clean_helpers.updateStatus ("")
                    
            'run routines
            If opts.Ellipses Then Call Clean.Ellipses(MyStoryNo)
            If opts.Spaces Then Call Clean.Spaces(MyStoryNo)
            If opts.Punctuation Then Call Clean.Punctuation(MyStoryNo)
            If opts.Hyphens Then Call Clean.Dashes(MyStoryNo)
            If opts.Quotes Then
                Call Clean.DoubleQuotes(MyStoryNo)
                Call Clean.SingleQuotes(MyStoryNo)
            End If
            
            If opts.TitleCase Then
                Call Clean.MakeTitleCase(MyStoryNo)
            End If
            
            If opts.CleanBreaks Then
                Call Clean.CleanBreaks(MyStoryNo)
            End If
            
            If opts.DeleteMarkup And TCrun_bool = False Then
                Call Clean.RemoveTrackChanges
                Call Clean.RemoveComments
                TCrun_bool = True
            End If
            
            If opts.DeleteObjects Then
                Call Clean.DeleteBookmarks
                Call Clean.DeleteObjects(MyStoryNo)
            End If
            
            If opts.RemoveHyperlinks Then
                Call Clean.RemoveHyperlinks(MyStoryNo)
            End If
            
            ' cleanup custom endnotes/footnotes
            If MyStoryNo = 2 Then
                If Clean_helpers.fnoteRefText = True Then
                    Call Clean.fixCustomFootnotes
                End If
            ElseIf MyStoryNo = 3 Then
                If Clean_helpers.enoteRefText = True Then
                    Call Clean.fixCustomEndnotes
                End If
            End If
        End If
    Next
    
    Call PublicVariables.DestroyCharacters
    Clean_helpers.ClearSearch
    Unload pBar
    
    Call Clean_helpers.MessageBox("Done", "Cleanup is complete!", vbOK)
    
    Exit Sub
    
ErrorHandler:
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERROR", Msg:="Sorry, an error occurred: " & Err.Number & " - " & Err.Description

End Sub
