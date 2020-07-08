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
    
    For Each StoryNo In ActiveDocument.StoryRanges
        
        If StoryNo.StoryType < 4 Then
        
            MyStoryNo = StoryNo.StoryType
            
            Set pBar = New Progress_Bar
            pBar.Caption = "RSuite Character Style"
            completeStatus = "Starting Character Style Replacement"
            pBar.Status.Caption = completeStatus
        
            Call Clean.LocalFormatting
            Call Clean.CheckSpecialCharactersPC
            Call CheckAppliedCharStyles
            
            Unload pBar
            
        End If
    Next
    
    Call Clean_helpers.MessageBox("Done", "Character Styles is complete!", vbOK)
    
    Exit Sub
    
ErrorHandler:
    
    If Err.Number = 5834 Then
        Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="NO TEMPLATE ATTACHED", Msg:="Macmillan RSuite styles not found." & vbNewLine & vbNewLine & "Please ensure you have a style template attached to this document."
    Else
        Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERRROR", Msg:="Sorry, an error occurred: " & Err.Number & " - " & Err.Description
    End If
    
End Sub

Sub StartCleanup(opts As tpOptions)
    On Error GoTo ErrorHandler
    
    Call PublicVariables.SetCharacters
    
    Dim StoryNo, StoryName As Variant

    Set pBar = New Progress_Bar
    pBar.Caption = "RSuite Cleanup Macros"
    completeStatus = "Starting Cleanup"
    pBar.Status.Caption = completeStatus

    'determine stories in document
    For Each StoryNo In ActiveDocument.StoryRanges
        
        'run on main (1) endnotes (2), and footnotes (3)
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
                    
            'run routines
            If opts.Ellipses Then Call Clean.Ellipses
            If opts.Spaces Then Call Clean.Spaces
            If opts.Punctuation Then Call Clean.Punctuation
            If opts.Hyphens Then Call Clean.Dashes
            If opts.Quotes Then
                Call Clean.DoubleQuotes
                Call Clean.SingleQuotes
            End If
            
            If opts.TitleCase Then
             Call Clean.MakeTitleCase
            End If
            
            If opts.CleanBreaks Then
                Call Clean.CleanBreaks
            End If
            
            If opts.DeleteMarkup Then
                Call Clean.RemoveTrackChanges
                Call Clean.RemoveComments
            End If
            
            If opts.DeleteObjects Then
                Call Clean.DeleteBookmarks
                Call Clean.DeleteObjects
            End If
            
            If opts.RemoveHyperlinks Then
                Call Clean.RemoveHyperlinks
            End If
                
        End If
    Next
    
    Call PublicVariables.DestroyCharacters
    Clean_helpers.ClearSearch
    Unload pBar
    
    Call Clean_helpers.MessageBox("Done", "Cleanup is complete!", vbOK)
    
    Exit Sub
    
ErrorHandler:
    Clean_helpers.MessageBox buttonType:=vbOKOnly, Title:="UNEXPECTED ERRROR", Msg:="Sorry, an error occurred: " & Err.Number & " - " & Err.Description

End Sub