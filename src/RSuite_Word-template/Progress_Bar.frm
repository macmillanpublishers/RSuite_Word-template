VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progress_Bar 
   Caption         =   "Macro Progress"
   ClientHeight    =   7504
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   OleObjectBlob   =   "Progress_Bar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progress_Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    #If Mac Then
        ' Nada! Can't run a useform modeless on Mac 2011, so a progress bar is useless.
        ' But lucky you, the Increment method will update the Mac Status Bar instead
        ' So go right on ahead with this, just DON'T use ProgressBar.Show method in the calling sub
        ' We'll show it below for PCs only
        If IsOldMac = True Then
            Me.Hide ' In case someone uses the Show method without the Load method
            Application.DisplayStatusBar = True
        End If
    #Else
        Me.Show
    #End If

End Sub




