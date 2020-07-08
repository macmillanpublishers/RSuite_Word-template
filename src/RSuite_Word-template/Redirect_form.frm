VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Redirect_form 
   Caption         =   "UserForm1"
   ClientHeight    =   3795
   ClientLeft      =   -8295
   ClientTop       =   -4200
   ClientWidth     =   5400
   OleObjectBlob   =   "Redirect_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Redirect_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub hyperlink_Click()
    followlink (hyperlink.Caption)
End Sub

Private Sub mail_Click()
    followlink ("mailto:" & mail.Caption)
End Sub

Private Sub followlink(Link As String)
    On Error GoTo HandleMe
    ActiveDocument.FollowHyperlink _
        Address:=Link, _
        NewWindow:=True
    Unload Me
    Exit Sub
HandleMe:
    MsgBox "Cannot open " & Link & "."
End Sub

