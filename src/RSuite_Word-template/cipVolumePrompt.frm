VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cipVolumePrompt 
   Caption         =   "UserForm1"
   ClientHeight    =   5628
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6792
   OleObjectBlob   =   "cipVolumePrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cipVolumePrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isCancelled As Boolean
Private frmTagChapters As Boolean
' many ideas for this form borrowed from:
'   https://stackoverflow.com/questions/43767656/pass-a-string-from-a-user-form-button-click-to-excel-vba

Public Property Get tagChapters() As Boolean
    tagChapters = frmTagChapters
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = isCancelled
End Property

Private Sub button1_Click()
    frmTagChapters = True
    Me.Hide
End Sub
Private Sub button2_Click()
    frmTagChapters = False
    Me.Hide
End Sub

Private Sub cbCancel_Click()
    isCancelled = True
    Me.Hide
End Sub

Private Sub LOC_hyperlink_Click()
    followlink (LOC_hyperlink.Caption)
End Sub

Private Sub followlink(Link As String)
    On Error GoTo HandleMe
    ActiveDocument.FollowHyperlink _
        Address:=Link, _
        NewWindow:=True
    Exit Sub
HandleMe:
    MsgBox "Cannot open " & Link & "."
End Sub

Private Sub text1_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        isCancelled = True
        Cancel = True
        Me.Hide
    End If
End Sub
