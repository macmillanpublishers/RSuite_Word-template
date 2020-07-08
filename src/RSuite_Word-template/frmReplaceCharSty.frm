VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReplaceCharSty 
   Caption         =   "Invalid Character Style"
   ClientHeight    =   4710
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   6075
   OleObjectBlob   =   "frmReplaceCharSty.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmReplaceCharSty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCancel_Click()
    endCharCheck = True
    Unload Me
End Sub

Private Sub cbRemoveAll_Click()
    On Error GoTo Handler
    
    Selection.Style = wdStyleDefaultParagraphFont
    
    Dim x As Integer
    x = UBound(removeStyles) + 1
    ReDim Preserve removeStyles(x)
    removeStyles(x) = frmReplaceCharSty.Tag
    
    Unload Me
    
    Exit Sub

Handler:
    If Err.Number = 9 Then
        x = 0
        Resume Next
    End If
    
End Sub

Private Sub cbRemoveOnce_Click()
    Selection.Style = wdStyleDefaultParagraphFont
    Unload Me
End Sub

Private Sub cbReplaceAll_Click()

    On Error GoTo Handler
    Selection.Style = frmReplaceCharSty.cbList.value
    
    Dim stylePair(1) As Variant
    stylePair(0) = frmReplaceCharSty.Tag
    stylePair(1) = frmReplaceCharSty.cbList.value
    
    Dim x As Integer
    x = UBound(replaceStyles) + 1
    ReDim Preserve replaceStyles(x)
    replaceStyles(x) = stylePair
    
    Unload Me
    Exit Sub

Handler:
    If Err.Number = 9 Then
        x = 0
        Resume Next
    End If

End Sub

Private Sub cbReplaceOnce_Click()
    Selection.Style = frmReplaceCharSty.cbList.value
    Unload Me
End Sub

Private Sub cbSkipAll_Click()

    On Error GoTo Handler
    
    Dim x As Integer
    x = UBound(skipStyles) + 1
    ReDim Preserve skipStyles(x)
    skipStyles(x) = frmReplaceCharSty.Tag
    
    Unload Me
    Exit Sub

Handler:
    If Err.Number = 9 Then
        x = 0
        Resume Next
    End If
End Sub

Private Sub cbSkipOnce_Click()
    Unload Me
End Sub


