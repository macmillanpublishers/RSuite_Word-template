VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CleanupOptions 
   Caption         =   "Cleanup Options"
   ClientHeight    =   7021
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6041
   OleObjectBlob   =   "CleanupOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CleanupOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCancel_Click()
    Unload CleanupOptions
End Sub

Private Sub cbOK_Click()

    Dim o As tpOptions
    o.Ellipses = cbEllipses.value
    o.Punctuation = cbPunct.value
    o.Quotes = cbCurly.value
    o.Spaces = cbSpaces.value
    o.Hyphens = cbHyphens.value
    o.CleanBreaks = chkCleanBreaks.value
    o.TitleCase = chkTitleCase.value
    o.DeleteMarkup = chkDeleteMarkup.value
    o.DeleteObjects = chkDeleteBookmarks.value
    o.RemoveHyperlinks = chkRemoveHyperlinks.value
    Unload CleanupOptions
    Clean_Start.StartCleanup opts:=o

End Sub

Private Sub chkCleanFormatting_Click()

    Dim thisVal As Boolean
    thisVal = chkCleanFormatting.value
    cbEllipses.value = thisVal
    cbPunct.value = thisVal
    cbCurly.value = thisVal
    cbSpaces.value = thisVal
    cbHyphens.value = thisVal

End Sub

Private Sub UserForm_Initialize()
    
    chkCleanFormatting.value = True
    chkTitleCase.value = True
    chkCleanBreaks.value = True
    chkDeleteMarkup.value = True
    chkDeleteBookmarks.value = True
    chkRemoveHyperlinks.value = True

End Sub
