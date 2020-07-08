VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public WithEvents appWord As Word.Application
Attribute appWord.VB_VarHelpID = -1

Public Sub appWord_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    myRibbon.Invalidate
End Sub

'Public Sub appWord_WindowSelectionChange(ByVal Sel As Selection)
'    myRibbon.InvalidateControl ("cboApplyStyles")
'End Sub


