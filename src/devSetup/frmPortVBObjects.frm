VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPortVBObjects 
   Caption         =   "UserForm1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "frmPortVBObjects.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPortVBObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdExportModules_Click()
Dim vbObjArray() As Variant
Dim i As Long
Dim lngCount As Long
Dim strDocname As String
Dim boolImport As Boolean
Dim modulesOnly As Boolean

    boolImport = False
    modulesOnly = True
    lngCount = 0

    'add selected items to array to pass to runExports
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
        ReDim Preserve vbObjArray(lngCount)
        ' capture just the document name
        strDocname = Split(ListBox1.List(i), " (")(1)
        strDocname = Left(strDocname, Len(strDocname) - 1)
        Debug.Print strDocname + "sd"
        vbObjArray(lngCount) = strDocname
        lngCount = lngCount + 1
        End If
    Next i
    
    If lngCount = 0 Then
        MsgBox "You must select at least one item to continue"
        Exit Sub
    End If
    
    Me.Hide
    ' call run Exports with boolImport as False
    Call ImportExport.runExports(vbObjArray, boolImport, modulesOnly)
    
End Sub

Private Sub cmdImport_Click()
Dim vbObjArray() As Variant
Dim i As Long
Dim lngCount As Long
Dim strDocname As String
Dim boolImport As Boolean
Dim modulesOnly As Boolean

    boolImport = True
    modulesOnly = False
    lngCount = 0
    
    'add selected items to array to pass to runExports
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            ReDim Preserve vbObjArray(lngCount)
            ' capture just the document name
            strDocname = Split(ListBox1.List(i), " (")(1)
            strDocname = Left(strDocname, Len(strDocname) - 1)
            Debug.Print strDocname + "sd"
            vbObjArray(lngCount) = strDocname
            lngCount = lngCount + 1
        End If
    Next i
    
    If lngCount = 0 Then
        MsgBox "You must select at least one item to continue"
        Exit Sub
    End If
    
    Me.Hide
    ' call run Exports with boolImport as True
    Call ImportExport.runExports(vbObjArray, boolImport, modulesOnly)
    
End Sub
Private Sub cmdExport_Click()
Dim vbObjArray() As Variant
Dim i As Long
Dim lngCount As Long
Dim strDocname As String
Dim boolImport As Boolean
Dim modulesOnly As Boolean

    boolImport = False
    modulesOnly = False
    lngCount = 0

    'add selected items to array to pass to runExports
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
        ReDim Preserve vbObjArray(lngCount)
        ' capture just the document name
        strDocname = Split(ListBox1.List(i), " (")(1)
        strDocname = Left(strDocname, Len(strDocname) - 1)
        Debug.Print strDocname + "sd"
        vbObjArray(lngCount) = strDocname
        lngCount = lngCount + 1
        End If
    Next i
    
    If lngCount = 0 Then
        MsgBox "You must select at least one item to continue"
        Exit Sub
    End If
    
    Me.Hide
    ' call run Exports with boolImport as False
    Call ImportExport.runExports(vbObjArray, boolImport, modulesOnly)
    
End Sub

Private Sub cmdSelectAll_Click()
Dim i As Long
  For i = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(i) = True
  Next i
End Sub


Private Sub UserForm_Initialize()
Dim i As Long
Dim j As Long
Dim arrVbaProjects() As String
Dim strProjName As Variant

    ' get list of open projects
    arrVbaProjects = getVbaProjectList()
    ' populate our listbox
    For Each strProjName In arrVbaProjects
        ListBox1.AddItem (strProjName)
    Next strProjName

lbl_Exit:
    Exit Sub
End Sub

