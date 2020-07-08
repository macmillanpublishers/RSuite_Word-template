Attribute VB_Name = "Install"
'Public Sub MoveAutoMacro()
'
'    On Error GoTo Handler:
'
'    Dim sPath, thisName, mydoc As String
'    Dim moduleCount, i As Integer
'    Dim amExists As Boolean
'
'    moduleCount = Application.Templates("Normal.dotm").VBProject.vbcomponents.Count
'    amExists = False
'    For i = 1 To moduleCount
'        thisName = Application.Templates("Normal.dotm").VBProject.vbcomponents(i).Name
'        If thisName = "AutoMacros" Then amExists = True
'    Next
'
'    If Not amExists Then
'
'        sPath = Environ("USERPROFILE") & Application.PathSeparator & "AppData" & Application.PathSeparator & _
'            "Roaming" & Application.PathSeparator & "Microsoft" & Application.PathSeparator & "Templates" & _
'            Application.PathSeparator & "Normal.dotm"
'
'        Application.OrganizerCopy _
'            Source:=ThisDocument.FullName, _
'            Destination:=sPath, Name:="AutoMacros", _
'            Object:=wdOrganizerObjectProjectItems
'
'
''        sPath = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator & _
''                "AutoMacros.bas"
''        ThisDocument.VBProject.vbcomponents("AutoMacros").Export sPath
''        MsgBox "Exported"
''        Application.Templates("Normal.dotm").VBProject.vbcomponents.Import sPath
'        MsgBox "Moved to Normal"
'
'    Else:
'        MsgBox "Already in Normal"
'    End If
'
'    Exit Sub
'
'Handler:
'    MsgBox Err.Number & " " & Err.Description
'End Sub

