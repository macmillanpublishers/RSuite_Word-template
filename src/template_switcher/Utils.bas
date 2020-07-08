Attribute VB_Name = "Utils"
Option Explicit

Type tpTemplate
    tName As String
    tNumber As String
End Type

Public Function GetVersion(ByVal fullPath As String)
    Dim installedVersion As String
    
    If IsItThere(fullPath) = False Then
        installedVersion = "none"
    Else
        Documents.Open FileName:=fullPath, ReadOnly:=True
        installedVersion = Documents(fullPath).CustomDocumentProperties("version")
        Documents(fullPath).Close
    End If
    
    If Left(installedVersion, 1) = "v" Then
        GetVersion = Right(installedVersion, Len(installedVersion) - 1)
    Else
        GetVersion = installedVersion
    End If

End Function

Public Function DocPropExists(ByRef objDoc As Document, ByVal PropName As String) As Boolean
  DocPropExists = False

' Note DocumentProperties returns a Collection
  Dim docProps As DocumentProperties
  Set docProps = objDoc.CustomDocumentProperties

  Dim varProp As Variant

  If docProps.Count > 0 Then
      For Each varProp In docProps
          If varProp.Name = PropName Then
              DocPropExists = True
              Exit Function
          End If
      Next varProp
  Else
      DocPropExists = False
  End If
  
  'MsgBox "4. Doc Prop " & PropName & " Exists: " & DocPropExists
  
End Function

Public Function SetStyleVersion(ByRef VersionNumber As String, ByVal templateNo As String)
  Dim strProps(), Prop As Variant
  strProps = Array(Array("Version", VersionNumber), Array("Template", templateNo))
  
  For Each Prop In strProps
    If Utils.DocPropExists(objDoc:=ActiveDocument, PropName:=Prop(0)) Then
        ActiveDocument.CustomDocumentProperties(Prop(0)).Value = Prop(1)
        'MsgBox "5. Overwriting existing property: " & Prop(0) & " " & Prop(1)
    Else
        ActiveDocument.CustomDocumentProperties.Add Name:=Prop(0), LinkToContent:=False, _
            Type:=msoPropertyTypeString, Value:=Prop(1)
        'MsgBox "5. Writing new property: " & Prop(0) & " " & Prop(1)
    End If
  Next

End Function


' ===== IsItThere =============================================================
' Check if file or directory exists on PC or Mac.
' Dir() doesn't work on Mac 2011 if file is longer than 32 char

Public Function isVisible(control As IRibbonControl, ByRef visible)
    
    visible = False
    #If Mac Then
        If control.Tag = "Mac" Then visible = True
        If control.Tag = "PC" Then visible = False
    #Else
        If control.Tag = "Mac" Then visible = False
        If control.Tag = "PC" Then visible = True
    #End If

End Function

Public Function IsItThere(ByVal Path As String) As Boolean

  On Error GoTo Handler
  'Remove trailing path separator from dir if it's there
  If Right(Path, 1) = Application.PathSeparator Then
    Path = Left(Path, Len(Path) - 1)
  End If
  
  Dim strCheckDir As String
  
  strCheckDir = vbNullString
  
  #If Mac Then
    #If MAC_OFFICE_VERSION >= 15 Then
        strCheckDir = Dir(Path, vbDirectory)
    #Else
        Dim strScript As String
        strScript = "tell application " & Chr(34) & "System Events" & Chr(34) & _
            "to return exists disk item (" & Chr(34) & Path & Chr(34) _
            & " as string)"
         IsItThere = MacScript(strScript)
         Exit Function
    #End If
#Else
    strCheckDir = Dir(Path, vbDirectory)
#End If

If strCheckDir = vbNullString Then
    IsItThere = False
Else
    IsItThere = True
End If

Exit Function
  
Handler:
    If Err.Number = 53 Then
        strCheckDir = vbNullString
        Resume Next
    Else:
        MsgBox Err.Number & ": " & Err.Description
        strCheckDir = vbNullString
        Resume Next
    End If
    
End Function

