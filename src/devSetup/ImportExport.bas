Attribute VB_Name = "ImportExport"
Option Base 0
Option Explicit

Sub z_Export_or_Import_VBA_Components()
Dim oFrm As frmPortVBObjects
    Set oFrm = New frmPortVBObjects
    oFrm.Show
    Unload oFrm
    Set oFrm = Nothing

lbl_Exit:
    Exit Sub
End Sub
Sub Open_All_Defined_VBA_Projects()
Dim dotm As Variant
    Call config.defineVBAProjectParams
    For Each dotm In dotms
        Documents.Open filename:=dotm.installedPath
    Next dotm
End Sub

Sub z_copyInstalledRSWTtoRepo()
Call config.Copy_Installed_Binary_Back_to_Repo("RSuite_Word-template.dotm")

End Sub

Function makeExportDir(exportDirPath As String, fso As Object) As String

    If fso.FolderExists(exportDirPath) = False Then
        On Error Resume Next
        MkDir exportDirPath
        On Error GoTo 0
    End If
    
    If fso.FolderExists(exportDirPath) = True Then
        makeExportDir = exportDirPath
    Else
        makeExportDir = "Error"
    End If
    
End Function

Sub runExports(vbObjArray As Variant, import As Boolean, modulesOnly As Boolean)
Dim strVbproj As Variant

    ' For imports this runs a backup export
    For Each strVbproj In vbObjArray
        Call ExportModules(strVbproj, import)
        ' if import is true, runs imports here
        If import = True Then
            Call ImportModules(strVbproj)
        ' if import is false, try copy binary back to repo
        ElseIf modulesOnly = False Then
            config.Copy_Installed_Binary_Back_to_Repo (strVbproj)
        End If
    Next strVbproj


End Sub

Function getVbaProjectList() As String()
Dim vbProj As VBIDE.VBProject
Dim arrProjectList() As String
Dim strDoc As Variant
Dim i As Long
i = 0

For Each vbProj In Application.VBE.VBProjects   'Loop through each project
    For Each strDoc In Documents              'Find the document name that matches
        If strDoc.VBProject Is vbProj Then
            ReDim Preserve arrProjectList(i)
            arrProjectList(i) = vbProj.Name + " (" + strDoc + ")"
            i = i + 1
        End If
    Next strDoc
Next vbProj
    
getVbaProjectList = arrProjectList
End Function
    Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the appropriate file extension based on the Type of
    ' the VBComponent.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                GetFileExtension = ".cls"
            Case vbext_ct_Document
                GetFileExtension = ".cls"
            Case vbext_ct_MSForm
                GetFileExtension = ".frm"
            Case vbext_ct_StdModule
                GetFileExtension = ".bas"
            Case Else
                GetFileExtension = ".bas"
        End Select
        
    End Function

Function FolderWithVBAProjectFiles(srcDoc As Word.Document, boolImport As Boolean) As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String
    Dim docPath As String
    Dim docBasename As String
    Dim cfgExportDirPath As String
    Dim stdExportDirPath As String
    Dim results As String
    
    Set fso = CreateObject("scripting.filesystemobject")
    
    ' in future, check for module in file to look for preferred export/import locale for doc
    cfgExportDirPath = config.configPathCheck(srcDoc.Name)
    If boolImport = True Then
        cfgExportDirPath = cfgExportDirPath + "_BACKUP_"
    End If
    
    ' default to rel folder in same dir as source doc
    docPath = srcDoc.Path
    docBasename = fso.GetBaseName(srcDoc.Name)
    stdExportDirPath = docPath + "\src-" + docBasename
    If boolImport = True Then
        stdExportDirPath = stdExportDirPath + "_BACKUP_"
    End If

    ' backup path is my Documents path
    Set WshShell = CreateObject("WScript.Shell")
    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    SpecialPath = SpecialPath & "VBAProjectFiles"
    If boolImport = True Then
        SpecialPath = SpecialPath + "_BACKUP_"
    End If
    
    
    results = makeExportDir(cfgExportDirPath, fso)
    ' on error retry backup locations
    If results <> cfgExportDirPath Then
       results = makeExportDir(stdExportDirPath, fso)
        If results <> stdExportDirPath Then
           results = makeExportDir(SpecialPath, fso)
        End If
    End If
    FolderWithVBAProjectFiles = results
    
End Function

    
  Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional filename As String, _
                Optional OverwriteExisting As Boolean = True) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This function exports the code module of a VBComponent to a text
    ' file. If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim fname As String
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(filename) = vbNullString Then
        fname = VBComp.Name & Extension
    Else
        fname = filename
        If InStr(1, fname, ".", vbBinaryCompare) = 0 Then
            fname = fname & Extension
        End If
    End If
    
    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        fname = FolderName & fname
    Else
        fname = FolderName & "\" & fname
    End If
    
    If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill fname
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If
    
    VBComp.Export filename:=fname
    ExportVBComponent = True
    
    End Function
    
    
Public Sub ExportModules(szSourceDocument As Variant, boolImport As Boolean)
    Dim bExport As Boolean
    Dim docSource As Word.Document
    Dim stdSrcPath As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ' Remnanat of orig version. LEaving in case we ever want a standalone version with no arg passed
    If IsMissing(szSourceDocument) Then
        szSourceDocument = ActiveDocument.Name
    End If
    Set docSource = Application.Documents(szSourceDocument)
    
    ''' The code below creates target folder if it does not exist
    ''' or deletes all files in the folder if it exist.
    If FolderWithVBAProjectFiles(docSource, boolImport) = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles(docSource, boolImport) & "\*.*"
    On Error GoTo 0

    If docSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles(docSource, boolImport) & "\"
    
    For Each cmpComponent In docSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                Debug.Print "compname " & szFileName
                ''' This is a document object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''docSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent
    If boolImport = False Then
        MsgBox "Export is ready"
    Else
        MsgBox "Backed up current VB components for " + szSourceDocument
    End If
End Sub

Public Sub ImportModules(szTargetDoc As Variant)
    Dim docTarget As Word.Document
    Dim objFSO As scripting.FileSystemObject
    Dim objFile As scripting.File
    'Dim szTargetDoc As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If szTargetDoc = ThisDocument.Name Then
        MsgBox "Select another destination document" & _
        "Not possible to import into: " + szTargetDoc
        Exit Sub
    End If

    Set docTarget = Application.Documents(szTargetDoc)

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles(docTarget, False) = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    If docTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles(docTarget, False) & "\"
        
    Set objFSO = New scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the target Document
    Call DeleteVBAModulesAndUserForms(docTarget)

    Set cmpComponents = docTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the target Document.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Function DeleteVBAModulesAndUserForms(docTarget As Word.Document)
        Dim vbProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set vbProj = docTarget.VBProject
        
        For Each VBComp In vbProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                vbProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function




