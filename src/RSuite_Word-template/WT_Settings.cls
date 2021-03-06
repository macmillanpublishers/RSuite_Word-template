VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WT_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
'       WT_Settings
' =============================================================================

' ***** USE *******************************************************************
' Static class to manage settings for the Word-template. Can access at any time
' by using it's name directly: WT_Settings.PropertyName

' ***** DEPENDENCIES **********************************************************
' Static class: there is a single default instance that is always available, so
' you don't need to (in fact, can't) create multiple objects from it.
' To set as a static class (only need to do once):
' 1. Export the class module.
' 2. Open the exported file in a text editor.
' 3. On line 8, set Attribute VB_PredeclaredId = True
' 4. Save .cls file.
' 5. Back in VB Editor, remove (and don't save) current class module.
' 6. Import the class module you edited above.

' ***** TO DO *****************************************************************
' [ ] Change all properties to read from config file

' *****************************************************************************
'       DECLARATIONS
' *****************************************************************************
Option Explicit

Private p_strInstallType As String
Private p_MasterConfig As Dictionary
Private p_GlobalConfig As Dictionary
Private p_DefaultConfigPath As String
Private p_TmpDir As String
Private p_Region As String
Private p_o2016 As String
Private p_showStatus As String
Private p_RegionConfig As Dictionary
Private p_LocalConfig As Dictionary


' *****************************************************************************
'     PROCEDURES
' *****************************************************************************

' ***** PUBLIC PROPERTIES *****************************************************

' ===== InstallType ===========================================================
' Word-template.dotm can be used in multiple settings, notably on the server to
' run bookmaker_validator vs. by end-users (via Macmillan Tools tab). This returns
' the current place the template is installed.

Public Property Get InstallType() As String
  If p_strInstallType = vbNullString Then
    Dim strLocalUser As String
    strLocalUser = Environ("USERNAME")

  ' username is "padwoadmin" on both servers
  ' TODO: change this to read from config file
    Select Case strLocalUser
      Case "padwoadmin"
        p_strInstallType = "server"
      Case Else
        p_strInstallType = "user"
    End Select

  Else
'    DebugPrint "InstallType already set"
  End If
  
'  DebugPrint "Current InstallType: " & p_strInstallType
  InstallType = p_strInstallType
End Property


' ===== DebugOn ===============================================================
' Are we in debugging mode? NOTE: can't use 'Debug' because that's reserved!

Public Property Get DebugOn() As Boolean
  If LocalConfig.Exists("debug") = True Then
    DebugOn = LocalConfig("debug")
  End If
End Property


' ===== StyleDir ==============================================================
' Returns path to RSuiteStyleTemplate directory on local machine. If you pass
' "GtUpdater.dotm" as FileName, it returns the Startup folder.


Public Property Get StyleDir(Optional FileName As String, Optional FileType As String) As String
  Dim strFullPath As String
  Dim strMacDocs As String
  Dim strStylesName As String
  strStylesName = "RSuiteStyleTemplate"

  #If Mac Then
    strMacDocs = MacScript("return (path to documents folder) as string")
    
    If WT_Settings.o2016 = True Then
      Dim pathArr As Variant
      Dim rootLength As Integer
      Dim newPath As String

      pathArr = Split(strMacDocs, ":")
      rootLength = Len(pathArr(0))
      newPath = Right(strMacDocs, Len(strMacDocs) - rootLength)
      newPath = Replace(newPath, ":", "/")
      
      strFullPath = newPath & strStylesName

    Else
      strFullPath = strMacDocs & strStylesName
    End If
    
    If FileType = "tools" Then
        'strFullPath = Application.StartupPath
        strFullPath = strFullPath & Application.PathSeparator
    ElseIf FileType = "styles" Then
        strFullPath = strFullPath & Application.PathSeparator & "StyleTemplate_auto-generate"
    End If
    
  #Else
    If FileType = "tools" Then
        strFullPath = Application.StartupPath
    ElseIf FileType = "styles" Then
        strFullPath = Environ("APPDATA") & Application.PathSeparator & strStylesName & Application.PathSeparator & "StyleTemplate_auto-generate"
    Else
        strFullPath = Environ("APPDATA") & Application.PathSeparator & strStylesName
    End If
  #End If

' Create directory if it doesn't exist yet
  If Utils.IsItThere(strFullPath) = False Then
    MkDir strFullPath
  End If

  StyleDir = strFullPath

End Property


' ===== Region ================================================================
' What region (us or uk) is user in?

Public Property Get Region() As String
  If p_Region = vbNullString Then
    p_Region = LocalConfig("region")
    If p_Region = vbNullString Then
      p_Region = AskForRegion
    End If
  End If
  Region = p_Region
End Property

' ===== Office 2016 ================================================================
' Is this Office 2016 for Mac?

Public Property Get o2016() As String
  If p_o2016 = vbNullString Then
    p_o2016 = False
    #If Mac Then
        Dim thisVersion As Variant
        thisVersion = Split(Application.Version, ".")
        If thisVersion(0) > 14 Then
            p_o2016 = True
        End If
    #End If
  End If
  o2016 = p_o2016
End Property

' ===== Ignore StatusBar ================================================================
' If this is Mac 2016, DisplayStatusBar does not work, so disable code using that

Public Property Get showStatus() As String
  If p_showStatus = vbNullString Then
    p_showStatus = True
    If WT_Settings.o2016 = True Then p_showStatus = False
  End If
  showStatus = p_showStatus
End Property

' ***** PRIVATE PROPERTIES *****************************************************

' ===== TmpDir ================================================================
' Returns path to TEMP directory, no trailing path separator

Public Property Get TmpDir() As String
  If p_TmpDir = vbNullString Then
    #If Mac Then
      If WT_Settings.o2016 = True Then
        p_TmpDir = "/Users/" & Environ("USER") & "/Downloads"
      Else
       p_TmpDir = MacScript("path to temporary items as string")
      End If
    #Else
      p_TmpDir = Environ("TEMP")
    #End If
  ' Remove trailing path separator, if any
    If Right(p_TmpDir, 1) = Application.PathSeparator Then
      p_TmpDir = Left(p_TmpDir, Len(p_TmpDir) - 1)
    End If
  End If
  TmpDir = p_TmpDir
End Property


' ===== DefaultConfigPath =====================================================
' Default path to config files on local machine.

Private Property Get DefaultConfigPath() As String
  If p_DefaultConfigPath = vbNullString Then
    p_DefaultConfigPath = StyleDir & Application.PathSeparator & "config"
    If Utils.IsItThere(p_DefaultConfigPath) = False Then
      MkDir p_DefaultConfigPath
    End If
  End If
    DefaultConfigPath = p_DefaultConfigPath
End Property

' ===== MasterConfig ==========================================================
' All config data, i.e. local data overwrites region, overwrites global.

Public Property Get MasterConfig() As Dictionary
  If p_MasterConfig Is Nothing Then
    Set p_MasterConfig = Me.GlobalConfig
    Set p_MasterConfig = MergeDictionary(p_MasterConfig, Me.RegionConfig)
    Set p_MasterConfig = MergeDictionary(p_MasterConfig, LocalConfig)
  End If
  Set MasterConfig = p_MasterConfig
End Property

' ===== GlobalConfig ==========================================================
' Dictionary of data in global_config.json file. If file isn't present, it is
' downloaded from the URL saved in the template's CustomDocumentProperties.

Public Property Get GlobalConfig() As Dictionary
  If p_GlobalConfig Is Nothing Then
    Dim strFileName As String
    Dim strFullPath As String
    strFileName = "global_config.json"
    strFullPath = DefaultConfigPath & Application.PathSeparator & strFileName

    If Utils.IsItThere(strFullPath) = False Then
      DownloadFromGithub strFileName
    End If
    Set p_GlobalConfig = ClassHelpers.ReadJson(strFullPath)
  End If
  
  Set GlobalConfig = p_GlobalConfig
End Property

' ===== RegionConfig ==========================================================
' Dictionary of data in global_config.json file. If file isn't present, it is
' downloaded from the URL saved in the template's CustomDocumentProperties.

Public Property Get RegionConfig() As Dictionary
  If p_RegionConfig Is Nothing Then
    Dim strFileName As String
    Dim strFullPath As String
    strFileName = Me.Region & "_region_config.json"
    strFullPath = DefaultConfigPath & Application.PathSeparator & strFileName

    If Utils.IsItThere(strFullPath) = False Then
      DownloadFromGithub strFileName
    End If
    Set p_RegionConfig = ClassHelpers.ReadJson(strFullPath)
  End If
  
  Set RegionConfig = p_RegionConfig
End Property

' ===== LocalConfig ===========================================================
' Local settings. Both a Set and a Get since we need to be able to edit locally.
'
Private Property Set LocalConfig(NewLocalConfig As Dictionary)
  Dim strLocalFile As String
  strLocalFile = DefaultConfigPath & Application.PathSeparator & "local_config.json"
  ClassHelpers.WriteJson strLocalFile, NewLocalConfig
  Set p_LocalConfig = NewLocalConfig
End Property

Public Property Get LocalConfig() As Dictionary
  If p_LocalConfig Is Nothing Then
    Dim strLocalFile As String
    strLocalFile = DefaultConfigPath & Application.PathSeparator & "local_config.json"
    
    If Utils.IsItThere(strLocalFile) = False Then
      Set p_LocalConfig = New Dictionary
    Else
      Set p_LocalConfig = ClassHelpers.ReadJson(strLocalFile)
    End If
  End If
  Set LocalConfig = p_LocalConfig
End Property

' *****************************************************************************
'       METHODS
'******************************************************************************

' ===== AskForRegion ==========================================================
' Prompts the user to pick their Region. Eventually turn it into a UserForm, and
' pull the region options from global_config.

Private Function AskForRegion() As String
'  Dim strMessage As String
'  strMessage = "Are you in the United States?" & vbNewLine & vbNewLine & _
'  vbTab & "US users: click YES" & vbNewLine & _
'  vbTab & "UK users: click NO"
'
'  Dim strTitle As String
'  strTitle = "Please Select Your Region"
'
'  If MsgBox(strMessage, vbYesNo + vbQuestion, strTitle) = vbYes Then
'    AskForRegion = "us"
'  Else
'    AskForRegion = "uk"
'  End If
  
  AskForRegion = "us"
' Now we have to add this to LocalConfig and write it to the JSON file
  Dim dictConfig As Dictionary
  Set dictConfig = LocalConfig
  dictConfig.item("region") = AskForRegion
' LocalConfig Set procedure sets the local variable and writes to JSON
  Set LocalConfig = dictConfig
End Function

' *****************************************************************************
'       EVENTS
'******************************************************************************

' ===== Class_Terminate =======================================================


