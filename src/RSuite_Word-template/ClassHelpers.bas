Attribute VB_Name = "ClassHelpers"
' =============================================================================
'     CLASS HELPERS
' =============================================================================
' By Erica Warren - erica.warren@macmillan.com
'
' ===== USE ===================================================================
' Some procedures needed for classes can't actually exist in the class for one
' reason or another, so we put them here.
'
' ===== DEPENDENCIES ==========================================================
' Obviously, the class in question must be a module in the same project.


Option Explicit

Private Const strClassHelpers As String = "ClassHelpers."

' =============================================================================
'     JSON HELPERS
' =============================================================================
' Not technically a class, but I don't feel like forking that repo now to add
' these additional functions, so I'll drop them here.

' ===== ReadJson ==============================================================
' To get from JSON file to Dictionary object, must read file to string, then
' convert string to Dictionary. This does all of that (and some error handling)


Public Function ReadJson(JsonPath As String) As Dictionary
  On Error GoTo ReadJsonError
  Dim dictJson As Dictionary
  
  If IsItThere(JsonPath) = True Then
    Dim strJson As String
    
    strJson = ReadTextFile(JsonPath, False)
    If strJson <> vbNullString Then
      Set dictJson = JsonConverter.ParseJson(strJson)
    Else
      ' If file exists but has no content, return empty dictionary
      Set dictJson = New Dictionary
    End If
  Else
    Err.Raise MacError.err_FileNotThere
  End If
  
  If dictJson Is Nothing Then
    DebugPrint "ReadJson fail"
  Else
'    DebugPrint "dictJson count: " & dictJson.Count
  End If
  
  Set ReadJson = dictJson
  Exit Function
  
ReadJsonError:
  Err.Source = strClassHelpers & "ReadJson"
  If ErrorChecker(Err, JsonPath) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function


' ===== WriteJson =============================================================
' JsonConverter.ConvertToJson returns a string, when we then need to write to
' a text file if we want the output. This combines those. Will overwrite the
' original file if already exists, will create file if it does not.


Public Sub WriteJson(JsonPath As String, JsonData As Dictionary)
  On Error GoTo WriteJsonError:

  Dim strJson As String
  strJson = JsonConverter.ConvertToJson(JsonData, Whitespace:=2)
  ' `OverwriteTextFile` validates directory
  Utils.OverwriteTextFile JsonPath, strJson
  Exit Sub
  
WriteJsonError:
  Err.Source = strClassHelpers & "WriteJson"
  If ErrorChecker(Err, JsonPath) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Sub


' ===== AddToJson =============================================================
' Adds the key/value pair to an already existing JSON file. Creates file if it
' doesn't exist yet. `NewValue` can be anything valid for JSON: string,
' number, boolean, dictionary, array. `JsonFile` is full path to file.

' NOTE!! If `NewKey` already exists, the value will be overwritten. Could change
' to check for existance and do something else instead (append number to key,
' add value to array, return false, whatever).

Public Sub AddToJson(JsonFile As String, NewKey As String, NewValue As Variant)
  On Error GoTo AddToJsonError
  Dim dictJson As Dictionary
  
  ' READ JSON FILE IF IT EXISTS
  ' Does the file exist yet?
  If Utils.IsItThere(JsonFile) = True Then
    Set dictJson = ReadJson(JsonFile)
  Else
    ' File doesn't exist yet, we'll be creating it
    Set dictJson = New Dictionary
  End If
  
  ' ADD NEW ITEM TO DICTIONARY
  ' `.Item("key")` method will add if key is new, overwrite if not
  If VBA.IsObject(NewValue) = True Then
    ' Need `Set` keyword for object
    Set dictJson.item(NewKey) = NewValue
  Else
    dictJson.item(NewKey) = NewValue
  End If

  ' WRITE UPDATED DICTIONARY (BACK) TO JSON FILE
  Call WriteJson(JsonFile, dictJson)

  Exit Sub

AddToJsonError:
  Err.Source = strClassHelpers & "AddToJson"
  If ErrorChecker(Err, JsonFile) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Sub

' =============================================================================
'       DICTIONARY HELPERS
' =============================================================================

' ===== MergeDictionary =======================================================
' Add all key:value pairs of one dictionary to another dictionary. Default is
' to overwrite value in DictOne if a key in DictTwo matches; Overwrite = False
' adds an integer to the key name and adds a new key:value pair.

Public Function MergeDictionary(DictOne As Dictionary, DictTwo As _
  Dictionary, Optional Overwrite As Boolean = True) As _
  Dictionary
  On Error GoTo MergeDictionaryError
  Dim key2 As Variant
  Dim lngCount As Long
  Dim strKey As String
  
  lngCount = 0
  
  ' Use .Item() not .Add, because .Add errors if same key is used
  For Each key2 In DictTwo.Keys
    If Overwrite = False Then
      If DictOne.Exists(key2) = True Then
        lngCount = lngCount + 1
        strKey = key2 & lngCount
      Else
        strKey = key2
      End If
    Else
      strKey = key2
    End If
    
    DictOne.item(strKey) = DictTwo(key2)
  
  Next key2
  
  Set MergeDictionary = DictOne
  Exit Function

MergeDictionaryError:
  Err.Source = strClassHelpers & "MergeDictionary"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function



' =============================================================================
'       PROGRESS BAR HELPERS
' =============================================================================
'
' ===== UpdateBarAndWait ======================================================
' Call this to update the progress bar. Can't be in the ProgressBar class,
' because that class can crash the program if it gets called
' again before it finishes the first call. This includes `DoEvents` which allows
' other work to finish before returning to the calling procedure.

Public Sub UpdateBarAndWait(ByVal currentStatus As String)
  
    On Error GoTo UpdateBarAndWaitError
    pBar.Status.Caption = completeStatus & vbNewLine & currentStatus
  Exit Sub

UpdateBarAndWaitError:
  Err.Source = strClassHelpers & "UpdateBarAndWaitError"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Sub



' =============================================================================
'       OTHER
' =============================================================================


' ===== FindStyleIndex =========================================================
' Adds index numbers for paragraphs with SearchStyle applied to DestinationCollection.
' Note that because DestinationCollection is passed ByRef we don't need to return
' it to the calling function (so we can ADD to same collection if we want to).

' PARAMS
' SearchStyle: name of style we're looking for
' DestinationCollection: collection to add indices to
' ReturnMultiple: True finds all instances, False only finds first one.

' TODO: see if there is a better way than ByRef (that doesn't involve writing a
' MergeCollections function) so we can keep better track of what's in it.

Public Sub FindStyleIndex(ByVal SearchStyle As String, ByRef DestinationCollection _
  As Collection, Optional ReturnMultiple As Boolean = True)
  Dim lngFoundIndex As Long

' MacroHelpers.ParaIndex works with current *Selection*, so use Selection.Find
' Start from beginning of document
  Selection.HomeKey unit:=wdStory

  MacroHelpers.zz_clearFind
  With Selection.Find
    .Format = True
    .Style = SearchStyle
    .Forward = True
    .Wrap = wdFindStop
    .Execute

    Do While .Found = True
      lngFoundIndex = MacroHelpers.ParaIndex(UseEnd:=True)
      DestinationCollection.Add lngFoundIndex
      If ReturnMultiple = True Then
        .Execute
      Else
        Exit Do
      End If
    Loop
  End With
End Sub

