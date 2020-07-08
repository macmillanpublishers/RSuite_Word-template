Attribute VB_Name = "Sandbox"
Private Sub testredirects()
    'try to download the file from Public Confluence page
    Dim WinHttpReq As Object
    Dim oStream As Object
    Dim myURL As String
    myURL = "https://github.com/macmillanpublishers/Word-template/raw/v2.9.5/README.md"
    
    myURL = "https://github.com/johnwangel/HCML/blob/master/HCML.dotm?raw=true"
    
    Dim destFile As String
    destFile = "C:\Users\johnatkins\Desktop\HCML.dotm"

   ' Set WinHttpReq = CreateObject("MSXML2.XMLHTTP.3.0")
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
'    WinHttpReq.Option(6) = True
    
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
      
    'DebugPrint "Http status for " & FileName & ": " & WinHttpReq.Status
    If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
  
      'if connection OK, download file to temp dir
      myURL = WinHttpReq.responseBody
      Set oStream = CreateObject("ADODB.Stream")
      oStream.Open
      oStream.Type = 1
      oStream.Write WinHttpReq.responseBody
      oStream.SaveToFile destFile, 2 ' 1 = no overwrite, 2 = overwrite
      oStream.Close
      
      Set oStream = Nothing
      Set WinHttpReq = Nothing

    Else
      Debug.Print WinHttpReq.Status
    End If
  
End Sub

Private Sub testDictSort()
  Dim dict1 As Dictionary
  Set dict1 = New Dictionary
  
  dict1.Add 100, "Last"
  dict1.Add 1, "First"
  dict1.Add 24, "Third"
  dict1.Add 2, "Second"
  
  Dim arrTemp() As Variant
  arrTemp = dict1.Keys
  
  Dim a As Long
  For a = LBound(arrTemp) To UBound(arrTemp)
    Debug.Print arrTemp(a)
  Next a
  
  WordBasic.SortArray arrTemp()

'  Dim B As Long
'  For B = LBound(arrTemp) To UBound(arrTemp)
'    Debug.Print arrTemp(B)
'  Next B

  Dim varKey As Variant
  For Each varKey In arrTemp
    Debug.Print varKey & ": " & dict1(varKey)
  Next varKey

End Sub


Private Sub CleanStringTest()
' https://msdn.microsoft.com/en-us/library/office/aa171835(v=office.11).aspx
  Dim strDoc As String
  strDoc = ActiveDocument.Range.Text
  strDoc = CleanString(strDoc)
  ActiveDocument.Range.Text = strDoc
  MacroHelpers.zz_clearFind
  With ActiveDocument.Range.Find
    .Text = "^13{2,}"
    .Replacement.Text = "^p"
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  End With
  With ActiveDocument.Range.Find
    .Text = "^32{2,}"
    .Replacement.Text = " "
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  End With
End Sub



