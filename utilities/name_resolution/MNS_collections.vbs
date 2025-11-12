' Collection helpers for name resolution diagnostics

Function CreateArrayList()
  Set CreateArrayList = CreateObject("System.Collections.ArrayList")
End Function

Function EnsureArrayList(list)
  If list Is Nothing Then
    Set list = CreateArrayList()
  End If
  Set EnsureArrayList = list
End Function

Function CreateTextDictionary()
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  dict.CompareMode = 1
  On Error GoTo 0
  Set CreateTextDictionary = dict
End Function

Function GetCollectionCount(list)
  GetCollectionCount = 0
  If list Is Nothing Then Exit Function
  On Error Resume Next
  GetCollectionCount = list.Count
  If Err.Number <> 0 Then
    Err.Clear
    GetCollectionCount = 0
  End If
  On Error GoTo 0
End Function

Function GetDictionaryCount(dict)
  GetDictionaryCount = 0
  If dict Is Nothing Then Exit Function
  On Error Resume Next
  GetDictionaryCount = dict.Count
  If Err.Number <> 0 Then
    Err.Clear
    GetDictionaryCount = 0
  End If
  On Error GoTo 0
End Function

