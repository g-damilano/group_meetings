' Helper routines for collection handling and template ordering

Function CreateCaseInsensitiveDictionary()
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  dict.CompareMode = 1
  On Error GoTo 0
  Set CreateCaseInsensitiveDictionary = dict
End Function

Function GetDictionaryValue(dict, key)
  If dict Is Nothing Then
    GetDictionaryValue = ""
    Exit Function
  End If

  If dict.Exists(LCase(key)) Then
    GetDictionaryValue = dict(LCase(key))
  ElseIf dict.Exists(UCase(key)) Then
    GetDictionaryValue = dict(UCase(key))
  Else
    GetDictionaryValue = ""
  End If
End Function

Function CloneDictionary(source)
  Dim clone: Set clone = CreateCaseInsensitiveDictionary()
  If Not source Is Nothing Then
    Dim key
    For Each key In source.Keys
      clone(key) = source(key)
    Next
  End If
  Set CloneDictionary = clone
End Function

Function ArrayFromArrayList(list)
  Dim arr()
  ReDim arr(list.Count - 1)
  Dim i
  For i = 0 To list.Count - 1
    arr(i) = list(i)
  Next
  ArrayFromArrayList = arr
End Function

Function ContainsInArrayList(list, value)
  Dim i
  For i = 0 To list.Count - 1
    If LCase(list(i)) = LCase(value) Then
      ContainsInArrayList = True
      Exit Function
    End If
  Next
  ContainsInArrayList = False
End Function

Function Capitalize(text)
  If Len(text) = 0 Then
    Capitalize = text
  ElseIf Left(text, 1) = "(" Then
    Capitalize = text
  Else
    Capitalize = UCase(Left(text, 1)) & Mid(text, 2)
  End If
End Function

Function TemplateSortValue(tmpl)
  Dim code: code = tmpl("Code")
  If Len(code) = 0 Then code = tmpl("FileName")
  TemplateSortValue = LCase(code)
End Function

Function SortTemplateList(list)
  Dim arr()
  ReDim arr(list.Count - 1)
  Dim i
  For i = 0 To list.Count - 1
    Set arr(i) = list(i)
  Next

  Dim j, keyTemplate
  For i = 1 To UBound(arr)
    Set keyTemplate = arr(i)
    j = i - 1
    Do
      If Not (j >= 0 And TemplateSortValue(arr(j)) > TemplateSortValue(keyTemplate)) Then Exit Do
      Set arr(j + 1) = arr(j)
      j = j - 1
    Loop
    Set arr(j + 1) = keyTemplate
  Next

  Dim ordered: Set ordered = CreateObject("System.Collections.ArrayList")
  For i = 0 To UBound(arr)
    ordered.Add arr(i)
  Next
  Set SortTemplateList = ordered
End Function

Function BuildOrderedDictionary(list)
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  Dim ordered: Set ordered = SortTemplateList(list)
  Dim idx
  For idx = 0 To ordered.Count - 1
    Dim tmpl: Set tmpl = ordered(idx)
    dict.Add CStr(idx + 1), tmpl
  Next
  Set BuildOrderedDictionary = dict
End Function

Function ParseNumberList(text)
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  dict.CompareMode = 1
  Dim cleaned: cleaned = Replace(Replace(text, ",", " "), ";", " ")
  Dim parts: parts = Split(cleaned, " ")
  Dim part
  For Each part In parts
    part = Trim(part)
    If Len(part) > 0 And IsNumeric(part) Then
      dict(CStr(CInt(part))) = True
    End If
  Next
  Set ParseNumberList = dict
End Function
