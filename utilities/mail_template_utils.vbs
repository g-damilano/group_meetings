' Template parsing and placeholder helpers

Function ReadFileSmart(filePath)
  On Error Resume Next
  Dim stream: Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile filePath
  Dim content: content = stream.ReadText
  stream.Close
  Set stream = Nothing
  If Err.Number = 0 Then
    ReadFileSmart = content
    Exit Function
  End If
  Err.Clear
  On Error GoTo 0

  Dim ts
  On Error Resume Next
  Set ts = fso.OpenTextFile(filePath, 1)
  If Err.Number = 0 Then
    ReadFileSmart = ts.ReadAll
    ts.Close
  Else
    ReadFileSmart = ""
  End If
  On Error GoTo 0
End Function

Function NormalizeLineEndings(text)
  text = Replace(text, vbCrLf, vbLf)
  text = Replace(text, vbCr, vbLf)
  NormalizeLineEndings = text
End Function

Function TrimLeadingBlankLines(text)
  Dim result: result = text
  Do
    If Not (Len(result) > 0 And (Left(result, 1) = vbCr Or Left(result, 1) = vbLf)) Then Exit Do
    result = Mid(result, 2)
  Loop
  TrimLeadingBlankLines = result
End Function

Function BuildContextFromMeta(meta)
  Dim context: Set context = CloneDictionary(linkMapContext)
  If context Is Nothing Then Set context = CreateCaseInsensitiveDictionary()
  If Not meta Is Nothing Then
    Dim key
    For Each key In meta.Keys
      context(UCase(key)) = meta(key)
    Next
  End If
  Set BuildContextFromMeta = context
End Function

Sub PromptForEmptyTags(meta, templateName)
  If meta Is Nothing Then Exit Sub
  Dim key
  For Each key In meta.Keys
    If Len(Trim(meta(key))) = 0 Then
      Dim prompt: prompt = "Provide value for '" & UCase(key) & "' in template '" & templateName & "': "
      meta(key) = PromptLine(prompt)
    End If
  Next
End Sub

Sub EnsurePlaceholderValues(textItems, context, templateName)
  If context Is Nothing Then Exit Sub
  Dim tokens: tokens = ExtractPlaceholders(textItems)
  If IsArray(tokens) Then
    Dim upper
    On Error Resume Next
    upper = UBound(tokens)
    If Err.Number <> 0 Then
      Err.Clear
      upper = -1
    End If
    On Error GoTo 0
    Dim i
    For i = 0 To upper
      Dim token: token = tokens(i)
      If Len(token) > 0 Then
        Dim lookup: lookup = UCase(token)
        If context.Exists(lookup) Then
          If Len(CStr(context(lookup))) = 0 Then
            context(lookup) = PromptLine("Enter value for [" & lookup & "] in '" & templateName & "': ")
          End If
        Else
          LogEvent "Placeholder [" & lookup & "] has no known value for template '" & templateName & "'"
        End If
      End If
    Next
  End If
End Sub

Function ExtractPlaceholders(textItems)
  Dim collector: Set collector = CreateCaseInsensitiveDictionary()
  Dim re: Set re = CreateObject("VBScript.RegExp")
  re.Global = True
  re.IgnoreCase = False
  re.Pattern = "\[([A-Z0-9_]+)\]"

  If Not IsArray(textItems) Then
    ExtractPlaceholders = Array()
    Exit Function
  End If

  Dim i, text, matches, m
  For i = 0 To UBound(textItems)
    text = textItems(i)
    If Not IsEmpty(text) And Not IsNull(text) And Len(text) > 0 Then
      Set matches = re.Execute(CStr(text))
      For Each m In matches
        collector(UCase(m.SubMatches(0))) = True
      Next
    End If
  Next

  Dim arr
  If collector.Count = 0 Then
    arr = Array()
  Else
    Dim list: Set list = CreateObject("System.Collections.ArrayList")
    Dim key
    For Each key In collector.Keys
      list.Add key
    Next
    arr = ArrayFromArrayList(list)
  End If
  ExtractPlaceholders = arr
End Function

Function ApplyContextToText(text, context)
  Dim result: result = text
  If IsNull(result) Or IsEmpty(result) Then result = ""
  If context Is Nothing Then
    ApplyContextToText = result
    Exit Function
  End If

  Dim key, value
  For Each key In context.Keys
    value = context(key)
    If IsNull(value) Then value = ""
    result = Replace(result, "[" & key & "]", CStr(value))
  Next
  ApplyContextToText = result
End Function

Function MergeRecipients(baseList, extraArray)
  Dim seen: Set seen = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  seen.CompareMode = 1
  On Error GoTo 0
  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  AppendAddressesFromString list, seen, baseList
  Dim i
  If IsArray(extraArray) Then
    For i = LBound(extraArray) To UBound(extraArray)
      AppendAddress list, seen, extraArray(i)
    Next
  End If
  If list.Count = 0 Then
    MergeRecipients = ""
  Else
    MergeRecipients = Join(ArrayFromArrayList(list), "; ")
  End If
End Function

Sub AppendAddressesFromString(list, seen, text)
  If Len(Trim(text)) = 0 Then Exit Sub
  Dim normalized: normalized = Replace(text, ",", ";")
  Dim parts: parts = Split(normalized, ";")
  Dim i
  For i = 0 To UBound(parts)
    AppendAddress list, seen, parts(i)
  Next
End Sub

Sub AppendAddress(list, seen, address)
  Dim trimmed: trimmed = Trim(address)
  If Len(trimmed) = 0 Then Exit Sub
  Dim key: key = LCase(trimmed)
  If Not seen.Exists(key) Then
    seen.Add key, True
    list.Add trimmed
  End If
End Sub

Function SafeText(value)
  If IsObject(value) Then
    SafeText = ""
    Exit Function
  End If

  If IsNull(value) Or IsEmpty(value) Then
    SafeText = ""
  Else
    On Error Resume Next
    SafeText = CStr(value)
    If Err.Number <> 0 Then
      Err.Clear
      SafeText = ""
    End If
    On Error GoTo 0
  End If
End Function
