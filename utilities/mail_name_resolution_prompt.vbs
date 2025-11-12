' Console prompt helpers for mail name resolution workflows

Function PromptNameResolutionLine(promptText)
  Dim prompt: prompt = promptText
  If Len(prompt) = 0 Then prompt = "Enter lookup value: "

  WScript.StdOut.Write prompt

  Dim raw: raw = ""
  On Error Resume Next
  raw = WScript.StdIn.ReadLine()
  If Err.Number <> 0 Then
    Err.Clear
    raw = ""
  End If
  On Error GoTo 0

  If IsNull(raw) Then raw = ""
  If IsEmpty(raw) Then raw = ""

  PromptNameResolutionLine = Trim(CStr(raw))
End Function

Function ParseRecipientLookupList(rawText)
  Dim normalized: normalized = ""
  If Not IsNull(rawText) Then
    If Not IsEmpty(rawText) Then normalized = CStr(rawText)
  End If
  normalized = Trim(normalized)

  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  Dim dedupe: Set dedupe = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  dedupe.CompareMode = 1
  On Error GoTo 0

  If Len(normalized) = 0 Then
    ParseRecipientLookupList = Array()
    Exit Function
  End If

  normalized = Replace(normalized, vbCrLf, ";")
  normalized = Replace(normalized, vbCr, ";")
  normalized = Replace(normalized, vbLf, ";")
  normalized = Replace(normalized, ",", ";")

  Dim parts: parts = Split(normalized, ";")
  Dim i, value, key
  For i = 0 To UBound(parts)
    value = Trim(parts(i))
    If Len(value) > 0 Then
      key = LCase(value)
      If Not dedupe.Exists(key) Then
        dedupe.Add key, True
        list.Add value
      End If
    End If
  Next

  ParseRecipientLookupList = ArrayFromList(list)
End Function

Function PromptRecipientLookupValues()
  PromptRecipientLookupValues = PromptRecipientLookupValuesWithOptions( _
    "Enter target email addresses (use ';' to separate multiple addresses): ", _
    False, _
    "At least one lookup value is required.")
End Function

Function PromptRecipientLookupValuesWithOptions(promptText, allowEmpty, emptyWarning)
  Dim prompt: prompt = promptText
  If Len(prompt) = 0 Then
    prompt = "Enter lookup value: "
  End If

  Dim warning: warning = emptyWarning
  Dim loopAllowEmpty: loopAllowEmpty
  loopAllowEmpty = CBool(allowEmpty)

  Do
    Dim raw: raw = PromptNameResolutionLine(prompt)
    Dim parsed: parsed = ParseRecipientLookupList(raw)
    If HasRecipientLookupValues(parsed) Then
      PromptRecipientLookupValuesWithOptions = parsed
      Exit Function
    End If

    If loopAllowEmpty Then
      PromptRecipientLookupValuesWithOptions = parsed
      Exit Function
    End If

    If Len(warning) > 0 Then WScript.Echo warning
  Loop
End Function

Function HasRecipientLookupValues(values)
  HasRecipientLookupValues = False
  If Not IsArray(values) Then Exit Function

  On Error Resume Next
  Dim upperBound: upperBound = UBound(values)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If upperBound >= 0 Then HasRecipientLookupValues = True
End Function

Function ArrayFromList(list)
  Dim result()
  If list Is Nothing Then
    ArrayFromList = Array()
    Exit Function
  End If

  If list.Count = 0 Then
    ArrayFromList = Array()
    Exit Function
  End If

  ReDim result(list.Count - 1)

  Dim i
  For i = 0 To list.Count - 1
    result(i) = list(i)
  Next

  ArrayFromList = result
End Function
