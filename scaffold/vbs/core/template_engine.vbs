' Template metadata and placeholder rendering helpers.
' Supports [TOKEN] and {{TOKEN}} placeholders with caller-supplied context dictionaries.

Function SC_ReadTextFile(filePath)
  SC_ReadTextFile = ""

  On Error Resume Next
  Dim stream
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile filePath
  If Err.Number = 0 Then SC_ReadTextFile = stream.ReadText
  stream.Close
  Set stream = Nothing
  If Err.Number = 0 Then
    On Error GoTo 0
    Exit Function
  End If
  Err.Clear
  On Error GoTo 0

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(filePath) Then Exit Function

  On Error Resume Next
  Dim ts
  Set ts = fso.OpenTextFile(filePath, 1)
  If Err.Number = 0 Then
    SC_ReadTextFile = ts.ReadAll
    ts.Close
  Else
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SC_NormalizeLineEndings(textValue)
  Dim result
  result = CStr(textValue)
  result = Replace(result, vbCrLf, vbLf)
  result = Replace(result, vbCr, vbLf)
  SC_NormalizeLineEndings = result
End Function

Function SC_ExtractPlaceholders(textValue)
  Dim placeholders
  Set placeholders = SC_CreateDictionary()

  Dim regex
  Set regex = CreateObject("VBScript.RegExp")
  regex.Global = True
  regex.IgnoreCase = False
  regex.Pattern = "\[([A-Za-z0-9_]+)\]|\{\{([A-Za-z0-9_]+)\}\}"

  Dim matches
  Set matches = regex.Execute(CStr(textValue))
  Dim match
  For Each match In matches
    Dim token
    token = ""
    If Len(match.SubMatches(0)) > 0 Then
      token = UCase(match.SubMatches(0))
    ElseIf Len(match.SubMatches(1)) > 0 Then
      token = UCase(match.SubMatches(1))
    End If
    If Len(token) > 0 Then
      If Not placeholders.Exists(token) Then placeholders.Add token, True
    End If
  Next

  Set SC_ExtractPlaceholders = placeholders
End Function

Function SC_ApplyContextToText(textValue, context)
  Dim result
  result = CStr(textValue)
  If Not IsObject(context) Then
    SC_ApplyContextToText = result
    Exit Function
  End If

  Dim key
  For Each key In context.Keys
    Dim value
    value = ""
    If Not IsNull(context(key)) Then value = CStr(context(key))
    result = Replace(result, "[" & CStr(key) & "]", value, 1, -1, vbTextCompare)
    result = Replace(result, "{{" & CStr(key) & "}}", value, 1, -1, vbTextCompare)
  Next

  SC_ApplyContextToText = result
End Function

Function SC_ParseTemplateMetadata(textValue)
  Dim metadata
  Set metadata = SC_CreateDictionary()

  Dim lines
  lines = Split(SC_NormalizeLineEndings(textValue), vbLf)
  Dim i
  For i = LBound(lines) To UBound(lines)
    Dim line
    line = Trim(CStr(lines(i)))
    If Left(line, 4) = "<!--" And InStr(1, line, ":", vbTextCompare) > 0 Then
      line = Trim(Mid(line, 5))
      If Right(line, 3) = "-->" Then line = Trim(Left(line, Len(line) - 3))
      Dim colonPos
      colonPos = InStr(1, line, ":", vbTextCompare)
      If colonPos > 1 Then
        Dim metaKey
        Dim metaValue
        metaKey = UCase(Trim(Left(line, colonPos - 1)))
        metaValue = Trim(Mid(line, colonPos + 1))
        If Len(metaKey) > 0 Then metadata(metaKey) = metaValue
      End If
    ElseIf Len(line) > 0 And Left(line, 4) <> "<!--" Then
      Exit For
    End If
  Next

  Set SC_ParseTemplateMetadata = metadata
End Function

Function SC_RenderTemplateString(templateText, context)
  Dim rendered
  Set rendered = SC_CreateDictionary()
  Set rendered("metadata") = SC_ParseTemplateMetadata(templateText)
  Set rendered("placeholders") = SC_ExtractPlaceholders(templateText)
  rendered("body") = SC_ApplyContextToText(templateText, context)
  Set SC_RenderTemplateString = rendered
End Function

Function SC_RenderTemplateFile(templatePath, context)
  Set SC_RenderTemplateFile = SC_RenderTemplateString(SC_ReadTextFile(templatePath), context)
End Function

