' Console display helpers for name resolution workflows

Function ParseLookupExpression(text)
  Dim result: Set result = CreateObject("Scripting.Dictionary")
  result("Field") = "auto"
  result("Value") = ""

  Dim trimmed: trimmed = Trim(text)
  If Len(trimmed) = 0 Then
    Set ParseLookupExpression = result
    Exit Function
  End If

  Dim separator: separator = InStr(trimmed, ":")
  If separator = 0 Then separator = InStr(trimmed, "=")
  If separator > 0 Then
    Dim potentialField: potentialField = Trim(Left(trimmed, separator - 1))
    Dim potentialValue: potentialValue = Trim(Mid(trimmed, separator + 1))
    If Len(potentialField) > 0 And Len(potentialValue) > 0 Then
      result("Field") = LCase(potentialField)
      result("Value") = potentialValue
      Set ParseLookupExpression = result
      Exit Function
    End If
  End If

  If InStr(trimmed, "@") > 0 Then
    result("Field") = "email"
  ElseIf InStr(trimmed, " ") > 0 Or InStr(trimmed, ",") > 0 Then
    result("Field") = "name"
  Else
    result("Field") = "auto"
  End If
  result("Value") = trimmed

  Set ParseLookupExpression = result
End Function

Function DescribeLookupField(fieldName, lookupValue)
  Dim lowered: lowered = LCase(Trim(fieldName))
  Select Case lowered
    Case "email", "mail", "address"
      DescribeLookupField = "email"
    Case "first", "given", "forename"
      DescribeLookupField = "given name"
    Case "surname", "last", "lastname", "family", "familyname"
      DescribeLookupField = "surname"
    Case "display", "name"
      DescribeLookupField = "display name"
    Case "fullname", "full", "full name", "full-name", "full_name"
      DescribeLookupField = "full name"
    Case Else
      Dim trimmedValue: trimmedValue = Trim(CStr(lookupValue))
      If InStr(trimmedValue, "@") > 0 Then
        DescribeLookupField = "email"
      ElseIf InStr(trimmedValue, " ") > 0 Or InStr(trimmedValue, ",") > 0 Then
        DescribeLookupField = "display name"
      Else
        DescribeLookupField = "auto-detected field"
      End If
  End Select
End Function

Sub DisplayResolution(rawInput)
  If Len(rawInput) = 0 Then Exit Sub

  Dim parsed: Set parsed = ParseLookupExpression(rawInput)
  Dim lookupField: lookupField = "auto"
  Dim lookupValue: lookupValue = ""
  If Not parsed Is Nothing Then
    If parsed.Exists("Field") Then lookupField = parsed("Field")
    If parsed.Exists("Value") Then lookupValue = parsed("Value")
  End If

  lookupValue = Trim(lookupValue)
  If Len(lookupValue) = 0 Then Exit Sub

  Dim fieldDescription: fieldDescription = DescribeLookupField(lookupField, lookupValue)

  WScript.Echo ""
  WScript.Echo "Resolving '" & lookupValue & "' by " & fieldDescription & "..."

  Dim info: Set info = ResolveNameLookup(lookupValue, lookupField)

  If LCase(lookupField) = "email" Then
    On Error Resume Next
    Call MaybeRefineNameUsingGivenName(lookupValue, info)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
  End If

  If info Is Nothing Then
    WScript.Echo "  No resolution information returned."
    Exit Sub
  End If

  Dim foundText: foundText = DescribeBoolean(info, "Found")
  WScript.Echo "  Resolution status: " & foundText
  Call DisplayResolutionSummary("Resolved name for " & fieldDescription & " '" & lookupValue & "':", info, "  ")
End Sub

Sub DisplayResolutionSummary(headerText, info, Optional indent)
  Dim linePrefix: linePrefix = ""
  If Not IsMissing(indent) Then linePrefix = indent
  If Len(linePrefix) = 0 Then linePrefix = ""

  Dim resolvedHeader: resolvedHeader = headerText
  If Len(resolvedHeader) = 0 Then resolvedHeader = "Resolved name:"

  If Len(linePrefix) > 0 Then
    WScript.Echo linePrefix & resolvedHeader
  Else
    WScript.Echo resolvedHeader
  End If

  Dim detailPrefix: detailPrefix = linePrefix & "  "
  Call DisplayResolutionFields(info, detailPrefix)
End Sub

Sub DisplayResolutionFields(info, Optional indent)
  Dim linePrefix: linePrefix = ""
  If Not IsMissing(indent) Then linePrefix = indent
  If Len(linePrefix) = 0 Then linePrefix = ""

  If info Is Nothing Then
    WScript.Echo linePrefix & "(No resolution information returned.)"
    Exit Sub
  End If

  WScript.Echo linePrefix & "Directory source: " & DescribeField(info, "DirectorySource", "(none)")
  WScript.Echo linePrefix & "Address lists checked: " & DescribeField(info, "DirectoryAttempted", "(none)")
  WScript.Echo linePrefix & "First name: " & DescribeField(info, "First", "(empty)")
  WScript.Echo linePrefix & "Surname: " & DescribeField(info, "Last", "(empty)")
  WScript.Echo linePrefix & "Email: " & DescribeField(info, "Email", "(empty)")
  WScript.Echo linePrefix & "Display name: " & DescribeField(info, "Full", "(empty)")
End Sub
