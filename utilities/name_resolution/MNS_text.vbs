' Text processing utilities for mail name resolution

Function CollectNameTokens(text)
  Dim list: Set list = CreateObject("System.Collections.ArrayList")

  If IsObject(text) Then
    Set CollectNameTokens = list
    Exit Function
  End If

  Dim raw: raw = ""
  If Not IsNull(text) Then
    If Not IsEmpty(text) Then raw = CStr(text)
  End If

  Dim trimmed: trimmed = Trim(raw)
  If Len(trimmed) = 0 Then
    Set CollectNameTokens = list
    Exit Function
  End If

  Dim segments: segments = Split(trimmed, " ")
  Dim i, token
  For i = LBound(segments) To UBound(segments)
    token = Trim(segments(i))
    If Len(token) > 0 Then list.Add token
  Next

  Set CollectNameTokens = list
End Function

Function CollectFlexibleNameTokens(text)
  Dim raw: raw = ""

  If IsObject(text) Then
    Set CollectFlexibleNameTokens = CollectNameTokens("")
    Exit Function
  End If

  If Not IsNull(text) Then
    If Not IsEmpty(text) Then raw = CStr(text)
  End If

  Dim normalized: normalized = Trim(raw)
  If Len(normalized) = 0 Then
    Set CollectFlexibleNameTokens = CollectNameTokens(normalized)
    Exit Function
  End If

  normalized = Replace(normalized, ",", " ")
  normalized = Replace(normalized, ".", " ")
  normalized = Replace(normalized, "_", " ")
  normalized = Replace(normalized, "-", " ")
  normalized = Replace(normalized, "/", " ")
  normalized = Replace(normalized, "\", " ")
  normalized = Replace(normalized, vbTab, " ")
  normalized = Replace(normalized, Chr(160), " ")

  Do While InStr(normalized, "  ") > 0
    normalized = Replace(normalized, "  ", " ")
  Loop

  Set CollectFlexibleNameTokens = CollectNameTokens(normalized)
End Function

Function ExtractDisplayNameAndEmail(text)
  Dim result: Set result = CreateObject("Scripting.Dictionary")

  result("Display") = ""
  result("Email") = ""
  result("HasAngleFormat") = False

  Dim raw: raw = Trim(SafeText(text))
  If Len(raw) = 0 Then
    result("Display") = ""
    Set ExtractDisplayNameAndEmail = result
    Exit Function
  End If

  result("Display") = raw

  Dim startPos: startPos = InStr(raw, "<")
  Dim endPos: endPos = InStrRev(raw, ">")

  If startPos > 0 And endPos > startPos Then
    Dim between: between = Mid(raw, startPos + 1, endPos - startPos - 1)
    Dim emailCandidate: emailCandidate = NormalizeEmailText(between)

    If Len(emailCandidate) > 0 Then
      result("Email") = emailCandidate
      result("HasAngleFormat") = True

      Dim beforePart: beforePart = Trim(Left(raw, startPos - 1))
      Dim afterPart: afterPart = Trim(Mid(raw, endPos + 1))
      Dim mergedDisplay: mergedDisplay = beforePart

      If Len(mergedDisplay) = 0 Then
        mergedDisplay = afterPart
      ElseIf Len(afterPart) > 0 Then
        mergedDisplay = Trim(mergedDisplay & " " & afterPart)
      End If

      result("Display") = Trim(mergedDisplay)
    End If
  End If

  Set ExtractDisplayNameAndEmail = result
End Function

Function JoinTokens(list, delimiter)
  Dim result: result = ""

  If list Is Nothing Then
    JoinTokens = ""
    Exit Function
  End If

  Dim i, value
  For i = 0 To list.Count - 1
    value = Trim(SafeText(list(i)))
    If Len(value) > 0 Then
      If Len(result) = 0 Then
        result = value
      Else
        result = result & delimiter & value
      End If
    End If
  Next

  JoinTokens = Trim(result)
End Function

Function BuildReversedTokenDisplay(list)
  BuildReversedTokenDisplay = ""
  If list Is Nothing Then Exit Function
  If list.Count < 2 Then Exit Function

  Dim lastPart: lastPart = Trim(SafeText(list(list.Count - 1)))
  Dim firstPart: firstPart = Trim(SafeText(list(0)))
  If Len(lastPart) = 0 Or Len(firstPart) = 0 Then Exit Function

  Dim middle: middle = ""
  If list.Count > 2 Then
    Dim i, token
    For i = 1 To list.Count - 2
      token = Trim(SafeText(list(i)))
      If Len(token) > 0 Then
        If Len(middle) = 0 Then
          middle = token
        Else
          middle = middle & " " & token
        End If
      End If
    Next
  End If

  If Len(middle) > 0 Then
    BuildReversedTokenDisplay = Trim(lastPart & ", " & firstPart & " " & middle)
  Else
    BuildReversedTokenDisplay = Trim(lastPart & ", " & firstPart)
  End If
End Function

Sub AddCandidateValue(dict, value)
  If dict Is Nothing Then Exit Sub
  Dim text: text = Trim(SafeText(value))
  If Len(text) = 0 Then Exit Sub
  On Error Resume Next
  If Not dict.Exists(text) Then dict.Add text, True
  On Error GoTo 0
End Sub

Function ParseFullNameComponents(fullName)
  Dim result: Set result = CreateObject("Scripting.Dictionary")

  Dim raw: raw = ""
  If Not IsObject(fullName) Then
    If Not IsNull(fullName) Then
      If Not IsEmpty(fullName) Then raw = CStr(fullName)
    End If
  End If

  Dim trimmed: trimmed = Trim(raw)
  result("Display") = trimmed
  result("First") = ""
  result("Last") = ""
  result("AlternateFirst") = ""
  result("AlternateLast") = ""
  result("Normalized") = trimmed
  result("ReverseDisplay") = ""

  If Len(trimmed) = 0 Then
    Set ParseFullNameComponents = result
    Exit Function
  End If

  Dim firstName: firstName = ""
  Dim lastName: lastName = ""
  Dim altFirst: altFirst = ""
  Dim altLast: altLast = ""
  Dim tokens: Set tokens = Nothing

  If InStr(trimmed, ",") > 0 Then
    Dim parts: parts = Split(trimmed, ",")
    Dim leftPart: leftPart = Trim(parts(0))
    Dim rightPart: rightPart = ""
    If UBound(parts) >= 1 Then rightPart = Trim(parts(1))
    Set tokens = CollectNameTokens(rightPart)
    lastName = leftPart
    If Not tokens Is Nothing Then
      If tokens.Count > 0 Then firstName = tokens(0)
    End If
  Else
    Set tokens = CollectNameTokens(trimmed)
    If Not tokens Is Nothing Then
      If tokens.Count = 1 Then
        firstName = tokens(0)
      ElseIf tokens.Count >= 2 Then
        firstName = tokens(0)
        lastName = tokens(tokens.Count - 1)
        altFirst = tokens(tokens.Count - 1)
        altLast = tokens(0)
      End If
    End If
  End If

  If Len(lastName) = 0 And Not tokens Is Nothing Then
    If tokens.Count >= 2 Then lastName = tokens(tokens.Count - 1)
  End If

  If Len(firstName) = 0 And Not tokens Is Nothing Then
    If tokens.Count > 0 Then firstName = tokens(0)
  End If

  result("First") = firstName
  result("Last") = lastName
  result("AlternateFirst") = altFirst
  result("AlternateLast") = altLast

  Dim normalized: normalized = Trim(firstName & " " & lastName)
  If Len(normalized) > 0 Then
    result("Normalized") = normalized
  End If

  If Len(lastName) > 0 And Len(firstName) > 0 Then
    result("ReverseDisplay") = Trim(lastName & ", " & firstName)
  Else
    result("ReverseDisplay") = trimmed
  End If

  Set ParseFullNameComponents = result
End Function

Function NormalizeEmailText(value)
  Dim trimmed: trimmed = Trim(SafeText(value))
  If Len(trimmed) = 0 Then
    NormalizeEmailText = ""
    Exit Function
  End If

  Dim lowered: lowered = SafeLower(trimmed)
  If Left(lowered, 6) = "smtp:" Then
    trimmed = Mid(trimmed, 6)
  ElseIf Left(lowered, 7) = "mailto:" Then
    trimmed = Mid(trimmed, 8)
  ElseIf Left(lowered, 4) = "sip:" Then
    trimmed = Mid(trimmed, 5)
  End If

  NormalizeEmailText = Trim(trimmed)
End Function

Sub ApplyEmailCandidate(info, candidate)
  If info Is Nothing Then Exit Sub

  Dim trimmed: trimmed = NormalizeEmailText(candidate)
  If Len(trimmed) = 0 Then Exit Sub

  Dim existing: existing = ""
  If info.Exists("Email") Then existing = NormalizeEmailText(info("Email"))

  If Len(existing) = 0 Then
    If InStr(trimmed, "@") > 0 Then info("Email") = trimmed
    Exit Sub
  End If

  If InStr(existing, "@") = 0 And InStr(trimmed, "@") > 0 Then
    info("Email") = trimmed
    Exit Sub
  End If

  If SafeLower(existing) = SafeLower(trimmed) Then Exit Sub

  If IsValidEmail(trimmed) And Not IsValidEmail(existing) Then
    info("Email") = trimmed
  End If
End Sub

Sub ApplyFullNameCandidate(info, candidate, emailAddress)
  If info Is Nothing Then Exit Sub

  Dim proposed: proposed = Trim(SafeText(candidate))
  If Len(proposed) = 0 Then Exit Sub

  Dim existing: existing = ""
  If info.Exists("Full") Then existing = Trim(SafeText(info("Full")))

  If Len(existing) = 0 Then
    info("Full") = proposed
    Exit Sub
  End If

  Dim normalizedEmail: normalizedEmail = NormalizeEmailText(emailAddress)
  If Len(normalizedEmail) > 0 Then
    If SafeLower(existing) = SafeLower(normalizedEmail) Then
      info("Full") = proposed
      Exit Sub
    End If
  End If

  Dim existingLooksLikeEmail: existingLooksLikeEmail = InStr(existing, "@") > 0
  Dim proposedLooksLikeEmail: proposedLooksLikeEmail = InStr(proposed, "@") > 0

  If existingLooksLikeEmail And Not proposedLooksLikeEmail Then
    info("Full") = proposed
    Exit Sub
  End If

  Dim existingHasNameSeparators
  existingHasNameSeparators = (InStr(existing, " ") > 0) Or (InStr(existing, ",") > 0)
  Dim proposedHasNameSeparators
  proposedHasNameSeparators = (InStr(proposed, " ") > 0) Or (InStr(proposed, ",") > 0)

  If Not existingHasNameSeparators And proposedHasNameSeparators Then
    info("Full") = proposed
    Exit Sub
  End If

  If Not proposedLooksLikeEmail And Len(proposed) > Len(existing) Then
    info("Full") = proposed
  End If
End Sub

Sub ApplyFirstNameCandidate(info, candidate, aliasValue, emailAddress)
  If info Is Nothing Then Exit Sub
  Dim proposed: proposed = SafeText(candidate)
  If Len(proposed) = 0 Then Exit Sub

  Dim existing: existing = ""
  If info.Exists("First") Then existing = SafeText(info("First"))

  If ShouldReplaceFirstName(existing, proposed, aliasValue, emailAddress) Or Len(existing) = 0 Then
    info("First") = candidate
  End If
End Sub

Function ShouldReplaceFirstName(existingValue, newValue, aliasValue, emailAddress)
  Dim existingLower: existingLower = SafeLower(existingValue)
  Dim newLower: newLower = SafeLower(newValue)

  If Len(newLower) = 0 Then
    ShouldReplaceFirstName = False
    Exit Function
  End If

  If Len(existingLower) = 0 Then
    ShouldReplaceFirstName = True
    Exit Function
  End If

  If existingLower = newLower Then
    ShouldReplaceFirstName = False
    Exit Function
  End If

  Dim aliasLower: aliasLower = SafeLower(aliasValue)
  If Len(aliasLower) > 0 Then
    If existingLower = aliasLower Then
      ShouldReplaceFirstName = True
      Exit Function
    End If
  End If

  Dim localLower: localLower = SafeLower(ExtractEmailLocalPart(emailAddress))
  If Len(localLower) > 0 Then
    If existingLower = localLower Then
      ShouldReplaceFirstName = True
      Exit Function
    End If
  End If

  ShouldReplaceFirstName = False
End Function

Function ExtractEmailLocalPart(emailAddress)
  ExtractEmailLocalPart = ""
  If Len(emailAddress) = 0 Then Exit Function
  Dim atPos: atPos = InStr(emailAddress, "@")
  If atPos > 0 Then
    ExtractEmailLocalPart = Left(emailAddress, atPos - 1)
  Else
    ExtractEmailLocalPart = emailAddress
  End If
End Function

Function SafeLower(value)
  SafeLower = ""
  If IsNull(value) Then Exit Function
  If IsEmpty(value) Then Exit Function
  Dim text: text = CStr(value)
  If Len(text) = 0 Then Exit Function
  SafeLower = LCase(text)
End Function

Function GetAccessorString(accessor, propertyUris)
  GetAccessorString = ""
  If accessor Is Nothing Then Exit Function
  If Not IsArray(propertyUris) Then Exit Function

  Dim i, candidate
  On Error Resume Next
  For i = LBound(propertyUris) To UBound(propertyUris)
    candidate = accessor.GetProperty(propertyUris(i))
    If Err.Number = 0 Then
      If Not IsNull(candidate) And Not IsEmpty(candidate) Then
        candidate = CStr(candidate)
        If Len(candidate) > 0 Then
          GetAccessorString = candidate
          Exit For
        End If
      End If
    Else
      Err.Clear
    End If
  Next
  On Error GoTo 0
End Function
