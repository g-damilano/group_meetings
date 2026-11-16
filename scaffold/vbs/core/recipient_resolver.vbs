' Recipient parsing and resolution helpers.
' The core resolver is deliberately conservative: direct SMTP parsing and caller-supplied
' directory maps are supported offline; live GAL/O365 lookup belongs behind a separate adapter.

Function SC_CleanDisplayName(value)
  Dim text
  text = Trim(CStr(value))
  If Len(text) >= 2 Then
    If (Left(text, 1) = """" And Right(text, 1) = """") Or _
       (Left(text, 1) = "'" And Right(text, 1) = "'") Then
      text = Mid(text, 2, Len(text) - 2)
    End If
  End If
  SC_CleanDisplayName = Trim(text)
End Function

Function SC_CollapseWhitespace(value)
  Dim text
  text = Trim(CStr(value))
  Do While InStr(1, text, "  ", vbTextCompare) > 0
    text = Replace(text, "  ", " ")
  Loop
  SC_CollapseWhitespace = text
End Function

Function SC_IsNameTitle(token)
  Select Case LCase(Replace(Trim(CStr(token)), ".", ""))
    Case "dr", "prof", "professor", "mr", "mrs", "ms", "miss", "mx"
      SC_IsNameTitle = True
    Case Else
      SC_IsNameTitle = False
  End Select
End Function

Function SC_IsFamilyNameParticle(token)
  Select Case LCase(Trim(CStr(token)))
    Case "da", "de", "del", "der", "di", "la", "le", "van", "von"
      SC_IsFamilyNameParticle = True
    Case Else
      SC_IsFamilyNameParticle = False
  End Select
End Function

Function SC_RemoveLeadingNameTitle(value)
  Dim text
  text = SC_CollapseWhitespace(value)
  Dim parts
  parts = Split(text, " ")
  If UBound(parts) >= 1 Then
    If SC_IsNameTitle(parts(0)) Then
      Dim i
      Dim remainder
      remainder = ""
      For i = 1 To UBound(parts)
        If Len(remainder) > 0 Then remainder = remainder & " "
        remainder = remainder & parts(i)
      Next
      text = remainder
    End If
  End If
  SC_RemoveLeadingNameTitle = text
End Function

Function SC_FirstGivenName(value)
  Dim text
  text = SC_RemoveLeadingNameTitle(value)
  If Len(text) = 0 Then
    SC_FirstGivenName = ""
  Else
    SC_FirstGivenName = Split(text, " ")(0)
  End If
End Function

Function SC_FamilyNameFromGivenFamily(value)
  Dim text
  text = SC_RemoveLeadingNameTitle(value)
  Dim parts
  parts = Split(text, " ")
  If UBound(parts) < 1 Then
    SC_FamilyNameFromGivenFamily = ""
    Exit Function
  End If

  Dim startIndex
  startIndex = UBound(parts)
  Do While startIndex > 0 And SC_IsFamilyNameParticle(parts(startIndex - 1))
    startIndex = startIndex - 1
  Loop

  Dim i
  Dim surname
  surname = ""
  For i = startIndex To UBound(parts)
    If Len(surname) > 0 Then surname = surname & " "
    surname = surname & parts(i)
  Next
  SC_FamilyNameFromGivenFamily = surname
End Function

Function SC_ParseDisplayNameParts(value)
  Dim parts
  Set parts = SC_CreateDictionary()
  parts("first_name") = ""
  parts("last_name") = ""
  parts("surname") = ""
  parts("full_name") = ""
  parts("name_parse_status") = "not_attempted"
  parts("name_parse_source") = ""
  parts("name_parse_order") = ""

  Dim text
  text = SC_CleanDisplayName(value)
  If InStr(1, text, "<", vbTextCompare) > 1 Then text = SC_DisplayNameFromRecipientText(text)
  text = SC_CollapseWhitespace(text)
  If Len(text) = 0 Or SC_IsValidEmailAddress(text) Then
    Set SC_ParseDisplayNameParts = parts
    Exit Function
  End If

  Dim commaPos
  commaPos = InStr(1, text, ",", vbTextCompare)
  If commaPos > 0 Then
    Dim familyName
    Dim givenName
    familyName = SC_CollapseWhitespace(Left(text, commaPos - 1))
    givenName = SC_FirstGivenName(Mid(text, commaPos + 1))
    parts("first_name") = givenName
    parts("last_name") = familyName
    parts("surname") = familyName
    If Len(givenName) > 0 And Len(familyName) > 0 Then
      parts("full_name") = givenName & " " & familyName
      parts("name_parse_status") = "parsed"
      parts("name_parse_order") = "family-given"
    Else
      parts("full_name") = text
      parts("name_parse_status") = "partial"
    End If
    Set SC_ParseDisplayNameParts = parts
    Exit Function
  End If

  Dim firstName
  Dim lastName
  firstName = SC_FirstGivenName(text)
  lastName = SC_FamilyNameFromGivenFamily(text)
  parts("first_name") = firstName
  parts("last_name") = lastName
  parts("surname") = lastName
  parts("full_name") = text
  If Len(firstName) > 0 And Len(lastName) > 0 Then
    parts("name_parse_status") = "parsed"
    parts("name_parse_order") = "given-family"
  ElseIf Len(firstName) > 0 Then
    parts("name_parse_status") = "partial"
    parts("name_parse_order") = "single-token"
  End If
  Set SC_ParseDisplayNameParts = parts
End Function

Sub SC_ApplyRecipientNameParts(spec, displayName, sourceName)
  If Not IsObject(spec) Then Exit Sub
  Dim parts
  Set parts = SC_ParseDisplayNameParts(displayName)

  If Len(Trim(CStr(spec("first_name")))) = 0 Then spec("first_name") = parts("first_name")
  If Len(Trim(CStr(spec("last_name")))) = 0 Then spec("last_name") = parts("last_name")
  If Len(Trim(CStr(spec("surname")))) = 0 Then spec("surname") = parts("surname")
  If Len(Trim(CStr(spec("full_name")))) = 0 Then spec("full_name") = parts("full_name")
  If Len(Trim(CStr(spec("name_parse_status")))) = 0 Or spec("name_parse_status") = "not_attempted" Then spec("name_parse_status") = parts("name_parse_status")
  If Len(Trim(CStr(spec("name_parse_source")))) = 0 Then spec("name_parse_source") = CStr(sourceName)
  If Len(Trim(CStr(spec("name_parse_order")))) = 0 Then spec("name_parse_order") = parts("name_parse_order")
End Sub

Function SC_NormalizeEmailAddress(value)
  Dim text
  text = Trim(CStr(value))
  If LCase(Left(text, 7)) = "mailto:" Then text = Mid(text, 8)

  Dim ltPos
  Dim gtPos
  ltPos = InStr(1, text, "<", vbTextCompare)
  gtPos = InStr(1, text, ">", vbTextCompare)
  If ltPos > 0 And gtPos > ltPos Then
    text = Mid(text, ltPos + 1, gtPos - ltPos - 1)
  End If

  SC_NormalizeEmailAddress = LCase(Trim(text))
End Function

Function SC_DisplayNameFromRecipientText(value)
  Dim text
  text = Trim(CStr(value))
  Dim ltPos
  ltPos = InStr(1, text, "<", vbTextCompare)
  If ltPos > 1 Then
    SC_DisplayNameFromRecipientText = SC_CleanDisplayName(Left(text, ltPos - 1))
  Else
    SC_DisplayNameFromRecipientText = ""
  End If
End Function

Function SC_IsValidEmailAddress(value)
  Dim candidate
  candidate = SC_NormalizeEmailAddress(value)
  If Len(candidate) = 0 Then
    SC_IsValidEmailAddress = False
    Exit Function
  End If

  Dim regex
  Set regex = Nothing
  On Error Resume Next
  Set regex = CreateObject("VBScript.RegExp")
  If Err.Number <> 0 Then
    Err.Clear
    Set regex = Nothing
  End If
  On Error GoTo 0

  If regex Is Nothing Then
    Dim atPos
    Dim dotPos
    atPos = InStr(1, candidate, "@", vbTextCompare)
    dotPos = InStrRev(candidate, ".")
    SC_IsValidEmailAddress = (atPos > 1 And dotPos > atPos + 1 And dotPos < Len(candidate))
    Exit Function
  End If

  regex.Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$"
  regex.IgnoreCase = True
  regex.Global = False
  SC_IsValidEmailAddress = regex.Test(candidate)
End Function

Function SC_RecipientLookupKey(value)
  Dim text
  text = LCase(SC_CleanDisplayName(value))
  If InStr(1, text, "<", vbTextCompare) > 1 Then
    text = SC_DisplayNameFromRecipientText(text)
  End If

  Dim i
  Dim ch
  Dim result
  result = ""
  For i = 1 To Len(text)
    ch = Mid(text, i, 1)
    If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
      result = result & ch
    End If
  Next
  SC_RecipientLookupKey = result
End Function

Function SC_DirectoryLookup(directory, lookupValue)
  SC_DirectoryLookup = ""
  If Not IsObject(directory) Then Exit Function

  Dim rawKey
  rawKey = Trim(CStr(lookupValue))
  Dim lowerKey
  lowerKey = LCase(rawKey)
  Dim upperKey
  upperKey = UCase(rawKey)

  On Error Resume Next
  If directory.Exists(rawKey) Then
    SC_DirectoryLookup = CStr(directory(rawKey))
  ElseIf directory.Exists(lowerKey) Then
    SC_DirectoryLookup = CStr(directory(lowerKey))
  ElseIf directory.Exists(upperKey) Then
    SC_DirectoryLookup = CStr(directory(upperKey))
  End If
  Err.Clear
  On Error GoTo 0

  If Len(SC_DirectoryLookup) > 0 Then Exit Function

  ' Fallback only: display-name-style lookup is convenient but not structured data.
  ' Exact caller-supplied aliases and direct SMTP parsing must win first.
  Dim lookupKey
  lookupKey = SC_RecipientLookupKey(rawKey)
  If Len(lookupKey) = 0 Then Exit Function

  Dim matchCount
  Dim matchedValue
  Dim key
  matchCount = 0
  matchedValue = ""

  On Error Resume Next
  For Each key In directory.Keys
    If SC_RecipientLookupKey(key) = lookupKey Then
      matchCount = matchCount + 1
      matchedValue = CStr(directory(key))
    End If
  Next
  Err.Clear
  On Error GoTo 0

  If matchCount = 1 Then SC_DirectoryLookup = matchedValue
End Function

Function SC_ResolveRecipient(rawValue, directory)
  Dim spec
  Set spec = SC_CreateRecipientSpec(rawValue)

  Dim normalizedEmail
  normalizedEmail = SC_NormalizeEmailAddress(rawValue)
  If SC_IsValidEmailAddress(normalizedEmail) Then
    spec("display_name") = SC_DisplayNameFromRecipientText(rawValue)
    spec("smtp_address") = normalizedEmail
    spec("resolution_status") = "direct"
    spec("directory_source") = "direct"
    spec("fallback_used") = False
    If Len(spec("display_name")) > 0 Then SC_ApplyRecipientNameParts spec, spec("display_name"), "direct-display"
    Set SC_ResolveRecipient = spec
    Exit Function
  End If

  Dim directoryValue
  directoryValue = SC_DirectoryLookup(directory, rawValue)
  If Len(directoryValue) > 0 And SC_IsValidEmailAddress(directoryValue) Then
    Dim directoryDisplayName
    directoryDisplayName = SC_DisplayNameFromRecipientText(directoryValue)
    If Len(directoryDisplayName) = 0 Then directoryDisplayName = Trim(CStr(rawValue))
    spec("display_name") = directoryDisplayName
    spec("smtp_address") = SC_NormalizeEmailAddress(directoryValue)
    spec("resolution_status") = "resolved"
    spec("directory_source") = "provided-directory"
    spec("fallback_used") = False
    If Len(SC_DisplayNameFromRecipientText(directoryValue)) > 0 Then SC_ApplyRecipientNameParts spec, directoryDisplayName, "directory-display"
    Set SC_ResolveRecipient = spec
    Exit Function
  End If

  spec("display_name") = Trim(CStr(rawValue))
  spec("smtp_address") = Trim(CStr(rawValue))
  spec("resolution_status") = "unresolved"
  spec("directory_source") = ""
  spec("fallback_used") = True
  Set SC_ResolveRecipient = spec
End Function

Function SC_ShouldSplitRecipientListOnComma(currentText)
  Dim text
  text = Trim(CStr(currentText))
  SC_ShouldSplitRecipientListOnComma = (Len(text) > 0 And (InStr(1, text, "@", vbTextCompare) > 0 Or Right(text, 1) = ">"))
End Function

Function SC_SplitRecipientListText(rawValue)
  Dim recipients
  Set recipients = SC_CreateDictionary()

  Dim text
  text = CStr(rawValue)
  Dim current
  current = ""
  Dim inQuotes
  inQuotes = False
  Dim angleDepth
  angleDepth = 0

  Dim i
  For i = 1 To Len(text)
    Dim ch
    ch = Mid(text, i, 1)
    If ch = """" Then
      inQuotes = Not inQuotes
      current = current & ch
    ElseIf ch = "<" And Not inQuotes Then
      angleDepth = angleDepth + 1
      current = current & ch
    ElseIf ch = ">" And Not inQuotes Then
      If angleDepth > 0 Then angleDepth = angleDepth - 1
      current = current & ch
    ElseIf (ch = ";" Or (ch = "," And SC_ShouldSplitRecipientListOnComma(current))) And Not inQuotes And angleDepth = 0 Then
      If Len(Trim(current)) > 0 Then recipients.Add CStr(recipients.Count), Trim(current)
      current = ""
    Else
      current = current & ch
    End If
  Next

  If Len(Trim(current)) > 0 Then recipients.Add CStr(recipients.Count), Trim(current)
  Set SC_SplitRecipientListText = recipients
End Function

Function SC_ResolveRecipientList(rawValue, directory)
  Dim recipients
  Set recipients = SC_CreateDictionary()

  Dim parts
  Set parts = SC_SplitRecipientListText(rawValue)

  Dim key
  For Each key In parts.Keys
    Dim itemText
    itemText = Trim(CStr(parts(key)))
    If Len(itemText) > 0 Then
      recipients.Add CStr(recipients.Count), SC_ResolveRecipient(itemText, directory)
    End If
  Next

  Set SC_ResolveRecipientList = recipients
End Function

Function SC_JoinResolvedRecipientAddresses(recipients)
  Dim list
  Set list = CreateObject("System.Collections.ArrayList")

  If IsObject(recipients) Then
    Dim key
    For Each key In recipients.Keys
      Dim spec
      Set spec = recipients(key)
      If Len(Trim(CStr(spec("smtp_address")))) > 0 Then list.Add Trim(CStr(spec("smtp_address")))
    Next
  End If

  If list.Count = 0 Then
    SC_JoinResolvedRecipientAddresses = ""
  Else
    Dim values()
    ReDim values(list.Count - 1)
    Dim i
    For i = 0 To list.Count - 1
      values(i) = list(i)
    Next
    SC_JoinResolvedRecipientAddresses = Join(values, "; ")
  End If
End Function
