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
  normalized = Replace(normalized, vbTab, ";")

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
  Dim loopAllowEmpty: loopAllowEmpty = CBool(allowEmpty)

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

Function PromptRecipientResolutionPackage()
  Set PromptRecipientResolutionPackage = PromptRecipientResolutionPackageWithOptions( _
    "Enter target email addresses (use ';' to separate multiple addresses): ", _
    False, _
    "At least one lookup value is required.", _
    "auto")
End Function

Function PromptRecipientResolutionPackageWithOptions(promptText, allowEmpty, emptyWarning, lookupField)
  Dim rawValues
  rawValues = PromptRecipientLookupValuesWithOptions(promptText, allowEmpty, emptyWarning)

  If Not IsArray(rawValues) Then
    rawValues = Array()
  End If

  Set PromptRecipientResolutionPackageWithOptions = BuildRecipientResolutionPackage(rawValues, lookupField)
End Function

Function BuildRecipientResolutionPackage(values, lookupField)
  Dim package: Set package = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  package.CompareMode = 1
  On Error GoTo 0

  Dim normalizedField: normalizedField = NormalizeLookupField(lookupField)
  package("LookupField") = normalizedField
  package("HasValues") = False
  package("RawValues") = Array()
  package("Entries") = Array()
  package("Addresses") = Array()

  If Not IsArray(values) Then
    Set BuildRecipientResolutionPackage = package
    Exit Function
  End If

  Dim entryList: Set entryList = CreateObject("System.Collections.ArrayList")
  Dim addressList: Set addressList = CreateObject("System.Collections.ArrayList")
  Dim rawList: Set rawList = CreateObject("System.Collections.ArrayList")

  Dim i, rawValue
  For i = LBound(values) To UBound(values)
    rawValue = Trim(SafeText(values(i)))
    If Len(rawValue) = 0 Then
      ' Skip empty entries
    Else
      rawList.Add rawValue
      Dim entry: Set entry = CreateRecipientResolutionEntry(rawValue, normalizedField)
      entryList.Add entry
      addressList.Add entry("PreferredAddress")
    End If
  Next

  If rawList.Count > 0 Then package("HasValues") = True
  package("RawValues") = ArrayFromList(rawList)
  package("Entries") = ArrayFromList(entryList)
  package("Addresses") = ArrayFromList(addressList)

  Set BuildRecipientResolutionPackage = package
End Function

Function HasRecipientResolutionEntries(package)
  HasRecipientResolutionEntries = False
  If package Is Nothing Then Exit Function

  On Error Resume Next
  If package.Exists("HasValues") Then
    If CBool(package("HasValues")) Then
      HasRecipientResolutionEntries = True
      On Error GoTo 0
      Exit Function
    End If
  End If
  On Error GoTo 0

  Dim addresses: addresses = GetRecipientResolutionAddresses(package)
  If IsArray(addresses) Then
    On Error Resume Next
    Dim upper: upper = UBound(addresses)
    If Err.Number = 0 Then
      If upper >= 0 Then HasRecipientResolutionEntries = True
    Else
      Err.Clear
    End If
    On Error GoTo 0
  End If
End Function

Function GetRecipientResolutionAddresses(package)
  If package Is Nothing Then
    GetRecipientResolutionAddresses = Array()
    Exit Function
  End If

  If package.Exists("Addresses") Then
    GetRecipientResolutionAddresses = package("Addresses")
  Else
    GetRecipientResolutionAddresses = Array()
  End If
End Function

Function GetRecipientResolutionEntries(package)
  If package Is Nothing Then
    GetRecipientResolutionEntries = Array()
    Exit Function
  End If

  If package.Exists("Entries") Then
    GetRecipientResolutionEntries = package("Entries")
  Else
    GetRecipientResolutionEntries = Array()
  End If
End Function

Function GetRecipientResolutionCount(package)
  GetRecipientResolutionCount = 0
  Dim addresses: addresses = GetRecipientResolutionAddresses(package)
  If Not IsArray(addresses) Then Exit Function

  On Error Resume Next
  Dim upper: upper = UBound(addresses)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If upper >= 0 Then GetRecipientResolutionCount = upper + 1
End Function

Function CreateRecipientResolutionEntry(rawValue, lookupField)
  Dim entry: Set entry = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  entry.CompareMode = 1
  On Error GoTo 0

  Dim trimmed: trimmed = Trim(SafeText(rawValue))
  entry("RawValue") = trimmed

  Dim normalizedField: normalizedField = NormalizeLookupField(lookupField)
  entry("LookupField") = normalizedField

  Dim parsedDisplay: parsedDisplay = trimmed
  Dim parsedEmail: parsedEmail = ""
  Dim hasAngle: hasAngle = False
  Dim parsedContact: Set parsedContact = ExtractDisplayNameAndEmail(trimmed)
  If Not parsedContact Is Nothing Then
    If parsedContact.Exists("Display") Then parsedDisplay = SafeText(parsedContact("Display"))
    If parsedContact.Exists("Email") Then parsedEmail = NormalizeEmailText(parsedContact("Email"))
    On Error Resume Next
    hasAngle = CBool(parsedContact("HasAngleFormat"))
    If Err.Number <> 0 Then
      Err.Clear
      hasAngle = False
    End If
    On Error GoTo 0
  End If

  entry("ParsedDisplay") = parsedDisplay
  entry("ParsedEmail") = parsedEmail
  entry("HasAngleFormat") = hasAngle

  Dim info: Set info = ResolveNameLookup(trimmed, normalizedField)
  If info Is Nothing Then
    Set info = CreateEmptyResolutionInfo()
  End If

  If InStr(trimmed, "@") > 0 Then
    On Error Resume Next
    Call MaybeRefineNameUsingGivenName(trimmed, info)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
  End If

  Dim resolvedEmail: resolvedEmail = ""
  If info.Exists("Email") Then resolvedEmail = NormalizeEmailText(info("Email"))
  entry("ResolvedEmail") = resolvedEmail

  Dim preferred: preferred = DeterminePreferredRecipientAddress(trimmed, resolvedEmail, parsedEmail)
  entry("PreferredAddress") = preferred

  entry("Found") = ResolutionWasSuccessful(info)

  Set entry("Resolution") = info

  Set CreateRecipientResolutionEntry = entry
End Function

Function DeterminePreferredRecipientAddress(rawValue, resolvedEmail, parsedEmail)
  Dim candidate: candidate = NormalizeEmailText(resolvedEmail)
  If Len(candidate) = 0 Then candidate = NormalizeEmailText(parsedEmail)
  If Len(candidate) = 0 Then
    If InStr(rawValue, "@") > 0 Then
      candidate = NormalizeEmailText(rawValue)
    End If
  End If
  If Len(candidate) = 0 Then candidate = Trim(SafeText(rawValue))
  DeterminePreferredRecipientAddress = candidate
End Function

Function NormalizeLookupField(fieldName)
  Dim normalized: normalized = Trim(SafeText(fieldName))
  If Len(normalized) = 0 Then
    NormalizeLookupField = "auto"
  Else
    NormalizeLookupField = LCase(normalized)
  End If
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
  If list Is Nothing Then
    ArrayFromList = Array()
    Exit Function
  End If

  If list.Count = 0 Then
    ArrayFromList = Array()
    Exit Function
  End If

  On Error Resume Next
  Dim copied: copied = list.ToArray()
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    ArrayFromList = Array()
    Exit Function
  End If
  On Error GoTo 0

  ArrayFromList = copied
End Function
