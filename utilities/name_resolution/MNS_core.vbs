' Core resolution entry points for mail name resolution

Function CreateEmptyResolutionInfo()
  Dim info: Set info = CreateObject("Scripting.Dictionary")
  info("Found") = False
  info("First") = ""
  info("Last") = ""
  info("Full") = ""
  info("Email") = ""
  info("DirectorySource") = ""
  info("DirectoryAttempted") = ""
  Set CreateEmptyResolutionInfo = info
End Function

Function AcquireResolutionNamespace()
  Set AcquireResolutionNamespace = Nothing

  Dim olApp
  On Error Resume Next
  Set olApp = GetOutlookApp()
  If Err.Number <> 0 Then
    Err.Clear
    Set olApp = Nothing
  End If
  On Error GoTo 0

  If olApp Is Nothing Then Exit Function

  Dim ns
  On Error Resume Next
  Set ns = olApp.GetNamespace("MAPI")
  If Err.Number <> 0 Then
    Err.Clear
    Set ns = Nothing
  End If
  On Error GoTo 0

  If ns Is Nothing Then Exit Function

  Set AcquireResolutionNamespace = ns
End Function

Function InitializeResolutionInfo(searchValue, info, trimmed)
  Set info = CreateEmptyResolutionInfo()
  trimmed = Trim(SafeText(searchValue))
  InitializeResolutionInfo = (Len(trimmed) > 0)
End Function

Sub EnsureFullNameField(info)
  If info Is Nothing Then Exit Sub

  Dim existingFull: existingFull = Trim(SafeText(info("Full")))
  If Len(existingFull) > 0 Then Exit Sub

  Dim combined: combined = Trim(SafeText(info("First")) & " " & SafeText(info("Last")))
  If Len(combined) > 0 Then info("Full") = combined
End Sub

Function ResolveNameLookup(searchValue, searchField)
  Dim fieldName: fieldName = SafeLower(Trim(searchField))
  If Len(fieldName) = 0 Then fieldName = "auto"

  Select Case fieldName
    Case "email", "mail", "smtp", "address"
      Set ResolveNameLookup = ResolveNameFromEmail(searchValue)
    Case "first", "given", "forename"
      Set ResolveNameLookup = ResolveNameFromGivenName(searchValue)
    Case "surname", "last", "lastname", "family", "familyname"
      Set ResolveNameLookup = ResolveNameFromSurname(searchValue)
    Case "display", "name"
      Set ResolveNameLookup = ResolveNameFromDisplayName(searchValue)
    Case "fullname", "full", "full name", "full-name", "full_name"
      Set ResolveNameLookup = ResolveNameFromFullName(searchValue)
    Case Else
      Set ResolveNameLookup = ResolveNameUsingHeuristics(searchValue)
  End Select
End Function

Function ResolveNameFromEmail(emailAddress)
  Dim info: Set info = CreateEmptyResolutionInfo()

  Dim trimmedEmail: trimmedEmail = Trim(SafeText(emailAddress))
  If Len(trimmedEmail) > 0 Then info("Email") = trimmedEmail

  Dim ns: Set ns = AcquireResolutionNamespace()
  If ns Is Nothing Then
    Set ResolveNameFromEmail = info
    Exit Function
  End If

  On Error Resume Next
  Dim recip: Set recip = ns.CreateRecipient(emailAddress)
  recip.Resolve
  If recip.Resolved Then
    Call RecordRecipientResolutionResponse("Email recipient resolve ('" & trimmedEmail & "')", recip)
    info("Found") = True
    Call ApplyFullNameCandidate(info, recip.Name, emailAddress)
    info("DirectorySource") = "Recipient resolve"
    Call ApplyEmailCandidate(info, emailAddress)
    Dim exch: Set exch = recip.AddressEntry.GetExchangeUser()
    Dim primaryAlias: primaryAlias = ExtractEmailLocalPart(emailAddress)
    If Not exch Is Nothing Then
      Dim primarySmtp
      primarySmtp = ""
      primarySmtp = SafeText(exch.PrimarySmtpAddress)
      Call ApplyEmailCandidate(info, primarySmtp)
      Call ApplyFirstNameCandidate(info, exch.FirstName, primaryAlias, emailAddress)
      If Len(exch.LastName) > 0 And Len(info("Last")) = 0 Then info("Last") = exch.LastName
      Call ApplyFullNameCandidate(info, exch.Name, emailAddress)
      Call PopulateNameFromPropertyAccessor(exch, info, primaryAlias, emailAddress)
    End If
    Call PopulateNameFromPropertyAccessor(recip.AddressEntry, info, primaryAlias, emailAddress)
    Dim resolvedAddress
    resolvedAddress = ""
    On Error Resume Next
    resolvedAddress = recip.Address
    If Err.Number <> 0 Then
      Err.Clear
      resolvedAddress = ""
    End If
    On Error GoTo 0
    Call ApplyEmailCandidate(info, resolvedAddress)
  End If
  On Error GoTo 0

  If Len(info("First")) = 0 Or Len(info("Last")) = 0 Then
    Call PopulateNameFromAddressLists(ns, emailAddress, info)
  End If

  If Len(info("First")) = 0 Or Len(info("Last")) = 0 Then
    Dim best: best = info("Full")
    Dim parts
    If Len(best) > 0 Then
      parts = Split(best, " ")
      If UBound(parts) >= 1 Then
        If Len(info("First")) = 0 Then info("First") = parts(0)
        If Len(info("Last")) = 0 Then info("Last") = parts(UBound(parts))
      End If
    End If
  End If

  If Len(info("First")) = 0 Or Len(info("Last")) = 0 Then
    Dim localPart
    localPart = emailAddress
    If InStr(localPart, "@") > 0 Then localPart = Left(localPart, InStr(localPart, "@") - 1)
    localPart = Replace(localPart, ".", " ")
    localPart = Replace(localPart, "_", " ")
    localPart = Replace(localPart, "-", " ")
    Dim segs: segs = Split(localPart, " ")
    Dim cleanedList: Set cleanedList = CreateObject("System.Collections.ArrayList")
    Dim i, segment
    For i = 0 To UBound(segs)
      segment = Trim(segs(i))
      If Len(segment) > 0 Then cleanedList.Add segment
    Next
    If cleanedList.Count >= 2 Then
      If Len(info("First")) = 0 Then info("First") = cleanedList(0)
      If Len(info("Last")) = 0 Then info("Last") = cleanedList(cleanedList.Count - 1)
    ElseIf cleanedList.Count = 1 Then
      If Len(info("First")) = 0 Then info("First") = cleanedList(0)
    End If
  End If

  Call EnsureFullNameField(info)

  Set ResolveNameFromEmail = info
End Function

Function ResolveNameUsingHeuristics(searchValue)
  Set ResolveNameUsingHeuristics = ResolveNameUsingHeuristicsWithOptions(searchValue, True)
End Function

Function ResolveNameUsingHeuristicsWithOptions(searchValue, allowSplit)
  Dim trimmed: trimmed = Trim(searchValue)
  If Len(trimmed) = 0 Then
    Set ResolveNameUsingHeuristicsWithOptions = CreateEmptyResolutionInfo()
    Exit Function
  End If

  If IsValidEmail(trimmed) Then
    Set ResolveNameUsingHeuristicsWithOptions = ResolveNameFromEmail(trimmed)
    Exit Function
  End If

  If allowSplit Then
    If InStr(trimmed, ";") > 0 Then
      Dim parts: parts = Split(trimmed, ";")
      Dim i
      For i = LBound(parts) To UBound(parts)
        Dim candidate: candidate = Trim(parts(i))
        If Len(candidate) = 0 Then
          ' Skip empty segments
        Else
          Dim parsedCandidate: Set parsedCandidate = ExtractDisplayNameAndEmail(candidate)
          Dim candidateHasAngle: candidateHasAngle = False
          On Error Resume Next
          candidateHasAngle = CBool(parsedCandidate("HasAngleFormat"))
          On Error GoTo 0
          If candidateHasAngle Then
            Dim preferredInfo: Set preferredInfo = ResolveNameUsingHeuristicsWithOptions(candidate, False)
            If Not preferredInfo Is Nothing Then
              Set ResolveNameUsingHeuristicsWithOptions = preferredInfo
              Exit Function
            End If
          End If
        End If
      Next

      Dim lastInfo: Set lastInfo = Nothing
      For i = LBound(parts) To UBound(parts)
        Dim fallbackCandidate: fallbackCandidate = Trim(parts(i))
        If Len(fallbackCandidate) = 0 Then
          ' Skip empty segments
        Else
          Dim candidateInfo: Set candidateInfo = ResolveNameUsingHeuristicsWithOptions(fallbackCandidate, False)
          If ResolutionWasSuccessful(candidateInfo) Then
            Set ResolveNameUsingHeuristicsWithOptions = candidateInfo
            Exit Function
          End If
          Set lastInfo = candidateInfo
        End If
      Next

      If Not lastInfo Is Nothing Then
        Set ResolveNameUsingHeuristicsWithOptions = lastInfo
        Exit Function
      End If
    End If
  End If

  Dim normalizedInput: normalizedInput = trimmed
  Dim componentSourceText: componentSourceText = trimmed
  Dim parsedDisplay: parsedDisplay = ""
  Dim extractedEmail: extractedEmail = ""
  Dim angleRecognized: angleRecognized = False
  Dim parsedContact: Set parsedContact = ExtractDisplayNameAndEmail(trimmed)
  If Not parsedContact Is Nothing Then
    On Error Resume Next
    parsedDisplay = Trim(SafeText(parsedContact("Display")))
    extractedEmail = NormalizeEmailText(parsedContact("Email"))
    angleRecognized = CBool(parsedContact("HasAngleFormat"))
    On Error GoTo 0

    If Len(parsedDisplay) > 0 Then
      normalizedInput = parsedDisplay
      componentSourceText = parsedDisplay
    ElseIf Len(extractedEmail) > 0 Then
      normalizedInput = extractedEmail
      componentSourceText = parsedDisplay
    End If
  End If

  trimmed = normalizedInput

  Dim components: Set components = ParseFullNameComponents(componentSourceText)
  Dim tokens: Set tokens = CollectFlexibleNameTokens(trimmed)
  Dim tokenCount: tokenCount = 0
  If Not tokens Is Nothing Then tokenCount = tokens.Count

  Dim angleInfo: Set angleInfo = Nothing
  If Len(extractedEmail) > 0 Then
    Set angleInfo = ResolveNameFromEmail(extractedEmail)
    Call ApplyEmailCandidate(angleInfo, extractedEmail)
    Call MergeParsedContactIntoInfo(angleInfo, components, componentSourceText)
    If ResolutionWasSuccessful(angleInfo) Then
      Set ResolveNameUsingHeuristicsWithOptions = angleInfo
      Exit Function
    End If
  End If

  If tokenCount >= 2 Then
    Dim candidateSet: Set candidateSet = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    candidateSet.CompareMode = 1
    On Error GoTo 0

    Call AddCandidateValue(candidateSet, trimmed)

    If Not components Is Nothing Then
      Call AddCandidateValue(candidateSet, components("Display"))
      Call AddCandidateValue(candidateSet, components("Normalized"))
      Call AddCandidateValue(candidateSet, components("ReverseDisplay"))
      Call AddCandidateValue(candidateSet, Trim(SafeText(components("First")) & " " & SafeText(components("Last"))))
      Call AddCandidateValue(candidateSet, Trim(SafeText(components("Last")) & " " & SafeText(components("First"))))
      Call AddCandidateValue(candidateSet, Trim(SafeText(components("AlternateFirst")) & " " & SafeText(components("AlternateLast"))))
    End If

    Dim normalizedTokens: normalizedTokens = JoinTokens(tokens, " ")
    Call AddCandidateValue(candidateSet, normalizedTokens)

    Dim reversedTokens: reversedTokens = BuildReversedTokenDisplay(tokens)
    Call AddCandidateValue(candidateSet, reversedTokens)

    Dim firstToken: firstToken = ""
    Dim lastToken: lastToken = ""
    If tokenCount > 0 Then firstToken = SafeText(tokens(0))
    If tokenCount > 0 Then lastToken = SafeText(tokens(tokenCount - 1))

    If Len(firstToken) > 0 And Len(lastToken) > 0 Then
      Call AddCandidateValue(candidateSet, Trim(firstToken & " " & lastToken))
      Call AddCandidateValue(candidateSet, Trim(lastToken & " " & firstToken))
      Call AddCandidateValue(candidateSet, Trim(lastToken & ", " & firstToken))
    End If

    Dim key
    For Each key In candidateSet.Keys
      Dim candidateValue: candidateValue = Trim(SafeText(key))
      If Len(candidateValue) > 0 Then
        Dim resolved: Set resolved = ResolveNameFromFullName(candidateValue)
        If ResolutionWasSuccessful(resolved) Then
          If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(resolved, extractedEmail)
          Call MergeParsedContactIntoInfo(resolved, components, componentSourceText)
          Set ResolveNameUsingHeuristicsWithOptions = resolved
          Exit Function
        End If
      End If
    Next

    For Each key In candidateSet.Keys
      Dim candidateDisplay: candidateDisplay = Trim(SafeText(key))
      If Len(candidateDisplay) > 0 Then
        Dim displayInfo: Set displayInfo = ResolveNameFromDisplayName(candidateDisplay)
        If ResolutionWasSuccessful(displayInfo) Then
          If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(displayInfo, extractedEmail)
          Call MergeParsedContactIntoInfo(displayInfo, components, componentSourceText)
          Set ResolveNameUsingHeuristicsWithOptions = displayInfo
          Exit Function
        End If
      End If
    Next

    If Len(SafeText(firstToken)) > 0 Then
      Dim firstTokenInfo: Set firstTokenInfo = ResolveNameFromGivenName(firstToken)
      If ResolutionWasSuccessful(firstTokenInfo) Then
        If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(firstTokenInfo, extractedEmail)
        Call MergeParsedContactIntoInfo(firstTokenInfo, components, componentSourceText)
        Set ResolveNameUsingHeuristicsWithOptions = firstTokenInfo
        Exit Function
      End If
    End If

    If Len(SafeText(lastToken)) > 0 Then
      Dim lastTokenInfo: Set lastTokenInfo = ResolveNameFromSurname(lastToken)
      If ResolutionWasSuccessful(lastTokenInfo) Then
        If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(lastTokenInfo, extractedEmail)
        Call MergeParsedContactIntoInfo(lastTokenInfo, components, componentSourceText)
        Set ResolveNameUsingHeuristicsWithOptions = lastTokenInfo
        Exit Function
      End If
    End If

    If Not components Is Nothing Then
      Dim componentFirst: componentFirst = SafeText(components("First"))
      Dim componentLast: componentLast = SafeText(components("Last"))

      If Len(componentFirst) > 0 Then
        Dim componentGiven: Set componentGiven = ResolveNameFromGivenName(componentFirst)
        If ResolutionWasSuccessful(componentGiven) Then
          If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(componentGiven, extractedEmail)
          Call MergeParsedContactIntoInfo(componentGiven, components, componentSourceText)
          Set ResolveNameUsingHeuristicsWithOptions = componentGiven
          Exit Function
        End If
      End If

      If Len(componentLast) > 0 Then
        Dim componentSurname: Set componentSurname = ResolveNameFromSurname(componentLast)
        If ResolutionWasSuccessful(componentSurname) Then
          If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(componentSurname, extractedEmail)
          Call MergeParsedContactIntoInfo(componentSurname, components, componentSourceText)
          Set ResolveNameUsingHeuristicsWithOptions = componentSurname
          Exit Function
        End If
      End If

      Dim normalizedFull: normalizedFull = Trim(componentFirst & " " & componentLast)
      If Len(normalizedFull) > 0 And SafeLower(normalizedFull) <> SafeLower(trimmed) Then
        Dim normalizedInfo: Set normalizedInfo = ResolveNameFromFullName(normalizedFull)
        If ResolutionWasSuccessful(normalizedInfo) Then
          If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(normalizedInfo, extractedEmail)
          Call MergeParsedContactIntoInfo(normalizedInfo, components, componentSourceText)
          Set ResolveNameUsingHeuristicsWithOptions = normalizedInfo
          Exit Function
        End If
      End If
    End If
  End If

  Dim fallbackInfo: Set fallbackInfo = ResolveNameFromGivenName(trimmed)
  If ResolutionWasSuccessful(fallbackInfo) Then
    If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(fallbackInfo, extractedEmail)
    Call MergeParsedContactIntoInfo(fallbackInfo, components, componentSourceText)
    Set ResolveNameUsingHeuristicsWithOptions = fallbackInfo
    Exit Function
  End If

  Dim surnameInfo: Set surnameInfo = ResolveNameFromSurname(trimmed)
  If ResolutionWasSuccessful(surnameInfo) Then
    If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(surnameInfo, extractedEmail)
    Call MergeParsedContactIntoInfo(surnameInfo, components, componentSourceText)
    Set ResolveNameUsingHeuristicsWithOptions = surnameInfo
    Exit Function
  End If

  Dim displayFallback: Set displayFallback = ResolveNameFromDisplayName(trimmed)
  If ResolutionWasSuccessful(displayFallback) Then
    If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(displayFallback, extractedEmail)
    Call MergeParsedContactIntoInfo(displayFallback, components, componentSourceText)
    Set ResolveNameUsingHeuristicsWithOptions = displayFallback
    Exit Function
  End If

  If Not angleInfo Is Nothing Then
    If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(angleInfo, extractedEmail)
    Call MergeParsedContactIntoInfo(angleInfo, components, componentSourceText)
    If angleRecognized Then
      angleInfo("Found") = True
      Dim existingSource: existingSource = Trim(SafeText(angleInfo("DirectorySource")))
      If Len(existingSource) = 0 Or SafeLower(existingSource) = "(not found)" Then
        angleInfo("DirectorySource") = "Parsed recipient input"
      End If
      If Len(Trim(SafeText(angleInfo("DirectoryAttempted")))) = 0 Then
        angleInfo("DirectoryAttempted") = "(not attempted)"
      End If
    End If
    Set ResolveNameUsingHeuristicsWithOptions = angleInfo
    Exit Function
  End If

  If Len(extractedEmail) > 0 Then Call ApplyEmailCandidate(fallbackInfo, extractedEmail)
  Call MergeParsedContactIntoInfo(fallbackInfo, components, componentSourceText)
  Set ResolveNameUsingHeuristicsWithOptions = fallbackInfo
End Function

Sub MergeParsedContactIntoInfo(info, components, displayText)
  If info Is Nothing Then Exit Sub

  Dim displayCandidate: displayCandidate = Trim(SafeText(displayText))
  Dim firstCandidate: firstCandidate = ""
  Dim lastCandidate: lastCandidate = ""
  Dim normalizedCandidate: normalizedCandidate = ""

  If Not components Is Nothing Then
    On Error Resume Next
    firstCandidate = Trim(SafeText(components("First")))
    lastCandidate = Trim(SafeText(components("Last")))
    normalizedCandidate = Trim(SafeText(components("Normalized")))
    If Len(displayCandidate) = 0 Then displayCandidate = Trim(SafeText(components("Display")))
    On Error GoTo 0
  End If

  Dim existingFirst: existingFirst = Trim(SafeText(info("First")))
  If Len(existingFirst) = 0 And Len(firstCandidate) > 0 Then info("First") = firstCandidate

  Dim existingLast: existingLast = Trim(SafeText(info("Last")))
  If Len(existingLast) = 0 And Len(lastCandidate) > 0 Then info("Last") = lastCandidate

  Dim existingFull: existingFull = Trim(SafeText(info("Full")))
  If Len(existingFull) = 0 Then
    If Len(displayCandidate) > 0 Then
      info("Full") = displayCandidate
    ElseIf Len(normalizedCandidate) > 0 Then
      info("Full") = normalizedCandidate
    ElseIf Len(firstCandidate) > 0 Or Len(lastCandidate) > 0 Then
      info("Full") = Trim(firstCandidate & " " & lastCandidate)
    End If
  End If
End Sub

Function ResolveNameFromGivenName(givenName)
  Dim info, trimmed
  If Not InitializeResolutionInfo(givenName, info, trimmed) Then
    Set ResolveNameFromGivenName = info
    Exit Function
  End If

  Dim ns: Set ns = AcquireResolutionNamespace()
  If ns Is Nothing Then
    Set ResolveNameFromGivenName = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsByGivenName(ns, trimmed, info, "given name '" & trimmed & "'")
  Call EnsureFullNameField(info)

  Set ResolveNameFromGivenName = info
End Function

Function ResolveNameFromSurname(surname)
  Dim info, trimmed
  If Not InitializeResolutionInfo(surname, info, trimmed) Then
    Set ResolveNameFromSurname = info
    Exit Function
  End If

  Dim ns: Set ns = AcquireResolutionNamespace()
  If ns Is Nothing Then
    Set ResolveNameFromSurname = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsBySurname(ns, trimmed, info, "surname '" & trimmed & "'")
  Call EnsureFullNameField(info)

  Set ResolveNameFromSurname = info
End Function

Function ResolveNameFromDisplayName(displayName)
  Dim info, trimmed
  If Not InitializeResolutionInfo(displayName, info, trimmed) Then
    Set ResolveNameFromDisplayName = info
    Exit Function
  End If

  Dim ns: Set ns = AcquireResolutionNamespace()
  If ns Is Nothing Then
    Set ResolveNameFromDisplayName = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsByDisplayName(ns, trimmed, info, "name '" & trimmed & "'")

  If Not ResolutionWasSuccessful(info) Then
    Call AttemptRecipientResolveForValue(ns, trimmed, info, "display name")

    If Not ResolutionWasSuccessful(info) Then
      WScript.Echo "  Recipient resolve did not identify a match; invoking display-name enumeration fallback for '" & trimmed & "'."
      Call EnumerateDirectoryMatchesByDisplayName(ns, trimmed, info, "name '" & trimmed & "'")
    End If
  End If

  Call EnsureFullNameField(info)

  Set ResolveNameFromDisplayName = info
End Function

Function ResolveNameFromFullName(fullName)
  Dim info, trimmed
  If Not InitializeResolutionInfo(fullName, info, trimmed) Then
    Set ResolveNameFromFullName = info
    Exit Function
  End If

  Dim ns: Set ns = AcquireResolutionNamespace()
  If ns Is Nothing Then
    Set ResolveNameFromFullName = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsByFullName(ns, trimmed, info, "full name '" & trimmed & "'")

  If Not ResolutionWasSuccessful(info) Then
    Call AttemptRecipientResolveForValue(ns, trimmed, info, "full name")
  End If

  Call EnsureFullNameField(info)

  If ResolutionWasSuccessful(info) Then
    Set ResolveNameFromFullName = info
    Exit Function
  End If

  Dim fallback: Set fallback = ResolveNameFromDisplayName(trimmed)
  If ResolutionWasSuccessful(fallback) Then
    Set ResolveNameFromFullName = fallback
  Else
    Set ResolveNameFromFullName = info
  End If
End Function

Function ResolutionWasSuccessful(info)
  ResolutionWasSuccessful = False
  If info Is Nothing Then Exit Function
  On Error Resume Next
  If info.Exists("Found") Then
    If CBool(info("Found")) Then ResolutionWasSuccessful = True
  End If
  On Error GoTo 0
End Function
