' Directory resolution utilities for personal deliveries

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

  On Error Resume Next
  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    ResolveNameFromEmail = info
    Exit Function
  End If
  Dim ns: Set ns = olApp.GetNamespace("MAPI")
  Dim recip: Set recip = ns.CreateRecipient(emailAddress)
  recip.Resolve
  If recip.Resolved Then
    info("Found") = True
    info("Full") = recip.Name
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
      If Len(info("Full")) = 0 Then info("Full") = exch.Name
      Call PopulateNameFromPropertyAccessor(exch, info, primaryAlias, emailAddress)
    End If
    Call PopulateNameFromPropertyAccessor(recip.AddressEntry, info, primaryAlias, emailAddress)
    Dim resolvedAddress: resolvedAddress
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

  If Len(info("Full")) = 0 Then
    info("Full") = Trim(info("First") & " " & info("Last"))
  End If

  Set ResolveNameFromEmail = info
End Function

Function ResolveNameUsingHeuristics(searchValue)
  Dim trimmed: trimmed = Trim(searchValue)
  If Len(trimmed) = 0 Then
    Set ResolveNameUsingHeuristics = CreateEmptyResolutionInfo()
    Exit Function
  End If

  If IsValidEmail(trimmed) Then
    Set ResolveNameUsingHeuristics = ResolveNameFromEmail(trimmed)
    Exit Function
  End If

  Dim components: Set components = ParseFullNameComponents(trimmed)
  Dim tokens: Set tokens = CollectFlexibleNameTokens(trimmed)
  Dim tokenCount: tokenCount = 0
  If Not tokens Is Nothing Then tokenCount = tokens.Count

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
          Set ResolveNameUsingHeuristics = resolved
          Exit Function
        End If
      End If
    Next

    For Each key In candidateSet.Keys
      Dim candidateDisplay: candidateDisplay = Trim(SafeText(key))
      If Len(candidateDisplay) > 0 Then
        Dim displayInfo: Set displayInfo = ResolveNameFromDisplayName(candidateDisplay)
        If ResolutionWasSuccessful(displayInfo) Then
          Set ResolveNameUsingHeuristics = displayInfo
          Exit Function
        End If
      End If
    Next

    If Len(SafeText(firstToken)) > 0 Then
      Dim firstTokenInfo: Set firstTokenInfo = ResolveNameFromGivenName(firstToken)
      If ResolutionWasSuccessful(firstTokenInfo) Then
        Set ResolveNameUsingHeuristics = firstTokenInfo
        Exit Function
      End If
    End If

    If Len(SafeText(lastToken)) > 0 Then
      Dim lastTokenInfo: Set lastTokenInfo = ResolveNameFromSurname(lastToken)
      If ResolutionWasSuccessful(lastTokenInfo) Then
        Set ResolveNameUsingHeuristics = lastTokenInfo
        Exit Function
      End If
    End If

    If Not components Is Nothing Then
      Dim componentFirst: componentFirst = SafeText(components("First"))
      Dim componentLast: componentLast = SafeText(components("Last"))

      If Len(componentFirst) > 0 Then
        Dim componentGiven: Set componentGiven = ResolveNameFromGivenName(componentFirst)
        If ResolutionWasSuccessful(componentGiven) Then
          Set ResolveNameUsingHeuristics = componentGiven
          Exit Function
        End If
      End If

      If Len(componentLast) > 0 Then
        Dim componentSurname: Set componentSurname = ResolveNameFromSurname(componentLast)
        If ResolutionWasSuccessful(componentSurname) Then
          Set ResolveNameUsingHeuristics = componentSurname
          Exit Function
        End If
      End If

      Dim normalizedFull: normalizedFull = Trim(componentFirst & " " & componentLast)
      If Len(normalizedFull) > 0 And SafeLower(normalizedFull) <> SafeLower(trimmed) Then
        Dim normalizedInfo: Set normalizedInfo = ResolveNameFromFullName(normalizedFull)
        If ResolutionWasSuccessful(normalizedInfo) Then
          Set ResolveNameUsingHeuristics = normalizedInfo
          Exit Function
        End If
      End If
    End If
  End If

  Dim fallbackInfo: Set fallbackInfo = ResolveNameFromGivenName(trimmed)
  If ResolutionWasSuccessful(fallbackInfo) Then
    Set ResolveNameUsingHeuristics = fallbackInfo
    Exit Function
  End If

  Dim surnameInfo: Set surnameInfo = ResolveNameFromSurname(trimmed)
  If ResolutionWasSuccessful(surnameInfo) Then
    Set ResolveNameUsingHeuristics = surnameInfo
    Exit Function
  End If

  Dim displayFallback: Set displayFallback = ResolveNameFromDisplayName(trimmed)
  If ResolutionWasSuccessful(displayFallback) Then
    Set ResolveNameUsingHeuristics = displayFallback
    Exit Function
  End If

  Set ResolveNameUsingHeuristics = fallbackInfo
End Function

Function ResolveNameFromGivenName(givenName)
  Dim info: Set info = CreateEmptyResolutionInfo()
  Dim trimmed: trimmed = Trim(givenName)
  If Len(trimmed) = 0 Then
    Set ResolveNameFromGivenName = info
    Exit Function
  End If

  On Error Resume Next
  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    Set ResolveNameFromGivenName = info
    Exit Function
  End If
  Dim ns: Set ns = olApp.GetNamespace("MAPI")
  On Error GoTo 0
  If ns Is Nothing Then
    Set ResolveNameFromGivenName = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsByGivenName(ns, trimmed, info, "given name '" & trimmed & "'")

  If Len(info("Full")) = 0 Then
    info("Full") = Trim(info("First") & " " & info("Last"))
  End If

  Set ResolveNameFromGivenName = info
End Function

Function ResolveNameFromSurname(surname)
  Dim info: Set info = CreateEmptyResolutionInfo()
  Dim trimmed: trimmed = Trim(surname)
  If Len(trimmed) = 0 Then
    Set ResolveNameFromSurname = info
    Exit Function
  End If

  On Error Resume Next
  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    Set ResolveNameFromSurname = info
    Exit Function
  End If
  Dim ns: Set ns = olApp.GetNamespace("MAPI")
  On Error GoTo 0
  If ns Is Nothing Then
    Set ResolveNameFromSurname = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsBySurname(ns, trimmed, info, "surname '" & trimmed & "'")

  If Len(info("Full")) = 0 Then
    info("Full") = Trim(info("First") & " " & info("Last"))
  End If

  Set ResolveNameFromSurname = info
End Function

Function ResolveNameFromDisplayName(displayName)
  Dim info: Set info = CreateEmptyResolutionInfo()
  Dim trimmed: trimmed = Trim(displayName)
  If Len(trimmed) = 0 Then
    Set ResolveNameFromDisplayName = info
    Exit Function
  End If

  On Error Resume Next
  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    Set ResolveNameFromDisplayName = info
    Exit Function
  End If
  Dim ns: Set ns = olApp.GetNamespace("MAPI")
  On Error GoTo 0
  If ns Is Nothing Then
    Set ResolveNameFromDisplayName = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsByDisplayName(ns, trimmed, info, "name '" & trimmed & "'")

  If Not ResolutionWasSuccessful(info) Then
    Call AttemptRecipientResolveForValue(ns, trimmed, info, "display name")
  End If

  If Len(info("Full")) = 0 Then
    info("Full") = Trim(info("First") & " " & info("Last"))
  End If

  Set ResolveNameFromDisplayName = info
End Function

Function ResolveNameFromFullName(fullName)
  Dim info: Set info = CreateEmptyResolutionInfo()
  Dim trimmed: trimmed = Trim(fullName)
  If Len(trimmed) = 0 Then
    Set ResolveNameFromFullName = info
    Exit Function
  End If

  On Error Resume Next
  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    Set ResolveNameFromFullName = info
    Exit Function
  End If
  Dim ns: Set ns = olApp.GetNamespace("MAPI")
  On Error GoTo 0
  If ns Is Nothing Then
    Set ResolveNameFromFullName = info
    Exit Function
  End If

  Call PopulateNameFromAddressListsByFullName(ns, trimmed, info, "full name '" & trimmed & "'")

  If Not ResolutionWasSuccessful(info) Then
    Call AttemptRecipientResolveForValue(ns, trimmed, info, "full name")
  End If

  If Len(info("Full")) = 0 Then
    info("Full") = Trim(info("First") & " " & info("Last"))
  End If

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

Sub RecordDirectoryAttempt(info, label)
  If info Is Nothing Then Exit Sub

  Dim attemptLabel: attemptLabel = ""
  If Not IsNull(label) Then
    If Not IsEmpty(label) Then attemptLabel = CStr(label)
  End If

  attemptLabel = Trim(attemptLabel)
  If Len(attemptLabel) = 0 Then Exit Sub

  Dim existing: existing = ""
  On Error Resume Next
  If info.Exists("DirectoryAttempted") Then existing = CStr(info("DirectoryAttempted"))
  On Error GoTo 0

  Dim shouldAppend: shouldAppend = True
  If Len(existing) > 0 Then
    Dim parts: parts = Split(existing, ",")
    Dim i, part
    For i = LBound(parts) To UBound(parts)
      part = Trim(parts(i))
      If Len(part) > 0 Then
        If SafeLower(part) = SafeLower(attemptLabel) Then
          shouldAppend = False
          Exit For
        End If
      End If
    Next
  End If

  If shouldAppend Then
    If Len(existing) = 0 Then
      info("DirectoryAttempted") = attemptLabel
    Else
      info("DirectoryAttempted") = existing & ", " & attemptLabel
    End If
  End If
End Sub

Sub PopulateNameFromAddressLists(ns, emailAddress, info)
  If ns Is Nothing Then Exit Sub

  Dim listNames
  listNames = BuildDirectorySearchList(ns)

  Dim i, addressList, entry, exch, contact
  On Error Resume Next
  For i = LBound(listNames) To UBound(listNames)
    Set addressList = Nothing
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    If Not addressList Is Nothing Then
      Call RecordDirectoryAttempt(info, listNames(i))
      Set entry = Nothing
      Set exch = Nothing
      Set contact = Nothing
      Set entry = addressList.AddressEntries.Find("[SMTPAddress] = '" & emailAddress & "'")
      If Err.Number <> 0 Then
        Err.Clear
        Set entry = Nothing
      End If
      If entry Is Nothing Then
        Set entry = addressList.AddressEntries.Find("[EmailAddress] = '" & emailAddress & "'")
        If Err.Number <> 0 Then
          Err.Clear
          Set entry = Nothing
        End If
      End If
      If entry Is Nothing Then
        Set contact = FindContactInAddressList(addressList, emailAddress)
        If Not contact Is Nothing Then
          info("Found") = True
          Call PopulateNameFromContactItem(contact, info)
          info("DirectorySource") = listNames(i)
          ReportAddressListLookup emailAddress, listNames(i), entry, exch, info, contact
          Exit For
        End If
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name
        Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        Dim listAlias: listAlias = ExtractEmailLocalPart(emailAddress)
        If Not exch Is Nothing Then
          Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
          Call ApplyFirstNameCandidate(info, exch.FirstName, listAlias, emailAddress)
          If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
          If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
          Call PopulateNameFromPropertyAccessor(exch, info, listAlias, emailAddress)
        End If
        Call PopulateNameFromPropertyAccessor(entry, info, listAlias, emailAddress)
        info("DirectorySource") = listNames(i)
        ReportAddressListLookup emailAddress, listNames(i), entry, exch, info, contact
        Exit For
      End If
    End If
  Next
  If Len(info("DirectorySource")) = 0 And Len(SafeText(info("DirectoryAttempted"))) > 0 Then
    info("DirectorySource") = "(not found)"
  End If
  On Error GoTo 0
End Sub

Sub MaybeRefineNameUsingGivenName(emailAddress, info)
  If info Is Nothing Then Exit Sub

  Dim firstName: firstName = ""
  If info.Exists("First") Then firstName = SafeText(info("First"))
  If Len(firstName) = 0 Then Exit Sub

  Dim lastName: lastName = ""
  If info.Exists("Last") Then lastName = SafeText(info("Last"))
  If Len(lastName) > 0 Then Exit Sub

  Dim localPart: localPart = ExtractEmailLocalPart(emailAddress)
  If Len(localPart) = 0 Then Exit Sub

  If SafeLower(firstName) <> SafeLower(localPart) Then Exit Sub

  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then Exit Sub

  Dim ns: Set ns = olApp.GetNamespace("MAPI")
  If ns Is Nothing Then Exit Sub

  WScript.Echo "  First name matches email username. Searching directory for '" & firstName & "'."

  Call AttemptAliasRecipientResolve(ns, firstName, info, emailAddress)

  Dim refinedLast: refinedLast = SafeText(info("Last"))
  Dim refinedFirst: refinedFirst = SafeLower(SafeText(info("First")))
  Dim targetAlias: targetAlias = SafeLower(firstName)

  If Len(refinedLast) = 0 Or refinedFirst = targetAlias Then
    Call PopulateNameFromAddressListsByGivenName(ns, firstName, info, emailAddress)
  End If
End Sub

Sub PopulateNameFromAddressListsByGivenName(ns, givenName, info, emailAddress)
  If ns Is Nothing Then Exit Sub
  If Len(givenName) = 0 Then Exit Sub

  Dim trimmed: trimmed = Trim(givenName)
  If Len(trimmed) = 0 Then Exit Sub

  Dim listNames: listNames = Array("Global Address List", "Offline Global Address List")
  Dim i
  For i = LBound(listNames) To UBound(listNames)
    Dim addressList: Set addressList = Nothing
    On Error Resume Next
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If Not addressList Is Nothing Then
      WScript.Echo "  Attempting given-name lookup in " & listNames(i) & " for '" & trimmed & "'."

      Call RecordDirectoryAttempt(info, listNames(i))

      Dim beforeFirst: beforeFirst = SafeText(info("First"))
      Dim beforeLast: beforeLast = SafeText(info("Last"))
      Dim beforeFull: beforeFull = SafeText(info("Full"))

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing

      Set entry = FindAddressEntryByFields(addressList, Array("Name", "FirstName", "Alias"), trimmed)

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        On Error GoTo 0

      If Not exch Is Nothing Then
        Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
        Call ApplyFirstNameCandidate(info, exch.FirstName, givenName, emailAddress)
        If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
        If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
          Call PopulateNameFromPropertyAccessor(exch, info, givenName, emailAddress)
        End If
        Call PopulateNameFromPropertyAccessor(entry, info, givenName, emailAddress)

        Dim afterFirst: afterFirst = SafeText(info("First"))
        Dim afterLast: afterLast = SafeText(info("Last"))
        Dim afterFull: afterFull = SafeText(info("Full"))

        Dim improved: improved = False
        If Len(afterLast) > 0 And SafeLower(afterLast) <> SafeLower(beforeLast) Then improved = True
        If Not improved Then
          If Len(afterFirst) > 0 And SafeLower(afterFirst) <> SafeLower(beforeFirst) Then improved = True
        End If
        If Not improved Then
          If Len(afterFull) > 0 And SafeLower(afterFull) <> SafeLower(beforeFull) Then improved = True
        End If

        If improved Then
          info("DirectorySource") = listNames(i)
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing)
          Exit For
        End If
      End If
    End If
  Next
End Sub

Sub PopulateNameFromAddressListsBySurname(ns, surname, info, emailAddress)
  If ns Is Nothing Then Exit Sub
  If Len(surname) = 0 Then Exit Sub

  Dim trimmed: trimmed = Trim(surname)
  If Len(trimmed) = 0 Then Exit Sub

  Dim listNames: listNames = Array("Global Address List", "Offline Global Address List")
  Dim i
  For i = LBound(listNames) To UBound(listNames)
    Dim addressList: Set addressList = Nothing
    On Error Resume Next
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If Not addressList Is Nothing Then
      WScript.Echo "  Attempting surname lookup in " & listNames(i) & " for '" & trimmed & "'."

      Call RecordDirectoryAttempt(info, listNames(i))

      Dim beforeFirst: beforeFirst = SafeText(info("First"))
      Dim beforeLast: beforeLast = SafeText(info("Last"))
      Dim beforeFull: beforeFull = SafeText(info("Full"))

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing

      Set entry = FindAddressEntryByFields(addressList, Array("LastName", "Surname", "Name"), trimmed)

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        On Error GoTo 0

      If Not exch Is Nothing Then
        Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
        Call ApplyFirstNameCandidate(info, exch.FirstName, "", "")
        If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
        If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
          Call PopulateNameFromPropertyAccessor(exch, info, "", "")
        End If
        Call PopulateNameFromPropertyAccessor(entry, info, "", "")

        Dim afterFirst: afterFirst = SafeText(info("First"))
        Dim afterLast: afterLast = SafeText(info("Last"))
        Dim afterFull: afterFull = SafeText(info("Full"))

        Dim improved: improved = False
        If Len(afterLast) > 0 And SafeLower(afterLast) <> SafeLower(beforeLast) Then improved = True
        If Not improved Then
          If Len(afterFirst) > 0 And SafeLower(afterFirst) <> SafeLower(beforeFirst) Then improved = True
        End If
        If Not improved Then
          If Len(afterFull) > 0 And SafeLower(afterFull) <> SafeLower(beforeFull) Then improved = True
        End If

        If improved Then
          info("DirectorySource") = listNames(i)
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing)
          Exit For
        End If
      End If
    End If
  Next
End Sub

Sub PopulateNameFromAddressListsByDisplayName(ns, displayName, info, emailAddress)
  If ns Is Nothing Then Exit Sub
  If Len(displayName) = 0 Then Exit Sub

  Dim trimmed: trimmed = Trim(displayName)
  If Len(trimmed) = 0 Then Exit Sub

  Dim listNames: listNames = Array("Global Address List", "Offline Global Address List")
  Dim i
  For i = LBound(listNames) To UBound(listNames)
    Dim addressList: Set addressList = Nothing
    On Error Resume Next
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If Not addressList Is Nothing Then
      WScript.Echo "  Attempting display-name lookup in " & listNames(i) & " for '" & trimmed & "'."

      Call RecordDirectoryAttempt(info, listNames(i))

      Dim beforeFirst: beforeFirst = SafeText(info("First"))
      Dim beforeLast: beforeLast = SafeText(info("Last"))
      Dim beforeFull: beforeFull = SafeText(info("Full"))

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing

      Set entry = FindAddressEntryByFields(addressList, Array("Name", "DisplayName"), trimmed)

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        On Error GoTo 0

      If Not exch Is Nothing Then
        Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
        Call ApplyFirstNameCandidate(info, exch.FirstName, "", "")
        If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
        If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
          Call PopulateNameFromPropertyAccessor(exch, info, "", "")
        End If
        Call PopulateNameFromPropertyAccessor(entry, info, "", "")

        Dim afterFirst: afterFirst = SafeText(info("First"))
        Dim afterLast: afterLast = SafeText(info("Last"))
        Dim afterFull: afterFull = SafeText(info("Full"))

        Dim improved: improved = False
        If Len(afterLast) > 0 And SafeLower(afterLast) <> SafeLower(beforeLast) Then improved = True
        If Not improved Then
          If Len(afterFirst) > 0 And SafeLower(afterFirst) <> SafeLower(beforeFirst) Then improved = True
        End If
        If Not improved Then
          If Len(afterFull) > 0 And SafeLower(afterFull) <> SafeLower(beforeFull) Then improved = True
        End If

        If improved Then
          info("DirectorySource") = listNames(i)
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing)
          Exit For
        End If
      End If
    End If
  Next
End Sub

Sub PopulateNameFromAddressListsByFullName(ns, fullName, info, emailAddress)
  If ns Is Nothing Then Exit Sub
  If info Is Nothing Then Exit Sub

  Dim trimmed: trimmed = Trim(fullName)
  If Len(trimmed) = 0 Then Exit Sub

  Dim components: Set components = ParseFullNameComponents(trimmed)

  Dim firstName: firstName = ""
  Dim lastName: lastName = ""
  Dim alternateFirst: alternateFirst = ""
  Dim alternateLast: alternateLast = ""
  Dim normalized: normalized = trimmed
  Dim reverseDisplay: reverseDisplay = ""

  If Not components Is Nothing Then
    If components.Exists("First") Then firstName = SafeText(components("First"))
    If components.Exists("Last") Then lastName = SafeText(components("Last"))
    If components.Exists("AlternateFirst") Then alternateFirst = SafeText(components("AlternateFirst"))
    If components.Exists("AlternateLast") Then alternateLast = SafeText(components("AlternateLast"))
    If components.Exists("Normalized") Then
      Dim normalizedCandidate: normalizedCandidate = SafeText(components("Normalized"))
      If Len(normalizedCandidate) > 0 Then normalized = normalizedCandidate
    End If
    If components.Exists("ReverseDisplay") Then reverseDisplay = SafeText(components("ReverseDisplay"))
  End If

  If Len(reverseDisplay) = 0 Then
    If Len(lastName) > 0 And Len(firstName) > 0 Then
      reverseDisplay = Trim(lastName & ", " & firstName)
    Else
      reverseDisplay = trimmed
    End If
  End If

  Dim listNames: listNames = Array("Global Address List", "Offline Global Address List")
  Dim i
  For i = LBound(listNames) To UBound(listNames)
    Dim addressList: Set addressList = Nothing
    On Error Resume Next
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If Not addressList Is Nothing Then
      WScript.Echo "  Attempting full-name lookup in " & listNames(i) & " for '" & trimmed & "'."

      Call RecordDirectoryAttempt(info, listNames(i))

      Dim beforeFirst: beforeFirst = SafeText(info("First"))
      Dim beforeLast: beforeLast = SafeText(info("Last"))
      Dim beforeFull: beforeFull = SafeText(info("Full"))

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing

      If Len(firstName) > 0 And Len(lastName) > 0 Then
        Set entry = FindAddressEntryByFullName(addressList, firstName, lastName)
      End If

      If entry Is Nothing Then
        If Len(alternateFirst) > 0 And Len(alternateLast) > 0 Then
          Set entry = FindAddressEntryByFullName(addressList, alternateFirst, alternateLast)
        End If
      End If

      If entry Is Nothing Then
        Dim candidates: candidates = Array(normalized, trimmed, reverseDisplay)
        Dim j
        For j = LBound(candidates) To UBound(candidates)
          Dim candidate: candidate = SafeText(candidates(j))
          If Len(candidate) > 0 Then
            If entry Is Nothing Then
              Set entry = FindAddressEntryByFields(addressList, Array("Name", "DisplayName"), candidate)
            End If
          End If
        Next
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        On Error GoTo 0

        If Not exch Is Nothing Then
          Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
          Call ApplyFirstNameCandidate(info, exch.FirstName, "", "")
          If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
          If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
          Call PopulateNameFromPropertyAccessor(exch, info, "", "")
        End If

        Call PopulateNameFromPropertyAccessor(entry, info, "", "")

        Dim afterFirst: afterFirst = SafeText(info("First"))
        Dim afterLast: afterLast = SafeText(info("Last"))
        Dim afterFull: afterFull = SafeText(info("Full"))

        Dim improved: improved = False
        If Len(afterLast) > 0 And SafeLower(afterLast) <> SafeLower(beforeLast) Then improved = True
        If Not improved Then
          If Len(afterFirst) > 0 And SafeLower(afterFirst) <> SafeLower(beforeFirst) Then improved = True
        End If
        If Not improved Then
          If Len(afterFull) > 0 And SafeLower(afterFull) <> SafeLower(beforeFull) Then improved = True
        End If

        If improved Then
          info("DirectorySource") = listNames(i)
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing)
          Exit For
        End If
      End If
    End If
  Next
End Sub

Sub AttemptAliasRecipientResolve(ns, aliasValue, info, emailAddress)
  If ns Is Nothing Then Exit Sub

  Dim trimmedAlias: trimmedAlias = Trim(aliasValue)
  If Len(trimmedAlias) = 0 Then Exit Sub

  Dim beforeFirst: beforeFirst = SafeText(info("First"))
  Dim beforeLast: beforeLast = SafeText(info("Last"))
  Dim beforeFull: beforeFull = SafeText(info("Full"))

  Dim recipient: Set recipient = Nothing
  On Error Resume Next
  Set recipient = ns.CreateRecipient(trimmedAlias)
  If Err.Number <> 0 Then
    Err.Clear
    Set recipient = Nothing
  End If
  On Error GoTo 0

  If recipient Is Nothing Then Exit Sub

  On Error Resume Next
  recipient.Resolve
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0

  If Not recipient.Resolved Then Exit Sub

  info("Found") = True

  Dim exch: Set exch = Nothing
  On Error Resume Next
  Set exch = recipient.AddressEntry.GetExchangeUser()
  If Err.Number <> 0 Then
    Err.Clear
    Set exch = Nothing
  End If
  On Error GoTo 0

  If Len(SafeText(info("Full"))) = 0 Then info("Full") = recipient.Name

  If Not exch Is Nothing Then
    Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
    Call ApplyFirstNameCandidate(info, exch.FirstName, trimmedAlias, emailAddress)
    If Len(SafeText(info("Last"))) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
    If Len(SafeText(info("Full"))) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
    Call PopulateNameFromPropertyAccessor(exch, info, trimmedAlias, emailAddress)
  End If

  Call PopulateNameFromPropertyAccessor(recipient.AddressEntry, info, trimmedAlias, emailAddress)

  Dim afterFirst: afterFirst = SafeText(info("First"))
  Dim afterLast: afterLast = SafeText(info("Last"))
  Dim afterFull: afterFull = SafeText(info("Full"))

  Dim improved: improved = False
  If Len(afterLast) > 0 And SafeLower(afterLast) <> SafeLower(beforeLast) Then improved = True
  If Not improved Then
    If Len(afterFirst) > 0 And SafeLower(afterFirst) <> SafeLower(beforeFirst) Then improved = True
  End If
  If Not improved Then
    If Len(afterFull) > 0 And SafeLower(afterFull) <> SafeLower(beforeFull) Then improved = True
  End If

  Dim aliasLabel: aliasLabel = "Recipient resolve (alias '" & trimmedAlias & "')"
  Call RecordDirectoryAttempt(info, aliasLabel)

  If improved Then
    info("DirectorySource") = aliasLabel
    WScript.Echo "  Alias recipient resolve succeeded for '" & trimmedAlias & "'."
    Call ReportAddressListLookup(emailAddress, aliasLabel, recipient.AddressEntry, exch, info, Nothing)
  End If
End Sub

Sub AttemptRecipientResolveForValue(ns, rawValue, info, descriptor)
  If ns Is Nothing Then Exit Sub
  If info Is Nothing Then Exit Sub

  Dim trimmed: trimmed = Trim(rawValue)
  If Len(trimmed) = 0 Then Exit Sub

  Dim beforeFirst: beforeFirst = SafeText(info("First"))
  Dim beforeLast: beforeLast = SafeText(info("Last"))
  Dim beforeFull: beforeFull = SafeText(info("Full"))

  Dim label: label = "Recipient resolve (" & descriptor & " '" & trimmed & "')"
  Call RecordDirectoryAttempt(info, label)

  On Error Resume Next
  Dim recipient: Set recipient = ns.CreateRecipient(trimmed)
  If Err.Number <> 0 Then
    Err.Clear
    Set recipient = Nothing
  End If

  Dim resolvedOk: resolvedOk = False
  Dim exch: Set exch = Nothing

  If Not recipient Is Nothing Then
    recipient.Resolve
    If Err.Number = 0 Then
      resolvedOk = recipient.Resolved
    Else
      Err.Clear
    End If
  End If

  If resolvedOk Then
    Set exch = recipient.AddressEntry.GetExchangeUser()
    If Err.Number <> 0 Then
      Err.Clear
      Set exch = Nothing
    End If
  End If
  On Error GoTo 0

  If Not resolvedOk Then Exit Sub

  info("Found") = True

  If Len(SafeText(info("Full"))) = 0 Then info("Full") = recipient.Name

  If Not exch Is Nothing Then
    Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
    Call ApplyFirstNameCandidate(info, exch.FirstName, "", "")
    If Len(SafeText(info("Last"))) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
    If Len(SafeText(info("Full"))) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
    Call PopulateNameFromPropertyAccessor(exch, info, "", "")
  End If

  Call PopulateNameFromPropertyAccessor(recipient.AddressEntry, info, "", "")

  Dim afterFirst: afterFirst = SafeText(info("First"))
  Dim afterLast: afterLast = SafeText(info("Last"))
  Dim afterFull: afterFull = SafeText(info("Full"))

  Dim improved: improved = False
  If Len(afterLast) > 0 And SafeLower(afterLast) <> SafeLower(beforeLast) Then improved = True
  If Not improved Then
    If Len(afterFirst) > 0 And SafeLower(afterFirst) <> SafeLower(beforeFirst) Then improved = True
  End If
  If Not improved Then
    If Len(afterFull) > 0 And SafeLower(afterFull) <> SafeLower(beforeFull) Then improved = True
  End If

  If improved Then
    info("DirectorySource") = label
    WScript.Echo "  Recipient resolve succeeded for " & descriptor & " '" & trimmed & "'."
    Call ReportAddressListLookup(trimmed, label, recipient.AddressEntry, exch, info, Nothing)
  End If
End Sub

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

Function FindAddressEntryByFullName(addressList, firstName, lastName)
  Set FindAddressEntryByFullName = Nothing
  If addressList Is Nothing Then Exit Function

  Dim trimmedFirst: trimmedFirst = Trim(firstName)
  Dim trimmedLast: trimmedLast = Trim(lastName)
  If Len(trimmedFirst) = 0 Or Len(trimmedLast) = 0 Then Exit Function

  Dim escapedFirst: escapedFirst = EscapeRestrictionValue(trimmedFirst)
  Dim escapedLast: escapedLast = EscapeRestrictionValue(trimmedLast)

  Dim filters: filters = Array( _
    "[FirstName] = '" & escapedFirst & "' AND [LastName] = '" & escapedLast & "'", _
    "[GivenName] = '" & escapedFirst & "' AND [Surname] = '" & escapedLast & "'", _
    "[GivenName] = '" & escapedFirst & "' AND [LastName] = '" & escapedLast & "'", _
    "[FirstName] = '" & escapedFirst & "' AND [Surname] = '" & escapedLast & "'" _
  )

  Dim i, filterText, entry
  On Error Resume Next
  For i = LBound(filters) To UBound(filters)
    filterText = filters(i)
    If Len(filterText) > 0 Then
      Set entry = addressList.AddressEntries.Find(filterText)
      If Err.Number = 0 Then
        If Not entry Is Nothing Then
          Set FindAddressEntryByFullName = entry
          Exit For
        End If
      Else
        Err.Clear
      End If
    End If
    Set entry = Nothing
  Next
  On Error GoTo 0
End Function

Function EscapeRestrictionValue(value)
  Dim text: text = CStr(value)
  EscapeRestrictionValue = Replace(text, "'", "''")
End Function

Function FindAddressEntryByFields(addressList, fieldNames, searchValue)
  Set FindAddressEntryByFields = Nothing
  If addressList Is Nothing Then Exit Function
  If Not IsArray(fieldNames) Then Exit Function

  Dim trimmed: trimmed = Trim(searchValue)
  If Len(trimmed) = 0 Then Exit Function

  Dim query: query = EscapeRestrictionValue(trimmed)
  Dim i, field, entry

  On Error Resume Next
  For i = LBound(fieldNames) To UBound(fieldNames)
    field = Trim(CStr(fieldNames(i)))
    If Len(field) > 0 Then
      Set entry = addressList.AddressEntries.Find("[" & field & "] = '" & query & "'")
      If Err.Number = 0 Then
        If Not entry Is Nothing Then
          Set FindAddressEntryByFields = entry
          Exit For
        End If
      Else
        Err.Clear
      End If
    End If
    Set entry = Nothing
  Next
  On Error GoTo 0
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

Sub ReportAddressListLookup(emailAddress, listName, entry, exch, info, contact)
  WScript.Echo "  Directory lookup for " & emailAddress & " via " & listName & ":"

  Dim entryName
  entryName = ""
  If Not entry Is Nothing Then
    On Error Resume Next
    entryName = entry.Name
    If Err.Number <> 0 Then
      Err.Clear
      entryName = ""
    End If
    On Error GoTo 0
  End If
  If Len(entryName) > 0 Then
    WScript.Echo "    Address entry name: " & entryName
  Else
    WScript.Echo "    Address entry name: (empty)"
  End If

  Dim contactName, contactFirst, contactLast
  contactName = ""
  contactFirst = ""
  contactLast = ""
  If Not contact Is Nothing Then
    On Error Resume Next
    contactName = contact.FullName
    If Err.Number <> 0 Then
      Err.Clear
      contactName = ""
    End If
    contactFirst = contact.FirstName
    If Err.Number <> 0 Then
      Err.Clear
      contactFirst = ""
    End If
    contactLast = contact.LastName
    If Err.Number <> 0 Then
      Err.Clear
      contactLast = ""
    End If
    On Error GoTo 0
  End If

  If Len(contactName) > 0 Or Len(contactFirst) > 0 Or Len(contactLast) > 0 Then
    If Len(contactName) > 0 Then
      WScript.Echo "    Contact full name: " & contactName
    Else
      WScript.Echo "    Contact full name: (empty)"
    End If
    If Len(contactFirst) > 0 Then
      WScript.Echo "    Contact first name: " & contactFirst
    Else
      WScript.Echo "    Contact first name: (empty)"
    End If
    If Len(contactLast) > 0 Then
      WScript.Echo "    Contact surname: " & contactLast
    Else
      WScript.Echo "    Contact surname: (empty)"
    End If
  End If

  Dim exchName, exchFirst, exchLast
  exchName = ""
  exchFirst = ""
  exchLast = ""
  If Not exch Is Nothing Then
    On Error Resume Next
    exchName = exch.Name
    If Err.Number <> 0 Then
      Err.Clear
      exchName = ""
    End If
    exchFirst = exch.FirstName
    If Err.Number <> 0 Then
      Err.Clear
      exchFirst = ""
    End If
    exchLast = exch.LastName
    If Err.Number <> 0 Then
      Err.Clear
      exchLast = ""
    End If
    On Error GoTo 0
  End If

  If Len(exchName) > 0 Or Len(exchFirst) > 0 Or Len(exchLast) > 0 Then
    If Len(exchName) > 0 Then
      WScript.Echo "    Exchange display name: " & exchName
    Else
      WScript.Echo "    Exchange display name: (empty)"
    End If
    If Len(exchFirst) > 0 Then
      WScript.Echo "    Exchange first name: " & exchFirst
    Else
      WScript.Echo "    Exchange first name: (empty)"
    End If
    If Len(exchLast) > 0 Then
      WScript.Echo "    Exchange last name: " & exchLast
    Else
      WScript.Echo "    Exchange last name: (empty)"
    End If
  Else
    WScript.Echo "    Exchange information: (not available)"
  End If

  Dim resolvedFirst, resolvedLast, resolvedFull
  resolvedFirst = ""
  resolvedLast = ""
  resolvedFull = ""
  Dim resolvedEmail: resolvedEmail = ""
  If Not info Is Nothing Then
    If info.Exists("First") Then resolvedFirst = SafeText(info("First"))
    If info.Exists("Last") Then resolvedLast = SafeText(info("Last"))
    If info.Exists("Full") Then resolvedFull = SafeText(info("Full"))
    If info.Exists("Email") Then resolvedEmail = SafeText(info("Email"))
  End If

  If Len(resolvedFirst) > 0 Then
    WScript.Echo "    Resolved first name: " & resolvedFirst
  Else
    WScript.Echo "    Resolved first name: (empty)"
  End If
  If Len(resolvedLast) > 0 Then
    WScript.Echo "    Resolved surname: " & resolvedLast
  Else
    WScript.Echo "    Resolved surname: (empty)"
  End If
  If Len(resolvedFull) > 0 Then
    WScript.Echo "    Resolved full name: " & resolvedFull
  Else
    WScript.Echo "    Resolved full name: (empty)"
  End If
  If Len(resolvedEmail) > 0 Then
    WScript.Echo "    Resolved email: " & resolvedEmail
  Else
    WScript.Echo "    Resolved email: (empty)"
  End If
End Sub

Function BuildDirectorySearchList(ns)
  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  list.Add "Global Address List"
  list.Add "Offline Global Address List"

  If Not ns Is Nothing Then
    On Error Resume Next
    Dim addressList, folder
    For Each addressList In ns.AddressLists
      Err.Clear
      Set folder = addressList.GetContactsFolder()
      If Err.Number = 0 Then
        If list.IndexOf(addressList.Name) = -1 Then list.Add addressList.Name
      Else
        Err.Clear
      End If
    Next
    On Error GoTo 0
  End If

  If list.Count = 0 Then
    BuildDirectorySearchList = Array()
  Else
    BuildDirectorySearchList = list.ToArray()
  End If
End Function

Function FindContactInAddressList(addressList, emailAddress)
  Set FindContactInAddressList = Nothing
  If addressList Is Nothing Then Exit Function
  Dim lowered: lowered = SafeLower(emailAddress)
  If Len(lowered) = 0 Then Exit Function

  On Error Resume Next
  Dim contactsFolder: Set contactsFolder = addressList.GetContactsFolder()
  If Err.Number <> 0 Then
    Err.Clear
    Set contactsFolder = Nothing
  End If
  If contactsFolder Is Nothing Then
    On Error GoTo 0
    Exit Function
  End If

  Dim items: Set items = contactsFolder.Items
  If Err.Number <> 0 Then
    Err.Clear
    Set items = Nothing
  End If
  If items Is Nothing Then
    On Error GoTo 0
    Exit Function
  End If

  Dim contact
  For Each contact In items
    If ContactMatchesEmail(contact, lowered) Then
      Set FindContactInAddressList = contact
      Exit For
    End If
  Next
  On Error GoTo 0
End Function

Function ContactMatchesEmail(contact, loweredTarget)
  ContactMatchesEmail = False
  If contact Is Nothing Then Exit Function
  If Len(loweredTarget) = 0 Then Exit Function

  On Error Resume Next
  Dim candidates: candidates = Array( _
    SafeLower(contact.Email1Address), _
    SafeLower(contact.Email2Address), _
    SafeLower(contact.Email3Address))
  On Error GoTo 0

  Dim i
  For i = LBound(candidates) To UBound(candidates)
    If Len(candidates(i)) > 0 Then
      If candidates(i) = loweredTarget Then
        ContactMatchesEmail = True
        Exit For
      End If
    End If
  Next
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

Sub PopulateNameFromContactItem(contact, info)
  If contact Is Nothing Then Exit Sub
  If info Is Nothing Then Exit Sub

  On Error Resume Next
  Dim firstName: firstName = contact.FirstName
  If Err.Number <> 0 Then
    Err.Clear
    firstName = ""
  End If
  Dim lastName: lastName = contact.LastName
  If Err.Number <> 0 Then
    Err.Clear
    lastName = ""
  End If
  Dim fullName: fullName = contact.FullName
  If Err.Number <> 0 Then
    Err.Clear
    fullName = ""
  End If
  On Error GoTo 0

  If Len(firstName) > 0 Then Call ApplyFirstNameCandidate(info, firstName, "", "")
  If Len(info("Last")) = 0 And Len(lastName) > 0 Then info("Last") = lastName
  If Len(info("Full")) = 0 And Len(fullName) > 0 Then info("Full") = fullName

  Dim emailSet
  emailSet = Array(SafeText(contact.Email1Address), SafeText(contact.Email2Address), SafeText(contact.Email3Address))
  Dim i
  For i = LBound(emailSet) To UBound(emailSet)
    Call ApplyEmailCandidate(info, emailSet(i))
  Next

  Call PopulateNameFromPropertyAccessor(contact, info, "", "")
End Sub

Function SafeLower(value)
  SafeLower = ""
  If IsNull(value) Then Exit Function
  If IsEmpty(value) Then Exit Function
  Dim text: text = CStr(value)
  If Len(text) = 0 Then Exit Function
  SafeLower = LCase(text)
End Function

Sub PopulateNameFromPropertyAccessor(addressEntry, info, aliasValue, emailAddress)
  If addressEntry Is Nothing Then Exit Sub

  On Error Resume Next
  Dim accessor
  Set accessor = addressEntry.PropertyAccessor
  If Err.Number <> 0 Then
    Err.Clear
    Set accessor = Nothing
  End If
  If accessor Is Nothing Then
    On Error GoTo 0
    Exit Sub
  End If

  Dim given, surname, display, foundAny
  Dim currentFirst: currentFirst = ""
  If info.Exists("First") Then currentFirst = SafeText(info("First"))
  given = GetAccessorString(accessor, Array( _
    "http://schemas.microsoft.com/mapi/proptag/0x3A06001F", _
    "http://schemas.microsoft.com/mapi/proptag/0x3A06001E"))
  If Len(given) > 0 Then
    Dim shouldReplace: shouldReplace = ShouldReplaceFirstName(currentFirst, given, aliasValue, emailAddress)
    If shouldReplace Or Len(currentFirst) = 0 Then
      info("First") = given
      foundAny = True
      currentFirst = SafeText(info("First"))
    End If
  End If

  surname = GetAccessorString(accessor, Array( _
    "http://schemas.microsoft.com/mapi/proptag/0x3A11001F", _
    "http://schemas.microsoft.com/mapi/proptag/0x3A11001E"))
  If Len(info("Last")) = 0 And Len(surname) > 0 Then
    info("Last") = surname
    foundAny = True
  End If

  display = GetAccessorString(accessor, Array( _
    "http://schemas.microsoft.com/mapi/proptag/0x3001001F", _
    "http://schemas.microsoft.com/mapi/proptag/0x3001001E"))
  If Len(info("Full")) = 0 And Len(display) > 0 Then
    info("Full") = display
    foundAny = True
  End If

  Dim emailValue
  emailValue = GetAccessorString(accessor, Array( _
    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F", _
    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E", _
    "http://schemas.microsoft.com/mapi/proptag/0x3003001F", _
    "http://schemas.microsoft.com/mapi/proptag/0x3003001E"))
  Call ApplyEmailCandidate(info, emailValue)

  If Len(Trim(SafeText(info("Email")))) = 0 Then
    Dim proxyValue
    proxyValue = GetAccessorString(accessor, Array( _
      "http://schemas.microsoft.com/mapi/proptag/0x39FF001F", _
      "http://schemas.microsoft.com/mapi/proptag/0x39FF001E"))
    Call ApplyEmailCandidate(info, proxyValue)
  End If

  If Len(Trim(SafeText(info("Email")))) = 0 Then
    On Error Resume Next
    Dim directAddress
    directAddress = addressEntry.Address
    If Err.Number <> 0 Then
      Err.Clear
      directAddress = ""
    End If
    On Error GoTo 0
    Call ApplyEmailCandidate(info, directAddress)
  End If

  If foundAny Then info("Found") = True
  On Error GoTo 0
End Sub

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
