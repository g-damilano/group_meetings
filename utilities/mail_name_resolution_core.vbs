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

    If Not ResolutionWasSuccessful(info) Then
      Call EnumerateDirectoryMatchesByDisplayName(ns, trimmed, info, "name '" & trimmed & "'")
    End If
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
