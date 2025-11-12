' Directory population helpers for name resolution diagnostics

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
  If Len(display) > 0 Then
    Dim beforeFull: beforeFull = SafeText(info("Full"))
    Call ApplyFullNameCandidate(info, display, emailAddress)
    If SafeLower(Trim(SafeText(info("Full")))) <> SafeLower(Trim(beforeFull)) Then
      foundAny = True
    End If
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
      If Err.Number = 0 Then
        If Not entry Is Nothing Then Call RecordLookupResponse(listNames(i), "SMTPAddress", entry)
      Else
        Err.Clear
        Set entry = Nothing
      End If
      If entry Is Nothing Then
        Set entry = addressList.AddressEntries.Find("[EmailAddress] = '" & emailAddress & "'")
        If Err.Number = 0 Then
          If Not entry Is Nothing Then Call RecordLookupResponse(listNames(i), "EmailAddress", entry)
        Else
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
          ReportAddressListLookup emailAddress, listNames(i), entry, exch, info, contact, Nothing
          Exit For
        End If
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name
        If exch Is Nothing Then Set exch = entry.GetExchangeUser()
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
        ReportAddressListLookup emailAddress, listNames(i), entry, exch, info, contact, Nothing
        Exit For
      End If
    End If
  Next
  If Len(info("DirectorySource")) = 0 And Len(SafeText(info("DirectoryAttempted"))) > 0 Then
    info("DirectorySource") = "(not found)"
  End If
  On Error GoTo 0
End Sub

Sub PopulateNameFromAddressListsByGivenName(ns, givenName, info, emailAddress)
  If ns Is Nothing Then Exit Sub
  If Len(givenName) = 0 Then Exit Sub

  Dim trimmed: trimmed = Trim(givenName)
  If Len(trimmed) = 0 Then Exit Sub

  Dim listNames: listNames = Array("Global Address List", "Offline Global Address List")
  Dim callerProfile: Set callerProfile = GetCurrentUserProfileInfo()
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

      Dim matches: Set matches = CollectAddressEntriesByFields(addressList, Array("Name", "FirstName", "Alias"), trimmed)
      Dim selection: Set selection = Nothing
      If Not matches Is Nothing Then
        If matches.Count > 0 Then Set selection = ChooseBestAddressEntry(matches, callerProfile)
      End If

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing
      Dim selectionSummary: Set selectionSummary = Nothing

      If Not selection Is Nothing Then
        If selection.Exists("Entry") Then Set entry = selection("Entry")
        If selection.Exists("Exchange") Then Set exch = selection("Exchange")
        Set selectionSummary = BuildSelectionSummary(selection)
      End If

      If entry Is Nothing Then
        Set entry = FindAddressEntryByFields(addressList, Array("Name", "FirstName", "Alias"), trimmed)
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        If exch Is Nothing Then Set exch = entry.GetExchangeUser()
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
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing, selectionSummary)
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
  Dim callerProfile: Set callerProfile = GetCurrentUserProfileInfo()
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

      Dim matches: Set matches = CollectAddressEntriesByFields(addressList, Array("LastName", "Surname", "Name"), trimmed)
      Dim selection: Set selection = Nothing
      If Not matches Is Nothing Then
        If matches.Count > 0 Then Set selection = ChooseBestAddressEntry(matches, callerProfile)
      End If

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing
      Dim selectionSummary: Set selectionSummary = Nothing

      If Not selection Is Nothing Then
        If selection.Exists("Entry") Then Set entry = selection("Entry")
        If selection.Exists("Exchange") Then Set exch = selection("Exchange")
        Set selectionSummary = BuildSelectionSummary(selection)
      End If

      If entry Is Nothing Then
        Set entry = FindAddressEntryByFields(addressList, Array("LastName", "Surname", "Name"), trimmed)
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        If exch Is Nothing Then Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        On Error GoTo 0

        If Not exch Is Nothing Then
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
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing, selectionSummary)
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
  Dim callerProfile: Set callerProfile = GetCurrentUserProfileInfo()
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

      Dim matches: Set matches = CollectAddressEntriesByFields(addressList, Array("Name", "DisplayName"), trimmed)
      Dim selection: Set selection = Nothing
      If Not matches Is Nothing Then
        If matches.Count > 0 Then Set selection = ChooseBestAddressEntry(matches, callerProfile)
      End If

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing
      Dim selectionSummary: Set selectionSummary = Nothing

      If Not selection Is Nothing Then
        If selection.Exists("Entry") Then Set entry = selection("Entry")
        If selection.Exists("Exchange") Then Set exch = selection("Exchange")
        Set selectionSummary = BuildSelectionSummary(selection)
      End If

      If entry Is Nothing Then
        Set entry = FindAddressEntryByFields(addressList, Array("Name", "DisplayName"), trimmed)
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        If exch Is Nothing Then Set exch = entry.GetExchangeUser()
        If Err.Number <> 0 Then
          Err.Clear
          Set exch = Nothing
        End If
        On Error GoTo 0

        If Not exch Is Nothing Then
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
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing, selectionSummary)
          Exit For
        End If
      End If
    End If
  Next
End Sub

Sub EnumerateDirectoryMatchesByDisplayName(ns, displayName, info, descriptor)
  If ns Is Nothing Then Exit Sub
  If info Is Nothing Then Exit Sub

  Dim trimmed: trimmed = Trim(displayName)
  If Len(trimmed) = 0 Then Exit Sub

  WScript.Echo "  Enumerating directory matches for display name fallback '" & trimmed & "'."

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

  Dim candidateValues: Set candidateValues = CreateObject("Scripting.Dictionary")
  If candidateValues Is Nothing Then Exit Sub
  On Error Resume Next
  candidateValues.CompareMode = 1
  On Error GoTo 0

  Call AddCandidateValue(candidateValues, trimmed)
  Call AddCandidateValue(candidateValues, normalized)
  Call AddCandidateValue(candidateValues, reverseDisplay)

  If Len(firstName) > 0 And Len(lastName) > 0 Then
    Call AddCandidateValue(candidateValues, Trim(firstName & " " & lastName))
    Call AddCandidateValue(candidateValues, Trim(lastName & " " & firstName))
    Call AddCandidateValue(candidateValues, Trim(lastName & ", " & firstName))
    Call AddCandidateValue(candidateValues, Trim(lastName & "," & firstName))
  End If

  If Len(alternateFirst) > 0 And Len(alternateLast) > 0 Then
    Call AddCandidateValue(candidateValues, Trim(alternateFirst & " " & alternateLast))
    Call AddCandidateValue(candidateValues, Trim(alternateLast & " " & alternateFirst))
    Call AddCandidateValue(candidateValues, Trim(alternateLast & ", " & alternateFirst))
    Call AddCandidateValue(candidateValues, Trim(alternateLast & "," & alternateFirst))
  End If

  Dim listNames: listNames = BuildDirectorySearchList(ns)
  If Not IsArray(listNames) Then Exit Sub

  Dim candidateCount: candidateCount = 0
  On Error Resume Next
  candidateCount = candidateValues.Count
  If Err.Number <> 0 Then
    Err.Clear
    candidateCount = 0
  End If
  On Error GoTo 0

  WScript.Echo "    Candidate permutations generated: " & candidateCount

  Dim listCount
  listCount = 0
  On Error Resume Next
  listCount = UBound(listNames) - LBound(listNames) + 1
  If Err.Number <> 0 Then
    Err.Clear
    listCount = 0
  End If
  On Error GoTo 0
  If listCount = 0 Then Exit Sub

  Dim beforeFirst: beforeFirst = SafeText(info("First"))
  Dim beforeLast: beforeLast = SafeText(info("Last"))
  Dim beforeFull: beforeFull = SafeText(info("Full"))

  Dim aggregate: Set aggregate = CreateObject("System.Collections.ArrayList")

  Dim seen: Set seen = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  seen.CompareMode = 1
  On Error GoTo 0

  Dim sourceMap: Set sourceMap = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  sourceMap.CompareMode = 1
  On Error GoTo 0

  Dim callerProfile: Set callerProfile = GetCurrentUserProfileInfo()

  Dim i, listName
  Dim matches
  Dim candidate
  For i = LBound(listNames) To UBound(listNames)
    listName = listNames(i)

    Dim addressList: Set addressList = Nothing
    On Error Resume Next
    Set addressList = ns.AddressLists(listName)
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If Not addressList Is Nothing Then
      WScript.Echo "    Inspecting address list: " & listName
      Call RecordDirectoryAttempt(info, listName)

      Dim valueKey
      For Each valueKey In candidateValues.Keys
        Dim candidateValue: candidateValue = Trim(SafeText(valueKey))
        If Len(candidateValue) > 0 Then
          Set matches = CollectAddressEntriesByFields(addressList, Array("Name", "DisplayName", "FileAs"), candidateValue)
          Call MergeAddressEntryMatches(aggregate, matches, seen, sourceMap, listName)

          Dim candidateMatchCount: candidateMatchCount = 0
          If Not matches Is Nothing Then
            On Error Resume Next
            candidateMatchCount = matches.Count
            If Err.Number <> 0 Then
              Err.Clear
              candidateMatchCount = 0
            End If
            On Error GoTo 0
          End If

          Dim aggregateCount: aggregateCount = 0
          On Error Resume Next
          aggregateCount = aggregate.Count
          If Err.Number <> 0 Then
            Err.Clear
            aggregateCount = 0
          End If
          On Error GoTo 0

          WScript.Echo "      Candidate '" & candidateValue & "' produced " & candidateMatchCount & " matches (aggregate unique: " & aggregateCount & ")."
        End If
      Next

      If Len(firstName) > 0 And Len(lastName) > 0 Then
        Set matches = CollectAddressEntriesByFullName(addressList, firstName, lastName)
        Call MergeAddressEntryMatches(aggregate, matches, seen, sourceMap, listName)

        Dim fullMatchCount: fullMatchCount = 0
        If Not matches Is Nothing Then
          On Error Resume Next
          fullMatchCount = matches.Count
          If Err.Number <> 0 Then
            Err.Clear
            fullMatchCount = 0
          End If
          On Error GoTo 0
        End If

        Dim aggregateCountFull: aggregateCountFull = 0
        On Error Resume Next
        aggregateCountFull = aggregate.Count
        If Err.Number <> 0 Then
          Err.Clear
          aggregateCountFull = 0
        End If
        On Error GoTo 0

        WScript.Echo "      Full-name candidate '" & Trim(firstName & " " & lastName) & "' produced " & fullMatchCount & " matches (aggregate unique: " & aggregateCountFull & ")."
      End If

      If Len(alternateFirst) > 0 And Len(alternateLast) > 0 Then
        Set matches = CollectAddressEntriesByFullName(addressList, alternateFirst, alternateLast)
        Call MergeAddressEntryMatches(aggregate, matches, seen, sourceMap, listName)

        Dim altMatchCount: altMatchCount = 0
        If Not matches Is Nothing Then
          On Error Resume Next
          altMatchCount = matches.Count
          If Err.Number <> 0 Then
            Err.Clear
            altMatchCount = 0
          End If
          On Error GoTo 0
        End If

        Dim aggregateCountAlt: aggregateCountAlt = 0
        On Error Resume Next
        aggregateCountAlt = aggregate.Count
        If Err.Number <> 0 Then
          Err.Clear
          aggregateCountAlt = 0
        End If
        On Error GoTo 0

        WScript.Echo "      Alternate-name candidate '" & Trim(alternateFirst & " " & alternateLast) & "' produced " & altMatchCount & " matches (aggregate unique: " & aggregateCountAlt & ")."
      End If
    End If
  Next

  Dim aggregateTotal: aggregateTotal = 0
  On Error Resume Next
  aggregateTotal = aggregate.Count
  If Err.Number <> 0 Then
    Err.Clear
    aggregateTotal = 0
  End If
  On Error GoTo 0

  If aggregateTotal = 0 Then
    WScript.Echo "  Display-name enumeration found no address entries for '" & trimmed & "'."
    Exit Sub
  End If

  WScript.Echo "  Display-name enumeration aggregated " & aggregateTotal & " unique entries."

  Dim selection: Set selection = Nothing
  Set selection = ChooseBestAddressEntry(aggregate, callerProfile)

  Dim selectionSummary: Set selectionSummary = Nothing
  If Not selection Is Nothing Then
    Set selectionSummary = BuildSelectionSummary(selection)
  End If

  Dim entry: Set entry = Nothing
  Dim exch: Set exch = Nothing

  If Not selection Is Nothing Then
    If selection.Exists("Entry") Then Set entry = selection("Entry")
    If selection.Exists("Exchange") Then Set exch = selection("Exchange")
  End If

  If entry Is Nothing Then
    Set entry = aggregate(0)
  End If

  If entry Is Nothing Then
    WScript.Echo "  Display-name enumeration could not select a best candidate despite aggregated matches."
    Exit Sub
  End If

  WScript.Echo "  Display-name enumeration selected candidate: " & DescribeAddressEntry(entry)

  info("Found") = True
  If Len(SafeText(info("Full"))) = 0 Then info("Full") = entry.Name

  On Error Resume Next
  If exch Is Nothing Then Set exch = entry.GetExchangeUser()
  If Err.Number <> 0 Then
    Err.Clear
    Set exch = Nothing
  End If
  On Error GoTo 0

  If Not exch Is Nothing Then
    Call ApplyFirstNameCandidate(info, exch.FirstName, "", "")
    If Len(SafeText(info("Last"))) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
    If Len(SafeText(info("Full"))) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
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

  Dim sourceLabel: sourceLabel = ""
  Dim chosenKey: chosenKey = BuildAddressEntryTrackingKey(entry)
  If Len(chosenKey) > 0 Then
    On Error Resume Next
    If sourceMap.Exists(chosenKey) Then sourceLabel = CStr(sourceMap(chosenKey))
    On Error GoTo 0
  End If
  If Len(Trim(sourceLabel)) = 0 Then sourceLabel = "Display name enumeration"

  Dim lookupLabel: lookupLabel = descriptor
  If Len(Trim(lookupLabel)) = 0 Then lookupLabel = "display name '" & trimmed & "'"

  If improved Then
    info("DirectorySource") = sourceLabel
    Call ReportAddressListLookup(lookupLabel, sourceLabel, entry, exch, info, Nothing, selectionSummary)
    WScript.Echo "  Display-name enumeration updated resolution fields using " & sourceLabel & "."
  Else
    WScript.Echo "  Display-name enumeration selected '" & DescribeAddressEntry(entry) & "' but did not improve resolved fields."
  End If
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
  Dim callerProfile: Set callerProfile = GetCurrentUserProfileInfo()
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

      Dim aggregate: Set aggregate = CreateObject("System.Collections.ArrayList")
      Dim aggregateSeen: Set aggregateSeen = CreateObject("Scripting.Dictionary")
      On Error Resume Next
      aggregateSeen.CompareMode = 1
      On Error GoTo 0

      Dim matches
      If Len(firstName) > 0 And Len(lastName) > 0 Then
        Set matches = CollectAddressEntriesByFullName(addressList, firstName, lastName)
        Call AppendAddressEntryResults(aggregate, matches, aggregateSeen)
      End If

      If Len(alternateFirst) > 0 And Len(alternateLast) > 0 Then
        Set matches = CollectAddressEntriesByFullName(addressList, alternateFirst, alternateLast)
        Call AppendAddressEntryResults(aggregate, matches, aggregateSeen)
      End If

      Dim candidates: candidates = Array(normalized, trimmed, reverseDisplay)
      Dim j
      For j = LBound(candidates) To UBound(candidates)
        Dim candidate: candidate = SafeText(candidates(j))
        If Len(candidate) > 0 Then
          Set matches = CollectAddressEntriesByFields(addressList, Array("Name", "DisplayName"), candidate)
          Call AppendAddressEntryResults(aggregate, matches, aggregateSeen)
        End If
      Next

      Dim selection: Set selection = Nothing
      If Not aggregate Is Nothing Then
        If aggregate.Count > 0 Then Set selection = ChooseBestAddressEntry(aggregate, callerProfile)
      End If

      Dim entry, exch
      Set entry = Nothing
      Set exch = Nothing
      Dim selectionSummary: Set selectionSummary = Nothing

      If Not selection Is Nothing Then
        If selection.Exists("Entry") Then Set entry = selection("Entry")
        If selection.Exists("Exchange") Then Set exch = selection("Exchange")
        Set selectionSummary = BuildSelectionSummary(selection)
      End If

      If entry Is Nothing Then
        If Len(firstName) > 0 And Len(lastName) > 0 Then
          Set entry = FindAddressEntryByFullName(addressList, firstName, lastName)
        End If
      End If

      If entry Is Nothing Then
        If Len(alternateFirst) > 0 And Len(alternateLast) > 0 Then
          Set entry = FindAddressEntryByFullName(addressList, alternateFirst, alternateLast)
        End If
      End If

      If entry Is Nothing Then
        For j = LBound(candidates) To UBound(candidates)
          Dim fallbackCandidate: fallbackCandidate = SafeText(candidates(j))
          If Len(fallbackCandidate) > 0 Then
            If entry Is Nothing Then
              Set entry = FindAddressEntryByFields(addressList, Array("Name", "DisplayName"), fallbackCandidate)
            End If
          End If
        Next
      End If

      If Not entry Is Nothing Then
        info("Found") = True
        If Len(info("Full")) = 0 Then info("Full") = entry.Name

        On Error Resume Next
        If exch Is Nothing Then Set exch = entry.GetExchangeUser()
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
          Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing, selectionSummary)
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

  Call RecordRecipientResolutionResponse("Alias recipient resolve ('" & trimmedAlias & "')", recipient)

  info("Found") = True

  Dim exch: Set exch = Nothing
  On Error Resume Next
  Set exch = recipient.AddressEntry.GetExchangeUser()
  If Err.Number <> 0 Then
    Err.Clear
    Set exch = Nothing
  End If
  On Error GoTo 0

  Call ApplyFullNameCandidate(info, recipient.Name, emailAddress)

  If Not exch Is Nothing Then
    Call ApplyEmailCandidate(info, SafeText(exch.PrimarySmtpAddress))
    Call ApplyFirstNameCandidate(info, exch.FirstName, trimmedAlias, emailAddress)
    If Len(SafeText(info("Last"))) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
    Call ApplyFullNameCandidate(info, exch.Name, emailAddress)
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
    Call ReportAddressListLookup(emailAddress, aliasLabel, recipient.AddressEntry, exch, info, Nothing, Nothing)
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
    Call RecordRecipientResolutionResponse(label, recipient)
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
    Call ReportAddressListLookup(trimmed, label, recipient.AddressEntry, exch, info, Nothing, Nothing)
  End If
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

