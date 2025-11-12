' Directory lookup utilities for mail name resolution

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

Sub ReportAddressListLookup(emailAddress, listName, entry, exch, info, contact, selectionDetails)
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

  Dim selectionInfo
  Set selectionInfo = Nothing
  If IsObject(selectionDetails) Then
    On Error Resume Next
    Set selectionInfo = selectionDetails
    If Err.Number <> 0 Then
      Err.Clear
      Set selectionInfo = Nothing
    End If
    On Error GoTo 0
  End If

  If Not selectionInfo Is Nothing Then
    Dim candidateCount: candidateCount = 0
    If selectionInfo.Exists("CandidateCount") Then
      On Error Resume Next
      candidateCount = CLng(selectionInfo("CandidateCount"))
      If Err.Number <> 0 Then
        Err.Clear
        candidateCount = 0
      End If
      On Error GoTo 0
    End If

    If candidateCount > 1 Then
      WScript.Echo "    Candidate disambiguation: " & candidateCount & " candidates considered."
    ElseIf candidateCount = 1 Then
      WScript.Echo "    Candidate disambiguation: single candidate considered."
    End If

    Dim reasonText: reasonText = ""
    If selectionInfo.Exists("Reason") Then reasonText = Trim(SafeText(selectionInfo("Reason")))
    If Len(reasonText) > 0 Then
      WScript.Echo "    Candidate selection: " & reasonText
    End If

    If selectionInfo.Exists("Score") Then
      If IsNumeric(selectionInfo("Score")) Then
        WScript.Echo "    Candidate score: " & selectionInfo("Score")
      End If
    End If

    Dim selectedDept: selectedDept = ""
    Dim selectedCompany: selectedCompany = ""
    If selectionInfo.Exists("Department") Then selectedDept = Trim(SafeText(selectionInfo("Department")))
    If selectionInfo.Exists("Company") Then selectedCompany = Trim(SafeText(selectionInfo("Company")))

    Dim orgSummary: orgSummary = ""
    If Len(selectedDept) > 0 Then orgSummary = "department '" & selectedDept & "'"
    If Len(selectedCompany) > 0 Then
      If Len(orgSummary) > 0 Then
        orgSummary = orgSummary & ", "
      End If
      orgSummary = orgSummary & "company '" & selectedCompany & "'"
    End If
    If Len(orgSummary) > 0 Then
      WScript.Echo "    Selected candidate org: " & orgSummary
    End If

    If candidateCount > 1 And selectionInfo.Exists("UsedProfile") Then
      Dim profileUsed: profileUsed = False
      On Error Resume Next
      profileUsed = CBool(selectionInfo("UsedProfile"))
      If Err.Number <> 0 Then
        Err.Clear
        profileUsed = False
      End If
      On Error GoTo 0

      If profileUsed Then
        WScript.Echo "    Caller profile data was used for disambiguation."
      Else
        WScript.Echo "    Caller profile data was not available for disambiguation."
      End If
    End If
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

Function CollectAddressEntriesByFilters(addressList, filters)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
  Set CollectAddressEntriesByFilters = results

  If addressList Is Nothing Then Exit Function
  If Not IsArray(filters) Then Exit Function

  Dim seen: Set seen = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  seen.CompareMode = 1
  On Error GoTo 0

  Dim i, filterText, entry
  On Error Resume Next
  For i = LBound(filters) To UBound(filters)
    filterText = Trim(SafeText(filters(i)))
    If Len(filterText) > 0 Then
      Set entry = addressList.AddressEntries.Find(filterText)
      If Err.Number = 0 Then
        Do While Not entry Is Nothing
          Call AddAddressEntryToCollection(results, entry, seen)
          Set entry = addressList.AddressEntries.FindNext()
          If Err.Number <> 0 Then
            Err.Clear
            Exit Do
          End If
        Loop
      Else
        Err.Clear
      End If
    End If
    Set entry = Nothing
  Next
  On Error GoTo 0
End Function

Function CollectAddressEntriesByFields(addressList, fieldNames, searchValue)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
  Set CollectAddressEntriesByFields = results

  If addressList Is Nothing Then Exit Function
  If Not IsArray(fieldNames) Then Exit Function

  Dim trimmed: trimmed = Trim(searchValue)
  If Len(trimmed) = 0 Then Exit Function

  Dim query: query = EscapeRestrictionValue(trimmed)

  Dim seen: Set seen = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  seen.CompareMode = 1
  On Error GoTo 0

  Dim i, field, entry
  On Error Resume Next
  For i = LBound(fieldNames) To UBound(fieldNames)
    field = Trim(CStr(fieldNames(i)))
    If Len(field) > 0 Then
      Set entry = addressList.AddressEntries.Find("[" & field & "] = '" & query & "'")
      If Err.Number = 0 Then
        Do While Not entry Is Nothing
          Call AddAddressEntryToCollection(results, entry, seen)
          Set entry = addressList.AddressEntries.FindNext()
          If Err.Number <> 0 Then
            Err.Clear
            Exit Do
          End If
        Loop
      Else
        Err.Clear
      End If
    End If
    Set entry = Nothing
  Next
  On Error GoTo 0
End Function

Sub AddAddressEntryToCollection(targetList, entry, seen)
  If targetList Is Nothing Then Exit Sub
  If entry Is Nothing Then Exit Sub

  Dim key: key = ""

  If Not seen Is Nothing Then
    key = BuildAddressEntryTrackingKey(entry)

    If Len(key) > 0 Then
      If Not seen.Exists(key) Then
        seen.Add key, True
        targetList.Add entry
      End If
      Exit Sub
    End If
  End If

  targetList.Add entry
End Sub

Sub MergeAddressEntryMatches(targetList, matches, seen, sourceMap, listName)
  If targetList Is Nothing Then Exit Sub
  If sourceMap Is Nothing Then Exit Sub

  Dim candidate
  If matches Is Nothing Then Exit Sub

  For Each candidate In matches
    Dim beforeCount: beforeCount = 0
    On Error Resume Next
    beforeCount = targetList.Count
    If Err.Number <> 0 Then
      Err.Clear
      beforeCount = 0
    End If
    On Error GoTo 0

    Call AddAddressEntryToCollection(targetList, candidate, seen)

    Dim afterCount: afterCount = beforeCount
    On Error Resume Next
    afterCount = targetList.Count
    If Err.Number <> 0 Then
      Err.Clear
      afterCount = beforeCount
    End If
    On Error GoTo 0

    If afterCount > beforeCount Then
      Dim entryKey: entryKey = BuildAddressEntryTrackingKey(candidate)
      If Len(entryKey) > 0 Then
        On Error Resume Next
        If Not sourceMap.Exists(entryKey) Then sourceMap.Add entryKey, listName
        On Error GoTo 0
      End If
    End If
  Next
End Sub

Function DescribeAddressEntry(entry)
  DescribeAddressEntry = "(none)"
  If entry Is Nothing Then Exit Function

  Dim namePart: namePart = ""
  Dim addressPart: addressPart = ""

  On Error Resume Next
  namePart = SafeText(entry.Name)
  If Err.Number <> 0 Then
    Err.Clear
    namePart = ""
  End If
  addressPart = SafeText(entry.Address)
  If Err.Number <> 0 Then
    Err.Clear
    addressPart = ""
  End If
  On Error GoTo 0

  Dim summary: summary = Trim(namePart)
  If Len(summary) = 0 Then summary = "(unnamed entry)"
  If Len(addressPart) > 0 Then summary = summary & " <" & addressPart & ">"

  DescribeAddressEntry = summary
End Function

Function BuildAddressEntryTrackingKey(entry)
  BuildAddressEntryTrackingKey = ""
  If entry Is Nothing Then Exit Function

  Dim key: key = ""

  On Error Resume Next
  key = SafeLower(entry.ID)
  If Err.Number <> 0 Then
    Err.Clear
    key = ""
  End If
  On Error GoTo 0

  If Len(key) = 0 Then
    Dim namePart: namePart = SafeLower(SafeText(entry.Name))
    Dim addrPart: addrPart = SafeLower(SafeText(entry.Address))
    key = Trim(namePart & "|" & addrPart)
  End If

  BuildAddressEntryTrackingKey = key
End Function

Sub AppendAddressEntryResults(targetList, sourceList, seen)
  If targetList Is Nothing Then Exit Sub
  If sourceList Is Nothing Then Exit Sub

  Dim entry
  For Each entry In sourceList
    Call AddAddressEntryToCollection(targetList, entry, seen)
  Next
End Sub

Function FindAddressEntryByFullName(addressList, firstName, lastName)
  Set FindAddressEntryByFullName = Nothing

  Dim matches: Set matches = CollectAddressEntriesByFullName(addressList, firstName, lastName)
  If matches Is Nothing Then Exit Function
  If matches.Count = 0 Then Exit Function

  Set FindAddressEntryByFullName = matches(0)
End Function

Function CollectAddressEntriesByFullName(addressList, firstName, lastName)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
  Set CollectAddressEntriesByFullName = results

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

  Set results = CollectAddressEntriesByFilters(addressList, filters)
  Set CollectAddressEntriesByFullName = results
End Function

Function EscapeRestrictionValue(value)
  Dim text: text = CStr(value)
  EscapeRestrictionValue = Replace(text, "'", "''")
End Function

Function FindAddressEntryByFields(addressList, fieldNames, searchValue)
  Set FindAddressEntryByFields = Nothing

  Dim matches: Set matches = CollectAddressEntriesByFields(addressList, fieldNames, searchValue)
  If matches Is Nothing Then Exit Function
  If matches.Count = 0 Then Exit Function

  Set FindAddressEntryByFields = matches(0)
End Function

Function BuildAddressEntryCandidate(entry)
  Set BuildAddressEntryCandidate = Nothing
  If entry Is Nothing Then Exit Function

  Dim candidate: Set candidate = CreateObject("Scripting.Dictionary")
  Set candidate("Entry") = entry
  Set candidate("Exchange") = Nothing
  candidate("Department") = ""
  candidate("Company") = ""
  Set candidate("MatchDetails") = CreateObject("System.Collections.ArrayList")

  Dim exch
  On Error Resume Next
  If exch Is Nothing Then Set exch = entry.GetExchangeUser()
  If Err.Number <> 0 Then
    Err.Clear
    Set exch = Nothing
  End If
  On Error GoTo 0

  If Not exch Is Nothing Then
    Set candidate("Exchange") = exch
    candidate("Department") = Trim(SafeText(exch.Department))
    candidate("Company") = Trim(SafeText(exch.CompanyName))
  End If

  If Len(SafeText(candidate("Department"))) = 0 Or Len(SafeText(candidate("Company"))) = 0 Then
    Dim accessor
    On Error Resume Next
    Set accessor = entry.PropertyAccessor
    If Err.Number <> 0 Then
      Err.Clear
      Set accessor = Nothing
    End If
    On Error GoTo 0

    If Not accessor Is Nothing Then
      If Len(SafeText(candidate("Department"))) = 0 Then
        candidate("Department") = Trim(SafeText(GetAccessorString(accessor, Array( _
          "http://schemas.microsoft.com/mapi/proptag/0x3A18001F", _
          "http://schemas.microsoft.com/mapi/proptag/0x3A18001E"))))
      End If
      If Len(SafeText(candidate("Company"))) = 0 Then
        candidate("Company") = Trim(SafeText(GetAccessorString(accessor, Array( _
          "http://schemas.microsoft.com/mapi/proptag/0x3A16001F", _
          "http://schemas.microsoft.com/mapi/proptag/0x3A16001E"))))
      End If
    End If
  End If

  Set BuildAddressEntryCandidate = candidate
End Function

Function CallerProfileHasOrgDetails(profile)
  CallerProfileHasOrgDetails = False
  If profile Is Nothing Then Exit Function

  Dim dept: dept = ""
  Dim company: company = ""

  If profile.Exists("Department") Then dept = Trim(SafeText(profile("Department")))
  If profile.Exists("Company") Then company = Trim(SafeText(profile("Company")))

  If Len(dept) > 0 Or Len(company) > 0 Then CallerProfileHasOrgDetails = True
End Function

Function ScoreEntryCandidate(candidate, profile)
  ScoreEntryCandidate = 0
  If candidate Is Nothing Then Exit Function
  If profile Is Nothing Then Exit Function
  If Not CallerProfileHasOrgDetails(profile) Then Exit Function

  Dim details: Set details = CreateObject("System.Collections.ArrayList")

  Dim callerDept: callerDept = ""
  Dim callerCompany: callerCompany = ""
  If profile.Exists("Department") Then callerDept = SafeLower(Trim(SafeText(profile("Department"))))
  If profile.Exists("Company") Then callerCompany = SafeLower(Trim(SafeText(profile("Company"))))
  Dim candidateDept: candidateDept = SafeLower(Trim(SafeText(candidate("Department"))))
  Dim candidateCompany: candidateCompany = SafeLower(Trim(SafeText(candidate("Company"))))

  Dim score: score = 0

  If Len(callerDept) > 0 And Len(candidateDept) > 0 Then
    If callerDept = candidateDept Then
      score = score + 80
      details.Add "department '" & SafeText(candidate("Department")) & "'"
    ElseIf InStr(callerDept, candidateDept) > 0 Or InStr(candidateDept, callerDept) > 0 Then
      score = score + 40
      details.Add "similar department '" & SafeText(candidate("Department")) & "'"
    End If
  End If

  If Len(callerCompany) > 0 And Len(candidateCompany) > 0 Then
    If callerCompany = candidateCompany Then
      score = score + 30
      details.Add "company '" & SafeText(candidate("Company")) & "'"
    ElseIf InStr(callerCompany, candidateCompany) > 0 Or InStr(candidateCompany, callerCompany) > 0 Then
      score = score + 15
      details.Add "similar company '" & SafeText(candidate("Company")) & "'"
    End If
  End If

  Set candidate("MatchDetails") = details
  ScoreEntryCandidate = score
End Function

Function ChooseBestAddressEntry(entries, profile)
  Set ChooseBestAddressEntry = Nothing

  If entries Is Nothing Then Exit Function

  Dim result: Set result = CreateObject("Scripting.Dictionary")
  result("CandidateCount") = 0
  result("Score") = 0
  result("Reason") = ""
  result("Department") = ""
  result("Company") = ""
  result("UsedProfile") = False
  Set result("Entry") = Nothing
  Set result("Exchange") = Nothing

  Dim count: count = 0
  On Error Resume Next
  count = entries.Count
  If Err.Number <> 0 Then
    Err.Clear
    count = 0
  End If
  On Error GoTo 0

  result("CandidateCount") = count
  If count = 0 Then
    Set ChooseBestAddressEntry = result
    Exit Function
  End If

  Dim useProfile: useProfile = CallerProfileHasOrgDetails(profile)
  result("UsedProfile") = useProfile

  Dim bestCandidate: Set bestCandidate = BuildAddressEntryCandidate(entries(0))
  Dim bestScore: bestScore = 0
  If useProfile Then
    bestScore = ScoreEntryCandidate(bestCandidate, profile)
  End If

  Dim i, candidate, candidateScore
  If count > 1 Then
    For i = 1 To count - 1
      Set candidate = BuildAddressEntryCandidate(entries(i))
      If useProfile Then
        candidateScore = ScoreEntryCandidate(candidate, profile)
        If candidateScore > bestScore Then
          Set bestCandidate = candidate
          bestScore = candidateScore
        End If
      End If
    Next
  End If

  result("Score") = bestScore
  If bestCandidate.Exists("Department") Then result("Department") = SafeText(bestCandidate("Department"))
  If bestCandidate.Exists("Company") Then result("Company") = SafeText(bestCandidate("Company"))

  Dim reason: reason = ""
  If count = 1 Then
    reason = "Single candidate returned."
  ElseIf Not useProfile Then
    reason = "Multiple candidates found; caller profile unavailable so first match selected."
  ElseIf bestScore > 0 Then
    Dim details: Set details = Nothing
    If bestCandidate.Exists("MatchDetails") Then
      Set details = bestCandidate("MatchDetails")
    End If
    Dim detailText: detailText = ""
    If Not details Is Nothing Then detailText = JoinTokens(details, ", ")
    If Len(detailText) > 0 Then
      reason = "Selected using caller profile match on " & detailText & "."
    Else
      reason = "Selected using caller profile scoring."
    End If
  Else
    If count > 1 Then
      reason = "Caller profile provided but no department/company match; defaulting to first candidate."
    Else
      reason = "Single candidate returned."
    End If
  End If
  result("Reason") = reason

  If bestCandidate.Exists("Entry") Then Set result("Entry") = bestCandidate("Entry")
  If bestCandidate.Exists("Exchange") Then Set result("Exchange") = bestCandidate("Exchange")

  Set ChooseBestAddressEntry = result
End Function

Function BuildSelectionSummary(selection)
  Set BuildSelectionSummary = Nothing
  If selection Is Nothing Then Exit Function
  If Not IsObject(selection) Then Exit Function

  Dim summary: Set summary = CreateObject("Scripting.Dictionary")

  If selection.Exists("CandidateCount") Then summary("CandidateCount") = selection("CandidateCount")
  If selection.Exists("Reason") Then summary("Reason") = selection("Reason")
  If selection.Exists("Score") Then summary("Score") = selection("Score")
  If selection.Exists("Department") Then summary("Department") = selection("Department")
  If selection.Exists("Company") Then summary("Company") = selection("Company")
  If selection.Exists("UsedProfile") Then summary("UsedProfile") = selection("UsedProfile")

  Set BuildSelectionSummary = summary
End Function

Function GetCurrentUserProfileInfo()
  Dim profile: Set profile = CreateObject("Scripting.Dictionary")
  profile("HasProfile") = False
  profile("Department") = ""
  profile("Company") = ""

  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    Set GetCurrentUserProfileInfo = profile
    Exit Function
  End If

  Dim ns
  On Error Resume Next
  Set ns = olApp.GetNamespace("MAPI")
  If Err.Number <> 0 Then
    Err.Clear
    Set ns = Nothing
  End If
  On Error GoTo 0
  If ns Is Nothing Then
    Set GetCurrentUserProfileInfo = profile
    Exit Function
  End If

  Dim currentUser
  On Error Resume Next
  Set currentUser = ns.CurrentUser
  If Err.Number <> 0 Then
    Err.Clear
    Set currentUser = Nothing
  End If
  On Error GoTo 0
  If currentUser Is Nothing Then
    Set GetCurrentUserProfileInfo = profile
    Exit Function
  End If

  Dim addressEntry
  On Error Resume Next
  Set addressEntry = currentUser.AddressEntry
  If Err.Number <> 0 Then
    Err.Clear
    Set addressEntry = Nothing
  End If
  On Error GoTo 0
  If addressEntry Is Nothing Then
    Set GetCurrentUserProfileInfo = profile
    Exit Function
  End If

  Dim department: department = ""
  Dim company: company = ""

  Dim exch
  On Error Resume Next
  Set exch = addressEntry.GetExchangeUser()
  If Err.Number <> 0 Then
    Err.Clear
    Set exch = Nothing
  End If
  On Error GoTo 0

  If Not exch Is Nothing Then
    department = SafeText(exch.Department)
    company = SafeText(exch.CompanyName)
  End If

  Dim accessor
  If Len(Trim(SafeText(department))) = 0 Or Len(Trim(SafeText(company))) = 0 Then
    On Error Resume Next
    Set accessor = addressEntry.PropertyAccessor
    If Err.Number <> 0 Then
      Err.Clear
      Set accessor = Nothing
    End If
    On Error GoTo 0

    If Not accessor Is Nothing Then
      If Len(Trim(SafeText(department))) = 0 Then
        department = SafeText(GetAccessorString(accessor, Array( _
          "http://schemas.microsoft.com/mapi/proptag/0x3A18001F", _
          "http://schemas.microsoft.com/mapi/proptag/0x3A18001E")))
      End If
      If Len(Trim(SafeText(company))) = 0 Then
        company = SafeText(GetAccessorString(accessor, Array( _
          "http://schemas.microsoft.com/mapi/proptag/0x3A16001F", _
          "http://schemas.microsoft.com/mapi/proptag/0x3A16001E")))
      End If
    End If
  End If

  profile("Department") = Trim(SafeText(department))
  profile("Company") = Trim(SafeText(company))
  If Len(profile("Department")) > 0 Or Len(profile("Company")) > 0 Then profile("HasProfile") = True

  Set GetCurrentUserProfileInfo = profile
End Function
