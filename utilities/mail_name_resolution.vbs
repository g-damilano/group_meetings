' Directory resolution utilities for personal deliveries

Function ResolveNameFromEmail(emailAddress)
  Dim info: Set info = CreateObject("Scripting.Dictionary")
  info("Found") = False
  info("First") = ""
  info("Last") = ""
  info("Full") = ""
  info("DirectorySource") = ""
  info("DirectoryAttempted") = ""

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
    Dim exch: Set exch = recip.AddressEntry.GetExchangeUser()
    If Not exch Is Nothing Then
      If Len(exch.FirstName) > 0 Then info("First") = exch.FirstName
      If Len(exch.LastName) > 0 Then info("Last") = exch.LastName
      If Len(info("Full")) = 0 Then info("Full") = exch.Name
      Call PopulateNameFromPropertyAccessor(exch, info)
    End If
    Call PopulateNameFromPropertyAccessor(recip.AddressEntry, info)
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

Sub PopulateNameFromAddressLists(ns, emailAddress, info)
  If ns Is Nothing Then Exit Sub

  Dim listNames
  listNames = BuildDirectorySearchList(ns)

  Dim i, addressList, entry, exch, contact
  Dim attempts: attempts = ""
  On Error Resume Next
  For i = LBound(listNames) To UBound(listNames)
    Set addressList = Nothing
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    If Not addressList Is Nothing Then
      If Len(attempts) > 0 Then
        attempts = attempts & ", "
      End If
      attempts = attempts & listNames(i)
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
          info("DirectoryAttempted") = attempts
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
        If Not exch Is Nothing Then
          If Len(info("First")) = 0 And Len(exch.FirstName) > 0 Then info("First") = exch.FirstName
          If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
          If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
          Call PopulateNameFromPropertyAccessor(exch, info)
        End If
        Call PopulateNameFromPropertyAccessor(entry, info)
        info("DirectorySource") = listNames(i)
        info("DirectoryAttempted") = attempts
        ReportAddressListLookup emailAddress, listNames(i), entry, exch, info, contact
        Exit For
      End If
    End If
  Next
  If Len(info("DirectoryAttempted")) = 0 Then
    info("DirectoryAttempted") = attempts
  End If
  If Len(info("DirectorySource")) = 0 And Len(attempts) > 0 Then
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
  Call PopulateNameFromAddressListsByGivenName(ns, firstName, info, emailAddress)
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
    If addressList Is Nothing Then GoTo NextList

    WScript.Echo "  Attempting given-name lookup in " & listNames(i) & " for '" & trimmed & "'."

    Dim beforeFirst: beforeFirst = SafeText(info("First"))
    Dim beforeLast: beforeLast = SafeText(info("Last"))
    Dim beforeFull: beforeFull = SafeText(info("Full"))

    Dim entry, exch
    Set entry = Nothing
    Set exch = Nothing

    On Error Resume Next
    Dim query: query = EscapeRestrictionValue(trimmed)
    Set entry = addressList.AddressEntries.Find("[Name] = '" & query & "'")
    If Err.Number <> 0 Then
      Err.Clear
      Set entry = Nothing
    End If
    If entry Is Nothing Then
      Set entry = addressList.AddressEntries.Find("[FirstName] = '" & query & "'")
      If Err.Number <> 0 Then
        Err.Clear
        Set entry = Nothing
      End If
    End If
    On Error GoTo 0

    If entry Is Nothing Then GoTo NextList

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
      If Len(info("First")) = 0 And Len(exch.FirstName) > 0 Then info("First") = exch.FirstName
      If Len(info("Last")) = 0 And Len(exch.LastName) > 0 Then info("Last") = exch.LastName
      If Len(info("Full")) = 0 And Len(exch.Name) > 0 Then info("Full") = exch.Name
      Call PopulateNameFromPropertyAccessor(exch, info)
    End If
    Call PopulateNameFromPropertyAccessor(entry, info)

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

    Dim attempts: attempts = SafeText(info("DirectoryAttempted"))
    If InStr(1, attempts, listNames(i), vbTextCompare) = 0 Then
      If Len(attempts) > 0 Then
        attempts = attempts & ", " & listNames(i)
      Else
        attempts = listNames(i)
      End If
      info("DirectoryAttempted") = attempts
    End If

    If improved Then
      info("DirectorySource") = listNames(i)
      Call ReportAddressListLookup(emailAddress, listNames(i), entry, exch, info, Nothing)
      Exit For
    End If

NextList:
  Next
End Sub

Function EscapeRestrictionValue(value)
  Dim text: text = CStr(value)
  EscapeRestrictionValue = Replace(text, "'", "''")
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
  If Not info Is Nothing Then
    If info.Exists("First") Then resolvedFirst = SafeText(info("First"))
    If info.Exists("Last") Then resolvedLast = SafeText(info("Last"))
    If info.Exists("Full") Then resolvedFull = SafeText(info("Full"))
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

  If Len(info("First")) = 0 And Len(firstName) > 0 Then info("First") = firstName
  If Len(info("Last")) = 0 And Len(lastName) > 0 Then info("Last") = lastName
  If Len(info("Full")) = 0 And Len(fullName) > 0 Then info("Full") = fullName

  Call PopulateNameFromPropertyAccessor(contact, info)
End Sub

Function SafeLower(value)
  SafeLower = ""
  If IsNull(value) Then Exit Function
  If IsEmpty(value) Then Exit Function
  Dim text: text = CStr(value)
  If Len(text) = 0 Then Exit Function
  SafeLower = LCase(text)
End Function

Sub PopulateNameFromPropertyAccessor(addressEntry, info)
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
  given = GetAccessorString(accessor, Array( _
    "http://schemas.microsoft.com/mapi/proptag/0x3A06001F", _
    "http://schemas.microsoft.com/mapi/proptag/0x3A06001E"))
  If Len(info("First")) = 0 And Len(given) > 0 Then
    info("First") = given
    foundAny = True
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
