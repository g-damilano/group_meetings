' Directory resolution utilities for personal deliveries

Function ResolveNameFromEmail(emailAddress)
  Dim info: Set info = CreateObject("Scripting.Dictionary")
  info("Found") = False
  info("First") = ""
  info("Last") = ""
  info("Full") = ""

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
  listNames = Array("Global Address List", "Offline Global Address List")

  Dim i, addressList, entry, exch
  On Error Resume Next
  For i = LBound(listNames) To UBound(listNames)
    Set addressList = Nothing
    Set addressList = ns.AddressLists(listNames(i))
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    If Not addressList Is Nothing Then
      Set entry = Nothing
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
        ReportAddressListLookup emailAddress, listNames(i), entry, exch, info
        Exit For
      End If
    End If
  Next
  On Error GoTo 0
End Sub

Sub ReportAddressListLookup(emailAddress, listName, entry, exch, info)
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
