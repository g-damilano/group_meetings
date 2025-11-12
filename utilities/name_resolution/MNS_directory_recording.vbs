' Directory lookup response recording helpers for name resolution diagnostics

Dim gLookupResponseRecorder

Sub SetLookupResponseRecorder(recorder)
  Set gLookupResponseRecorder = Nothing
  If IsObject(recorder) Then
    Set gLookupResponseRecorder = recorder
  End If
End Sub

Sub ClearLookupResponseRecorder()
  Set gLookupResponseRecorder = Nothing
End Sub

Sub RecordLookupResponse(listName, fieldName, entry)
  If Not IsObject(gLookupResponseRecorder) Then Exit Sub

  Dim summary: summary = BuildLookupResponseSummary(listName, fieldName, entry)

  On Error Resume Next
  gLookupResponseRecorder.Add summary
  On Error GoTo 0
End Sub

Sub RecordRecipientResolutionResponse(label, recipient)
  If Not IsObject(gLookupResponseRecorder) Then Exit Sub

  Dim entry: Set entry = Nothing
  Dim entrySummary: entrySummary = "(no address entry)"

  If Not recipient Is Nothing Then
    On Error Resume Next
    Set entry = recipient.AddressEntry
    If Err.Number <> 0 Then
      Err.Clear
      Set entry = Nothing
    End If
    On Error GoTo 0
  End If

  If Not entry Is Nothing Then entrySummary = DescribeAddressEntry(entry)

  Dim resolvedName: resolvedName = ""
  Dim resolvedAddress: resolvedAddress = ""

  If Not recipient Is Nothing Then
    On Error Resume Next
    resolvedName = SafeText(recipient.Name)
    If Err.Number <> 0 Then
      Err.Clear
      resolvedName = ""
    End If
    resolvedAddress = SafeText(recipient.Address)
    If Err.Number <> 0 Then
      Err.Clear
      resolvedAddress = ""
    End If
    On Error GoTo 0
  End If

  Dim detail: detail = Trim(label)
  If Len(resolvedName) > 0 Then detail = detail & " -> Name: " & resolvedName
  If Len(resolvedAddress) > 0 Then detail = detail & ", Address: " & resolvedAddress
  detail = detail & " (" & entrySummary & ")"

  On Error Resume Next
  gLookupResponseRecorder.Add detail
  On Error GoTo 0
End Sub

Function BuildLookupResponseSummary(listName, fieldName, entry)
  Dim effectiveList: effectiveList = Trim(SafeText(listName))
  If Len(effectiveList) = 0 Then effectiveList = "(unknown list)"

  Dim effectiveField: effectiveField = Trim(SafeText(fieldName))
  If Len(effectiveField) = 0 Then effectiveField = "(unknown field)"

  Dim entrySummary: entrySummary = DescribeAddressEntry(entry)
  If Len(entrySummary) = 0 Then entrySummary = "(no entry information)"

  BuildLookupResponseSummary = "List '" & effectiveList & "' field '" & effectiveField & "' => " & entrySummary
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

  If Len(contactName) > 0 Then WScript.Echo "    Contact full name: " & contactName
  If Len(contactFirst) > 0 Then WScript.Echo "    Contact first name: " & contactFirst
  If Len(contactLast) > 0 Then WScript.Echo "    Contact surname: " & contactLast

  Dim exchName, exchFirst, exchLast, exchSmtp
  exchName = ""
  exchFirst = ""
  exchLast = ""
  exchSmtp = ""
  If Not exch Is Nothing Then
    On Error Resume Next
    exchName = exch.Name
    If Err.Number <> 0 Then exchName = ""
    exchFirst = exch.FirstName
    If Err.Number <> 0 Then exchFirst = ""
    exchLast = exch.LastName
    If Err.Number <> 0 Then exchLast = ""
    exchSmtp = exch.PrimarySmtpAddress
    If Err.Number <> 0 Then exchSmtp = ""
    On Error GoTo 0
  End If

  If Len(exchName) > 0 Then WScript.Echo "    Exchange full name: " & exchName
  If Len(exchFirst) > 0 Then WScript.Echo "    Exchange first name: " & exchFirst
  If Len(exchLast) > 0 Then WScript.Echo "    Exchange surname: " & exchLast
  If Len(exchSmtp) > 0 Then WScript.Echo "    Exchange primary SMTP: " & exchSmtp

  If Not info Is Nothing Then
    Dim infoFirst: infoFirst = ""
    Dim infoLast: infoLast = ""
    Dim infoFull: infoFull = ""
    On Error Resume Next
    infoFirst = SafeText(info("First"))
    infoLast = SafeText(info("Last"))
    infoFull = SafeText(info("Full"))
    On Error GoTo 0
    If Len(infoFirst) > 0 Then WScript.Echo "    Resolved first name: " & infoFirst
    If Len(infoLast) > 0 Then WScript.Echo "    Resolved surname: " & infoLast
    If Len(infoFull) > 0 Then WScript.Echo "    Resolved full name: " & infoFull
  End If

  If Not selectionDetails Is Nothing Then
    WScript.Echo "    Selection details:"
    Dim key
    For Each key In selectionDetails.Keys
      WScript.Echo "      - " & key & ": " & SafeText(selectionDetails(key))
    Next
  End If
End Sub
