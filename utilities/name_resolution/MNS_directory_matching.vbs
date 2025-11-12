' Directory match aggregation helpers for name resolution diagnostics

Function CollectAddressEntriesByFilters(addressList, filters)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
  Set CollectAddressEntriesByFilters = results

  If addressList Is Nothing Then Exit Function
  If Not IsArray(filters) Then Exit Function

  Dim listLabel: listLabel = ""
  On Error Resume Next
  listLabel = SafeText(addressList.Name)
  If Err.Number <> 0 Then
    Err.Clear
    listLabel = ""
  End If
  On Error GoTo 0

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
          Call RecordLookupResponse(listLabel, filterText, entry)
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

  Dim listLabel: listLabel = ""
  On Error Resume Next
  listLabel = SafeText(addressList.Name)
  If Err.Number <> 0 Then
    Err.Clear
    listLabel = ""
  End If
  On Error GoTo 0

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
          Call RecordLookupResponse(listLabel, field, entry)
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

' Shared directory diagnostics helpers

Sub AppendReasonIfMissing(reasonList, reasonText)
  If reasonList Is Nothing Then Exit Sub
  If Len(reasonText) = 0 Then Exit Sub

  Dim i
  For i = 0 To reasonList.Count - 1
    If reasonList(i) = reasonText Then Exit Sub
  Next

  reasonList.Add reasonText
End Sub

Function EnsureAddressEntryKey(entry)
  Dim key: key = BuildAddressEntryTrackingKey(entry)
  If Len(key) = 0 Then key = SafeLower(DescribeAddressEntry(entry))
  EnsureAddressEntryKey = key
End Function

Function EnsureReasonList(reasonMap, entryKey)
  Set EnsureReasonList = Nothing
  If reasonMap Is Nothing Then Exit Function
  If Len(entryKey) = 0 Then Exit Function

  Dim bucket
  If reasonMap.Exists(entryKey) Then
    Set bucket = reasonMap(entryKey)
  Else
    Set bucket = CreateArrayList()
    reasonMap.Add entryKey, bucket
  End If

  Set EnsureReasonList = bucket
End Function

Function AddEntryWithReasonCollection(aggregate, seen, reasonMap, entry, listName, reasons)
  AddEntryWithReasonCollection = 0
  If aggregate Is Nothing Then Exit Function
  If entry Is Nothing Then Exit Function
  If reasons Is Nothing Then Exit Function

  Dim beforeCount: beforeCount = GetCollectionCount(aggregate)
  Call AddAddressEntryToCollection(aggregate, entry, seen)
  Dim afterCount: afterCount = GetCollectionCount(aggregate)

  Dim entryKey: entryKey = EnsureAddressEntryKey(entry)
  Dim reasonList: Set reasonList = EnsureReasonList(reasonMap, entryKey)

  If Not reasonList Is Nothing Then
    Dim idx
    For idx = 0 To reasons.Count - 1
      Dim reasonText: reasonText = SafeText(reasons(idx))
      If Len(reasonText) > 0 Then
        If Len(listName) > 0 Then reasonText = "[" & listName & "] " & reasonText
        Call AppendReasonIfMissing(reasonList, reasonText)
      End If
    Next
  End If

  AddEntryWithReasonCollection = afterCount - beforeCount
End Function

Function BuildAddressEntrySearchBlob(entry)
  BuildAddressEntrySearchBlob = ""
  If entry Is Nothing Then Exit Function

  Dim parts: Set parts = CreateArrayList()
  Dim value

  value = SafeText(entry.Name)
  If Len(value) > 0 Then parts.Add value

  value = SafeText(entry.Address)
  If Len(value) > 0 Then parts.Add value

  On Error Resume Next
  Dim exch, dist, contact
  Set exch = entry.GetExchangeUser()
  If Err.Number = 0 Then
    value = SafeText(exch.Name)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(exch.FirstName)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(exch.LastName)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(exch.PrimarySmtpAddress)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(exch.Department)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(exch.CompanyName)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(exch.Alias)
    If Len(value) > 0 Then parts.Add value
  Else
    Err.Clear
  End If

  Set dist = entry.GetExchangeDistributionList()
  If Err.Number = 0 Then
    value = SafeText(dist.Name)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(dist.PrimarySmtpAddress)
    If Len(value) > 0 Then parts.Add value
  Else
    Err.Clear
  End If

  Set contact = entry.GetContact()
  If Err.Number = 0 Then
    value = SafeText(contact.FullName)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(contact.Email1Address)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(contact.Email2Address)
    If Len(value) > 0 Then parts.Add value
    value = SafeText(contact.Email3Address)
    If Len(value) > 0 Then parts.Add value
  Else
    Err.Clear
  End If
  On Error GoTo 0

  Dim combined: combined = ""
  Dim i
  For i = 0 To parts.Count - 1
    value = SafeLower(parts(i))
    If Len(value) > 0 Then
      If Len(combined) > 0 Then combined = combined & " "
      combined = combined & value
    End If
  Next

  BuildAddressEntrySearchBlob = Trim(combined)
End Function

Function EnumeratePartialMatches(addressList, probes, tokens, aggregate, seen, reasonMap, listName)
  Set EnumeratePartialMatches = Nothing
  If addressList Is Nothing Then Exit Function

  Dim entries
  On Error Resume Next
  Set entries = addressList.AddressEntries
  If Err.Number <> 0 Then
    Err.Clear
    Set entries = Nothing
  End If
  On Error GoTo 0
  If entries Is Nothing Then Exit Function

  Dim candidateMaps: Set candidateMaps = CreateTextDictionary()
  Dim tokenEntryKeys: Set tokenEntryKeys = CreateTextDictionary()

  Dim entry
  Dim candidateBucket
  For Each entry In entries
    Dim blobLower: blobLower = BuildAddressEntrySearchBlob(entry)
    If Len(blobLower) = 0 Then
      ' Skip entries with no searchable text.
    Else
      Dim entryReasons: Set entryReasons = Nothing
      Dim j
      For j = 0 To probes.Count - 1
        Dim probe: Set probe = probes(j)
        Dim candidateLower: candidateLower = SafeText(probe("Lower"))
        Dim candidateValue: candidateValue = SafeText(probe("Value"))
        Dim candidateReason: candidateReason = SafeText(probe("Reason"))

        If Len(candidateLower) > 0 Then
          If InStr(blobLower, candidateLower) > 0 Then
            Set entryReasons = EnsureArrayList(entryReasons)
            Call AppendReasonIfMissing(entryReasons, "Partial match on '" & candidateValue & "' (" & candidateReason & ")")

            If candidateMaps.Exists(candidateLower) Then
              Set candidateBucket = candidateMaps(candidateLower)
            Else
              Set candidateBucket = CreateTextDictionary()
              candidateMaps.Add candidateLower, candidateBucket
            End If

            Dim entryKey: entryKey = EnsureAddressEntryKey(entry)
            If Len(entryKey) > 0 Then
              If Not candidateBucket.Exists(entryKey) Then candidateBucket.Add entryKey, True
            End If
          End If
        End If
      Next

      If Not tokens Is Nothing Then
        If tokens.Count > 0 Then
          If ContainsAllTokens(blobLower, tokens) Then
            Set entryReasons = EnsureArrayList(entryReasons)
            Call AppendReasonIfMissing(entryReasons, "Contains all lookup tokens")

            Dim tokenKey: tokenKey = EnsureAddressEntryKey(entry)
            If Len(tokenKey) > 0 Then
              If Not tokenEntryKeys.Exists(tokenKey) Then tokenEntryKeys.Add tokenKey, True
            End If
          End If
        End If
      End If

      If Not entryReasons Is Nothing Then
        Call AddEntryWithReasonCollection(aggregate, seen, reasonMap, entry, listName, entryReasons)
      End If
    End If
  Next

  Dim stats: Set stats = CreateTextDictionary()
  Dim candidateCounts: Set candidateCounts = CreateTextDictionary()

  Dim candidateKey
  For Each candidateKey In candidateMaps.Keys
    Set candidateBucket = candidateMaps(candidateKey)
    candidateCounts(candidateKey) = GetDictionaryCount(candidateBucket)
  Next

  Set stats("CandidateCounts") = candidateCounts
  stats("TokenCount") = GetDictionaryCount(tokenEntryKeys)

  Set EnumeratePartialMatches = stats
End Function
Sub OutputMatchReasons(reasonMap, entry, indent)
  If Len(indent) = 0 Then indent = ""
  If reasonMap Is Nothing Then Exit Sub
  If entry Is Nothing Then Exit Sub

  Dim entryKey: entryKey = EnsureAddressEntryKey(entry)
  If Len(entryKey) = 0 Then Exit Sub
  If Not reasonMap.Exists(entryKey) Then Exit Sub

  Dim reasons: Set reasons = reasonMap(entryKey)
  If reasons Is Nothing Then Exit Sub
  If reasons.Count = 0 Then Exit Sub

  WScript.Echo indent & "Match reasons:"

  Dim i
  For i = 0 To reasons.Count - 1
    WScript.Echo indent & "  - " & reasons(i)
  Next
End Sub

Sub OutputSelectionSummary(summary, indent)
  If summary Is Nothing Then Exit Sub
  If Len(indent) = 0 Then indent = ""

  If summary.Exists("Reason") Then WScript.Echo indent & "Reason: " & SafeText(summary("Reason"))
  If summary.Exists("Score") Then WScript.Echo indent & "Score: " & SafeText(summary("Score"))
  If summary.Exists("CandidateCount") Then WScript.Echo indent & "Candidate count: " & SafeText(summary("CandidateCount"))
  If summary.Exists("Department") Then WScript.Echo indent & "Department preference: " & SafeText(summary("Department"))
  If summary.Exists("Company") Then WScript.Echo indent & "Company preference: " & SafeText(summary("Company"))
  If summary.Exists("UsedProfile") Then WScript.Echo indent & "Used caller profile: " & SafeText(summary("UsedProfile"))
End Sub

Function BuildCandidateProbeList(target, components, tokens)
  Dim results: Set results = CreateArrayList()
  Dim seen: Set seen = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  seen.CompareMode = 1
  On Error GoTo 0

  Call AddProbe(results, seen, target, "Original display name")

  If Not components Is Nothing Then
    Call AddProbe(results, seen, SafeText(components("Normalized")), "Normalized components")
    Call AddProbe(results, seen, SafeText(components("ReverseDisplay")), "Reverse display")

    Dim firstName: firstName = SafeText(components("First"))
    Dim lastName: lastName = SafeText(components("Last"))
    If Len(firstName) > 0 Then Call AddProbe(results, seen, firstName, "Parsed first name only")
    If Len(lastName) > 0 Then Call AddProbe(results, seen, lastName, "Parsed surname only")

    If Len(firstName) > 0 And Len(lastName) > 0 Then
      Call AddProbe(results, seen, Trim(firstName & " " & lastName), "First name + surname")
      Call AddProbe(results, seen, Trim(lastName & " " & firstName), "Surname + first name")
      Call AddProbe(results, seen, Trim(lastName & ", " & firstName), "Surname, first name")
      Call AddProbe(results, seen, Trim(lastName & "," & firstName), "Surname,first name (no space)")
    End If

    Dim altFirst: altFirst = SafeText(components("AlternateFirst"))
    Dim altLast: altLast = SafeText(components("AlternateLast"))
    If Len(altFirst) > 0 And Len(altLast) > 0 Then
      Call AddProbe(results, seen, Trim(altFirst & " " & altLast), "Alternate token order")
      Call AddProbe(results, seen, Trim(altLast & " " & altFirst), "Alternate reversed order")
      Call AddProbe(results, seen, Trim(altLast & ", " & altFirst), "Alternate surname, first")
      Call AddProbe(results, seen, Trim(altLast & "," & altFirst), "Alternate surname,first (no space)")
    End If
  End If

  If Not tokens Is Nothing Then
    Dim tokenString: tokenString = JoinTokens(tokens, " ")
    Call AddProbe(results, seen, tokenString, "Normalized tokens")

    Dim reversedTokens: reversedTokens = BuildReversedTokenDisplay(tokens)
    Call AddProbe(results, seen, reversedTokens, "Reversed token display")
  End If

  Set BuildCandidateProbeList = results
End Function

Sub AddProbe(list, seen, value, reason)
  Dim text: text = Trim(SafeText(value))
  If Len(text) = 0 Then Exit Sub
  Dim lowered: lowered = SafeLower(text)
  If Len(lowered) = 0 Then Exit Sub

  If Not seen Is Nothing Then
    If seen.Exists(lowered) Then Exit Sub
    On Error Resume Next
    seen.Add lowered, True
    If Err.Number <> 0 Then
      Err.Clear
      Exit Sub
    End If
    On Error GoTo 0
  End If

  Dim probe: Set probe = CreateObject("Scripting.Dictionary")
  probe("Value") = text
  probe("Reason") = reason
  probe("Lower") = lowered
  list.Add probe
End Sub
