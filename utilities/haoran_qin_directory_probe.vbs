Option Explicit
SetLocale 1033

Dim bootstrapFso, bootstrapBaseFolder, bootstrapModulePath
Dim bootstrapCurrentFolder, bootstrapCandidate, bootstrapParentFolder
Dim bootstrapStream, bootstrapContents

Set bootstrapFso = CreateObject("Scripting.FileSystemObject")
bootstrapBaseFolder = bootstrapFso.GetParentFolderName(WScript.ScriptFullName)
If Len(bootstrapBaseFolder) = 0 Then bootstrapBaseFolder = "."

bootstrapModulePath = bootstrapFso.BuildPath(bootstrapBaseFolder, "common_bootstrap.vbs")
If Not bootstrapFso.FileExists(bootstrapModulePath) Then
  bootstrapModulePath = bootstrapFso.BuildPath(bootstrapBaseFolder, "utilities\\common_bootstrap.vbs")
End If

If Not bootstrapFso.FileExists(bootstrapModulePath) Then
  bootstrapModulePath = ""
  bootstrapCurrentFolder = bootstrapBaseFolder
  Do While Len(bootstrapCurrentFolder) > 0
    bootstrapCandidate = bootstrapFso.BuildPath(bootstrapCurrentFolder, "utilities\\common_bootstrap.vbs")
    If bootstrapFso.FileExists(bootstrapCandidate) Then
      bootstrapModulePath = bootstrapCandidate
      Exit Do
    End If

    bootstrapParentFolder = bootstrapFso.GetParentFolderName(bootstrapCurrentFolder)
    If StrComp(bootstrapParentFolder, bootstrapCurrentFolder, vbTextCompare) = 0 Then Exit Do
    bootstrapCurrentFolder = bootstrapParentFolder
  Loop
End If

If Not bootstrapFso.FileExists(bootstrapModulePath) Then
  WScript.Echo "Fatal: unable to locate utilities\\common_bootstrap.vbs"
  WScript.Quit 1
End If

Set bootstrapStream = bootstrapFso.OpenTextFile(bootstrapModulePath, 1)
bootstrapContents = bootstrapStream.ReadAll
bootstrapStream.Close
ExecuteGlobal bootstrapContents

Call Main()

Sub Main()
  Call LoadBootstrapModules(Array( _
    "common_console.vbs", _
    "mail_collections.vbs", _
    "mail_template_utils.vbs", _
    "common_outlook.vbs", _
    "mail_input_validation.vbs", _
    "mail_name_resolution_text.vbs", _
    "mail_name_resolution_directory.vbs", _
    "mail_name_resolution_core.vbs"))

  If Not EnsureConsoleHost() Then Exit Sub

  Dim target: target = "Haoran Qin"
  Dim baselineEmail: baselineEmail = "haoran.qin22@imperial.ac.uk"
  Dim argCount: argCount = WScript.Arguments.Count
  If argCount > 0 Then
    Dim firstArg: firstArg = Trim(WScript.Arguments(0))
    If Len(firstArg) > 0 Then target = firstArg
    If argCount > 1 Then
      Dim secondArg: secondArg = Trim(WScript.Arguments(1))
      If Len(secondArg) > 0 Then baselineEmail = secondArg
    End If
  End If

  WScript.Echo "=== Haoran Qin Directory Diagnostic ==="
  WScript.Echo "Target lookup value: '" & target & "'"
  WScript.Echo "Baseline email value: '" & baselineEmail & "'"
  WScript.Echo "Timestamp: " & Now

  Dim olApp: Set olApp = GetOutlookApp()
  If olApp Is Nothing Then
    WScript.Echo "Unable to activate Outlook.Application; aborting diagnostics."
    Exit Sub
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
    WScript.Echo "Unable to obtain Outlook MAPI namespace; aborting diagnostics."
    Exit Sub
  End If

  Dim components: Set components = ParseFullNameComponents(target)
  Dim tokens: Set tokens = CollectFlexibleNameTokens(target)

  Call OutputComponentSummary(components, tokens)

  Dim baseline: Set baseline = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  baseline.CompareMode = 1
  On Error GoTo 0

  Call RunEmailBaselineCheck(ns, baselineEmail, baseline)

  Dim autoInfo: Set autoInfo = ResolveNameLookup(target, "auto")
  Call DumpResolutionBlock("Step 1 - ResolveNameLookup (auto)", autoInfo, "    ")

  Dim displayInfo: Set displayInfo = ResolveNameFromDisplayName(target)
  Call DumpResolutionBlock("Step 2 - ResolveNameFromDisplayName", displayInfo, "    ")

  Dim fullInfo: Set fullInfo = ResolveNameFromFullName(target)
  Call DumpResolutionBlock("Step 3 - ResolveNameFromFullName", fullInfo, "    ")

  Call RunRecipientResolveProbes(ns, target, components, tokens)
  Call RunGalDisplayEnumerationProbes(ns, target, components, tokens, baseline)
  Call ScanGalAddressListsForMatches(ns, target, components, tokens)
  Call SummarizeBaselineComparison(baseline)
  Call ScanContactFoldersForTokens(ns, tokens)
End Sub

Function BuildMatchTokenList(tokens)
  Set BuildMatchTokenList = Nothing
  If tokens Is Nothing Then Exit Function

  Dim result: Set result = CreateArrayList()
  Dim i
  For i = 0 To tokens.Count - 1
    Dim lowered: lowered = SafeLower(tokens(i))
    If Len(lowered) > 0 Then result.Add lowered
  Next

  If result.Count = 0 Then Exit Function

  Set BuildMatchTokenList = result
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

  stats("CandidateCounts") = candidateCounts
  stats("TokenCount") = GetDictionaryCount(tokenEntryKeys)

  Set EnumeratePartialMatches = stats
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

Function EnsureAddressEntryKey(entry)
  Dim key: key = BuildAddressEntryTrackingKey(entry)
  If Len(key) = 0 Then key = SafeLower(DescribeAddressEntry(entry))
  EnsureAddressEntryKey = key
End Function

Function GetCollectionCount(list)
  GetCollectionCount = 0
  If list Is Nothing Then Exit Function
  On Error Resume Next
  GetCollectionCount = list.Count
  If Err.Number <> 0 Then
    Err.Clear
    GetCollectionCount = 0
  End If
  On Error GoTo 0
End Function

Function GetDictionaryCount(dict)
  GetDictionaryCount = 0
  If dict Is Nothing Then Exit Function
  On Error Resume Next
  GetDictionaryCount = dict.Count
  If Err.Number <> 0 Then
    Err.Clear
    GetDictionaryCount = 0
  End If
  On Error GoTo 0
End Function

Function EnsureArrayList(list)
  If list Is Nothing Then
    Set list = CreateArrayList()
  End If
  Set EnsureArrayList = list
End Function

Function CreateArrayList()
  Set CreateArrayList = CreateObject("System.Collections.ArrayList")
End Function

Function CreateTextDictionary()
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  dict.CompareMode = 1
  On Error GoTo 0
  Set CreateTextDictionary = dict
End Function

Sub AppendReasonIfMissing(reasonList, reasonText)
  If reasonList Is Nothing Then Exit Sub
  If Len(reasonText) = 0 Then Exit Sub

  Dim i
  For i = 0 To reasonList.Count - 1
    If reasonList(i) = reasonText Then Exit Sub
  Next

  reasonList.Add reasonText
End Sub

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

Sub OutputComponentSummary(components, tokens)
  WScript.Echo ""
  WScript.Echo "Component analysis of target value:"

  If components Is Nothing Then
    WScript.Echo "  (No components parsed.)"
  Else
    WScript.Echo "  Display: " & SafeText(components("Display"))
    WScript.Echo "  Normalized: " & SafeText(components("Normalized"))
    WScript.Echo "  Reverse display: " & SafeText(components("ReverseDisplay"))
    WScript.Echo "  First name: " & SafeText(components("First"))
    WScript.Echo "  Surname: " & SafeText(components("Last"))
    WScript.Echo "  Alternate first name: " & SafeText(components("AlternateFirst"))
    WScript.Echo "  Alternate surname: " & SafeText(components("AlternateLast"))
  End If

  Dim tokenSummary
  tokenSummary = ""
  If Not tokens Is Nothing Then
    Dim i
    For i = 0 To tokens.Count - 1
      If Len(tokenSummary) = 0 Then
        tokenSummary = tokens(i)
      Else
        tokenSummary = tokenSummary & ", " & tokens(i)
      End If
    Next
  End If
  If Len(tokenSummary) = 0 Then
    WScript.Echo "  Extracted tokens: (none)"
  Else
    WScript.Echo "  Extracted tokens: " & tokenSummary
  End If
End Sub

Sub RunEmailBaselineCheck(ns, emailAddress, baseline)
  WScript.Echo ""
  WScript.Echo "Step 0 - Email baseline verification"

  If Len(Trim(emailAddress)) = 0 Then
    WScript.Echo "  (No baseline email supplied.)"
    Exit Sub
  End If

  WScript.Echo "  Target email: " & emailAddress

  Dim emailInfo: Set emailInfo = ResolveNameFromEmail(emailAddress)
  Call DumpResolutionBlock("  Email resolution via ResolveNameFromEmail", emailInfo, "    ")

  If Not baseline Is Nothing Then
    On Error Resume Next
    baseline.Add "EmailResolutionFull", DescribeField(emailInfo, "Full", "")
    baseline.Add "EmailResolutionFirst", DescribeField(emailInfo, "First", "")
    baseline.Add "EmailResolutionLast", DescribeField(emailInfo, "Last", "")
    baseline.Add "EmailResolutionEmail", DescribeField(emailInfo, "Email", "")
    On Error GoTo 0
  End If

  If ns Is Nothing Then
    WScript.Echo "  (Namespace unavailable; skipping recipient baseline probes.)"
    Exit Sub
  End If

  Dim recipient: Set recipient = ns.CreateRecipient(emailAddress)
  If recipient Is Nothing Then
    WScript.Echo "  (Failed to create recipient for baseline email.)"
  Else
    If recipient.Resolve Then
      WScript.Echo "  Recipient.Resolve succeeded for baseline email."
      Call OutputAddressEntrySummary(recipient.AddressEntry, "    ")
      If Not baseline Is Nothing Then
        Call RecordBaselineEntry(baseline, "Recipient", recipient.AddressEntry)
      End If
    Else
      WScript.Echo "  Recipient.Resolve FAILED for baseline email."
    End If
  End If

  Dim listNames: listNames = GetGalListNames()
  Dim i
  For i = LBound(listNames) To UBound(listNames)
    Dim listName: listName = listNames(i)
    Dim addressList: Set addressList = OpenAddressList(ns, listName)
    If addressList Is Nothing Then
      WScript.Echo "  Address list '" & listName & "' unavailable for email cross-check."
    Else
      WScript.Echo "  Address list email lookup: " & listName
      Dim combined: Set combined = CreateObject("System.Collections.ArrayList")
      Dim seen: Set seen = CreateObject("Scripting.Dictionary")
      On Error Resume Next
      seen.CompareMode = 1
      On Error GoTo 0

      Call AppendAddressEntries(combined, CollectAddressEntriesByProperty(addressList, "SMTPAddress", emailAddress), seen)
      Call AppendAddressEntries(combined, CollectAddressEntriesByProperty(addressList, "EmailAddress", emailAddress), seen)
      Call AppendAddressEntries(combined, CollectAddressEntriesByProperty(addressList, "Address", emailAddress), seen)

      If combined Is Nothing Then
        WScript.Echo "    (Lookup helper returned Nothing.)"
      ElseIf combined.Count = 0 Then
        WScript.Echo "    No entries matched the email address."
      Else
        Dim entry
        For Each entry In combined
          Call OutputAddressEntrySummary(entry, "    ")
          If Not baseline Is Nothing Then
            Call RecordBaselineEntry(baseline, listName, entry)
          End If
        Next
      End If
    End If
  Next
End Sub

Sub RunRecipientResolveProbes(ns, target, components, tokens)
  WScript.Echo ""
  WScript.Echo "Step 4 - Recipient resolution probes"

  Dim probes: Set probes = BuildCandidateProbeList(target, components, tokens)
  If probes Is Nothing Then
    WScript.Echo "  (No probes generated.)"
    Exit Sub
  End If

  If probes.Count = 0 Then
    WScript.Echo "  (No probes generated.)"
    Exit Sub
  End If

  Dim i
  For i = 0 To probes.Count - 1
    Dim probe: Set probe = probes(i)
    Dim candidate: candidate = SafeText(probe("Value"))
    Dim reason: reason = SafeText(probe("Reason"))

    WScript.Echo "  Probe " & (i + 1) & " of " & probes.Count & ": '" & candidate & "' (" & reason & ")"

    Dim info: Set info = CreateEmptyResolutionInfo()
    Call AttemptRecipientResolveForValue(ns, candidate, info, reason)
    Call DumpResolutionDetails(info, "    ")
  Next
End Sub

Sub RunGalDisplayEnumerationProbes(ns, target, components, tokens, baseline)
  WScript.Echo ""
  WScript.Echo "Step 5 - GAL display-name enumeration probes"

  Dim probes: Set probes = BuildCandidateProbeList(target, components, tokens)
  If probes Is Nothing Then
    WScript.Echo "  (No probes generated.)"
    Exit Sub
  End If
  If probes.Count = 0 Then
    WScript.Echo "  (No probes generated.)"
    Exit Sub
  End If

  Dim i
  For i = 0 To probes.Count - 1
    Dim probe: Set probe = probes(i)
    Dim candidate: candidate = SafeText(probe("Value"))
    Dim reason: reason = SafeText(probe("Reason"))

    WScript.Echo "  Enumerating address lists for '" & candidate & "' (" & reason & ")"
    Dim listNames: listNames = GetGalListNames()
    Dim listIndex
    For listIndex = LBound(listNames) To UBound(listNames)
      Dim listName: listName = listNames(listIndex)
      Dim addressList: Set addressList = OpenAddressList(ns, listName)
      If addressList Is Nothing Then
        WScript.Echo "    [" & listName & "] list unavailable."
      Else
        Dim matches: Set matches = CollectAddressEntriesByFields(addressList, Array("Name", "DisplayName", "FileAs", "Alias"), candidate)
        If matches Is Nothing Then
          WScript.Echo "    [" & listName & "] helper returned Nothing."
        ElseIf matches.Count = 0 Then
          WScript.Echo "    [" & listName & "] no matches."
        Else
          WScript.Echo "    [" & listName & "] matches: " & matches.Count
          Dim entry
          For Each entry In matches
            Call OutputAddressEntrySummary(entry, "      ")
            If Not baseline Is Nothing Then Call RecordBaselineEntry(baseline, listName & " display probe", entry)
          Next
        End If
      End If
    Next
  Next
End Sub

Sub ScanGalAddressListsForMatches(ns, target, components, tokens)
  WScript.Echo ""
  WScript.Echo "Step 6 - GAL match scans with disambiguation"

  If ns Is Nothing Then
    WScript.Echo "  (Namespace unavailable.)"
    Exit Sub
  End If

  Dim listNames: listNames = GetGalListNames()
  If Not IsArray(listNames) Then
    WScript.Echo "  (No GAL lists available for scanning.)"
    Exit Sub
  End If

  Dim probes: Set probes = BuildCandidateProbeList(target, components, tokens)
  If probes Is Nothing Then
    WScript.Echo "  (No probes generated.)"
    Exit Sub
  End If

  Dim normalizedTokens: Set normalizedTokens = BuildMatchTokenList(tokens)
  Dim callerProfile: Set callerProfile = GetCurrentUserProfileInfo()

  Dim i
  For i = LBound(listNames) To UBound(listNames)
    Dim listName: listName = listNames(i)
    Dim addressList
    Set addressList = Nothing

    On Error Resume Next
    Set addressList = ns.AddressLists(listName)
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If addressList Is Nothing Then
      WScript.Echo "  Address list '" & listName & "' could not be opened."
    Else
      WScript.Echo "  Address list: " & listName

      Dim aggregate: Set aggregate = CreateObject("System.Collections.ArrayList")
      Dim seen: Set seen = CreateTextDictionary()
      Dim reasonMap: Set reasonMap = CreateTextDictionary()

      Dim j
      For j = 0 To probes.Count - 1
        Dim probe: Set probe = probes(j)
        Dim candidate: candidate = SafeText(probe("Value"))
        Dim reason: reason = SafeText(probe("Reason"))
        Dim exactMatches: Set exactMatches = CollectAddressEntriesByFields(addressList, Array("Name", "DisplayName", "FileAs"), candidate)

        Dim matchCount: matchCount = GetCollectionCount(exactMatches)
        Dim newUnique: newUnique = 0

        If Not exactMatches Is Nothing Then
          Dim entry
          For Each entry In exactMatches
            Dim exactReasons: Set exactReasons = CreateArrayList()
            Call exactReasons.Add("Exact field match for '" & candidate & "' (" & reason & ")")
            newUnique = newUnique + AddEntryWithReasonCollection(aggregate, seen, reasonMap, entry, listName, exactReasons)
          Next
        End If

        WScript.Echo "    Exact candidate '" & candidate & "' matched " & matchCount & " entr(ies) (" & newUnique & " new unique)."
      Next

      Dim partialStats: Set partialStats = EnumeratePartialMatches(addressList, probes, normalizedTokens, aggregate, seen, reasonMap, listName)

      If Not partialStats Is Nothing Then
        Dim candidateCounts
        Set candidateCounts = Nothing
        If partialStats.Exists("CandidateCounts") Then Set candidateCounts = partialStats("CandidateCounts")

        For j = 0 To probes.Count - 1
          Dim probePartial: Set probePartial = probes(j)
          Dim partialCandidate: partialCandidate = SafeText(probePartial("Value"))
          Dim candidateLower: candidateLower = SafeText(probePartial("Lower"))
          Dim partialCount: partialCount = 0
          If IsObject(candidateCounts) Then
            If candidateCounts.Exists(candidateLower) Then partialCount = candidateCounts(candidateLower)
          End If
          WScript.Echo "    Partial candidate '" & partialCandidate & "' matched " & partialCount & " unique entr(ies)."
        Next

        Dim tokenSummary: tokenSummary = ""
        If Not normalizedTokens Is Nothing Then tokenSummary = JoinTokens(normalizedTokens, ", ")
        If Len(tokenSummary) = 0 Then tokenSummary = "(no tokens)"

        Dim tokenCount: tokenCount = 0
        If partialStats.Exists("TokenCount") Then tokenCount = partialStats("TokenCount")
        WScript.Echo "    Token combination '" & tokenSummary & "' matched " & tokenCount & " unique entr(ies)."
      End If

      Dim aggregateCount: aggregateCount = GetCollectionCount(aggregate)
      If aggregateCount = 0 Then
        WScript.Echo "    No exact or partial matches located for any probe candidates."
      Else
        WScript.Echo "    Aggregated " & aggregateCount & " unique match(es) across exact and partial searches."

        Dim idx
        For idx = 0 To aggregateCount - 1
          Dim currentEntry: Set currentEntry = aggregate(idx)
          WScript.Echo "    Match " & (idx + 1) & ": " & DescribeAddressEntry(currentEntry)
          Call OutputMatchReasons(reasonMap, currentEntry, "      ")
        Next

        Dim selection: Set selection = ChooseBestAddressEntry(aggregate, callerProfile)
        Dim selectionEntry: Set selectionEntry = Nothing
        Dim selectionSummary: Set selectionSummary = Nothing

        If Not selection Is Nothing Then
          If selection.Exists("Entry") Then Set selectionEntry = selection("Entry")
          Set selectionSummary = BuildSelectionSummary(selection)
        End If

        If selectionEntry Is Nothing Then
          If aggregateCount > 0 Then
            Set selectionEntry = aggregate(0)
            If selectionEntry Is Nothing Then
              WScript.Echo "    Disambiguation did not return a candidate."
            Else
              WScript.Echo "    Disambiguation did not provide a scored candidate; defaulting to first match: " & DescribeAddressEntry(selectionEntry)
            End If
          Else
            WScript.Echo "    Disambiguation did not return a candidate."
          End If
        Else
          WScript.Echo "    Disambiguation selected: " & DescribeAddressEntry(selectionEntry)
          Call OutputSelectionSummary(selectionSummary, "      ")
        End If
      End If
    End If
  Next
End Sub

Sub ScanContactFoldersForTokens(ns, tokens)
  WScript.Echo ""
  WScript.Echo "Step 7 - Contact folder substring scans"

  If ns Is Nothing Then
    WScript.Echo "  (Namespace unavailable.)"
    Exit Sub
  End If

  If tokens Is Nothing Then
    WScript.Echo "  (No tokens extracted from lookup value.)"
    Exit Sub
  End If

  Dim normalizedTokens: Set normalizedTokens = CreateObject("System.Collections.ArrayList")
  Dim i
  For i = 0 To tokens.Count - 1
    normalizedTokens.Add SafeLower(tokens(i))
  Next

  If normalizedTokens.Count = 0 Then
    WScript.Echo "  (No tokens extracted from lookup value.)"
    Exit Sub
  End If

  Dim listNames: listNames = BuildDirectorySearchList(ns)
  If Not IsArray(listNames) Then
    WScript.Echo "  (No address lists available for scanning.)"
    Exit Sub
  End If

  Dim listIndex
  For listIndex = LBound(listNames) To UBound(listNames)
    Dim listName: listName = listNames(listIndex)
    Dim addressList
    Set addressList = Nothing

    On Error Resume Next
    Set addressList = ns.AddressLists(listName)
    If Err.Number <> 0 Then
      Err.Clear
      Set addressList = Nothing
    End If
    On Error GoTo 0

    If addressList Is Nothing Then
      WScript.Echo "  Address list '" & listName & "' could not be opened for contact scanning."
    Else
      Dim contactsFolder
      Set contactsFolder = Nothing
      On Error Resume Next
      Set contactsFolder = addressList.GetContactsFolder()
      If Err.Number <> 0 Then
        Err.Clear
        Set contactsFolder = Nothing
      End If
      On Error GoTo 0

      If contactsFolder Is Nothing Then
        WScript.Echo "  Address list '" & listName & "' does not expose a contacts folder."
      Else
        WScript.Echo "  Contact folder: " & listName
        Dim items
        Set items = Nothing
        On Error Resume Next
        Set items = contactsFolder.Items
        If Err.Number <> 0 Then
          Err.Clear
          Set items = Nothing
        End If
        On Error GoTo 0

        If items Is Nothing Then
          WScript.Echo "    Unable to enumerate contacts for this folder."
        Else
          Dim matches: Set matches = CreateObject("System.Collections.ArrayList")
          Dim contact
          Dim skippedNonContacts: skippedNonContacts = 0
          For Each contact In items
            If IsContactLikeItem(contact) Then
              If ContactItemMatchesTokens(contact, normalizedTokens) Then
                matches.Add contact
              End If
            Else
              skippedNonContacts = skippedNonContacts + 1
            End If
          Next

          If matches.Count = 0 Then
            WScript.Echo "    No contact entries matched all search tokens."
          Else
            WScript.Echo "    Matched " & matches.Count & " contact(s) containing all tokens:"
            Dim m
            For m = 0 To matches.Count - 1
              Call OutputContactSummary(matches(m), "      ")
            Next
          End If

          If skippedNonContacts > 0 Then
            WScript.Echo "    Skipped " & skippedNonContacts & " non-contact item(s) during scanning."
          End If
        End If
      End If
    End If
  Next
End Sub

Function ContactItemMatchesTokens(contact, tokens)
  ContactItemMatchesTokens = False
  If contact Is Nothing Then Exit Function
  If tokens Is Nothing Then Exit Function
  If tokens.Count = 0 Then Exit Function
  If Not IsContactLikeItem(contact) Then Exit Function

  Dim fields: fields = Array( _
    SafeLower(ContactFieldValue(contact, "FullName")), _
    SafeLower(ContactFieldValue(contact, "FileAs")), _
    SafeLower(ContactFieldValue(contact, "FirstName")), _
    SafeLower(ContactFieldValue(contact, "LastName")), _
    SafeLower(ContactFieldValue(contact, "Email1Address")), _
    SafeLower(ContactFieldValue(contact, "Email2Address")), _
    SafeLower(ContactFieldValue(contact, "Email3Address")), _
    SafeLower(ContactFieldValue(contact, "CompanyName")), _
    SafeLower(ContactFieldValue(contact, "Department")) _
  )

  Dim i
  For i = LBound(fields) To UBound(fields)
    If ContainsAllTokens(fields(i), tokens) Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
  Next

  Dim firstLower: firstLower = SafeLower(ContactFieldValue(contact, "FirstName"))
  Dim lastLower: lastLower = SafeLower(ContactFieldValue(contact, "LastName"))

  If tokens.Count = 2 Then
    Dim t0: t0 = SafeLower(tokens(0))
    Dim t1: t1 = SafeLower(tokens(1))
    If InStr(firstLower, t0) > 0 And InStr(lastLower, t1) > 0 Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
    If InStr(firstLower, t1) > 0 And InStr(lastLower, t0) > 0 Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
  End If
End Function

Function ContainsAllTokens(text, tokens)
  ContainsAllTokens = False
  If tokens Is Nothing Then Exit Function
  If tokens.Count = 0 Then Exit Function

  Dim haystack: haystack = SafeLower(text)
  If Len(haystack) = 0 Then Exit Function

  Dim i
  For i = 0 To tokens.Count - 1
    Dim needle: needle = SafeLower(tokens(i))
    If Len(needle) = 0 Then
      ContainsAllTokens = False
      Exit Function
    End If
    If InStr(haystack, needle) = 0 Then
      ContainsAllTokens = False
      Exit Function
    End If
  Next

  ContainsAllTokens = True
End Function

Sub OutputContactSummary(contact, indent)
  If Len(indent) = 0 Then indent = ""
  If contact Is Nothing Then
    WScript.Echo indent & "(Contact reference is Nothing)"
    Exit Sub
  End If

  If Not IsContactLikeItem(contact) Then
    WScript.Echo indent & "(Item is not a contact: " & DescribeItemClass(contact) & ")"
    Exit Sub
  End If

  WScript.Echo indent & "Contact: " & SafeText(ContactFieldValue(contact, "FullName"))
  WScript.Echo indent & "  FileAs: " & SafeText(ContactFieldValue(contact, "FileAs"))
  WScript.Echo indent & "  First name: " & SafeText(ContactFieldValue(contact, "FirstName"))
  WScript.Echo indent & "  Surname: " & SafeText(ContactFieldValue(contact, "LastName"))
  WScript.Echo indent & "  Primary email: " & SafeText(ContactFieldValue(contact, "Email1Address"))
  WScript.Echo indent & "  Secondary email: " & SafeText(ContactFieldValue(contact, "Email2Address"))
  WScript.Echo indent & "  Tertiary email: " & SafeText(ContactFieldValue(contact, "Email3Address"))
  WScript.Echo indent & "  Company: " & SafeText(ContactFieldValue(contact, "CompanyName"))
  WScript.Echo indent & "  Department: " & SafeText(ContactFieldValue(contact, "Department"))
End Sub

Function IsContactLikeItem(item)
  IsContactLikeItem = False
  If item Is Nothing Then Exit Function

  Dim classValue: classValue = -1
  On Error Resume Next
  classValue = item.Class
  If Err.Number <> 0 Then
    Err.Clear
    classValue = -1
  End If

  Dim messageClass: messageClass = ""
  messageClass = item.MessageClass
  If Err.Number <> 0 Then
    Err.Clear
    messageClass = ""
  End If
  On Error GoTo 0

  If classValue = 40 Or classValue = 69 Then
    IsContactLikeItem = True
    Exit Function
  End If

  Dim lowered: lowered = LCase(messageClass)
  If Len(lowered) = 0 Then Exit Function
  If Left(lowered, 11) = "ipm.contact" Then
    IsContactLikeItem = True
  ElseIf lowered = "ipm.distlist" Then
    IsContactLikeItem = True
  End If
End Function

Function ContactFieldValue(contact, fieldName)
  Dim value: value = ""
  If contact Is Nothing Then
    ContactFieldValue = ""
    Exit Function
  End If

  On Error Resume Next
  Select Case LCase(fieldName)
    Case "fullname"
      value = contact.FullName
    Case "fileas"
      value = contact.FileAs
    Case "firstname"
      value = contact.FirstName
    Case "lastname"
      value = contact.LastName
    Case "email1address"
      value = contact.Email1Address
    Case "email2address"
      value = contact.Email2Address
    Case "email3address"
      value = contact.Email3Address
    Case "companyname"
      value = contact.CompanyName
    Case "department"
      value = contact.Department
    Case Else
      value = ""
  End Select
  If Err.Number <> 0 Then
    Err.Clear
    value = ""
  End If
  On Error GoTo 0

  ContactFieldValue = SafeText(value)
End Function

Function DescribeItemClass(item)
  DescribeItemClass = "unknown"
  If item Is Nothing Then Exit Function

  Dim classValue: classValue = "?"
  Dim messageClass: messageClass = ""

  On Error Resume Next
  classValue = item.Class
  If Err.Number <> 0 Then
    Err.Clear
    classValue = "?"
  End If
  messageClass = item.MessageClass
  If Err.Number <> 0 Then
    Err.Clear
    messageClass = ""
  End If
  On Error GoTo 0

  If Len(messageClass) > 0 Then
    DescribeItemClass = "Class=" & classValue & ", MessageClass='" & messageClass & "'"
  Else
    DescribeItemClass = "Class=" & classValue
  End If
End Function

Sub OutputAddressEntrySummary(entry, indent)
  If Len(indent) = 0 Then indent = ""
  If entry Is Nothing Then
    WScript.Echo indent & "(Address entry reference is Nothing)"
    Exit Sub
  End If

  Dim summary: summary = DescribeAddressEntry(entry)
  WScript.Echo indent & summary

  Dim exch
  On Error Resume Next
  Set exch = entry.GetExchangeUser()
  If Err.Number <> 0 Then
    Err.Clear
    Set exch = Nothing
  End If
  On Error GoTo 0

  Dim info: Set info = CreateEmptyResolutionInfo()
  info("Full") = SafeText(entry.Name)
  Call PopulateNameFromPropertyAccessor(entry, info, "", "")

  If Not exch Is Nothing Then
    Call PopulateNameFromPropertyAccessor(exch, info, "", "")
  End If

  WScript.Echo indent & "  First name: " & DescribeField(info, "First", "(empty)")
  WScript.Echo indent & "  Surname: " & DescribeField(info, "Last", "(empty)")
  WScript.Echo indent & "  Display name: " & DescribeField(info, "Full", "(empty)")
  WScript.Echo indent & "  Email: " & DescribeField(info, "Email", "(empty)")

  If Not exch Is Nothing Then
    WScript.Echo indent & "  Exchange department: " & SafeText(exch.Department)
    WScript.Echo indent & "  Exchange company: " & SafeText(exch.CompanyName)
  End If
End Sub

Sub DumpResolutionBlock(title, info, indent)
  WScript.Echo ""
  WScript.Echo title
  Call DumpResolutionDetails(info, indent)
End Sub

Sub DumpResolutionDetails(info, indent)
  If Len(indent) = 0 Then indent = ""

  If info Is Nothing Then
    WScript.Echo indent & "(No resolution information returned.)"
    Exit Sub
  End If

  WScript.Echo indent & "Found: " & DescribeBoolean(info, "Found")
  WScript.Echo indent & "Directory source: " & DescribeField(info, "DirectorySource", "(none)")
  WScript.Echo indent & "Address lists checked: " & DescribeField(info, "DirectoryAttempted", "(none)")
  WScript.Echo indent & "First name: " & DescribeField(info, "First", "(empty)")
  WScript.Echo indent & "Surname: " & DescribeField(info, "Last", "(empty)")
  WScript.Echo indent & "Email: " & DescribeField(info, "Email", "(empty)")
  WScript.Echo indent & "Display name: " & DescribeField(info, "Full", "(empty)")
End Sub

Function BuildCandidateProbeList(target, components, tokens)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
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

Function DescribeBoolean(info, key)
  Dim defaultText: defaultText = "Unknown"
  If info Is Nothing Then
    DescribeBoolean = defaultText
    Exit Function
  End If

  If info.Exists(key) Then
    If CBool(info(key)) Then
      DescribeBoolean = "Found"
    Else
      DescribeBoolean = "Not found"
    End If
  Else
    DescribeBoolean = defaultText
  End If
End Function

Function DescribeField(info, key, placeholder)
  Dim text: text = placeholder
  If Not info Is Nothing Then
    If info.Exists(key) Then
      Dim candidate: candidate = SafeText(info(key))
      If Len(candidate) > 0 Then text = candidate
    End If
  End If
  DescribeField = text
End Function

Sub SummarizeBaselineComparison(baseline)
  WScript.Echo ""
  WScript.Echo "Baseline summary and comparison"

  If baseline Is Nothing Then
    WScript.Echo "  (No baseline information recorded.)"
    Exit Sub
  End If

  Dim fullName: fullName = SafeText(baselineLookup(baseline, "EmailResolutionFull"))
  Dim firstName: firstName = SafeText(baselineLookup(baseline, "EmailResolutionFirst"))
  Dim lastName: lastName = SafeText(baselineLookup(baseline, "EmailResolutionLast"))
  Dim emailValue: emailValue = SafeText(baselineLookup(baseline, "EmailResolutionEmail"))

  If Len(fullName) > 0 Or Len(emailValue) > 0 Then
    WScript.Echo "  ResolveNameFromEmail produced:"
    WScript.Echo "    Display name: " & SafeText(fullName)
    WScript.Echo "    First name: " & SafeText(firstName)
    WScript.Echo "    Surname: " & SafeText(lastName)
    WScript.Echo "    Email: " & SafeText(emailValue)
  Else
    WScript.Echo "  (Email baseline resolution returned no fields.)"
  End If

  Dim key
  For Each key In baseline.Keys
    If StrComp(key, "EmailResolutionFull", vbTextCompare) = 0 Then
      ' already reported
    ElseIf StrComp(key, "EmailResolutionFirst", vbTextCompare) = 0 Then
      ' already reported
    ElseIf StrComp(key, "EmailResolutionLast", vbTextCompare) = 0 Then
      ' already reported
    ElseIf StrComp(key, "EmailResolutionEmail", vbTextCompare) = 0 Then
      ' already reported
    Else
      If IsObject(baseline(key)) Then
        Dim list: Set list = baseline(key)
        If TypeName(list) = "ArrayList" Then
          If list.Count > 0 Then
            WScript.Echo "  Entries recorded from " & key & ":"
            Dim idx
            For idx = 0 To list.Count - 1
              WScript.Echo "    " & list(idx)
            Next
          Else
            WScript.Echo "  Entries recorded from " & key & ": (none)"
          End If
        End If
      End If
    End If
  Next
End Sub

Function baselineLookup(store, key)
  baselineLookup = ""
  If store Is Nothing Then Exit Function
  Dim safeKey: safeKey = Trim(CStr(key))
  If Len(safeKey) = 0 Then Exit Function
  If store.Exists(safeKey) Then baselineLookup = store(safeKey)
End Function

Sub RecordBaselineEntry(baseline, bucket, entry)
  If baseline Is Nothing Then Exit Sub
  Dim key: key = Trim(CStr(bucket))
  If Len(key) = 0 Then key = "(unspecified)"
  If entry Is Nothing Then Exit Sub

  Dim summary: summary = DescribeAddressEntry(entry)
  Dim list
  Set list = Nothing
  If baseline.Exists(key) Then
    If IsObject(baseline(key)) Then
      Set list = baseline(key)
    Else
      Set list = Nothing
    End If
  End If

  If list Is Nothing Then
    Set list = CreateObject("System.Collections.ArrayList")
    If baseline.Exists(key) Then
      Set baseline(key) = list
    Else
      baseline.Add key, list
    End If
  End If

  If TypeName(list) = "ArrayList" Then
    If Not ContainsInArrayList(list, summary) Then list.Add summary
  End If
End Sub

Sub AppendAddressEntries(targetList, sourceList, seen)
  If targetList Is Nothing Then Exit Sub
  If sourceList Is Nothing Then Exit Sub

  Dim entry
  For Each entry In sourceList
    If entry Is Nothing Then
      ' skip
    ElseIf seen Is Nothing Then
      targetList.Add entry
    Else
      Dim key: key = BuildAddressEntryTrackingKey(entry)
      If Len(key) = 0 Then
        targetList.Add entry
      ElseIf Not seen.Exists(key) Then
        seen.Add key, True
        targetList.Add entry
      End If
    End If
  Next
End Sub

Function CollectAddressEntriesByProperty(addressList, propertyName, propertyValue)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
  Set CollectAddressEntriesByProperty = results

  If addressList Is Nothing Then Exit Function

  Dim trimmed: trimmed = Trim(propertyValue)
  If Len(trimmed) = 0 Then Exit Function

  Dim field: field = Trim(propertyName)
  If Len(field) = 0 Then Exit Function

  Dim query: query = EscapeRestrictionValue(trimmed)

  Dim entry
  On Error Resume Next
  Set entry = addressList.AddressEntries.Find("[" & field & "] = '" & query & "'")
  If Err.Number <> 0 Then
    Err.Clear
    Set entry = Nothing
  End If

  Do While Not entry Is Nothing
    Call results.Add(entry)
    Set entry = addressList.AddressEntries.FindNext()
    If Err.Number <> 0 Then
      Err.Clear
      Exit Do
    End If
  Loop
  On Error GoTo 0
End Function

Function GetGalListNames()
  GetGalListNames = Array("Global Address List", "Offline Global Address List")
End Function

Function OpenAddressList(ns, listName)
  Set OpenAddressList = Nothing
  If ns Is Nothing Then Exit Function

  Dim trimmed: trimmed = Trim(CStr(listName))
  If Len(trimmed) = 0 Then Exit Function

  On Error Resume Next
  Set OpenAddressList = ns.AddressLists(trimmed)
  If Err.Number <> 0 Then
    Err.Clear
    Set OpenAddressList = Nothing
  End If
  On Error GoTo 0
End Function

