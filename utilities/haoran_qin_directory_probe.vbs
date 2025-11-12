Option Explicit
SetLocale 1033

Call BootstrapIncludeEntry()

Sub BootstrapIncludeEntry()
  Dim entryFso: Set entryFso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = entryFso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."

  Dim entryPath
  entryPath = entryFso.BuildPath(scriptFolder, "utilities\\bootstrap_entry.vbs")
  If Not entryFso.FileExists(entryPath) Then
    entryPath = entryFso.BuildPath(scriptFolder, "bootstrap_entry.vbs")
  End If

  If Not entryFso.FileExists(entryPath) Then
    WScript.Echo "Fatal: unable to locate utilities\\bootstrap_entry.vbs"
    WScript.Quit 1
  End If

  Dim entryStream
  On Error Resume Next
  Set entryStream = entryFso.OpenTextFile(entryPath, 1)
  If Err.Number <> 0 Then
    Dim entryError: entryError = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "Fatal: unable to read bootstrap entry: " & entryError
    WScript.Quit 1
  End If
  On Error GoTo 0

  Dim entryContents: entryContents = entryStream.ReadAll
  entryStream.Close

  ExecuteGlobal entryContents
  Set entryFso = Nothing
End Sub

Call Main()

Sub Main()
  Dim scriptFolder: scriptFolder = GetScriptFolder()

  Call RequireBootstrapModules(Array( _
    "common_console.vbs", _
    "mail_collections.vbs", _
    "mail_template_utils.vbs", _
    "common_outlook.vbs", _
    "mail_input_validation.vbs", _
    "MNS_loader.vbs"))

  Call LoadNameResolutionModules(scriptFolder)

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

  Dim autoResponses: Set autoResponses = CreateArrayList()
  Call SetLookupResponseRecorder(autoResponses)
  Dim autoInfo: Set autoInfo = ResolveNameLookup(target, "auto")
  Call ClearLookupResponseRecorder()
  Call DumpResolutionBlock("Step 1 - ResolveNameLookup (auto)", autoInfo, "    ", autoResponses)

  Dim displayResponses: Set displayResponses = CreateArrayList()
  Call SetLookupResponseRecorder(displayResponses)
  Dim displayInfo: Set displayInfo = ResolveNameFromDisplayName(target)
  Call ClearLookupResponseRecorder()
  Call DumpResolutionBlock("Step 2 - ResolveNameFromDisplayName", displayInfo, "    ", displayResponses)

  Dim fullResponses: Set fullResponses = CreateArrayList()
  Call SetLookupResponseRecorder(fullResponses)
  Dim fullInfo: Set fullInfo = ResolveNameFromFullName(target)
  Call ClearLookupResponseRecorder()
  Call DumpResolutionBlock("Step 3 - ResolveNameFromFullName", fullInfo, "    ", fullResponses)

  Dim recipientResponses: Set recipientResponses = CreateArrayList()
  Call SetLookupResponseRecorder(recipientResponses)
  Call RunRecipientResolveProbes(ns, target, components, tokens)
  Call ClearLookupResponseRecorder()
  Call DumpRecordedOutlookResponses(recipientResponses, "    ", "Direct Outlook responses (Recipient resolve probes)")

  Dim enumerationResponses: Set enumerationResponses = CreateArrayList()
  Call SetLookupResponseRecorder(enumerationResponses)
  Call RunGalDisplayEnumerationProbes(ns, target, components, tokens, baseline)
  Call ClearLookupResponseRecorder()
  Call DumpRecordedOutlookResponses(enumerationResponses, "    ", "Direct Outlook responses (Display-name enumeration)")

  Dim scanResponses: Set scanResponses = CreateArrayList()
  Call SetLookupResponseRecorder(scanResponses)
  Call ScanGalAddressListsForMatches(ns, target, components, tokens)
  Call ClearLookupResponseRecorder()
  Call DumpRecordedOutlookResponses(scanResponses, "    ", "Direct Outlook responses (Address list scans)")
  Call SummarizeBaselineComparison(baseline)
  Call ScanContactFoldersForTokens(ns, tokens)
End Sub

Sub RequireBootstrapModules(modulePaths)
  Dim loadError: loadError = ""

  On Error Resume Next
  Execute "Call LoadBootstrapModules(modulePaths)"
  If Err.Number <> 0 Then
    loadError = Err.Description
    Err.Clear
  End If
  On Error GoTo 0

  If Len(loadError) > 0 Then
    WScript.Echo "Fatal: unable to invoke LoadBootstrapModules: " & loadError
    WScript.Quit 1
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

  Dim emailResponses: Set emailResponses = CreateArrayList()
  Call SetLookupResponseRecorder(emailResponses)
  Dim emailInfo: Set emailInfo = ResolveNameFromEmail(emailAddress)
  Call ClearLookupResponseRecorder()
  Call DumpResolutionBlock("  Email resolution via ResolveNameFromEmail", emailInfo, "    ", emailResponses)

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

      Dim partialStats: Set partialStats = Nothing
      If ShouldSkipPartialMatchScan(listName, normalizedTokens) Then
        WScript.Echo "    Skipping partial candidate enumeration for '" & listName & "' due to performance safeguards."
      Else
        Set partialStats = EnumeratePartialMatches(addressList, probes, normalizedTokens, aggregate, seen, reasonMap, listName)
      End If

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

Sub DumpResolutionBlock(title, info, indent, responses)
  WScript.Echo ""
  WScript.Echo title
  Call DumpResolutionDetails(info, indent)
  Call DumpRecordedOutlookResponses(responses, indent, "")
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
  WScript.Echo indent & "Raw resolution payload:"
  Call DumpVariantStructure(info, indent & "  ", 0)
End Sub

Sub DumpRecordedOutlookResponses(responses, indent, label)
  If Len(indent) = 0 Then indent = ""

  If responses Is Nothing Then Exit Sub

  Dim count: count = 0
  On Error Resume Next
  count = responses.Count
  If Err.Number <> 0 Then
    Err.Clear
    count = 0
  End If
  On Error GoTo 0

  If count = 0 Then Exit Sub

  Dim effectiveLabel: effectiveLabel = Trim(label)
  If Len(effectiveLabel) = 0 Then effectiveLabel = "Direct Outlook responses"

  WScript.Echo indent & effectiveLabel & ":"

  Dim i
  For i = 0 To count - 1
    WScript.Echo indent & "  - " & SafeText(responses(i))
  Next
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

Function ShouldSkipPartialMatchScan(listName, normalizedTokens)
  ShouldSkipPartialMatchScan = False

  Dim loweredName: loweredName = SafeLower(listName)
  If Len(loweredName) = 0 Then Exit Function

  If StrComp(loweredName, "offline global address list", vbTextCompare) = 0 Then
    ' Enumerating partial matches over the offline GAL can take several minutes because
    ' Outlook must hydrate every entry locally. We already scanned the online GAL, so
    ' repeating the expensive scan provides little additional value but causes the
    ' diagnostic script to appear frozen. Skip the partial scan for this list to keep
    ' the tooling responsive.
    ShouldSkipPartialMatchScan = True
    Exit Function
  End If

  If normalizedTokens Is Nothing Then Exit Function
  On Error Resume Next
  ShouldSkipPartialMatchScan = (normalizedTokens.Count = 0)
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


Sub DumpVariantStructure(value, indent, depth)
  If Len(indent) = 0 Then indent = ""
  If depth > 8 Then
    WScript.Echo indent & "(Maximum inspection depth reached.)"
    Exit Sub
  End If

  If value Is Nothing Then
    WScript.Echo indent & "(Value is Nothing)"
    Exit Sub
  End If

  Dim typeName: typeName = SafeTypeName(value)

  If IsObject(value) Then
    If StrComp(typeName, "Dictionary", vbTextCompare) = 0 Then
      Call DumpDictionaryVariant(value, indent, depth)
      Exit Sub
    End If

    If StrComp(typeName, "ArrayList", vbTextCompare) = 0 Then
      Call DumpArrayListVariant(value, indent, depth)
      Exit Sub
    End If

    WScript.Echo indent & DescribeVariant(value, True)
    Exit Sub
  End If

  If IsArray(value) Then
    Call DumpArrayVariant(value, indent, depth)
    Exit Sub
  End If

  WScript.Echo indent & DescribeVariant(value, True)
End Sub

Sub DumpDictionaryVariant(dictValue, indent, depth)
  Dim count
  On Error Resume Next
  count = dictValue.Count
  If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo indent & DescribeVariant(dictValue, True)
    Exit Sub
  End If
  On Error GoTo 0

  WScript.Echo indent & "(Dictionary with " & count & " entries)"
  If count = 0 Then Exit Sub

  Dim entryIndent: entryIndent = indent & "  "
  Dim key
  For Each key In dictValue
    Dim entryValue
    On Error Resume Next
    entryValue = dictValue(key)
    If Err.Number <> 0 Then
      Err.Clear
      WScript.Echo entryIndent & "- Key " & DescribeVariant(key, True) & " => (Unable to read value)"
    Else
      WScript.Echo entryIndent & "- Key " & DescribeVariant(key, True) & " => " & DescribeVariant(entryValue, True)
      If ShouldExpandVariant(entryValue, depth) Then
        Call DumpVariantStructure(entryValue, entryIndent & "    ", depth + 1)
      End If
    End If
    On Error GoTo 0
  Next
End Sub

Sub DumpArrayListVariant(listValue, indent, depth)
  Dim count
  On Error Resume Next
  count = listValue.Count
  If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo indent & DescribeVariant(listValue, True)
    Exit Sub
  End If
  On Error GoTo 0

  WScript.Echo indent & "(ArrayList with " & count & " entries)"
  If count = 0 Then Exit Sub

  Dim entryIndent: entryIndent = indent & "  "
  Dim idx
  For idx = 0 To count - 1
    Dim listItem
    On Error Resume Next
    listItem = listValue(idx)
    If Err.Number <> 0 Then
      Err.Clear
      WScript.Echo entryIndent & "[" & idx & "] => (Unable to read value)"
    Else
      WScript.Echo entryIndent & "[" & idx & "] => " & DescribeVariant(listItem, True)
      If ShouldExpandVariant(listItem, depth) Then
        Call DumpVariantStructure(listItem, entryIndent & "    ", depth + 1)
      End If
    End If
    On Error GoTo 0
  Next
End Sub

Sub DumpArrayVariant(arrayValue, indent, depth)
  Dim lowerBound, upperBound, length
  lowerBound = 0
  upperBound = -1
  length = 0

  On Error Resume Next
  lowerBound = LBound(arrayValue)
  upperBound = UBound(arrayValue)
  If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo indent & DescribeVariant(arrayValue, True)
    Exit Sub
  End If
  On Error GoTo 0

  length = upperBound - lowerBound + 1
  WScript.Echo indent & "(Array with " & length & " elements)"
  If length <= 0 Then Exit Sub

  Dim entryIndent: entryIndent = indent & "  "
  Dim idx
  For idx = lowerBound To upperBound
    Dim item
    item = arrayValue(idx)
    WScript.Echo entryIndent & "[" & idx & "] => " & DescribeVariant(item, True)
    If ShouldExpandVariant(item, depth) Then
      Call DumpVariantStructure(item, entryIndent & "    ", depth + 1)
    End If
  Next
End Sub

Function ShouldExpandVariant(value, depth)
  ShouldExpandVariant = False
  If depth >= 8 Then Exit Function

  If IsObject(value) Then
    Dim typeName: typeName = SafeTypeName(value)
    If StrComp(typeName, "Dictionary", vbTextCompare) = 0 Then
      ShouldExpandVariant = True
    ElseIf StrComp(typeName, "ArrayList", vbTextCompare) = 0 Then
      ShouldExpandVariant = True
    End If
  ElseIf IsArray(value) Then
    ShouldExpandVariant = True
  End If
End Function

Function SafeTypeName(value)
  On Error Resume Next
  Dim name: name = TypeName(value)
  If Err.Number <> 0 Then
    Err.Clear
    name = "Unknown"
  End If
  On Error GoTo 0
  SafeTypeName = name
End Function

Function DescribeVariant(value, includeQuotes)
  Dim typeName: typeName = SafeTypeName(value)
  Dim description: description = ""

  If IsObject(value) Then
    description = "<Object>"
  ElseIf IsArray(value) Then
    Dim lowerBound, upperBound
    lowerBound = 0
    upperBound = -1
    On Error Resume Next
    lowerBound = LBound(value)
    upperBound = UBound(value)
    If Err.Number <> 0 Then
      Err.Clear
      description = "<Array>"
    Else
      description = "<Array length=" & (upperBound - lowerBound + 1) & ">"
    End If
    On Error GoTo 0
  ElseIf IsNull(value) Then
    description = "(Null)"
  ElseIf IsEmpty(value) Then
    description = "(Empty)"
  ElseIf VarType(value) = vbBoolean Then
    If value Then
      description = "True"
    Else
      description = "False"
    End If
  ElseIf VarType(value) = vbString Then
    description = value
    If includeQuotes Then description = """" & description & """"
  Else
    On Error Resume Next
    description = CStr(value)
    If Err.Number <> 0 Then
      Err.Clear
      description = "<Unprintable>"
    End If
    On Error GoTo 0
  End If

  DescribeVariant = description & " (" & typeName & ")"
End Function

