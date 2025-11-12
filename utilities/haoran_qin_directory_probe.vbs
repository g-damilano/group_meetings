Option Explicit
SetLocale 1033

Call Main()

Sub Main()
  If Not EnsureConsoleHost() Then Exit Sub

  Dim scriptFolder: scriptFolder = GetScriptFolder()
  Call LoadModule(scriptFolder, "mail_collections.vbs")
  Call LoadModule(scriptFolder, "mail_template_utils.vbs")
  Call LoadModule(scriptFolder, "mail_input_validation.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_text.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_directory.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_core.vbs")

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
  Call ScanGalAddressListsForExactMatches(ns, target, components, tokens)
  Call SummarizeBaselineComparison(baseline)
  Call ScanContactFoldersForTokens(ns, tokens)
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

Sub ScanGalAddressListsForExactMatches(ns, target, components, tokens)
  WScript.Echo ""
  WScript.Echo "Step 6 - GAL exact-match scans"

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

      Dim anyMatches: anyMatches = False
      Dim j
      For j = 0 To probes.Count - 1
        Dim probe: Set probe = probes(j)
        Dim candidate: candidate = SafeText(probe("Value"))
        Dim reason: reason = SafeText(probe("Reason"))

        Dim matches: Set matches = CollectAddressEntriesByFields(addressList, Array("Name", "DisplayName", "FileAs"), candidate)
        If Not matches Is Nothing Then
          If matches.Count > 0 Then
            anyMatches = True
            WScript.Echo "    Candidate '" & candidate & "' (" & reason & ") produced " & matches.Count & " match(es):"

            Dim k
            For k = 0 To matches.Count - 1
              Call OutputAddressEntrySummary(matches(k), "      ")
            Next
          End If
        End If
      Next

      If Not anyMatches Then
        WScript.Echo "    No exact matches located for any probe candidates."
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

Function EnsureConsoleHost()
  Dim host: host = LCase(WScript.FullName)
  If InStr(host, "cscript.exe") = 0 Then
    WScript.Echo "Please run this script with cscript.exe to use the console interface."
    EnsureConsoleHost = False
  Else
    EnsureConsoleHost = True
  End If
End Function

Function GetScriptFolder()
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  GetScriptFolder = folder
End Function

Function GetOutlookApp()
  Dim app
  On Error Resume Next
  Set app = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("Outlook.Application")
  End If
  On Error GoTo 0
  Set GetOutlookApp = app
End Function

Sub LoadModule(baseFolder, relativePath)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim modulePath: modulePath = ResolveModulePath(fso, baseFolder, relativePath)
  If Len(modulePath) = 0 Then
    WScript.Echo "Warning: utility module not found: " & relativePath
    Exit Sub
  End If

  Dim ts: Set ts = fso.OpenTextFile(modulePath, 1)
  Dim contents: contents = ts.ReadAll
  ts.Close

  contents = RemoveByteOrderMark(contents)

  On Error Resume Next
  ExecuteGlobal contents
  If Err.Number <> 0 Then
    WScript.Echo "Warning: failed to load module '" & relativePath & "': " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Function ResolveModulePath(fso, baseFolder, relativePath)
  Dim modulePath: modulePath = ""

  If Not fso Is Nothing Then
    Dim currentFolder: currentFolder = baseFolder
    Do While Len(currentFolder) > 0
      Dim candidate: candidate = fso.BuildPath(currentFolder, relativePath)
      If fso.FileExists(candidate) Then
        modulePath = candidate
        Exit Do
      End If

      Dim parentFolder: parentFolder = fso.GetParentFolderName(currentFolder)
      If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do
      currentFolder = parentFolder
    Loop
  End If

  ResolveModulePath = modulePath
End Function

Function RemoveByteOrderMark(text)
  RemoveByteOrderMark = text
  If Len(text) = 0 Then Exit Function

  If Len(text) >= 3 Then
    Dim leading: leading = Mid(text, 1, 3)
    If StrComp(leading, Chr(239) & Chr(187) & Chr(191), vbBinaryCompare) = 0 Then
      RemoveByteOrderMark = Mid(text, 4)
      Exit Function
    End If
  End If

  Dim firstChar: firstChar = Mid(text, 1, 1)
  If Len(firstChar) = 0 Then Exit Function

  On Error Resume Next
  Dim codePoint: codePoint = AscW(firstChar)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If codePoint = &HFEFF Or codePoint = &HFFFE Then
    RemoveByteOrderMark = Mid(text, 2)
  End If
End Function
