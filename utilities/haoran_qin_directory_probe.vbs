Option Explicit
SetLocale 1033

Call Main()

Sub Main()
  If Not EnsureConsoleHost() Then Exit Sub

  Dim scriptFolder: scriptFolder = GetScriptFolder()
  Call LoadModule(scriptFolder, "mail_utils.vbs")
  Call LoadModule(scriptFolder, "mail_collections.vbs")
  Call LoadModule(scriptFolder, "mail_template_utils.vbs")
  Call LoadModule(scriptFolder, "mail_input_validation.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_text.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_directory.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_core.vbs")

  Dim target: target = "Haoran Qin"
  Dim argCount: argCount = WScript.Arguments.Count
  If argCount > 0 Then
    Dim firstArg: firstArg = Trim(WScript.Arguments(0))
    If Len(firstArg) > 0 Then target = firstArg
  End If

  WScript.Echo "=== Haoran Qin Directory Diagnostic ==="
  WScript.Echo "Target lookup value: '" & target & "'"
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

  Dim autoInfo: Set autoInfo = ResolveNameLookup(target, "auto")
  Call DumpResolutionBlock("Step 1 - ResolveNameLookup (auto)", autoInfo, "    ")

  Dim displayInfo: Set displayInfo = ResolveNameFromDisplayName(target)
  Call DumpResolutionBlock("Step 2 - ResolveNameFromDisplayName", displayInfo, "    ")

  Dim fullInfo: Set fullInfo = ResolveNameFromFullName(target)
  Call DumpResolutionBlock("Step 3 - ResolveNameFromFullName", fullInfo, "    ")

  Call RunRecipientResolveProbes(ns, target, components, tokens)
  Call RunDisplayEnumerationProbes(ns, target, components, tokens)
  Call ScanAddressListsForExactMatches(ns, target, components, tokens)
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

Sub RunDisplayEnumerationProbes(ns, target, components, tokens)
  WScript.Echo ""
  WScript.Echo "Step 5 - Display-name enumeration probes"

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
    Dim info: Set info = CreateEmptyResolutionInfo()
    Call EnumerateDirectoryMatchesByDisplayName(ns, candidate, info, reason)
    Call DumpResolutionDetails(info, "    ")
  Next
End Sub

Sub ScanAddressListsForExactMatches(ns, target, components, tokens)
  WScript.Echo ""
  WScript.Echo "Step 6 - Address list exact-match scans"

  If ns Is Nothing Then
    WScript.Echo "  (Namespace unavailable.)"
    Exit Sub
  End If

  Dim listNames: listNames = BuildDirectorySearchList(ns)
  If Not IsArray(listNames) Then
    WScript.Echo "  (No address lists available for scanning.)"
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
          For Each contact In items
            If ContactItemMatchesTokens(contact, normalizedTokens) Then
              matches.Add contact
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

  Dim fields: fields = Array( _
    SafeLower(SafeText(contact.FullName)), _
    SafeLower(SafeText(contact.FileAs)), _
    SafeLower(SafeText(contact.FirstName)), _
    SafeLower(SafeText(contact.LastName)), _
    SafeLower(SafeText(contact.Email1Address)), _
    SafeLower(SafeText(contact.Email2Address)), _
    SafeLower(SafeText(contact.Email3Address)), _
    SafeLower(SafeText(contact.CompanyName)), _
    SafeLower(SafeText(contact.Department)) _
  )

  Dim i
  For i = LBound(fields) To UBound(fields)
    If ContainsAllTokens(fields(i), tokens) Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
  Next

  Dim firstLower: firstLower = SafeLower(SafeText(contact.FirstName))
  Dim lastLower: lastLower = SafeLower(SafeText(contact.LastName))

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

  WScript.Echo indent & "Contact: " & SafeText(contact.FullName)
  WScript.Echo indent & "  FileAs: " & SafeText(contact.FileAs)
  WScript.Echo indent & "  First name: " & SafeText(contact.FirstName)
  WScript.Echo indent & "  Surname: " & SafeText(contact.LastName)
  WScript.Echo indent & "  Primary email: " & SafeText(contact.Email1Address)
  WScript.Echo indent & "  Secondary email: " & SafeText(contact.Email2Address)
  WScript.Echo indent & "  Tertiary email: " & SafeText(contact.Email3Address)
  WScript.Echo indent & "  Company: " & SafeText(contact.CompanyName)
  WScript.Echo indent & "  Department: " & SafeText(contact.Department)
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
