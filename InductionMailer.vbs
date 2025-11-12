Option Explicit
SetLocale 1033

'======================================================================
' Bootstrap shared utilities
'======================================================================
Dim bootstrapFso: Set bootstrapFso = CreateObject("Scripting.FileSystemObject")
Dim bootstrapModulePath: bootstrapModulePath = LocateBootstrapModule(bootstrapFso, WScript.ScriptFullName)

If Len(bootstrapModulePath) = 0 Then
  WScript.Echo "Fatal: unable to locate utilities\\common_bootstrap.vbs"
  WScript.Quit 1
End If

ExecuteGlobal ReadAllText(bootstrapFso, bootstrapModulePath)

Set bootstrapFso = Nothing

Function LocateBootstrapModule(fso, scriptPath)
  Dim startFolder: startFolder = fso.GetParentFolderName(scriptPath)
  If Len(startFolder) = 0 Then startFolder = "."

  Dim searchTargets: searchTargets = Array("utilities\\common_bootstrap.vbs", "common_bootstrap.vbs")
  Dim currentFolder: currentFolder = startFolder

  Do While Len(currentFolder) > 0
    Dim i
    For i = LBound(searchTargets) To UBound(searchTargets)
      Dim candidate: candidate = fso.BuildPath(currentFolder, searchTargets(i))
      If fso.FileExists(candidate) Then
        LocateBootstrapModule = candidate
        Exit Function
      End If
    Next

    Dim parentFolder: parentFolder = fso.GetParentFolderName(currentFolder)
    If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do
    currentFolder = parentFolder
  Loop

  LocateBootstrapModule = ""
End Function

Function ReadAllText(fso, filePath)
  Dim stream: Set stream = fso.OpenTextFile(filePath, 1)
  ReadAllText = stream.ReadAll
  stream.Close
End Function

'======================================================================
' InductionMailer.vbs — Console workflow for staged induction templates
'======================================================================

Const TEMPLATE_DIR_NAME = "templates"
Const LOG_FILE_NAME     = "InductionMailer.log"
Const DEFAULT_ACTION    = "display"   ' display | send | save
Const MAILER_LIB_FILE   = "mail_utils_mailer_dynamiclinks.vbs"
Const MAILER_LIB_EXT_SUBPATH = "mail_utils_utf8_plus_dynamiclinks_ext\mail_utils_mailer_dynamiclinks.vbs"
Const LINK_MAP_FILE = "links_map.txt"

Const HEADER_PREFIX = "@--"
Const HEADER_SUFFIX = "--"
Const HTML_HEADER_PREFIX = "<!--"
Const HTML_HEADER_SUFFIX = "-->"
Const DELIVERY_BROADCAST = "broadcast"
Const DELIVERY_PERSONAL  = "personal"

Dim fso, scriptFolder, templateFolder, logFilePath, linkMapContext

Call LoadUtilityModules()
Call Main()

'======================================================================
' Support module bootstrap
'======================================================================
Sub LoadUtilityModules()
  Dim modules
  modules = Array( _
    "utilities\common_console.vbs", _
    "utilities\mail_collections.vbs", _
    "utilities\mail_template_utils.vbs", _
    "utilities\mail_delivery_utils.vbs", _
    "utilities\mail_input_validation.vbs", _
    "utilities\mail_name_resolution_loader.vbs")

  Dim baseFolder: baseFolder = GetScriptFolder()

  On Error Resume Next
  Call LoadBootstrapModules(modules)
  If Err.Number <> 0 Then
    WScript.Echo "Fatal: unable to load utility modules: " & Err.Description
    Err.Clear
    WScript.Quit 1
  End If

  Call LoadNameResolutionModules(baseFolder)
  If Err.Number <> 0 Then
    WScript.Echo "Fatal: unable to load name resolution modules: " & Err.Description
    Err.Clear
    WScript.Quit 1
  End If
  On Error GoTo 0
End Sub

'======================================================================
' Entry point
'======================================================================
Sub Main()
  If Not EnsureConsoleHost() Then Exit Sub

  InitializeEnvironment

  If Not LoadMailerLibrary() Then
    LogEvent "Unable to load mail_utils_mailer_dynamiclinks.vbs"
    WScript.Echo "Fatal: mail utility library missing."
    Exit Sub
  End If

  Set linkMapContext = LoadLinkMapContext()

  Dim templates: Set templates = LoadTemplates(templateFolder)
  If templates.Count = 0 Then
    WScript.Echo "No templates were found under '" & templateFolder & "'."
    Exit Sub
  End If

  Dim recipientPackage: Set recipientPackage = PromptRecipientResolutionPackage()
  If Not HasRecipientResolutionEntries(recipientPackage) Then
    WScript.Echo "No valid target email addresses supplied."
    Exit Sub
  End If

  Dim action: action = DetermineAction()

  Dim mode: mode = PromptOperationMode()
  If Len(mode) = 0 Then
    WScript.Echo "Operation cancelled."
    Exit Sub
  End If

  Dim selectedTemplates: Set selectedTemplates = Nothing
  If mode = "stage" Then
    Set selectedTemplates = SelectTemplatesByStage(templates)
  Else
    Set selectedTemplates = SelectTemplatesIndividually(templates)
  End If

  If selectedTemplates Is Nothing Or selectedTemplates.Count = 0 Then
    WScript.Echo "No templates selected."
    Exit Sub
  End If

  Dim key
  For Each key In selectedTemplates.Keys
    ProcessTemplate selectedTemplates(key), recipientPackage, action
  Next

  WScript.Echo "All requested templates processed."
End Sub

Function PromptLine(prompt)
  WScript.StdOut.Write prompt
  PromptLine = Trim(WScript.StdIn.ReadLine())
End Function

Function PromptOperationMode()
  WScript.Echo "Select the operation mode:"& vbCrLf & _
               "  1) Stage mode (send every email in a stage)" & vbCrLf & _
               "  2) Email mode (pick one or more templates)"
  Dim mode: mode = ""
  Do
    Dim choice: choice = PromptLine("Enter choice (1/2) or press Enter to cancel: ")
    If Len(choice) = 0 Then
      mode = ""
      Exit Do
    End If
    Select Case choice
      Case "1": mode = "stage": Exit Do
      Case "2": mode = "email": Exit Do
      Case Else
        WScript.Echo "Invalid option. Please choose 1 or 2."
    End Select
  Loop
  PromptOperationMode = mode
End Function

'======================================================================
' Environment bootstrap
'======================================================================
Sub InitializeEnvironment()
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = GetScriptFolder()
  templateFolder = fso.BuildPath(scriptFolder, TEMPLATE_DIR_NAME)
  logFilePath = fso.BuildPath(scriptFolder, LOG_FILE_NAME)
  LogEvent "=== Induction Mailer (console) started ==="
  LogEvent "Template folder: " & templateFolder
  Set linkMapContext = CreateCaseInsensitiveDictionary()
End Sub

Function LoadMailerLibrary()
  Dim candidates
  candidates = Array( _
    fso.BuildPath(scriptFolder, MAILER_LIB_FILE), _
    fso.BuildPath(scriptFolder, MAILER_LIB_EXT_SUBPATH), _
    fso.BuildPath(templateFolder, MAILER_LIB_FILE))

  Dim i
  For i = LBound(candidates) To UBound(candidates)
    If fso.FileExists(candidates(i)) Then
      Dim ts, contents
      Set ts = fso.OpenTextFile(candidates(i), 1)
      contents = ts.ReadAll
      ts.Close
      contents = RemoveByteOrderMark(contents)
      ExecuteGlobal contents
      LogEvent "Loaded mailer library from " & candidates(i)
      EnsureLinkMapAvailable candidates(i)
      LoadMailerLibrary = True
      Exit Function
    End If
  Next

  LoadMailerLibrary = False
End Function

Sub EnsureLinkMapAvailable(libraryPath)
  On Error Resume Next
  Dim libFolder: libFolder = fso.GetParentFolderName(libraryPath)
  Dim sourceMap: sourceMap = fso.BuildPath(libFolder, LINK_MAP_FILE)
  Dim targetMap: targetMap = fso.BuildPath(scriptFolder, LINK_MAP_FILE)
  If fso.FileExists(sourceMap) Then
    If Not fso.FileExists(targetMap) Then
      fso.CopyFile sourceMap, targetMap
      If Err.Number = 0 Then
        LogEvent "Copied " & LINK_MAP_FILE & " to script folder"
      Else
        LogEvent "Failed to copy " & LINK_MAP_FILE & ": " & Err.Description
        Err.Clear
      End If
    End If
  End If
  On Error GoTo 0
End Sub

'======================================================================
' Template discovery and parsing
'======================================================================
Function LoadTemplates(rootFolder)
  Dim items: Set items = CreateCaseInsensitiveDictionary()
  If Not fso.FolderExists(rootFolder) Then
    Set LoadTemplates = items
    Exit Function
  End If
  Dim counter: counter = 0
  Call CollectTemplates(rootFolder, items, counter)
  Set LoadTemplates = items
End Function

Sub CollectTemplates(folderPath, dict, ByRef counter)
  Dim folder: Set folder = fso.GetFolder(folderPath)
  Dim file
  For Each file In folder.Files
    Dim ext: ext = LCase(fso.GetExtensionName(file.Name))
    If ext = "htm" Or ext = "html" Or ext = "txt" Then
      Dim tmpl: Set tmpl = ParseTemplate(file.Path)
      If Not tmpl Is Nothing Then
        counter = counter + 1
        dict.Add CStr(counter), tmpl
      End If
    End If
  Next
  Dim subFolder
  For Each subFolder In folder.SubFolders
    CollectTemplates subFolder.Path, dict, counter
  Next
End Sub

Function ParseTemplate(filePath)
  Dim resultText: resultText = ReadFileSmart(filePath)
  If Len(resultText) = 0 Then
    LogEvent "Skipped empty template: " & filePath
    Set ParseTemplate = Nothing
    Exit Function
  End If

  Dim pieces: pieces = ExtractHeaderAndBody(resultText)
  Dim metaText, bodyText
  metaText = NormalizeLineEndings(pieces(0))
  bodyText = TrimLeadingBlankLines(pieces(1))

  Dim meta: Set meta = ParseHeaderBlock(metaText)

  Dim template: Set template = CreateObject("Scripting.Dictionary")
  template("Path") = filePath
  template("Folder") = fso.GetParentFolderName(filePath)
  template("FileName") = fso.GetFileName(filePath)
  Set template("Meta") = meta
  template("Body") = bodyText
  template("DisplayName") = fso.GetFileName(filePath)
  template("Stage") = LCase(GetDictionaryValue(meta, "stage"))
  template("Code") = GetDictionaryValue(meta, "code")
  Set ParseTemplate = template
End Function

Function ExtractHeaderAndBody(resultText)
  Dim headerText: headerText = ""
  Dim bodyText: bodyText = resultText

  Dim startPos, endPos

  startPos = InStr(1, resultText, HTML_HEADER_PREFIX, vbTextCompare)
  If startPos > 0 Then
    endPos = InStr(startPos + Len(HTML_HEADER_PREFIX), resultText, HTML_HEADER_SUFFIX)
    If endPos > 0 Then
      headerText = Mid(resultText, startPos + Len(HTML_HEADER_PREFIX), endPos - startPos - Len(HTML_HEADER_PREFIX))
      bodyText = Mid(resultText, endPos + Len(HTML_HEADER_SUFFIX))
      ExtractHeaderAndBody = Array(headerText, bodyText)
      Exit Function
    End If
  End If

  startPos = InStr(1, resultText, HEADER_PREFIX, vbTextCompare)
  If startPos > 0 Then
    endPos = InStr(startPos + Len(HEADER_PREFIX), resultText, HEADER_SUFFIX)
    If endPos > 0 Then
      headerText = Mid(resultText, startPos + Len(HEADER_PREFIX), endPos - startPos - Len(HEADER_PREFIX))
      bodyText = Mid(resultText, endPos + Len(HEADER_SUFFIX))
    End If
  End If

  ExtractHeaderAndBody = Array(headerText, bodyText)
End Function

Function ParseHeaderBlock(blockText)
  Dim dict: Set dict = CreateCaseInsensitiveDictionary()
  If Len(Trim(blockText)) = 0 Then
    Set ParseHeaderBlock = dict
    Exit Function
  End If

  Dim lines: lines = Split(blockText, vbLf)
  Dim line, pos, key, value
  For Each line In lines
    pos = InStr(line, ":")
    If pos > 0 Then
      key = Trim(Left(line, pos - 1))
      value = Trim(Mid(line, pos + 1))
      If Len(key) > 0 Then
        dict(LCase(key)) = value
      End If
    End If
  Next

  Set ParseHeaderBlock = dict
End Function
'======================================================================
' Template selection
'======================================================================
Function BuildTemplateLabel(tmpl)
  Dim stage: stage = tmpl("Stage")
  Dim code: code = tmpl("Code")
  Dim label: label = tmpl("FileName")
  Dim prefix: prefix = ""
  If Len(stage) > 0 Then prefix = Capitalize(stage)
  If Len(code) > 0 Then
    If Len(prefix) > 0 Then
      prefix = prefix & " - " & code
    Else
      prefix = code
    End If
  End If
  If Len(prefix) > 0 Then
    BuildTemplateLabel = prefix & " | " & label
  Else
    BuildTemplateLabel = label
  End If
End Function

Function BuildStageIndex(templates)
  Dim stageMap: Set stageMap = CreateCaseInsensitiveDictionary()
  Dim key, template, stageName
  For Each key In templates.Keys
    Set template = templates(key)
    stageName = template("Stage")
    If Len(stageName) = 0 Then stageName = "(unspecified)"
    If Not stageMap.Exists(stageName) Then
      Dim list: Set list = CreateObject("System.Collections.ArrayList")
      Set stageMap(stageName) = list
    End If
    stageMap(stageName).Add template
  Next
  Set BuildStageIndex = stageMap
End Function

Function SelectTemplatesByStage(templates)
  Dim stageMap: Set stageMap = BuildStageIndex(templates)
  Dim orderedStages: orderedStages = Array("onboarding", "scheduling", "authorisation", "finalisation")

  If stageMap.Count = 0 Then
    WScript.Echo "No stage information available in templates."
    Set SelectTemplatesByStage = Nothing
    Exit Function
  End If

  Dim displayList: Set displayList = CreateObject("System.Collections.ArrayList")
  Dim stageName
  For Each stageName In orderedStages
    If stageMap.Exists(stageName) Then displayList.Add stageName
  Next
  For Each stageName In stageMap.Keys
    If Not ContainsInArrayList(displayList, stageName) Then displayList.Add stageName
  Next

  WScript.Echo "Available stages:"
  Dim i
  For i = 0 To displayList.Count - 1
    stageName = displayList(i)
    WScript.Echo "  " & (i + 1) & ") " & Capitalize(stageName) & " (" & stageMap(stageName).Count & " template(s))"
  Next

  Dim selection: Set selection = Nothing
  Do
    Dim choice: choice = PromptLine("Select a stage by number or press Enter to cancel: ")
    If Len(choice) = 0 Then
      Exit Do
    End If
    If IsNumeric(choice) Then
      i = CInt(choice) - 1
      If i >= 0 And i < displayList.Count Then
        stageName = displayList(i)
        WScript.Echo "Stage '" & Capitalize(stageName) & "' selected."
        Set selection = BuildOrderedDictionary(stageMap(stageName))
        Exit Do
      End If
    End If
    WScript.Echo "Invalid selection."
  Loop

  Set SelectTemplatesByStage = selection
End Function

Function SelectTemplatesIndividually(templates)
  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  Dim key
  For Each key In templates.Keys
    list.Add templates(key)
  Next

  Dim ordered: Set ordered = SortTemplateList(list)
  Dim selectionResult: Set selectionResult = Nothing

  If ordered.Count = 0 Then
    Set SelectTemplatesIndividually = Nothing
    Exit Function
  End If

  WScript.Echo "Available templates:"
  Dim i
  For i = 0 To ordered.Count - 1
    WScript.Echo "  " & (i + 1) & ") " & BuildTemplateLabel(ordered(i))
  Next

  Do
    Dim response: response = PromptLine("Select template numbers (comma-separated) or press Enter to cancel: ")
    If Len(response) = 0 Then
      Exit Do
    End If

    Dim selections: Set selections = ParseNumberList(response)
    If selections.Count = 0 Then
      WScript.Echo "No valid selections detected."
    Else
      Dim dict: Set dict = CreateObject("Scripting.Dictionary")
      Dim invalid: invalid = False
      Dim number
      For Each number In selections.Keys
        i = CInt(number) - 1
        If i < 0 Or i >= ordered.Count Then
          invalid = True
          Exit For
        End If
      Next

      If invalid Then
        WScript.Echo "One or more selections were out of range."
      Else
        Dim orderedKeys: Set orderedKeys = CreateObject("System.Collections.ArrayList")
        For Each number In selections.Keys
          orderedKeys.Add CInt(number)
        Next
        orderedKeys.Sort

        Dim idx
        For i = 0 To orderedKeys.Count - 1
          idx = orderedKeys(i) - 1
          dict.Add CStr(dict.Count + 1), ordered(idx)
        Next

        Set selectionResult = dict
        Exit Do
      End If
    End If
  Loop

  Set SelectTemplatesIndividually = selectionResult
End Function

'======================================================================
' Delivery helpers
'======================================================================
Function ResolveDeliveryMode(tmpl, meta)
  Dim delivery: delivery = LCase(GetDictionaryValue(meta, "delivery"))
  If delivery = DELIVERY_BROADCAST Or delivery = DELIVERY_PERSONAL Then
    ResolveDeliveryMode = delivery
    Exit Function
  End If

  Dim inferred: inferred = DELIVERY_PERSONAL
  Dim toField: toField = LCase(GetDictionaryValue(meta, "to"))

  If InStr(toField, "[targets]") > 0 Then
    inferred = DELIVERY_BROADCAST
  ElseIf InStr(toField, "[email]") > 0 Then
    inferred = DELIVERY_PERSONAL
  ElseIf InStr(toField, "[name]") > 0 Or InStr(toField, "[surname]") > 0 Or InStr(toField, "[full_name]") > 0 Then
    inferred = DELIVERY_PERSONAL
  ElseIf Len(Trim(toField)) = 0 Then
    inferred = DELIVERY_PERSONAL
  Else
    inferred = DELIVERY_BROADCAST
  End If

  LogEvent "Delivery mode not specified for '" & tmpl("DisplayName") & "'. Defaulting to " & inferred
  WScript.Echo "  Delivery mode not specified. Defaulting to " & inferred & "."

  ResolveDeliveryMode = inferred
End Function

'======================================================================
' Template processing
'======================================================================
Sub ProcessTemplate(tmpl, recipientPackage, action)
  WScript.Echo "Preparing template: " & tmpl("DisplayName")
  LogEvent "Processing template: " & tmpl("DisplayName")

  Dim meta: Set meta = CloneDictionary(tmpl("Meta"))
  Call PromptForEmptyTags(meta, tmpl("DisplayName"))
  Call NormalizeNameParts(meta)

  Dim delivery: delivery = ResolveDeliveryMode(tmpl, meta)
  meta("delivery") = delivery

  Dim subjectText: subjectText = GetDictionaryValue(meta, "subject")
  Dim toText: toText = GetDictionaryValue(meta, "to")
  Dim ccText: ccText = GetDictionaryValue(meta, "cc")
  Dim bccText: bccText = GetDictionaryValue(meta, "bcc")
  Dim attachmentsValue: attachmentsValue = GetDictionaryValue(meta, "attachments")

  Dim attachments: attachments = ResolveAttachments(attachmentsValue, tmpl("Folder"))

  Dim inlineInfo: Set inlineInfo = PrepareInlineImageSupport(tmpl("Body"), tmpl("Path"))
  Dim htmlBody: htmlBody = inlineInfo("HtmlBody")
  Dim inlineFolder: inlineFolder = inlineInfo("InlineFolder")

  Dim targets: targets = GetRecipientResolutionAddresses(recipientPackage)
  Dim entries: entries = GetRecipientResolutionEntries(recipientPackage)
  Dim recipientCount: recipientCount = GetRecipientResolutionCount(recipientPackage)

  Dim baseContext: Set baseContext = BuildContextFromMeta(meta)
  baseContext("TEMPLATE_NAME") = tmpl("DisplayName")
  baseContext("RECIPIENT_COUNT") = CStr(recipientCount)
  baseContext("TARGETS") = Join(targets, "; ")
  Call NormalizeNameParts(baseContext)

  If delivery = DELIVERY_BROADCAST Then
    HandleBroadcast tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext
  Else
    HandlePersonal tmpl, entries, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext
  End If
End Sub

Sub HandleBroadcast(tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext)
  Dim joinedTargets: joinedTargets = Join(targets, "; ")
  baseContext("TARGETS") = joinedTargets

  Dim context: Set context = CloneDictionary(baseContext)
  Call EnsurePlaceholderValues(Array(subjectText, htmlBody, toText, ccText, bccText), context, tmpl("DisplayName"))
  Call NormalizeNameParts(context)

  Dim finalSubject: finalSubject = ApplyContextToText(subjectText, context)
  Dim finalBody: finalBody = ApplyContextToText(htmlBody, context)
  Dim templatedTo: templatedTo = ApplyContextToText(toText, context)
  Dim finalTo: finalTo = MergeRecipients(templatedTo, targets)
  Dim finalCc: finalCc = ApplyContextToText(ccText, context)
  Dim finalBcc: finalBcc = ApplyContextToText(bccText, context)

  finalSubject = EnsureCapitalizedNameText(finalSubject, context)
  finalBody = EnsureCapitalizedNameText(finalBody, context)
  finalTo = EnsureCapitalizedNameText(finalTo, context)
  finalCc = EnsureCapitalizedNameText(finalCc, context)
  finalBcc = EnsureCapitalizedNameText(finalBcc, context)

  LogEvent "Broadcast delivery for " & tmpl("DisplayName") & " -> " & finalTo
  WScript.Echo "  Sending broadcast to " & finalTo

  Dim ok
  ok = CreateMail(finalTo, finalCc, finalBcc, finalSubject, finalBody, attachments, action, "", "", inlineFolder, 0, logFilePath)
  If Not ok Then
    WScript.Echo "  Failed to create broadcast email for " & tmpl("DisplayName")
    LogEvent "ERROR: Broadcast delivery failed for " & tmpl("DisplayName")
  End If
End Sub

Sub HandlePersonal(tmpl, entries, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext)
  Dim entryUpper: entryUpper = GetRecipientResolutionArrayUpperBound(entries)
  If entryUpper < 0 Then Exit Sub

  Dim i
  For i = 0 To entryUpper
    Dim entry
    Set entry = entries(i)
    If entry Is Nothing Then
      ' Skip missing entries
    Else
      Dim address: address = SafeText(GetDictionaryValue(entry, "PreferredAddress"))
      If Len(address) = 0 Then address = SafeText(GetDictionaryValue(entry, "RawValue"))

      Dim context: Set context = CloneDictionary(baseContext)
      context("EMAIL") = address

      If Not context.Exists("NAME") Then context("NAME") = ""
      If Not context.Exists("SURNAME") Then context("SURNAME") = ""
      If Not context.Exists("FULL_NAME") Then context("FULL_NAME") = ""
      If Not context.Exists("DIRECTORY_SOURCE") Then context("DIRECTORY_SOURCE") = ""
      If Not context.Exists("DIRECTORY_ATTEMPTED") Then context("DIRECTORY_ATTEMPTED") = ""

      Dim nameInfo: Set nameInfo = Nothing
      If entry.Exists("Resolution") Then Set nameInfo = entry("Resolution")
      If nameInfo Is Nothing Then Set nameInfo = CreateEmptyResolutionInfo()

      If nameInfo("Found") Then
        Dim resolvedFirst: resolvedFirst = SafeText(nameInfo("First"))
        Dim resolvedLast: resolvedLast = SafeText(nameInfo("Last"))
        Dim resolvedFull: resolvedFull = SafeText(nameInfo("Full"))
        Dim resolvedEmail: resolvedEmail = ""
        If nameInfo.Exists("Email") Then resolvedEmail = SafeText(nameInfo("Email"))

        If Len(resolvedFirst) > 0 Then
          resolvedFirst = CapitalizeNamePart(resolvedFirst)
          nameInfo("First") = resolvedFirst
          context("NAME") = resolvedFirst
        End If

        If Len(resolvedLast) > 0 Then
          resolvedLast = CapitalizeNamePart(resolvedLast)
          nameInfo("Last") = resolvedLast
          context("SURNAME") = resolvedLast
        End If

        If Len(resolvedFull) > 0 Then
          context("FULL_NAME") = resolvedFull
        End If
        If Len(resolvedEmail) > 0 Then
          context("EMAIL") = resolvedEmail
        End If
      End If

      If nameInfo.Exists("DirectorySource") Then context("DIRECTORY_SOURCE") = nameInfo("DirectorySource")
      If nameInfo.Exists("DirectoryAttempted") Then context("DIRECTORY_ATTEMPTED") = nameInfo("DirectoryAttempted")

      Call NormalizeNameParts(context)
      Call ReviewResolvedName(address, context)

      Call NormalizeNameParts(context)

      If Len(context("FULL_NAME")) = 0 Then
        context("FULL_NAME") = Trim(context("NAME") & " " & context("SURNAME"))
      End If

      Call EnsurePlaceholderValues(Array(subjectText, htmlBody, toText, ccText, bccText), context, tmpl("DisplayName"))
      Call NormalizeNameParts(context)

      Dim finalSubject: finalSubject = ApplyContextToText(subjectText, context)
      Dim finalBody: finalBody = ApplyContextToText(htmlBody, context)
      Dim templatedTo: templatedTo = ApplyContextToText(toText, context)
      Dim finalTo: finalTo = MergeRecipients(templatedTo, Array(address))
      Dim finalCc: finalCc = ApplyContextToText(ccText, context)
      Dim finalBcc: finalBcc = ApplyContextToText(bccText, context)

      finalSubject = EnsureCapitalizedNameText(finalSubject, context)
      finalBody = EnsureCapitalizedNameText(finalBody, context)
      finalTo = EnsureCapitalizedNameText(finalTo, context)
      finalCc = EnsureCapitalizedNameText(finalCc, context)
      finalBcc = EnsureCapitalizedNameText(finalBcc, context)

      LogEvent "Personal delivery for " & tmpl("DisplayName") & " -> " & finalTo
      WScript.Echo "  Sending personal email to " & finalTo

      Dim ok
      ok = CreateMail(finalTo, finalCc, finalBcc, finalSubject, finalBody, attachments, action, "", "", inlineFolder, 0, logFilePath)
      If Not ok Then
        WScript.Echo "  Failed to create personal email for " & finalTo
        LogEvent "ERROR: Personal delivery failed for " & finalTo
      End If
    End If
  Next
End Sub

Function GetRecipientResolutionArrayUpperBound(values)
  GetRecipientResolutionArrayUpperBound = -1
  If Not IsArray(values) Then Exit Function

  On Error Resume Next
  Dim upper: upper = UBound(values)
  If Err.Number <> 0 Then
    Err.Clear
  Else
    GetRecipientResolutionArrayUpperBound = upper
  End If
  On Error GoTo 0
End Function

Sub NormalizeNameParts(dict)
  If dict Is Nothing Then Exit Sub

  Call NormalizeNameKey(dict, "NAME")
  Call NormalizeNameKey(dict, "SURNAME")
End Sub

Sub NormalizeNameKey(dict, canonicalKey)
  If dict Is Nothing Then Exit Sub

  Dim actualKey: actualKey = FindDictionaryKey(dict, canonicalKey)
  If Len(actualKey) = 0 Then
    If dict.Exists(canonicalKey) Then
      actualKey = canonicalKey
    ElseIf dict.Exists(LCase(canonicalKey)) Then
      actualKey = LCase(canonicalKey)
    Else
      Exit Sub
    End If
  End If

  Dim normalized: normalized = CapitalizeNamePart(dict(actualKey))
  dict(actualKey) = normalized

  If StrComp(actualKey, canonicalKey, vbTextCompare) <> 0 Then
    dict(canonicalKey) = normalized
  End If
End Sub

Function GetNormalizedNameValue(dict, canonicalKey)
  If dict Is Nothing Then
    GetNormalizedNameValue = ""
    Exit Function
  End If

  Call NormalizeNameKey(dict, canonicalKey)

  Dim actualKey: actualKey = FindDictionaryKey(dict, canonicalKey)
  If Len(actualKey) = 0 Then
    If dict.Exists(canonicalKey) Then
      actualKey = canonicalKey
    ElseIf dict.Exists(LCase(canonicalKey)) Then
      actualKey = LCase(canonicalKey)
    Else
      GetNormalizedNameValue = ""
      Exit Function
    End If
  End If

  GetNormalizedNameValue = SafeText(dict(actualKey))
End Function

Function EnsureCapitalizedNameText(text, context)
  Dim result: result = SafeText(text)
  If Len(result) = 0 Then
    EnsureCapitalizedNameText = result
    Exit Function
  End If

  Dim keys: keys = Array("NAME", "SURNAME")
  Dim i
  For i = 0 To UBound(keys)
    Dim key: key = keys(i)
    Dim normalized: normalized = GetNormalizedNameValue(context, key)
    If Len(normalized) > 0 Then
      Dim capitalized: capitalized = CapitalizeNamePart(normalized)
      If Len(capitalized) > 0 Then
        result = ReplaceWholeWordCaseInsensitive(result, normalized, capitalized)
      End If
    End If
  Next

  EnsureCapitalizedNameText = result
End Function

Function FindDictionaryKey(dict, target)
  Dim key
  For Each key In dict.Keys
    If StrComp(CStr(key), target, vbTextCompare) = 0 Then
      FindDictionaryKey = CStr(key)
      Exit Function
    End If
  Next
  FindDictionaryKey = ""
End Function

Function CapitalizeNamePart(value)
  Dim text: text = Trim(SafeText(value))
  If Len(text) = 0 Then
    CapitalizeNamePart = ""
  Else
    Dim firstChar: firstChar = UCase(Left(text, 1))
    If Len(text) > 1 Then
      CapitalizeNamePart = firstChar & Mid(text, 2)
    Else
      CapitalizeNamePart = firstChar
    End If
  End If
End Function

Function ReplaceWholeWordCaseInsensitive(text, findValue, replacement)
  Dim source: source = SafeText(text)
  Dim target: target = SafeText(findValue)

  If Len(target) = 0 Then
    ReplaceWholeWordCaseInsensitive = source
    Exit Function
  End If

  Dim result: result = source
  Dim startPos: startPos = 1
  Dim targetLen: targetLen = Len(target)

  Do While targetLen > 0
    Dim pos: pos = InStr(startPos, result, target, vbTextCompare)
    If pos = 0 Then Exit Do

    Dim beforeChar: beforeChar = ""
    Dim afterChar: afterChar = ""

    If pos > 1 Then beforeChar = Mid(result, pos - 1, 1)
    If pos + targetLen <= Len(result) Then afterChar = Mid(result, pos + targetLen, 1)

    If IsWordBoundary(beforeChar, afterChar) Then
      result = Left(result, pos - 1) & replacement & Mid(result, pos + targetLen)
      startPos = pos + Len(replacement)
    Else
      startPos = pos + targetLen
    End If
  Loop

  ReplaceWholeWordCaseInsensitive = result
End Function

Function IsWordBoundary(beforeChar, afterChar)
  If Len(beforeChar) = 0 Or Not IsWordCharacter(beforeChar) Then
    If Len(afterChar) = 0 Or Not IsWordCharacter(afterChar) Then
      IsWordBoundary = True
      Exit Function
    End If
  End If
  IsWordBoundary = False
End Function

Function IsWordCharacter(ch)
  If Len(ch) = 0 Then
    IsWordCharacter = False
  Else
    Dim code: code = AscW(ch)
    IsWordCharacter = _
      (code >= 48 And code <= 57) _
      Or (code >= 65 And code <= 90) _
      Or (code >= 97 And code <= 122) _
      Or (code >= &H00C0 And code <= &H02AF) _
      Or (code >= &H0370 And code <= &H1FFF)
  End If
End Function

Sub ReviewResolvedName(address, context)
  Call NormalizeNameParts(context)

  Dim currentFirst: currentFirst = SafeText(context("NAME"))
  Dim currentLast: currentLast = SafeText(context("SURNAME"))
  Dim currentFull: currentFull = SafeText(context("FULL_NAME"))
  Dim directorySource: directorySource = ""
  Dim directoryAttempted: directoryAttempted = ""
  If context.Exists("DIRECTORY_SOURCE") Then directorySource = SafeText(context("DIRECTORY_SOURCE"))
  If context.Exists("DIRECTORY_ATTEMPTED") Then directoryAttempted = SafeText(context("DIRECTORY_ATTEMPTED"))

  WScript.Echo "  Resolved name for " & address & ":"
  If Len(directorySource) > 0 Then
    WScript.Echo "    Directory source: " & directorySource
  Else
    WScript.Echo "    Directory source: (none)"
  End If
  If Len(directoryAttempted) > 0 Then
    WScript.Echo "    Address lists checked: " & directoryAttempted
  Else
    WScript.Echo "    Address lists checked: (none)"
  End If
  If Len(currentFirst) > 0 Then
    WScript.Echo "    First name: " & currentFirst
  Else
    WScript.Echo "    First name: (empty)"
  End If
  If Len(currentLast) > 0 Then
    WScript.Echo "    Surname: " & currentLast
  Else
    WScript.Echo "    Surname: (empty)"
  End If
  If Len(currentFull) > 0 Then
    WScript.Echo "    Display name: " & currentFull
  Else
    WScript.Echo "    Display name: (empty)"
  End If

  Dim requireReview: requireReview = False
  If Len(currentFirst) = 0 Or Len(currentLast) = 0 Or Len(currentFull) = 0 Then
    requireReview = True
  End If

  If requireReview Then
    Dim input
    input = PromptLine("    Enter first name or press Enter to keep current: ")
    If Len(input) > 0 Then
      context("NAME") = CapitalizeNamePart(input)
    End If

    input = PromptLine("    Enter surname or press Enter to keep current: ")
    If Len(input) > 0 Then
      context("SURNAME") = CapitalizeNamePart(input)
    End If

    Call NormalizeNameParts(context)

    Dim computedFull: computedFull = Trim(SafeText(context("NAME")) & " " & SafeText(context("SURNAME")))
    If Len(computedFull) = 0 Then computedFull = currentFull

    input = PromptLine("    Enter display name or press Enter to keep current: ")
    If Len(input) > 0 Then
      context("FULL_NAME") = input
    ElseIf Len(SafeText(context("FULL_NAME"))) = 0 Then
      context("FULL_NAME") = computedFull
    End If
  End If
End Sub
Function LoadLinkMapContext()
  Dim dict: Set dict = CreateCaseInsensitiveDictionary()
  If fso Is Nothing Then
    Set LoadLinkMapContext = dict
    Exit Function
  End If

  Dim mapPath: mapPath = fso.BuildPath(scriptFolder, LINK_MAP_FILE)
  If Len(mapPath) = 0 Or Not fso.FileExists(mapPath) Then
    LogEvent "Link map not found at " & mapPath
    Set LoadLinkMapContext = dict
    Exit Function
  End If

  On Error Resume Next
  Dim ts: Set ts = fso.OpenTextFile(mapPath, 1)
  If Err.Number <> 0 Then
    LogEvent "Unable to open link map: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Set LoadLinkMapContext = dict
    Exit Function
  End If
  On Error GoTo 0

  Do
    If ts.AtEndOfStream Then Exit Do
    Dim rawLine: rawLine = Trim(ts.ReadLine)
    If Len(rawLine) > 0 Then
      If Left(rawLine, 1) <> "#" Then
        Dim pos: pos = InStr(rawLine, "=")
        If pos > 0 Then
          Dim key: key = Trim(Left(rawLine, pos - 1))
          Dim value: value = Trim(Mid(rawLine, pos + 1))
          If Len(key) > 0 Then dict(UCase(key)) = value
        End If
      End If
    End If
  Loop
  ts.Close

  If dict.Count > 0 Then
    LogEvent "Loaded " & dict.Count & " link map entries"
  Else
    LogEvent "Link map was empty: " & mapPath
  End If

  Set LoadLinkMapContext = dict
End Function
'======================================================================
' Action helpers
'======================================================================
Function DetermineAction()
  Dim action: action = DEFAULT_ACTION
  If WScript.Arguments.Count > 0 Then
    Dim arg: arg = LCase(Trim(WScript.Arguments(0)))
    Select Case arg
      Case "display", "send", "save"
        action = arg
      Case Else
        LogEvent "Unknown action argument: " & arg & ". Using default '" & DEFAULT_ACTION & "'"
    End Select
  End If
  DetermineAction = action
End Function

'======================================================================
' Logging
'======================================================================
Sub LogEvent(message)
  On Error Resume Next
  Dim ts: Set ts = fso.OpenTextFile(logFilePath, 8, True)
  ts.WriteLine Now & "  " & message
  ts.Close
  On Error GoTo 0
End Sub
