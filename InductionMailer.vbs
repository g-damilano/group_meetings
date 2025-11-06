Option Explicit
SetLocale 1033

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

Call Main()

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

  Dim targets: targets = PromptTargetEmails()
  If UBound(targets) < 0 Then
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
    ProcessTemplate selectedTemplates(key), targets, action
  Next

  WScript.Echo "All requested templates processed."
End Sub

'======================================================================
' Console helpers
'======================================================================
Function EnsureConsoleHost()
  Dim host: host = LCase(WScript.FullName)
  If InStr(host, "cscript.exe") = 0 Then
    WScript.Echo "Please run this script with cscript.exe to use the console interface."
    EnsureConsoleHost = False
  Else
    EnsureConsoleHost = True
  End If
End Function

Function PromptLine(prompt)
  WScript.StdOut.Write prompt
  PromptLine = Trim(WScript.StdIn.ReadLine())
End Function

Function PromptTargetEmails()
  Dim responses: responses = Array()
  Do While True
    Dim input: input = PromptLine("Enter target email addresses (use ';' to separate multiple addresses): ")
    input = Replace(input, ",", ";")
    Dim parts: parts = Split(input, ";")
    Dim list: Set list = CreateObject("System.Collections.ArrayList")
    Dim seen: Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    seen.CompareMode = 1
    On Error GoTo 0
    Dim i, item, trimmed
    For i = 0 To UBound(parts)
      item = parts(i)
      trimmed = Trim(item)
      If Len(trimmed) > 0 Then
        If Not seen.Exists(LCase(trimmed)) Then
          seen.Add LCase(trimmed), True
          list.Add trimmed
        End If
      End If
    Next
    If list.Count > 0 Then
      responses = ArrayFromArrayList(list)
      Exit Do
    Else
      WScript.Echo "At least one email address is required."
    End If
  Loop
  PromptTargetEmails = responses
End Function

Function PromptOperationMode()
  WScript.Echo "Select the operation mode:"& vbCrLf & _
               "  1) Stage mode (send every email in a stage)" & vbCrLf & _
               "  2) Email mode (pick one or more templates)"
  Do While True
    Dim choice: choice = PromptLine("Enter choice (1/2) or press Enter to cancel: ")
    If Len(choice) = 0 Then
      PromptOperationMode = ""
      Exit Function
    End If
    Select Case choice
      Case "1": PromptOperationMode = "stage": Exit Function
      Case "2": PromptOperationMode = "email": Exit Function
      Case Else
        WScript.Echo "Invalid option. Please choose 1 or 2."
    End Select
  Loop
End Function

'======================================================================
' Environment bootstrap
'======================================================================
Sub InitializeEnvironment()
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
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

Function ReadFileSmart(filePath)
  On Error Resume Next
  Dim stream: Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile filePath
  Dim content: content = stream.ReadText
  stream.Close
  Set stream = Nothing
  If Err.Number = 0 Then
    ReadFileSmart = content
    Exit Function
  End If
  Err.Clear
  On Error GoTo 0

  Dim ts
  On Error Resume Next
  Set ts = fso.OpenTextFile(filePath, 1)
  If Err.Number = 0 Then
    ReadFileSmart = ts.ReadAll
    ts.Close
  Else
    ReadFileSmart = ""
  End If
  On Error GoTo 0
End Function

Function NormalizeLineEndings(text)
  text = Replace(text, vbCrLf, vbLf)
  text = Replace(text, vbCr, vbLf)
  NormalizeLineEndings = text
End Function

Function TrimLeadingBlankLines(text)
  Dim result: result = text
  Do While Len(result) > 0 And (Left(result, 1) = vbCr Or Left(result, 1) = vbLf)
    result = Mid(result, 2)
  Loop
  TrimLeadingBlankLines = result
End Function

Function GetDictionaryValue(dict, key)
  If dict.Exists(LCase(key)) Then
    GetDictionaryValue = dict(LCase(key))
  Else
    GetDictionaryValue = ""
  End If
End Function

Function CreateCaseInsensitiveDictionary()
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  dict.CompareMode = 1
  On Error GoTo 0
  Set CreateCaseInsensitiveDictionary = dict
End Function

'======================================================================
' Template selection
'======================================================================
Function Capitalize(text)
  If Len(text) = 0 Then
    Capitalize = text
  ElseIf Left(text, 1) = "(" Then
    Capitalize = text
  Else
    Capitalize = UCase(Left(text, 1)) & Mid(text, 2)
  End If
End Function

Function ContainsInArrayList(list, value)
  Dim i
  For i = 0 To list.Count - 1
    If LCase(list(i)) = LCase(value) Then
      ContainsInArrayList = True
      Exit Function
    End If
  Next
  ContainsInArrayList = False
End Function

Function TemplateSortValue(tmpl)
  Dim code: code = tmpl("Code")
  If Len(code) = 0 Then code = tmpl("FileName")
  TemplateSortValue = LCase(code)
End Function

Function SortTemplateList(list)
  Dim arr()
  ReDim arr(list.Count - 1)
  Dim i
  For i = 0 To list.Count - 1
    Set arr(i) = list(i)
  Next

  Dim j, keyTemplate
  For i = 1 To UBound(arr)
    Set keyTemplate = arr(i)
    j = i - 1
    Do While j >= 0 And TemplateSortValue(arr(j)) > TemplateSortValue(keyTemplate)
      Set arr(j + 1) = arr(j)
      j = j - 1
    Loop
    Set arr(j + 1) = keyTemplate
  Next

  Dim ordered: Set ordered = CreateObject("System.Collections.ArrayList")
  For i = 0 To UBound(arr)
    ordered.Add arr(i)
  Next
  Set SortTemplateList = ordered
End Function

Function BuildOrderedDictionary(list)
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  Dim ordered: Set ordered = SortTemplateList(list)
  Dim idx
  For idx = 0 To ordered.Count - 1
    Dim tmpl: Set tmpl = ordered(idx)
    dict.Add CStr(idx + 1), tmpl
  Next
  Set BuildOrderedDictionary = dict
End Function

Function ParseNumberList(text)
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  dict.CompareMode = 1
  Dim cleaned: cleaned = Replace(Replace(text, ",", " "), ";", " ")
  Dim parts: parts = Split(cleaned, " ")
  Dim part
  For Each part In parts
    part = Trim(part)
    If Len(part) > 0 And IsNumeric(part) Then
      dict(CStr(CInt(part))) = True
    End If
  Next
  Set ParseNumberList = dict
End Function

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

Function SelectTemplatesByStage(templates)
  Dim stageMap: Set stageMap = CreateObject("Scripting.Dictionary")
  stageMap.CompareMode = 1

  Dim orderedStages: orderedStages = Array("onboarding", "scheduling", "authorisation", "finalisation")
  Dim key, template
  For Each key In templates.Keys
    Set template = templates(key)
    Dim stage: stage = template("Stage")
    If Len(stage) = 0 Then stage = "(unspecified)"
    If Not stageMap.Exists(stage) Then
      stageMap(stage) = CreateObject("System.Collections.ArrayList")
    End If
    stageMap(stage).Add template
  Next

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

  Do While True
    Dim choice: choice = PromptLine("Select a stage by number or press Enter to cancel: ")
    If Len(choice) = 0 Then
      Set SelectTemplatesByStage = Nothing
      Exit Function
    End If
    If IsNumeric(choice) Then
      i = CInt(choice) - 1
      If i >= 0 And i < displayList.Count Then
        stageName = displayList(i)
        WScript.Echo "Stage '" & Capitalize(stageName) & "' selected."
        Set SelectTemplatesByStage = BuildOrderedDictionary(stageMap(stageName))
        Exit Function
      End If
    End If
    WScript.Echo "Invalid selection."
  Loop
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

  Do While True
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
' Template processing
'======================================================================
Sub ProcessTemplate(tmpl, targets, action)
  WScript.Echo "Preparing template: " & tmpl("DisplayName")
  LogEvent "Processing template: " & tmpl("DisplayName")

  Dim meta: Set meta = CloneDictionary(tmpl("Meta"))
  Call PromptForEmptyTags(meta, tmpl("DisplayName"))

  Dim delivery: delivery = LCase(GetDictionaryValue(meta, "delivery"))
  Do While delivery <> DELIVERY_BROADCAST And delivery <> DELIVERY_PERSONAL
    delivery = LCase(PromptLine("Specify delivery type for '" & tmpl("DisplayName") & "' (broadcast/personal): "))
  Loop
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

  Dim baseContext: Set baseContext = BuildContextFromMeta(meta)
  baseContext("TEMPLATE_NAME") = tmpl("DisplayName")
  baseContext("RECIPIENT_COUNT") = CStr(UBound(targets) + 1)
  baseContext("TARGETS") = Join(targets, "; ")

  If delivery = DELIVERY_BROADCAST Then
    HandleBroadcast tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext
  Else
    HandlePersonal tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext
  End If
End Sub

Sub HandleBroadcast(tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext)
  Dim joinedTargets: joinedTargets = Join(targets, "; ")
  baseContext("TARGETS") = joinedTargets

  Dim context: Set context = CloneDictionary(baseContext)
  Call EnsurePlaceholderValues(Array(subjectText, htmlBody, toText, ccText, bccText), context, tmpl("DisplayName"))

  Dim finalSubject: finalSubject = ApplyContextToText(subjectText, context)
  Dim finalBody: finalBody = ApplyContextToText(htmlBody, context)
  Dim templatedTo: templatedTo = ApplyContextToText(toText, context)
  Dim finalTo: finalTo = MergeRecipients(templatedTo, targets)
  Dim finalCc: finalCc = ApplyContextToText(ccText, context)
  Dim finalBcc: finalBcc = ApplyContextToText(bccText, context)

  LogEvent "Broadcast delivery for " & tmpl("DisplayName") & " -> " & finalTo
  WScript.Echo "  Sending broadcast to " & finalTo

  Dim ok
  ok = CreateMail(finalTo, finalCc, finalBcc, finalSubject, finalBody, attachments, action, "", "", inlineFolder, 0, logFilePath)
  If Not ok Then
    WScript.Echo "  Failed to create broadcast email for " & tmpl("DisplayName")
    LogEvent "ERROR: Broadcast delivery failed for " & tmpl("DisplayName")
  End If
End Sub

Sub HandlePersonal(tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext)
  Dim i
  For i = LBound(targets) To UBound(targets)
    Dim address: address = targets(i)
    Dim context: Set context = CloneDictionary(baseContext)
    context("EMAIL") = address

    If Not context.Exists("NAME") Then context("NAME") = ""
    If Not context.Exists("SURNAME") Then context("SURNAME") = ""
    If Not context.Exists("FULL_NAME") Then context("FULL_NAME") = ""

    Dim nameInfo: Set nameInfo = ResolveNameFromEmail(address)
    If nameInfo("Found") Then
      If Len(nameInfo("First")) > 0 Then context("NAME") = nameInfo("First")
      If Len(nameInfo("Last")) > 0 Then context("SURNAME") = nameInfo("Last")
      If Len(nameInfo("Full")) > 0 Then context("FULL_NAME") = nameInfo("Full")
    End If

    If Len(context("FULL_NAME")) = 0 Then
      context("FULL_NAME") = Trim(context("NAME") & " " & context("SURNAME"))
    End If

    Call EnsurePlaceholderValues(Array(subjectText, htmlBody, toText, ccText, bccText), context, tmpl("DisplayName"))

    Dim finalSubject: finalSubject = ApplyContextToText(subjectText, context)
    Dim finalBody: finalBody = ApplyContextToText(htmlBody, context)
    Dim templatedTo: templatedTo = ApplyContextToText(toText, context)
    Dim finalTo: finalTo = MergeRecipients(templatedTo, Array(address))
    Dim finalCc: finalCc = ApplyContextToText(ccText, context)
    Dim finalBcc: finalBcc = ApplyContextToText(bccText, context)

    LogEvent "Personal delivery for " & tmpl("DisplayName") & " -> " & finalTo
    WScript.Echo "  Sending personal email to " & finalTo

    Dim ok
    ok = CreateMail(finalTo, finalCc, finalBcc, finalSubject, finalBody, attachments, action, "", "", inlineFolder, 0, logFilePath)
    If Not ok Then
      WScript.Echo "  Failed to create personal email for " & finalTo
      LogEvent "ERROR: Personal delivery failed for " & finalTo
    End If
  Next
End Sub

Function BuildContextFromMeta(meta)
  Dim context: Set context = CloneDictionary(linkMapContext)
  If context Is Nothing Then Set context = CreateCaseInsensitiveDictionary()
  If Not meta Is Nothing Then
    Dim key
    For Each key In meta.Keys
      context(UCase(key)) = meta(key)
    Next
  End If
  Set BuildContextFromMeta = context
End Function

Function CloneDictionary(source)
  Dim clone: Set clone = CreateCaseInsensitiveDictionary()
  If Not source Is Nothing Then
    Dim key
    For Each key In source.Keys
      clone(key) = source(key)
    Next
  End If
  Set CloneDictionary = clone
End Function

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

  Do Until ts.AtEndOfStream
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

Sub PromptForEmptyTags(meta, templateName)
  If meta Is Nothing Then Exit Sub
  Dim key
  For Each key In meta.Keys
    If Len(Trim(meta(key))) = 0 Then
      Dim prompt: prompt = "Provide value for '" & UCase(key) & "' in template '" & templateName & "': "
      meta(key) = PromptLine(prompt)
    End If
  Next
End Sub

Sub EnsurePlaceholderValues(textItems, context, templateName)
  If context Is Nothing Then Exit Sub
  Dim tokens: tokens = ExtractPlaceholders(textItems)
  If IsArray(tokens) Then
    Dim upper
    On Error Resume Next
    upper = UBound(tokens)
    If Err.Number <> 0 Then
      Err.Clear
      upper = -1
    End If
    On Error GoTo 0
    Dim i
    For i = 0 To upper
      Dim token: token = tokens(i)
      If Len(token) > 0 Then
        Dim lookup: lookup = UCase(token)
        If context.Exists(lookup) Then
          If Len(CStr(context(lookup))) = 0 Then
            context(lookup) = PromptLine("Enter value for [" & lookup & "] in '" & templateName & "': ")
          End If
        Else
          LogEvent "Placeholder [" & lookup & "] has no known value for template '" & templateName & "'"
        End If
      End If
    Next
  End If
End Sub

Function ExtractPlaceholders(textItems)
  Dim collector: Set collector = CreateCaseInsensitiveDictionary()
  Dim re: Set re = CreateObject("VBScript.RegExp")
  re.Global = True
  re.IgnoreCase = False
  re.Pattern = "\[([A-Z0-9_]+)\]"

  If Not IsArray(textItems) Then
    ExtractPlaceholders = Array()
    Exit Function
  End If

  Dim i, text, matches, m
  For i = 0 To UBound(textItems)
    text = textItems(i)
    If Not IsEmpty(text) And Not IsNull(text) And Len(text) > 0 Then
      Set matches = re.Execute(CStr(text))
      For Each m In matches
        collector(UCase(m.SubMatches(0))) = True
      Next
    End If
  Next

  Dim arr
  If collector.Count = 0 Then
    arr = Array()
  Else
    Dim list: Set list = CreateObject("System.Collections.ArrayList")
    Dim key
    For Each key In collector.Keys
      list.Add key
    Next
    arr = ArrayFromArrayList(list)
  End If
  ExtractPlaceholders = arr
End Function

Function ApplyContextToText(text, context)
  Dim result: result = text
  If IsNull(result) Or IsEmpty(result) Then result = ""
  If context Is Nothing Then
    ApplyContextToText = result
    Exit Function
  End If

  Dim key, value
  For Each key In context.Keys
    value = context(key)
    If IsNull(value) Then value = ""
    result = Replace(result, "[" & key & "]", CStr(value))
  Next
  ApplyContextToText = result
End Function

Function SafeText(value)
  If IsObject(value) Then
    SafeText = ""
    Exit Function
  End If

  If IsNull(value) Or IsEmpty(value) Then
    SafeText = ""
  Else
    On Error Resume Next
    SafeText = CStr(value)
    If Err.Number <> 0 Then
      Err.Clear
      SafeText = ""
    End If
    On Error GoTo 0
  End If
End Function

Function ArrayFromArrayList(list)
  Dim arr()
  ReDim arr(list.Count - 1)
  Dim i
  For i = 0 To list.Count - 1
    arr(i) = list(i)
  Next
  ArrayFromArrayList = arr
End Function

Function MergeRecipients(baseList, extraArray)
  Dim seen: Set seen = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  seen.CompareMode = 1
  On Error GoTo 0
  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  AppendAddressesFromString list, seen, baseList
  Dim i
  If IsArray(extraArray) Then
    For i = LBound(extraArray) To UBound(extraArray)
      AppendAddress list, seen, extraArray(i)
    Next
  End If
  If list.Count = 0 Then
    MergeRecipients = ""
  Else
    MergeRecipients = Join(ArrayFromArrayList(list), "; ")
  End If
End Function

Sub AppendAddressesFromString(list, seen, text)
  If Len(Trim(text)) = 0 Then Exit Sub
  Dim normalized: normalized = Replace(text, ",", ";")
  Dim parts: parts = Split(normalized, ";")
  Dim i
  For i = 0 To UBound(parts)
    AppendAddress list, seen, parts(i)
  Next
End Sub

Sub AppendAddress(list, seen, address)
  Dim trimmed: trimmed = Trim(address)
  If Len(trimmed) = 0 Then Exit Sub
  Dim key: key = LCase(trimmed)
  If Not seen.Exists(key) Then
    seen.Add key, True
    list.Add trimmed
  End If
End Sub

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

Function ResolveAttachments(rawValue, baseFolder)
  Dim cleaned: cleaned = Trim(rawValue)
  cleaned = Replace(cleaned, ",", ";")
  If Len(cleaned) = 0 Then
    ResolveAttachments = Array()
    Exit Function
  End If

  Dim parts: parts = Split(cleaned, ";")
  Dim resolvedList: Set resolvedList = CreateObject("System.Collections.ArrayList")
  Dim i, candidate, item
  For i = 0 To UBound(parts)
    item = Trim(parts(i))
    If Len(item) > 0 Then
      candidate = item
      If Not fso.FileExists(candidate) Then
        If Left(candidate, 2) = ".\" Or Left(candidate, 2) = "./" Then
          candidate = fso.BuildPath(baseFolder, Mid(candidate, 3))
        ElseIf Not IsAbsolutePath(candidate) Then
          candidate = fso.BuildPath(baseFolder, candidate)
        End If
      End If
      If fso.FileExists(candidate) Then
        resolvedList.Add candidate
        LogEvent "Attachment queued: " & candidate
      Else
        LogEvent "Attachment not found: " & item
      End If
    End If
  Next

  If resolvedList.Count = 0 Then
    ResolveAttachments = Array()
  Else
    ResolveAttachments = ArrayFromArrayList(resolvedList)
  End If
End Function

Function IsAbsolutePath(path)
  If Len(path) < 2 Then
    IsAbsolutePath = False
  ElseIf Mid(path, 2, 1) = ":" Then
    IsAbsolutePath = True
  ElseIf Left(path, 2) = "\\" Then
    IsAbsolutePath = True
  Else
    IsAbsolutePath = False
  End If
End Function

Function PrepareInlineImageSupport(htmlBody, templatePath)
  Dim info: Set info = CreateObject("Scripting.Dictionary")
  info("HtmlBody") = htmlBody
  info("InlineFolder") = ""

  Dim baseFolder: baseFolder = fso.GetParentFolderName(templatePath)
  Dim baseName: baseName = fso.GetBaseName(templatePath)
  Dim inlineFolder: inlineFolder = fso.BuildPath(baseFolder, baseName & ".files")

  If Not fso.FolderExists(inlineFolder) Then
    Set PrepareInlineImageSupport = info
    Exit Function
  End If

  LogEvent "Inline image folder detected: " & inlineFolder

  Dim file
  For Each file In fso.GetFolder(inlineFolder).Files
    info("HtmlBody") = RewriteInlineReferences(info("HtmlBody"), baseName, file.Name)
  Next

  info("InlineFolder") = inlineFolder
  Set PrepareInlineImageSupport = info
End Function

Function RewriteInlineReferences(htmlText, templateBaseName, fileName)
  Dim sanitizedCid: sanitizedCid = SanitizeContentId(fileName)
  Dim encodedBase: encodedBase = Replace(templateBaseName, " ", "%20")
  Dim encodedFile: encodedFile = Replace(fileName, " ", "%20")
  Dim target: target = "cid:" & sanitizedCid
  Dim backslash: backslash = Chr(92)

  htmlText = Replace(htmlText, templateBaseName & ".files/" & fileName, target)
  htmlText = Replace(htmlText, templateBaseName & ".files" & backslash & fileName, target)
  htmlText = Replace(htmlText, encodedBase & ".files/" & encodedFile, target)
  htmlText = Replace(htmlText, encodedBase & ".files" & backslash & encodedFile, target)

  htmlText = Replace(htmlText, "=""" & fileName & """", "=""" & target & """")
  htmlText = Replace(htmlText, "='" & fileName & "'", "='" & target & "'")
  htmlText = Replace(htmlText, "=""" & encodedFile & """", "=""" & target & """")
  htmlText = Replace(htmlText, "='" & encodedFile & "'", "='" & target & "'")

  RewriteInlineReferences = htmlText
End Function

Function SanitizeContentId(fileName)
  Dim cid: cid = fso.GetFileName(fileName)
  cid = Replace(cid, " ", "_")
  SanitizeContentId = cid
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
