Option Explicit
SetLocale 1033

'======================================================================
' InductionMailer.vbs — Compose induction templates via dynamic link mailer
'======================================================================

'---------------------------
' Global objects and config
'---------------------------
Dim fso, scriptFolder, templateFolder, logFilePath
Const TEMPLATE_DIR_NAME = "templates"
Const LOG_FILE_NAME = "InductionMailer.log"
Const DEFAULT_ACTION = "display"   ' display | send | save
Const MAILER_LIB_FILE = "mail_utils_mailer_dynamiclinks.vbs"
Const MAILER_LIB_EXT_DIR = "mail_utils_utf8_plus_dynamiclinks_ext"
Const MAILER_LIB_EXT_SUBPATH = "mail_utils_utf8_plus_dynamiclinks_ext\\mail_utils_mailer_dynamiclinks.vbs"
Const LINK_MAP_FILE = "links_map.txt"

Call Main()

'---------------------------
' Entry point
'---------------------------
Sub Main()
  InitializeEnvironment

  If Not LoadMailerLibrary() Then
    ShowError "Unable to load mail_utils_mailer_dynamiclinks.vbs"
    Exit Sub
  End If

  Dim templates: Set templates = DiscoverTemplates(templateFolder)
  If templates.Count = 0 Then
    ShowError "No HTML templates found in " & templateFolder
    Exit Sub
  End If

  Dim selectionKey: selectionKey = PromptTemplateSelection(templates)
  If Len(selectionKey) = 0 Then
    LogEvent "User cancelled template selection."
    Exit Sub
  End If

  Dim templatePath: templatePath = templates(selectionKey)
  LogEvent "Selected template: " & templatePath

  Dim templateData: Set templateData = ParseTemplateFile(templatePath)
  If templateData Is Nothing Then
    ShowError "Could not parse template: " & templatePath
    Exit Sub
  End If

  Dim action: action = DetermineAction()
  Dim attachments: attachments = ResolveAttachments(templateData("Attachments"), templateData("TemplateFolder"))
  Dim htmlBody: htmlBody = templateData("HtmlBody")

  Dim inlineInfo: Set inlineInfo = PrepareInlineImageSupport(htmlBody, templatePath)
  htmlBody = inlineInfo("HtmlBody")
  Dim inlineFolder: inlineFolder = inlineInfo("InlineFolder")

  Dim ok
  ok = CreateMail(templateData("To"), _
                  templateData("Cc"), _
                  templateData("Bcc"), _
                  templateData("Subject"), _
                  htmlBody, _
                  attachments, _
                  action, _
                  "", _
                  "", _
                  inlineFolder, _
                  0, _
                  logFilePath)

  If ok Then
    LogEvent "CreateMail succeeded with action='" & action & "'"
  Else
    ShowError "CreateMail reported a failure. Check log for details."
  End If
End Sub

'---------------------------
' Environment helpers
'---------------------------
Sub InitializeEnvironment()
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
  templateFolder = fso.BuildPath(scriptFolder, TEMPLATE_DIR_NAME)
  logFilePath = fso.BuildPath(scriptFolder, LOG_FILE_NAME)
  LogEvent "=== Induction Mailer started ==="
  LogEvent "Template folder: " & templateFolder
End Sub

Function LoadMailerLibrary()
  Dim candidates
  candidates = Array(_
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

  LogEvent "Failed to locate mail_utils_mailer_dynamiclinks.vbs"
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
        LogEvent "Copied links_map.txt to script folder"
      Else
        LogEvent "Failed to copy links_map.txt: " & Err.Description
        Err.Clear
      End If
    End If
  End If
  On Error GoTo 0
End Sub

'---------------------------
' Template discovery and selection
'---------------------------
Function DiscoverTemplates(folderPath)
  Dim dict: Set dict = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  If Not fso.FolderExists(folderPath) Then
    LogEvent "Template folder missing: " & folderPath
    Set DiscoverTemplates = dict
    Exit Function
  End If
  On Error GoTo 0

  Dim files, file, idx
  Set files = fso.GetFolder(folderPath).Files

  Dim orderedPaths, sortFailed
  On Error Resume Next
  Set orderedPaths = CreateObject("System.Collections.ArrayList")
  If Err.Number <> 0 Then
    sortFailed = True
    Err.Clear
  End If
  On Error GoTo 0

  For Each file In files
    Dim ext: ext = LCase(fso.GetExtensionName(file.Name))
    If ext = "html" Or ext = "htm" Then
      If sortFailed Then
        dict.Add CStr(dict.Count + 1), file.Path
      Else
        orderedPaths.Add file.Path
      End If
    End If
  Next

  If Not sortFailed Then
    orderedPaths.Sort
    For idx = 0 To orderedPaths.Count - 1
      dict.Add CStr(idx + 1), orderedPaths(idx)
    Next
  End If

  Set DiscoverTemplates = dict
End Function

Function PromptTemplateSelection(templates)
  Dim key, prompt, i
  prompt = "Select a template by number:" & vbCrLf & vbCrLf
  For i = 1 To templates.Count
    key = CStr(i)
    If templates.Exists(key) Then
      prompt = prompt & key & ". " & fso.GetFileName(templates(key)) & vbCrLf
    End If
  Next
  prompt = prompt & vbCrLf & "Cancel to exit."

  Dim selection
  selection = Trim(InputBox(prompt, "Induction Mailer"))
  If Len(selection) = 0 Then
    PromptTemplateSelection = ""
  ElseIf templates.Exists(selection) Then
    PromptTemplateSelection = selection
  Else
    MsgBox "Invalid selection.", vbExclamation, "Induction Mailer"
    PromptTemplateSelection = PromptTemplateSelection(templates)
  End If
End Function

'---------------------------
' Template parsing
'---------------------------
Function ParseTemplateFile(filePath)
  On Error Resume Next
  Dim s: Set s = CreateObject("ADODB.Stream")
  s.Type = 2
  s.Charset = "utf-8"
  s.Open
  s.LoadFromFile filePath
  Dim content: content = s.ReadText
  s.Close
  Set s = Nothing
  If Err.Number <> 0 Then
    Dim fallback: fallback = ReadFileWithFallback(filePath)
    If fallback("Success") Then
      content = fallback("Content")
    Else
      LogEvent "Error reading template: " & filePath & " | " & Err.Description
      Err.Clear
      Set ParseTemplateFile = Nothing
      Exit Function
    End If
  End If
  On Error GoTo 0

  content = NormalizeLineEndings(content)
  Dim parsed: Set parsed = ExtractMetadataAndBody(content)
  If parsed Is Nothing Then
    LogEvent "Metadata extraction failed for: " & filePath
    Set ParseTemplateFile = Nothing
    Exit Function
  End If

  Dim data: Set data = CreateObject("Scripting.Dictionary")
  data("TemplateFolder") = fso.GetParentFolderName(filePath)
  data("To") = parsed("To")
  data("Cc") = parsed("Cc")
  data("Bcc") = parsed("Bcc")
  If Len(Trim(parsed("Subject"))) = 0 Then
    data("Subject") = fso.GetBaseName(filePath)
  Else
    data("Subject") = parsed("Subject")
  End If
  data("Attachments") = parsed("Attachments")
  data("HtmlBody") = parsed("HtmlBody")
  data("HtmlPath") = filePath
  Set ParseTemplateFile = data
End Function

Function NormalizeLineEndings(text)
  text = Replace(text, vbCrLf, vbLf)
  text = Replace(text, vbCr, vbLf)
  NormalizeLineEndings = text
End Function

Function TrimTrailingEmptyLines(text)
  Dim cleaned: cleaned = text
  Do While Len(cleaned) > 0 And Right(cleaned, 1) = vbLf
    cleaned = Left(cleaned, Len(cleaned) - 1)
  Loop
  TrimTrailingEmptyLines = cleaned
End Function

'---------------------------
' HTML metadata helpers
'---------------------------
Function ExtractMetadataAndBody(content)
  Dim data: Set data = CreateObject("Scripting.Dictionary")
  data("To") = ""
  data("Cc") = ""
  data("Bcc") = ""
  data("Subject") = ""
  data("Attachments") = ""

  Dim working: working = RemoveUtf8Bom(content)
  Dim trimmed: trimmed = LTrim(working)
  Dim leadingWs: leadingWs = Left(working, Len(working) - Len(trimmed))

  Dim metaText, htmlBody, metaEnd
  metaText = ""
  htmlBody = working

  If Left(trimmed, 4) = "<!--" Then
    metaEnd = InStr(5, trimmed, "-->")
    If metaEnd > 0 Then
      metaText = Mid(trimmed, 5, metaEnd - 5)
      htmlBody = leadingWs & Mid(trimmed, metaEnd + 3)
    End If
  End If

  If Len(metaText) > 0 Then
    Dim lines, line, pos, key, value
    metaText = NormalizeLineEndings(metaText)
    lines = Split(metaText, vbLf)
    For Each line In lines
      If Len(Trim(line)) > 0 Then
        pos = InStr(line, ":")
        If pos > 0 Then
          key = LCase(Trim(Left(line, pos - 1)))
          value = Trim(Mid(line, pos + 1))
          Select Case key
            Case "to": data("To") = value
            Case "cc": data("Cc") = value
            Case "bcc": data("Bcc") = value
            Case "subject": data("Subject") = value
            Case "attachments": data("Attachments") = value
          End Select
        End If
      End If
    Next
  End If

  data("HtmlBody") = TrimLeadingBlankLines(htmlBody)
  Set ExtractMetadataAndBody = data
End Function

Function TrimLeadingBlankLines(text)
  Dim result: result = text
  Do While Len(result) > 0 And (Left(result, 1) = vbLf Or Left(result, 1) = vbCr)
    result = Mid(result, 2)
  Loop
  TrimLeadingBlankLines = result
End Function

Function RemoveUtf8Bom(text)
  If Len(text) >= 3 Then
    If Asc(Mid(text, 1, 1)) = 239 And _
       Asc(Mid(text, 2, 1)) = 187 And _
       Asc(Mid(text, 3, 1)) = 191 Then
      RemoveUtf8Bom = Mid(text, 4)
      Exit Function
    End If
  End If
  RemoveUtf8Bom = text
End Function

Function ReadFileWithFallback(filePath)
  Dim result: Set result = CreateObject("Scripting.Dictionary")
  result("Success") = False
  result("Content") = ""

  Dim encodings: encodings = Array("utf-8", "utf-16le", "unicode", "windows-1252")
  Dim i

  For i = LBound(encodings) To UBound(encodings)
    On Error Resume Next
    Dim s: Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = encodings(i)
    s.Open
    s.LoadFromFile filePath
    Dim text: text = s.ReadText
    s.Close
    Set s = Nothing
    If Err.Number = 0 Then
      result("Success") = True
      result("Content") = text
      Exit For
    End If
    Err.Clear
    On Error GoTo 0
  Next

  On Error GoTo 0

  ReadFileWithFallback = result
End Function

'---------------------------
' Attachment and body helpers
'---------------------------
Function ResolveAttachments(rawValue, baseFolder)
  Dim cleaned: cleaned = Trim(rawValue)
  cleaned = Replace(cleaned, ",", ";")
  If Len(cleaned) = 0 Then
    ResolveAttachments = Array()
    Exit Function
  End If

  Dim parts: parts = Split(cleaned, ";")
  Dim resolved(), count, i, candidate
  count = -1

  For i = 0 To UBound(parts)
    Dim item: item = Trim(parts(i))
    If Len(item) > 0 Then
      candidate = item
      If Not fso.FileExists(candidate) Then
        If Left(candidate, 2) = ".\" Then
          candidate = fso.BuildPath(baseFolder, Mid(candidate, 3))
        ElseIf Not IsAbsolutePath(candidate) Then
          candidate = fso.BuildPath(baseFolder, candidate)
        End If
      End If

      If fso.FileExists(candidate) Then
        count = count + 1
        ReDim Preserve resolved(count)
        resolved(count) = candidate
        LogEvent "Queued attachment: " & candidate
      Else
        LogEvent "Attachment not found: " & item
      End If
    End If
  Next

  If count = -1 Then
    ResolveAttachments = Array()
  Else
    ResolveAttachments = resolved
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

'---------------------------
' Inline image helpers
'---------------------------
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

  htmlText = Replace(htmlText, templateBaseName & ".files/" & fileName, target)
  htmlText = Replace(htmlText, templateBaseName & ".files\" & fileName, target)
  htmlText = Replace(htmlText, encodedBase & ".files/" & encodedFile, target)
  htmlText = Replace(htmlText, encodedBase & ".files\" & encodedFile, target)

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

'---------------------------
' Action helpers
'---------------------------
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

'---------------------------
' Logging and messaging
'---------------------------
Sub LogEvent(message)
  On Error Resume Next
  Dim ts: Set ts = fso.OpenTextFile(logFilePath, 8, True)
  ts.WriteLine Now & "  " & message
  ts.Close
  On Error GoTo 0
End Sub

Sub ShowError(msg)
  MsgBox msg, vbCritical, "Induction Mailer"
  LogEvent "ERROR: " & msg
End Sub
