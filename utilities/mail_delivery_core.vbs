'====================================================================
' mail_delivery_core.vbs — Shared helpers for composing and delivering Outlook mail
'====================================================================
Option Explicit
SetLocale 1033

'--------------------------------------------------------------------
' Public entry point
'--------------------------------------------------------------------
Function MailDeliveryExecute(config)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim logPath: logPath = MailDeliveryResolveLogPath(config, fso)
  Dim logHandler: Set logHandler = MailDeliveryResolveLogHandler(config)

  MailDeliveryLog logPath, "Mail delivery core initialised."

  If Not MailDeliveryEnsureComposer() Then
    MailDeliveryLogViaHandler logHandler, logPath, "ERROR: Shared mail composer module could not be loaded."
    MailDeliveryExecute = False
    Exit Function
  End If

  Dim bodyInfo: Set bodyInfo = MailDeliveryResolveBodyContext(config, fso, logPath)
  Dim composerOptions: Set composerOptions = CreateObject("Scripting.Dictionary")

  composerOptions("To") = MailDeliveryGetOption(config, "To", "")
  composerOptions("Cc") = MailDeliveryGetOption(config, "Cc", "")
  composerOptions("Bcc") = MailDeliveryGetOption(config, "Bcc", "")
  composerOptions("Subject") = MailDeliveryGetOption(config, "Subject", "")
  composerOptions("BodyHtml") = bodyInfo("HtmlBody")
  composerOptions("Action") = LCase(Trim(CStr(MailDeliveryGetOption(config, "Action", "display"))))
  composerOptions("SavePath") = MailDeliveryGetOption(config, "SavePath", "")
  composerOptions("DeferMinutes") = MailDeliveryGetOption(config, "DeferMinutes", 0)
  composerOptions("LogPath") = logPath
  Set composerOptions("LogHandler") = logHandler

  MailDeliveryCopyIfExists config, composerOptions, "HeadContent"
  MailDeliveryCopyIfExists config, composerOptions, "BodyAttributes"
  MailDeliveryCopyIfExists config, composerOptions, "BodyTransforms"
  MailDeliveryCopyIfExists config, composerOptions, "InlineImageCidTransform"
  MailDeliveryCopyIfExists config, composerOptions, "SignatureProvider"

  Dim signatureLiteral
  signatureLiteral = MailDeliveryGetOption(config, "SignatureLiteral", "")
  If Len(signatureLiteral) > 0 Then composerOptions("SignatureProvider") = signatureLiteral

  Dim inlineFolder
  inlineFolder = MailDeliveryGetOption(config, "InlineImageFolder", "")
  If Len(inlineFolder) = 0 Then inlineFolder = bodyInfo("InlineFolder")
  If Len(inlineFolder) > 0 Then composerOptions("InlineImageFolder") = inlineFolder

  Dim attachments
  attachments = MailDeliveryResolveAttachmentsOption(config, fso, logPath)
  If Not IsEmpty(attachments) Then composerOptions("Attachments") = attachments

  Dim ok: ok = ComposeAndSendMail(composerOptions)
  If Not ok Then MailDeliveryLogViaHandler logHandler, logPath, "Mail delivery failed for subject: " & composerOptions("Subject")
  MailDeliveryExecute = ok
End Function

'--------------------------------------------------------------------
' Body resolution helpers
'--------------------------------------------------------------------
Function MailDeliveryResolveBodyContext(config, fso, logPath)
  Dim htmlBody: htmlBody = MailDeliveryGetOption(config, "BodyHtml", "")
  Dim templateSource: templateSource = MailDeliveryGetOption(config, "TemplateSource", "")
  Dim sourcePath: sourcePath = ""

  If Len(Trim(templateSource)) > 0 Then
    htmlBody = MailDeliveryLoadHtml(templateSource, fso)
    If fso.FileExists(templateSource) Then sourcePath = templateSource
  End If

  Dim linkMapPath: linkMapPath = MailDeliveryGetOption(config, "LinkMapPath", "")
  If Len(linkMapPath) > 0 Then htmlBody = ApplyLinkMap(htmlBody, linkMapPath)

  htmlBody = MailDeliveryApplyLinkCallback(htmlBody, config)
  htmlBody = MailDeliveryApplyLinkReplacements(htmlBody, config)

  Dim info: Set info = MailDeliveryPrepareInlineSupport(htmlBody, sourcePath, fso, logPath)
  Set MailDeliveryResolveBodyContext = info
End Function

Function MailDeliveryLoadHtml(templateSource, fso)
  Dim resolved
  resolved = Trim(CStr(templateSource))
  If Len(resolved) = 0 Then
    MailDeliveryLoadHtml = ""
  ElseIf fso.FileExists(resolved) Then
    MailDeliveryLoadHtml = ReadUTF8File(resolved)
  Else
    MailDeliveryLoadHtml = resolved
  End If
End Function

Function MailDeliveryApplyLinkCallback(htmlBody, config)
  Dim transformed: transformed = htmlBody
  If config Is Nothing Then
    MailDeliveryApplyLinkCallback = transformed
    Exit Function
  End If

  On Error Resume Next
  If config.Exists("LinkCallback") Then
    If IsObject(config("LinkCallback")) Then
      transformed = config("LinkCallback")(transformed)
    End If
  End If
  On Error GoTo 0

  MailDeliveryApplyLinkCallback = transformed
End Function

Function MailDeliveryApplyLinkReplacements(htmlBody, config)
  Dim result: result = htmlBody
  If config Is Nothing Then
    MailDeliveryApplyLinkReplacements = result
    Exit Function
  End If

  On Error Resume Next
  If config.Exists("LinkReplacements") Then
    If IsObject(config("LinkReplacements")) Then
      Dim replacements: Set replacements = config("LinkReplacements")
      Dim key
      For Each key In replacements.Keys
        result = Replace(result, key, replacements(key))
      Next
    End If
  End If
  On Error GoTo 0

  MailDeliveryApplyLinkReplacements = result
End Function

Function MailDeliveryPrepareInlineSupport(htmlBody, templatePath, fso, logPath)
  Dim info: Set info = CreateObject("Scripting.Dictionary")
  info("HtmlBody") = htmlBody
  info("InlineFolder") = ""

  If Len(templatePath) = 0 Then
    Set MailDeliveryPrepareInlineSupport = info
    Exit Function
  End If

  Dim baseFolder: baseFolder = fso.GetParentFolderName(templatePath)
  Dim baseName: baseName = fso.GetBaseName(templatePath)
  Dim inlineFolder: inlineFolder
  inlineFolder = fso.BuildPath(baseFolder, baseName & ".files")

  If Not fso.FolderExists(inlineFolder) Then
    Set MailDeliveryPrepareInlineSupport = info
    Exit Function
  End If

  MailDeliveryLog logPath, "Inline image folder detected: " & inlineFolder

  Dim file
  For Each file In fso.GetFolder(inlineFolder).Files
    info("HtmlBody") = MailDeliveryRewriteInlineReferences(info("HtmlBody"), baseName, file.Name)
  Next

  info("InlineFolder") = inlineFolder
  Set MailDeliveryPrepareInlineSupport = info
End Function

Function MailDeliveryRewriteInlineReferences(htmlText, templateBaseName, fileName)
  Dim sanitizedCid: sanitizedCid = MailDeliverySanitizeContentId(fileName)
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

  MailDeliveryRewriteInlineReferences = htmlText
End Function

Function MailDeliverySanitizeContentId(fileName)
  Dim cid: cid = fileName
  cid = Replace(cid, " ", "_")
  MailDeliverySanitizeContentId = cid
End Function

'--------------------------------------------------------------------
' Attachment helpers
'--------------------------------------------------------------------
Function MailDeliveryResolveAttachmentsOption(config, fso, logPath)
  Dim attachments
  attachments = MailDeliveryGetOption(config, "Attachments", Empty)
  If Not IsEmpty(attachments) Then
    MailDeliveryResolveAttachmentsOption = attachments
    Exit Function
  End If

  Dim rawList: rawList = MailDeliveryGetOption(config, "AttachmentList", "")
  If Len(rawList) = 0 Then
    MailDeliveryResolveAttachmentsOption = Empty
    Exit Function
  End If

  Dim baseFolder: baseFolder = MailDeliveryGetOption(config, "AttachmentBaseFolder", "")
  Dim resolved
  resolved = MailDeliveryResolveAttachmentList(rawList, baseFolder, fso, logPath)
  MailDeliveryResolveAttachmentsOption = resolved
End Function

Function MailDeliveryResolveAttachmentList(rawValue, baseFolder, fso, logPath)
  Dim cleaned: cleaned = Trim(CStr(rawValue))
  cleaned = Replace(cleaned, ",", ";")
  If Len(cleaned) = 0 Then
    MailDeliveryResolveAttachmentList = Empty
    Exit Function
  End If

  Dim folderBase: folderBase
  folderBase = Trim(CStr(baseFolder))
  If Len(folderBase) = 0 Then folderBase = MailDeliveryScriptFolder(fso)

  Dim parts: parts = Split(cleaned, ";")
  Dim resolved()
  Dim count: count = 0
  Dim i, item, candidate

  For i = 0 To UBound(parts)
    item = Trim(parts(i))
    If Len(item) > 0 Then
      candidate = MailDeliveryResolveAttachmentPath(item, folderBase, fso)
      If Len(candidate) > 0 And fso.FileExists(candidate) Then
        MailDeliveryAppendPath resolved, count, candidate
        MailDeliveryLog logPath, "Attachment queued: " & candidate
      Else
        MailDeliveryLog logPath, "Attachment not found: " & item
      End If
    End If
  Next

  If count = 0 Then
    MailDeliveryResolveAttachmentList = Empty
  Else
    ReDim Preserve resolved(count - 1)
    MailDeliveryResolveAttachmentList = resolved
  End If
End Function

Function MailDeliveryResolveAttachmentPath(item, baseFolder, fso)
  Dim candidate: candidate = item
  If MailDeliveryIsAbsolutePath(candidate) Then
    MailDeliveryResolveAttachmentPath = candidate
    Exit Function
  End If

  If Left(candidate, 2) = "./" Or Left(candidate, 2) = ".\" Then
    candidate = Mid(candidate, 3)
  End If

  If Len(baseFolder) = 0 Then baseFolder = MailDeliveryScriptFolder(fso)
  MailDeliveryResolveAttachmentPath = fso.BuildPath(baseFolder, candidate)
End Function

Sub MailDeliveryAppendPath(ByRef list, ByRef count, path)
  If count = 0 Then
    ReDim list(0)
  Else
    ReDim Preserve list(count)
  End If
  list(count) = path
  count = count + 1
End Sub

Function MailDeliveryIsAbsolutePath(path)
  Dim trimmed: trimmed = Trim(CStr(path))
  If Len(trimmed) < 2 Then
    MailDeliveryIsAbsolutePath = False
  ElseIf Mid(trimmed, 2, 1) = ":" Then
    MailDeliveryIsAbsolutePath = True
  ElseIf Left(trimmed, 2) = "\\" Then
    MailDeliveryIsAbsolutePath = True
  Else
    MailDeliveryIsAbsolutePath = False
  End If
End Function

'--------------------------------------------------------------------
' Logging helpers
'--------------------------------------------------------------------
Function MailDeliveryResolveLogPath(config, fso)
  Dim requested: requested = ""
  If Not config Is Nothing Then
    On Error Resume Next
    If config.Exists("LogPath") Then requested = Trim(CStr(config("LogPath")))
    On Error GoTo 0
  End If
  If Len(requested) = 0 Then requested = MailDeliveryDefaultLogPath(fso)
  MailDeliveryResolveLogPath = requested
End Function

Function MailDeliveryDefaultLogPath(fso)
  Dim folder: folder = MailDeliveryScriptFolder(fso)
  MailDeliveryDefaultLogPath = fso.BuildPath(folder, "mail_util.log")
End Function

Function MailDeliveryScriptFolder(fso)
  Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  MailDeliveryScriptFolder = folder
End Function

Function MailDeliveryResolveLogHandler(config)
  Dim handler
  Set handler = Nothing
  If Not config Is Nothing Then
    On Error Resume Next
    If config.Exists("LogHandler") Then Set handler = config("LogHandler")
    On Error GoTo 0
  End If
  If handler Is Nothing Then Set handler = GetRef("MailDeliveryLog")
  Set MailDeliveryResolveLogHandler = handler
End Function

Sub MailDeliveryLogViaHandler(logHandler, logPath, message)
  On Error Resume Next
  If Not logHandler Is Nothing Then
    logHandler logPath, message
  Else
    MailDeliveryLog logPath, message
  End If
  On Error GoTo 0
End Sub

Sub MailDeliveryLog(logPath, message)
  Dim target: target = Trim(CStr(logPath))
  If Len(target) = 0 Then
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    target = MailDeliveryDefaultLogPath(fso)
  End If

  On Error Resume Next
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(target, 8, True)
  f.WriteLine Now & "  " & message
  f.Close
  On Error GoTo 0
End Sub

'--------------------------------------------------------------------
' Composer bootstrap
'--------------------------------------------------------------------
Function MailDeliveryEnsureComposer()
  On Error Resume Next
  Dim existing
  Set existing = GetRef("ComposeAndSendMail")
  If Err.Number = 0 Then
    MailDeliveryEnsureComposer = True
    Set existing = Nothing
    On Error GoTo 0
    Exit Function
  End If
  Err.Clear
  On Error GoTo 0

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = MailDeliveryScriptFolder(fso)
  Dim parentFolder: parentFolder = fso.GetParentFolderName(scriptFolder)
  Dim grandParent: grandParent = ""
  If Len(parentFolder) > 0 Then grandParent = fso.GetParentFolderName(parentFolder)

  Dim candidates(2)
  candidates(0) = fso.BuildPath(scriptFolder, "utilities\mail_composer.vbs")
  If Len(parentFolder) > 0 Then
    candidates(1) = fso.BuildPath(parentFolder, "utilities\mail_composer.vbs")
  Else
    candidates(1) = ""
  End If
  If Len(grandParent) > 0 Then
    candidates(2) = fso.BuildPath(grandParent, "utilities\mail_composer.vbs")
  Else
    candidates(2) = ""
  End If

  Dim i, path, ts, contents
  For i = 0 To UBound(candidates)
    path = candidates(i)
    If Len(path) > 0 Then
      If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1)
        contents = ts.ReadAll
        ts.Close
        ExecuteGlobal contents
        MailDeliveryEnsureComposer = True
        Exit Function
      End If
    End If
  Next

  MailDeliveryEnsureComposer = False
End Function

'--------------------------------------------------------------------
' Signature helpers
'--------------------------------------------------------------------
Function MailDeliveryLoadSignatureAutoCharset(olMail)
  Dim fso, shell, sigPath, folder, file, sigHTML, sigFile, baseName, imgFolder, imgFile, att, imgName, cid
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set shell = CreateObject("WScript.Shell")
  sigPath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Signatures\"
  sigHTML = ""

  If fso.FolderExists(sigPath) Then
    Set folder = fso.GetFolder(sigPath)
    Dim found: found = False
    For Each file In folder.Files
      If LCase(fso.GetExtensionName(file.Name)) = "htm" Then
        sigFile = file.Path
        baseName = fso.GetBaseName(file.Name)
        sigHTML = MailDeliveryReadHtmlAuto(sigFile)
        found = True
        Exit For
      End If
    Next
    If Not found Then MailDeliveryLoadSignatureAutoCharset = "" : Exit Function
  Else
    MailDeliveryLoadSignatureAutoCharset = ""
    Exit Function
  End If

  imgFolder = sigPath & baseName & ".files"
  If fso.FolderExists(imgFolder) Then
    For Each imgFile In fso.GetFolder(imgFolder).Files
      imgName = fso.GetFileName(imgFile.Name)
      cid = MailDeliverySanitizeContentId(imgName)
      Set att = olMail.Attachments.Add(imgFile.Path)
      att.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, cid
      sigHTML = Replace(sigHTML, baseName & ".files/" & imgName, "cid:" & cid)
      sigHTML = Replace(sigHTML, baseName & ".files\" & imgName, "cid:" & cid)
    Next
  End If

  MailDeliveryLoadSignatureAutoCharset = sigHTML
End Function

Function MailDeliveryReadHtmlAuto(filePath)
  On Error Resume Next
  Dim cs: cs = MailDeliveryDetectHtmlCharset(filePath)
  Dim s: Set s = CreateObject("ADODB.Stream")
  s.Type = 2
  s.Charset = cs
  s.Open
  s.LoadFromFile filePath
  Dim txt: txt = s.ReadText
  s.Close

  If cs = "windows-1252" Then
    If (InStr(txt, "Ã") > 0) Or (InStr(txt, "â€”") > 0) Or (InStr(txt, "â€œ") > 0) Or (InStr(txt, "â€™") > 0) Then
      Dim s2: Set s2 = CreateObject("ADODB.Stream")
      s2.Type = 2
      s2.Charset = "utf-8"
      s2.Open
      s2.LoadFromFile filePath
      Dim t2: t2 = s2.ReadText
      s2.Close
      If MailDeliveryMarkerCount(t2) < MailDeliveryMarkerCount(txt) Then txt = t2
    End If
  End If
  On Error GoTo 0

  MailDeliveryReadHtmlAuto = txt
End Function

Function MailDeliveryDetectHtmlCharset(filePath)
  On Error Resume Next
  Dim s: Set s = CreateObject("ADODB.Stream")
  s.Type = 2
  s.Charset = "windows-1252"
  s.Open
  s.LoadFromFile filePath
  Dim header: header = LCase(Left(s.ReadText, 1024))
  s.Close
  Dim charset: charset = ""

  Dim re: Set re = CreateObject("VBScript.RegExp")
  re.Global = False
  re.IgnoreCase = True
  re.Pattern = "charset\s*=\s*([a-z0-9\-\_]+)"
  Dim m: Set m = re.Execute(header)
  If m.Count > 0 Then charset = LCase(m(0).SubMatches(0))

  If charset = "" Then
    MailDeliveryDetectHtmlCharset = "windows-1252"
  ElseIf charset = "utf8" Then
    MailDeliveryDetectHtmlCharset = "utf-8"
  Else
    MailDeliveryDetectHtmlCharset = charset
  End If
End Function

Function MailDeliveryMarkerCount(s)
  Dim c: c = 0
  c = c + MailDeliveryCountOf(s, "Ã")
  c = c + MailDeliveryCountOf(s, "Â")
  c = c + MailDeliveryCountOf(s, "â€”") + MailDeliveryCountOf(s, "â€“") + MailDeliveryCountOf(s, "â€œ") + MailDeliveryCountOf(s, "â€�") + MailDeliveryCountOf(s, "â€˜") + MailDeliveryCountOf(s, "â€™") + MailDeliveryCountOf(s, "â€¦")
  MailDeliveryMarkerCount = c
End Function

Function MailDeliveryCountOf(s, token)
  Dim p, n: p = 1: n = 0
  Do
    p = InStr(p, s, token)
    If p = 0 Then Exit Do
    n = n + 1
    p = p + Len(token)
  Loop
  MailDeliveryCountOf = n
End Function

Function MailDeliverySanitizeInlineCid(fileName)
  MailDeliverySanitizeInlineCid = MailDeliverySanitizeContentId(fileName)
End Function

'--------------------------------------------------------------------
' Generic helpers
'--------------------------------------------------------------------
Function MailDeliveryGetOption(config, key, defaultValue)
  On Error Resume Next
  If Not config Is Nothing Then
    If config.Exists(key) Then
      MailDeliveryGetOption = config(key)
      On Error GoTo 0
      Exit Function
    End If
  End If
  On Error GoTo 0
  MailDeliveryGetOption = defaultValue
End Function

Sub MailDeliveryCopyIfExists(source, target, key)
  On Error Resume Next
  If source Is Nothing Then Exit Sub
  If source.Exists(key) Then
    If IsObject(source(key)) Then
      Set target(key) = source(key)
    Else
      target(key) = source(key)
    End If
  End If
  On Error GoTo 0
End Sub
