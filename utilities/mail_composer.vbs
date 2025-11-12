'====================================================================
' mail_composer.vbs — Shared Outlook mail composition routines
'====================================================================
Option Explicit
SetLocale 1033

Const olMailItem = 0
Const olFormatHTML = 2
Const olMSG = 3
Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

Function ComposeAndSendMail(mailOptions)
  Dim logger, logPath, action, savePath, deferMinutes
  Dim toAddr, ccAddr, bccAddr, subjectText, bodyHtml, headExtra, bodyAttributes
  Dim attachments, inlineImgFolder, inlineCidTransform
  Dim signatureLiteral
  Dim success, olApp, olMail, fso
  Dim signatureProvider

  Set logger = Nothing
  logPath = ComposerGetOption(mailOptions, "LogPath", "")
  action = LCase(Trim(CStr(ComposerGetOption(mailOptions, "Action", "display"))))
  savePath = ComposerGetOption(mailOptions, "SavePath", "")
  deferMinutes = ComposerGetOption(mailOptions, "DeferMinutes", 0)
  toAddr = ComposerGetOption(mailOptions, "To", "")
  ccAddr = ComposerGetOption(mailOptions, "Cc", "")
  bccAddr = ComposerGetOption(mailOptions, "Bcc", "")
  subjectText = ComposerGetOption(mailOptions, "Subject", "")
  bodyHtml = ComposerGetOption(mailOptions, "BodyHtml", "")
  headExtra = ComposerGetOption(mailOptions, "HeadContent", "")
  bodyAttributes = ComposerGetOption(mailOptions, "BodyAttributes", "")
  attachments = ComposerGetOption(mailOptions, "Attachments", Empty)
  inlineImgFolder = ComposerGetOption(mailOptions, "InlineImageFolder", "")
  Set inlineCidTransform = Nothing
  Set signatureProvider = Nothing
  signatureLiteral = ""

  If Not mailOptions Is Nothing Then
    On Error Resume Next
    If mailOptions.Exists("LogHandler") Then Set logger = mailOptions("LogHandler")
    If mailOptions.Exists("InlineImageCidTransform") Then Set inlineCidTransform = mailOptions("InlineImageCidTransform")
    If mailOptions.Exists("SignatureProvider") Then
      If IsObject(mailOptions("SignatureProvider")) Then
        Set signatureProvider = mailOptions("SignatureProvider")
      Else
        signatureLiteral = CStr(mailOptions("SignatureProvider"))
      End If
    End If
    On Error GoTo 0
  End If

  Set fso = CreateObject("Scripting.FileSystemObject")

  On Error Resume Next
  Set olApp = GetOutlookApp()
  If Err.Number <> 0 Then
    ComposerLog logger, logPath, "ERROR: Unable to obtain Outlook application — " & Err.Description
    Err.Clear
    ComposeAndSendMail = False
    On Error GoTo 0
    Exit Function
  End If

  Set olMail = olApp.CreateItem(olMailItem)

  bodyHtml = ComposerApplyBodyTransforms(bodyHtml, mailOptions)

  Dim signatureHtml
  signatureHtml = ComposerResolveSignatureHtml(signatureProvider, signatureLiteral, olMail)

  Dim finalHtml
  finalHtml = ComposerBuildHtmlDocument(bodyHtml, signatureHtml, headExtra, bodyAttributes, mailOptions)

  With olMail
    .To = toAddr
    .CC = ccAddr
    .BCC = bccAddr
    .Subject = subjectText
    .BodyFormat = olFormatHTML
    .HTMLBody = finalHtml
  End With

  ComposerAddAttachments olMail, attachments, fso
  ComposerAddInlineImages olMail, inlineImgFolder, inlineCidTransform, fso

  If IsNumeric(deferMinutes) Then
    If CDbl(deferMinutes) > 0 Then
      olMail.DeferredDeliveryTime = Now + (CDbl(deferMinutes) / (24 * 60))
    End If
  End If

  Select Case action
    Case "send"
      olMail.Send
      success = (Err.Number = 0)
      If success Then
        ComposerLog logger, logPath, "Sent: " & subjectText
      End If
    Case "save"
      If Len(Trim(savePath)) > 0 Then
        olMail.SaveAs savePath, olMSG
        success = (Err.Number = 0)
        If success Then ComposerLog logger, logPath, "Saved as: " & savePath
      Else
        olMail.Save
        success = (Err.Number = 0)
        If success Then ComposerLog logger, logPath, "Saved to Drafts: " & subjectText
      End If
    Case Else
      olMail.Display
      success = (Err.Number = 0)
      If success Then ComposerLog logger, logPath, "Displayed: " & subjectText
  End Select

  If Err.Number <> 0 Then
    ComposerLog logger, logPath, "ERROR: " & Err.Number & " - " & Err.Description
    Err.Clear
    success = False
  End If

  ComposeAndSendMail = success
  On Error GoTo 0
End Function

Function ComposerApplyBodyTransforms(bodyHtml, mailOptions)
  Dim transformed
  transformed = bodyHtml

  If mailOptions Is Nothing Then
    ComposerApplyBodyTransforms = transformed
    Exit Function
  End If

  On Error Resume Next
  If mailOptions.Exists("BodyTransforms") Then
    Dim transforms, t
    transforms = mailOptions("BodyTransforms")
    If IsArray(transforms) Then
      For Each t In transforms
        If IsObject(t) Then transformed = t(transformed)
      Next
    ElseIf IsObject(transforms) Then
      transformed = transforms(transformed)
    End If
  End If
  On Error GoTo 0

  ComposerApplyBodyTransforms = transformed
End Function

Function ComposerResolveSignatureHtml(signatureProvider, signatureLiteral, olMail)
  Dim signatureHtml
  signatureHtml = ""

  If Not signatureProvider Is Nothing Then
    On Error Resume Next
    signatureHtml = signatureProvider(olMail)
    If Err.Number <> 0 Then
      signatureHtml = ""
      Err.Clear
    End If
    On Error GoTo 0
  ElseIf Len(signatureLiteral) > 0 Then
    signatureHtml = signatureLiteral
  End If

  ComposerResolveSignatureHtml = signatureHtml
End Function

Function ComposerBuildHtmlDocument(bodyHtml, signatureHtml, headExtra, bodyAttributes, mailOptions)
  Dim headBlock, bodyOpen
  headBlock = "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
  If Len(Trim(headExtra)) > 0 Then headBlock = headBlock & headExtra

  If Len(Trim(bodyAttributes)) = 0 Then
    bodyOpen = "<body style='font-family:Segoe UI, Arial, sans-serif;'>"
  Else
    bodyOpen = "<body " & bodyAttributes & ">"
  End If

  ComposerBuildHtmlDocument = "<html><head>" & headBlock & "</head>" & _
                              bodyOpen & bodyHtml & signatureHtml & "</body></html>"
End Function

Sub ComposerAddAttachments(olMail, attachments, fso)
  If IsEmpty(attachments) Then Exit Sub
  If IsArray(attachments) Then
    Dim i
    For i = LBound(attachments) To UBound(attachments)
      ComposerAddAttachmentPath olMail, attachments(i), fso
    Next
  Else
    Dim typeName
    typeName = TypeName(attachments)
    Select Case typeName
      Case "Collection"
        Dim item
        For Each item In attachments
          ComposerAddAttachmentPath olMail, item, fso
        Next
      Case "Dictionary"
        Dim key
        For Each key In attachments.Keys
          ComposerAddAttachmentPath olMail, attachments(key), fso
        Next
      Case Else
        ComposerAddAttachmentPath olMail, attachments, fso
    End Select
  End If
End Sub

Sub ComposerAddAttachmentPath(olMail, candidate, fso)
  Dim path
  path = ""
  If IsObject(candidate) Then Exit Sub
  If IsEmpty(candidate) Then Exit Sub
  path = Trim(CStr(candidate))
  If Len(path) = 0 Then Exit Sub
  If Not fso.FileExists(path) Then Exit Sub

  On Error Resume Next
  olMail.Attachments.Add path
  On Error GoTo 0
End Sub

Sub ComposerAddInlineImages(olMail, folderPath, inlineCidTransform, fso)
  If Len(Trim(folderPath)) = 0 Then Exit Sub
  If Not fso.FolderExists(folderPath) Then Exit Sub

  Dim folder, file, cid, att, fileName
  Set folder = fso.GetFolder(folderPath)
  For Each file In folder.Files
    fileName = fso.GetFileName(file.Name)
    cid = fileName
    If Not inlineCidTransform Is Nothing Then
      On Error Resume Next
      cid = inlineCidTransform(fileName)
      If Err.Number <> 0 Then
        cid = fileName
        Err.Clear
      End If
      On Error GoTo 0
    End If
    On Error Resume Next
    Set att = olMail.Attachments.Add(file.Path)
    If Err.Number = 0 Then att.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, cid
    Err.Clear
    On Error GoTo 0
  Next
End Sub

Function ApplyLinkMap(htmlBody, mapPath)
  Dim fso, ts, line, pos, key, value
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(mapPath) Then
    ApplyLinkMap = htmlBody
    Exit Function
  End If

  Set ts = fso.OpenTextFile(mapPath, 1)
  Dim mapDict: Set mapDict = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  mapDict.CompareMode = 1 ' TextCompare when available
  On Error GoTo 0

  Do Until ts.AtEndOfStream
    line = Trim(ts.ReadLine)
    If Len(line) > 0 And Left(line, 1) <> "#" Then
      pos = InStr(line, "=")
      If pos > 0 Then
        key = Trim(Left(line, pos - 1))
        value = Trim(Mid(line, pos + 1))
        If Len(key) > 0 Then mapDict(key) = value
      End If
    End If
  Loop
  ts.Close

  For Each key In mapDict.Keys
    value = mapDict(key)
    htmlBody = Replace(htmlBody, "{{" & key & "}}", value)
    htmlBody = Replace(htmlBody, "{{" & UCase(key) & "}}", value)
    htmlBody = Replace(htmlBody, "[" & key & "]", value)
    htmlBody = Replace(htmlBody, "[" & UCase(key) & "]", value)
  Next

  ApplyLinkMap = htmlBody
End Function

Function LoadHtmlFromSource(src)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(src) Then
    LoadHtmlFromSource = ReadUTF8File(src)
  Else
    LoadHtmlFromSource = src
  End If
End Function

Function ReadUTF8File(filePath)
  Dim stream: Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile filePath
  ReadUTF8File = stream.ReadText
  stream.Close
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

Sub ComposerLog(logger, logPath, message)
  On Error Resume Next
  If Not logger Is Nothing Then
    logger logPath, message
  End If
  On Error GoTo 0
End Sub

Function ComposerGetOption(options, key, defaultValue)
  On Error Resume Next
  If Not options Is Nothing Then
    If options.Exists(key) Then
      ComposerGetOption = options(key)
      On Error GoTo 0
      Exit Function
    End If
  End If
  On Error GoTo 0
  ComposerGetOption = defaultValue
End Function
