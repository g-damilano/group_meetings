'====================================================================
' mail_utils_mailer_dynamiclinks.vbs — Outlook Mailer with dynamic hyperlink substitution
'====================================================================
Option Explicit
SetLocale 1033

Function CreateMail(toAddr, ccAddr, bccAddr, subjectText, htmlBody, attachPaths, _
                    action, savePath, templateSource, inlineImgFolder, deferMinutes, logPath)
  On Error Resume Next
  If Len(Trim(logPath)) = 0 Then logPath = DefaultLogPath()
  LogAction logPath, "Mailer (dynamic links) active."

  If Not EnsureMailComposerLoaded() Then
    LogAction logPath, "ERROR: Shared mail composer module could not be loaded."
    CreateMail = False
    Exit Function
  End If

  Dim fso, scriptFolder, bodyHtml, mapPath, options
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."

  If Len(Trim(templateSource)) > 0 Then
    bodyHtml = LoadHtmlFromSource(templateSource)
  Else
    bodyHtml = htmlBody
  End If

  mapPath = fso.BuildPath(scriptFolder, "links_map.txt")
  bodyHtml = ApplyLinkMap(bodyHtml, mapPath)

  Set options = CreateObject("Scripting.Dictionary")
  options("To") = toAddr
  options("Cc") = ccAddr
  options("Bcc") = bccAddr
  options("Subject") = subjectText
  options("BodyHtml") = bodyHtml
  options("InlineImageFolder") = inlineImgFolder
  options("Action") = action
  options("SavePath") = savePath
  options("DeferMinutes") = deferMinutes
  options("LogPath") = logPath
  Set options("LogHandler") = GetRef("LogAction")
  Set options("SignatureProvider") = GetRef("LoadSignatureWithImages_AutoCharset")
  Set options("InlineImageCidTransform") = GetRef("SanitizeInlineCid")
  If Not IsEmpty(attachPaths) Then options("Attachments") = attachPaths

  Dim ok: ok = ComposeAndSendMail(options)

  If Not ok Then LogAction logPath, "CreateMail reported failure for subject: " & subjectText
  CreateMail = ok
  On Error GoTo 0
End Function

Function SanitizeInlineCid(fileName)
  SanitizeInlineCid = Replace(fileName, " ", "_")
End Function

Function EnsureMailComposerLoaded()
  On Error Resume Next
  Dim existing
  Set existing = GetRef("ComposeAndSendMail")
  If Err.Number = 0 Then
    EnsureMailComposerLoaded = True
    Set existing = Nothing
    On Error GoTo 0
    Exit Function
  End If
  Err.Clear
  On Error GoTo 0

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
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
        EnsureMailComposerLoaded = True
        Exit Function
      End If
    End If
  Next

  EnsureMailComposerLoaded = False
End Function

Function LoadSignatureWithImages_AutoCharset(olMail)
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
        sigHTML = ReadHtmlAuto(sigFile)
        found = True
        Exit For
      End If
    Next
    If Not found Then LoadSignatureWithImages_AutoCharset = "" : Exit Function
  Else
    LoadSignatureWithImages_AutoCharset = ""
    Exit Function
  End If

  imgFolder = sigPath & baseName & ".files"
  If fso.FolderExists(imgFolder) Then
    For Each imgFile In fso.GetFolder(imgFolder).Files
      imgName = fso.GetFileName(imgFile.Name)
      cid = Replace(imgName, " ", "_")
      Set att = olMail.Attachments.Add(imgFile.Path)
      att.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, cid
      sigHTML = Replace(sigHTML, baseName & ".files/" & imgName, "cid:" & cid)
      sigHTML = Replace(sigHTML, baseName & ".files\" & imgName, "cid:" & cid)
    Next
  End If

  LoadSignatureWithImages_AutoCharset = sigHTML
End Function

Function DetectHtmlCharset(filePath)
  On Error Resume Next
  Dim s : Set s = CreateObject("ADODB.Stream")
  s.Type = 2 : s.Charset = "windows-1252" : s.Open : s.LoadFromFile filePath
  Dim header : header = LCase(Left(s.ReadText, 1024))
  s.Close
  Dim charset : charset = ""

  Dim re: Set re = CreateObject("VBScript.RegExp")
  re.Global = False : re.IgnoreCase = True
  re.Pattern = "charset\s*=\s*([a-z0-9\-\_]+)"
  Dim m: Set m = re.Execute(header)
  If m.Count > 0 Then
    charset = LCase(m(0).SubMatches(0))
  End If

  If charset = "" Then
    DetectHtmlCharset = "windows-1252"
  ElseIf charset = "utf8" Then
    DetectHtmlCharset = "utf-8"
  Else
    DetectHtmlCharset = charset
  End If
End Function

Function ReadHtmlAuto(filePath)
  On Error Resume Next
  Dim cs: cs = DetectHtmlCharset(filePath)
  Dim s : Set s = CreateObject("ADODB.Stream")
  s.Type = 2 : s.Charset = cs : s.Open : s.LoadFromFile filePath
  Dim txt : txt = s.ReadText : s.Close

  If cs = "windows-1252" Then
    If (InStr(txt, "Ã")>0) Or (InStr(txt, "â€”")>0) Or (InStr(txt, "â€œ")>0) Or (InStr(txt, "â€™")>0) Then
      Dim s2: Set s2 = CreateObject("ADODB.Stream")
      s2.Type = 2 : s2.Charset = "utf-8" : s2.Open : s2.LoadFromFile filePath
      Dim t2: t2 = s2.ReadText : s2.Close
      If MarkerCount(t2) < MarkerCount(txt) Then txt = t2
    End If
  End If
  ReadHtmlAuto = txt
End Function

Function MarkerCount(s)
  Dim c: c = 0
  c = c + CountOf(s, "Ã")
  c = c + CountOf(s, "Â")
  c = c + CountOf(s, "â€”") + CountOf(s, "â€“") + CountOf(s, "â€œ") + CountOf(s, "â€�") + CountOf(s, "â€˜") + CountOf(s, "â€™") + CountOf(s, "â€¦")
  MarkerCount = c
End Function

Function CountOf(s, token)
  Dim p, n: p = 1: n = 0
  Do
    p = InStr(p, s, token)
    If p = 0 Then Exit Do
    n = n + 1
    p = p + Len(token)
  Loop
  CountOf = n
End Function

Function DefaultLogPath()
  Dim fso, scriptFolder
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
  DefaultLogPath = scriptFolder & "\mail_util.log"
End Function

Sub LogAction(logPath, message)
  On Error Resume Next
  Dim fso, f: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim ts: Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & message
  f.Close
End Sub
