'====================================================================
' mail_utils_mailer_dynamiclinks.vbs — Outlook Mailer with dynamic hyperlink substitution
'====================================================================
Option Explicit
SetLocale 1033

Const olMailItem = 0
Const olFormatHTML = 2
Const olMSG = 3
Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

Call IncludeCommonOutlook()

Sub IncludeCommonOutlook()
  Dim loaderFso: Set loaderFso = CreateObject("Scripting.FileSystemObject")
  Dim baseFolder: baseFolder = loaderFso.GetParentFolderName(WScript.ScriptFullName)
  If Len(baseFolder) = 0 Then baseFolder = "."

  Dim candidates()
  ReDim candidates(1)
  candidates(0) = loaderFso.BuildPath(baseFolder, "common_outlook.vbs")
  candidates(1) = loaderFso.BuildPath(baseFolder, "utilities\common_outlook.vbs")

  Dim parentFolder: parentFolder = loaderFso.GetParentFolderName(baseFolder)
  If Len(parentFolder) > 0 Then
    ReDim Preserve candidates(UBound(candidates) + 1)
    candidates(UBound(candidates)) = loaderFso.BuildPath(parentFolder, "utilities\common_outlook.vbs")
  End If

  Dim grandParent: grandParent = loaderFso.GetParentFolderName(parentFolder)
  If Len(grandParent) > 0 Then
    ReDim Preserve candidates(UBound(candidates) + 1)
    candidates(UBound(candidates)) = loaderFso.BuildPath(grandParent, "utilities\common_outlook.vbs")
  End If

  Dim i, candidatePath, loaded: loaded = False
  For i = LBound(candidates) To UBound(candidates)
    candidatePath = candidates(i)
    If Len(candidatePath) > 0 Then
      If loaderFso.FileExists(candidatePath) Then
        Dim ts: Set ts = loaderFso.OpenTextFile(candidatePath, 1)
        Dim contents: contents = ts.ReadAll
        ts.Close
        ExecuteGlobal contents
        loaded = True
        Exit For
      End If
    End If
  Next

  If Not loaded Then
    WScript.Echo "Fatal: unable to load common Outlook helpers."
    WScript.Quit 1
  End If
End Sub

Function CreateMail(toAddr, ccAddr, bccAddr, subjectText, htmlBody, attachPaths, _
                    action, savePath, templateSource, inlineImgFolder, deferMinutes, logPath)
  On Error Resume Next
  If Len(Trim(logPath)) = 0 Then logPath = DefaultLogPath()
  LogAction logPath, "Mailer (dynamic links) active."

  Dim ok, olApp, olMail, sigHTML, f, fso, bodyHtml, htmlWrapped, folder
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set olApp = GetOutlookApp()
  folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."

  If Len(Trim(templateSource)) > 0 Then
    bodyHtml = LoadHtmlFromSource(templateSource)
  Else
    bodyHtml = htmlBody
  End If

  ' Apply dynamic link mapping (links_map.txt)
  bodyHtml = ApplyLinkMap(bodyHtml, folder & "\links_map.txt")

  htmlWrapped = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>" & _
                "<body style='font-family:Segoe UI, Arial, sans-serif;'>" & bodyHtml

  Set olMail = olApp.CreateItem(olMailItem)
  sigHTML = LoadSignatureWithImages_AutoCharset(olMail)

  With olMail
    .To = toAddr
    .CC = ccAddr
    .BCC = bccAddr
    .Subject = subjectText
    .BodyFormat = olFormatHTML
    .HTMLBody = htmlWrapped & sigHTML & "</body></html>"

    If Not IsEmpty(attachPaths) Then
      For Each f In attachPaths
        If Len(Trim(f)) > 0 And fso.FileExists(f) Then .Attachments.Add f
      Next
    End If

    If Len(Trim(inlineImgFolder)) > 0 And fso.FolderExists(inlineImgFolder) Then
      EmbedImagesFromFolder olMail, inlineImgFolder
    End If

    If IsNumeric(deferMinutes) And CDbl(deferMinutes) > 0 Then
      .DeferredDeliveryTime = Now + (CDbl(deferMinutes) / (24 * 60))
    End If

    Select Case LCase(action)
      Case "send"   : .Send : LogAction logPath, "Sent: " & subjectText
      Case "save"
        If Len(Trim(savePath)) > 0 Then
          .SaveAs savePath, olMSG : LogAction logPath, "Saved as: " & savePath
        Else
          .Save : LogAction logPath, "Saved to Drafts: " & subjectText
        End If
      Case Else     : .Display : LogAction logPath, "Displayed: " & subjectText
    End Select
  End With

  If Err.Number <> 0 Then
    LogAction logPath, "ERROR: " & Err.Number & " - " & Err.Description
    ok = False
  Else
    ok = True
  End If

  CreateMail = ok
  On Error GoTo 0
End Function

'================== Dynamic Link Mapping ==================
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
  mapDict.CompareMode = 1
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
  Dim tokens, i, total
  tokens = Array("Ã", "Â", "â€”", "â€“", "â€œ", "â€�", "â€˜", "â€™", "â€¦")
  total = 0
  For i = 0 To UBound(tokens)
    total = total + CountOf(s, tokens(i))
  Next
  MarkerCount = total
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

Function ReadUTF8File(filePath)
  Dim s : Set s = CreateObject("ADODB.Stream")
  s.Type = 2 : s.Charset = "utf-8" : s.Open : s.LoadFromFile filePath
  ReadUTF8File = s.ReadText : s.Close
End Function

