'====================================================================
' mail_utils_mailer_fixsig.vbs — Outlook Mailer (no shield) + Signature charset auto-detect
'====================================================================
Option Explicit
SetLocale 1033

Const olMailItem = 0
Const olFormatHTML = 2
Const olMSG = 3
Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

Function CreateMail(toAddr, ccAddr, bccAddr, subjectText, htmlBody, attachPaths, _
                    action, savePath, templateSource, inlineImgFolder, deferMinutes, logPath)
  On Error Resume Next
  If Len(Trim(logPath)) = 0 Then logPath = DefaultLogPath()
  LogAction logPath, "Mailer (fixsig) active."

  Dim ok, olApp, olMail, sigHTML, f, fso, bodyHtml, htmlWrapped
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set olApp = GetOutlookApp()

  If Len(Trim(templateSource)) > 0 Then
    bodyHtml = LoadHtmlFromSource(templateSource) ' assumes UTF-8 template
  Else
    bodyHtml = htmlBody
  End If

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

Sub EmbedImagesFromFolder(olMail, folderPath)
  Dim fso, file, cid, att
  Set fso = CreateObject("Scripting.FileSystemObject")
  For Each file In fso.GetFolder(folderPath).Files
    cid = fso.GetFileName(file.Name)
    Set att = olMail.Attachments.Add(file.Path)
    att.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, cid
  Next
End Sub

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

' ================= Signature loader with charset auto-detect =================
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
        sigHTML = ReadHtmlAuto(sigFile) ' << auto-detect charset here
        found = True
        Exit For
      End If
    Next
    If Not found Then Exit Function
  Else
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

' Detect HTML charset from meta tag; default to windows-1252 for Outlook signatures
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
    ' Heuristic: many Outlook signatures are ANSI (1252)
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

  ' Fallback heuristic: if mojibake markers present and we used 1252, try UTF-8
  If cs = "windows-1252" Then
    If (InStr(txt, "Ã")>0) Or (InStr(txt, "â€”")>0) Or (InStr(txt, "â€œ")>0) Or (InStr(txt, "â€™")>0) Then
      Dim s2: Set s2 = CreateObject("ADODB.Stream")
      s2.Type = 2 : s2.Charset = "utf-8" : s2.Open : s2.LoadFromFile filePath
      Dim t2: t2 = s2.ReadText : s2.Close
      ' choose whichever has fewer markers
      If MarkerCount(t2) < MarkerCount(txt) Then txt = t2
    End If
  End If
  ReadHtmlAuto = txt
End Function

' Marker counter for heuristic
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

' ================= Template loader (UTF-8) =================
Function LoadHtmlFromSource(src)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If Left(LCase(src), 7) = "http://" Or Left(LCase(src), 8) = "https://" Then
    LoadHtmlFromSource = HttpGetUtf8(src)
  ElseIf fso.FileExists(src) Then
    LoadHtmlFromSource = ReadUTF8File(src)
  Else
    LoadHtmlFromSource = src
  End If
End Function

Function ReadUTF8File(filePath)
  Dim s : Set s = CreateObject("ADODB.Stream")
  s.Type = 2 : s.Charset = "utf-8" : s.Open : s.LoadFromFile filePath
  ReadUTF8File = s.ReadText : s.Close
End Function

Function HttpGetUtf8(url)
  Dim xhr, stm, txt
  Set xhr = CreateObject("MSXML2.XMLHTTP")
  xhr.Open "GET", url, False
  xhr.Send
  If xhr.Status = 200 Then
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.Write xhr.responseBody
    stm.Position = 0 : stm.Type = 2 : stm.Charset = "utf-8"
    txt = stm.ReadText : stm.Close
    HttpGetUtf8 = txt
  Else
    HttpGetUtf8 = "<p>Failed to load template. HTTP " & xhr.Status & "</p>"
  End If
End Function

' ================= Logging =================
Function DefaultLogPath()
  Dim fso, scriptFolder
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
  DefaultLogPath = scriptFolder & "\mail_util.log"
End Function

Sub LogAction(logPath, message)
  On Error Resume Next
  If Len(Trim(logPath)) = 0 Then logPath = DefaultLogPath()
  Dim fso, f: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim ts: Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & message
  f.Close
End Sub
