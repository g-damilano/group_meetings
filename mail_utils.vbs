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
  sigHTML = "" ' skip signature for simplicity

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
  Dim fso, ts, line, parts
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(mapPath) Then
    ApplyLinkMap = htmlBody
    Exit Function
  End If

  Set ts = fso.OpenTextFile(mapPath, 1)
  Dim mapDict: Set mapDict = CreateObject("Scripting.Dictionary")

  Do Until ts.AtEndOfStream
    line = Trim(ts.ReadLine)
    If Len(line) > 0 And InStr(line, "=") > 0 Then
      parts = Split(line, "=")
      If UBound(parts) = 1 Then mapDict(parts(0)) = parts(1)
    End If
  Loop
  ts.Close

  Dim key
  For Each key In mapDict.Keys
    htmlBody = Replace(htmlBody, "{{" & key & "}}", mapDict(key))
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
  Dim s : Set s = CreateObject("ADODB.Stream")
  s.Type = 2 : s.Charset = "utf-8" : s.Open : s.LoadFromFile filePath
  ReadUTF8File = s.ReadText : s.Close
End Function

