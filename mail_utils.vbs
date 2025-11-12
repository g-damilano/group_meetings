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
  If Not IsEmpty(attachPaths) Then options("Attachments") = attachPaths
  Set options("LogHandler") = GetRef("LogAction")

  Dim ok: ok = ComposeAndSendMail(options)

  If Not ok Then LogAction logPath, "CreateMail reported failure for subject: " & subjectText
  CreateMail = ok
  On Error GoTo 0
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
