'====================================================================
' mail_utils_mailer_dynamiclinks.vbs — Outlook Mailer with dynamic hyperlink substitution
'====================================================================
Option Explicit
SetLocale 1033

Function CreateMail(toAddr, ccAddr, bccAddr, subjectText, htmlBody, attachPaths, _
                    action, savePath, templateSource, inlineImgFolder, deferMinutes, logPath)
  Dim fallbackLogPath: fallbackLogPath = MailExtResolveLogPath(logPath)

  If Not EnsureMailDeliveryCoreLoaded() Then
    MailExtFallbackLog fallbackLogPath, "ERROR: mail_delivery_core.vbs could not be loaded."
    CreateMail = False
    Exit Function
  End If

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim config: Set config = CreateObject("Scripting.Dictionary")

  config("To") = toAddr
  config("Cc") = ccAddr
  config("Bcc") = bccAddr
  config("Subject") = subjectText
  config("BodyHtml") = htmlBody
  config("Action") = action
  config("SavePath") = savePath
  config("DeferMinutes") = deferMinutes
  config("LogPath") = logPath
  Set config("SignatureProvider") = GetRef("MailDeliveryLoadSignatureAutoCharset")
  Set config("InlineImageCidTransform") = GetRef("MailDeliverySanitizeInlineCid")

  If Len(Trim(templateSource)) > 0 Then config("TemplateSource") = templateSource
  If Len(Trim(inlineImgFolder)) > 0 Then config("InlineImageFolder") = inlineImgFolder
  If Not IsEmpty(attachPaths) Then config("Attachments") = attachPaths

  Dim scriptFolder: scriptFolder = MailDeliveryScriptFolder(fso)
  Dim linkMapPath: linkMapPath = fso.BuildPath(scriptFolder, "links_map.txt")
  If fso.FileExists(linkMapPath) Then config("LinkMapPath") = linkMapPath

  Dim resolvedLogPath: resolvedLogPath = MailDeliveryResolveLogPath(config, fso)
  config("LogPath") = resolvedLogPath

  MailDeliveryLog resolvedLogPath, "Mailer (dynamic links) active."

  Dim ok: ok = MailDeliveryExecute(config)
  CreateMail = ok
End Function

Function EnsureMailDeliveryCoreLoaded()
  On Error Resume Next
  Dim existing
  Set existing = GetRef("MailDeliveryExecute")
  If Err.Number = 0 Then
    EnsureMailDeliveryCoreLoaded = True
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
  candidates(0) = fso.BuildPath(scriptFolder, "utilities\mail_delivery_core.vbs")
  If Len(parentFolder) > 0 Then
    candidates(1) = fso.BuildPath(parentFolder, "utilities\mail_delivery_core.vbs")
  Else
    candidates(1) = ""
  End If
  If Len(grandParent) > 0 Then
    candidates(2) = fso.BuildPath(grandParent, "utilities\mail_delivery_core.vbs")
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
        EnsureMailDeliveryCoreLoaded = True
        Exit Function
      End If
    End If
  Next

  EnsureMailDeliveryCoreLoaded = False
End Function

Function MailExtResolveLogPath(candidate)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  If Len(Trim(candidate)) = 0 Then
    MailExtResolveLogPath = fso.BuildPath(folder, "mail_util.log")
  Else
    MailExtResolveLogPath = candidate
  End If
End Function

Sub MailExtFallbackLog(logPath, message)
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim stream: Set stream = fso.OpenTextFile(logPath, 8, True)
  stream.WriteLine Now & "  " & message
  stream.Close
End Sub
