' Attachment and inline-image helpers for scaffold mail composition.
' These helpers validate and prepare paths without sending mail. Outlook-specific attachment
' insertion is isolated in SC_AddAttachmentsToMail and SC_AddInlineImagesToMail.

Const SC_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

Sub SC_LogAssetWarning(logPath, message)
  If Len(Trim(CStr(logPath))) > 0 Then SC_LogWarning logPath, message
End Sub

Function SC_IsPathLike(value)
  Dim text
  text = Trim(CStr(value))
  SC_IsPathLike = (Len(text) > 0)
End Function

Function SC_ResolveExistingFilePath(candidate, baseFolder)
  SC_ResolveExistingFilePath = ""
  If Not SC_IsPathLike(candidate) Then Exit Function

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim path
  path = Trim(CStr(candidate))

  If fso.FileExists(path) Then
    SC_ResolveExistingFilePath = fso.GetAbsolutePathName(path)
    Exit Function
  End If

  If Len(Trim(CStr(baseFolder))) > 0 Then
    Dim combined
    combined = fso.BuildPath(CStr(baseFolder), path)
    If fso.FileExists(combined) Then SC_ResolveExistingFilePath = fso.GetAbsolutePathName(combined)
  End If
End Function

Sub SC_AddResolvedPath(paths, filePath)
  If Len(Trim(CStr(filePath))) = 0 Then Exit Sub

  Dim key
  For Each key In paths.Keys
    If StrComp(CStr(paths(key)), CStr(filePath), vbTextCompare) = 0 Then Exit Sub
  Next

  paths.Add CStr(paths.Count), CStr(filePath)
End Sub

Sub SC_ResolveAttachmentItem(paths, item, baseFolder, logPath)
  If IsObject(item) Or IsEmpty(item) Or IsNull(item) Then Exit Sub

  Dim resolved
  resolved = SC_ResolveExistingFilePath(item, baseFolder)
  If Len(resolved) > 0 Then
    SC_AddResolvedPath paths, resolved
  Else
    SC_LogAssetWarning logPath, "Attachment path not found: " & CStr(item)
  End If
End Sub

Sub SC_ResolveAttachmentText(paths, rawValue, baseFolder, logPath)
  Dim cleaned
  cleaned = Replace(CStr(rawValue), ",", ";")
  Dim parts
  parts = Split(cleaned, ";")

  Dim i
  For i = LBound(parts) To UBound(parts)
    SC_ResolveAttachmentItem paths, Trim(CStr(parts(i))), baseFolder, logPath
  Next
End Sub

Function SC_ResolveAttachmentList(attachments, baseFolder, logPath)
  Dim paths
  Set paths = SC_CreateDictionary()

  If IsEmpty(attachments) Or IsNull(attachments) Then
    Set SC_ResolveAttachmentList = paths
    Exit Function
  End If

  If IsArray(attachments) Then
    Dim i
    For i = LBound(attachments) To UBound(attachments)
      SC_ResolveAttachmentItem paths, attachments(i), baseFolder, logPath
    Next
  ElseIf IsObject(attachments) Then
    Dim objectTypeName
    objectTypeName = TypeName(attachments)
    If StrComp(objectTypeName, "Dictionary", vbTextCompare) = 0 Then
      Dim key
      For Each key In attachments.Keys
        SC_ResolveAttachmentItem paths, attachments(key), baseFolder, logPath
      Next
    Else
      Dim item
      For Each item In attachments
        SC_ResolveAttachmentItem paths, item, baseFolder, logPath
      Next
    End If
  Else
    SC_ResolveAttachmentText paths, attachments, baseFolder, logPath
  End If

  Set SC_ResolveAttachmentList = paths
End Function

Function SC_GetMessageSpecObject(messageSpec, key)
  Set SC_GetMessageSpecObject = Nothing
  On Error Resume Next
  If IsObject(messageSpec) Then
    If messageSpec.Exists(key) Then
      If IsObject(messageSpec(key)) Then Set SC_GetMessageSpecObject = messageSpec(key)
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_ResolveMessageAttachments(messageSpec, logPath)
  Dim paths
  Set paths = SC_CreateDictionary()

  Dim baseFolder
  baseFolder = CStr(SC_DictValue(messageSpec, "attachment_base_folder", ""))

  If IsObject(messageSpec) Then
    If messageSpec.Exists("attachments") Then
      If IsObject(messageSpec("attachments")) Then
        Set paths = SC_ResolveAttachmentList(messageSpec("attachments"), baseFolder, logPath)
      Else
        Set paths = SC_ResolveAttachmentList(messageSpec("attachments"), baseFolder, logPath)
      End If
    End If
  End If

  Set SC_ResolveMessageAttachments = paths
End Function

Function SC_IsSafeContentIdCharacter(ch)
  Dim code
  code = Asc(ch)
  SC_IsSafeContentIdCharacter = _
    (code >= 48 And code <= 57) Or _
    (code >= 65 And code <= 90) Or _
    (code >= 97 And code <= 122) Or _
    ch = "." Or ch = "_" Or ch = "-"
End Function

Function SC_SanitizeContentId(fileName)
  Dim raw
  raw = CStr(fileName)
  Dim output
  output = ""

  Dim i
  For i = 1 To Len(raw)
    Dim ch
    ch = Mid(raw, i, 1)
    If SC_IsSafeContentIdCharacter(ch) Then
      output = output & ch
    Else
      output = output & "_"
    End If
  Next

  If Len(output) = 0 Then output = "inline_asset"
  SC_SanitizeContentId = output
End Function

Function SC_EncodeSpaces(value)
  SC_EncodeSpaces = Replace(CStr(value), " ", "%20")
End Function

Function SC_RewriteInlineImageReferences(htmlText, templateBaseName, fileName)
  Dim cid
  cid = SC_SanitizeContentId(fileName)
  Dim target
  target = "cid:" & cid
  Dim result
  result = CStr(htmlText)
  Dim backslash
  backslash = Chr(92)

  result = Replace(result, templateBaseName & ".files/" & fileName, target)
  result = Replace(result, templateBaseName & ".files" & backslash & fileName, target)
  result = Replace(result, SC_EncodeSpaces(templateBaseName) & ".files/" & SC_EncodeSpaces(fileName), target)
  result = Replace(result, SC_EncodeSpaces(templateBaseName) & ".files" & backslash & SC_EncodeSpaces(fileName), target)
  result = Replace(result, "=""" & fileName & """", "=""" & target & """")
  result = Replace(result, "='" & fileName & "'", "='" & target & "'")
  result = Replace(result, "=""" & SC_EncodeSpaces(fileName) & """", "=""" & target & """")
  result = Replace(result, "='" & SC_EncodeSpaces(fileName) & "'", "='" & target & "'")

  SC_RewriteInlineImageReferences = result
End Function

Function SC_PrepareInlineImages(htmlBody, templatePath, inlineFolder, logPath)
  Dim info
  Set info = SC_CreateDictionary()
  info("html_body") = CStr(htmlBody)
  info("inline_folder") = ""
  info("inline_count") = 0
  Set info("inline_files") = SC_CreateDictionary()
  Set info("content_ids") = SC_CreateDictionary()

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  Dim resolvedFolder
  resolvedFolder = Trim(CStr(inlineFolder))
  Dim templateBaseName
  templateBaseName = ""

  If Len(Trim(CStr(templatePath))) > 0 Then
    templateBaseName = fso.GetBaseName(CStr(templatePath))
    If Len(resolvedFolder) = 0 Then
      Dim parentFolder
      parentFolder = fso.GetParentFolderName(CStr(templatePath))
      resolvedFolder = fso.BuildPath(parentFolder, templateBaseName & ".files")
    End If
  End If

  If Len(resolvedFolder) = 0 Then
    Set SC_PrepareInlineImages = info
    Exit Function
  End If

  If Not fso.FolderExists(resolvedFolder) Then
    SC_LogAssetWarning logPath, "Inline image folder not found: " & resolvedFolder
    Set SC_PrepareInlineImages = info
    Exit Function
  End If

  If Len(templateBaseName) = 0 Then templateBaseName = fso.GetBaseName(resolvedFolder)
  info("inline_folder") = fso.GetAbsolutePathName(resolvedFolder)

  Dim file
  For Each file In fso.GetFolder(resolvedFolder).Files
    Dim cid
    cid = SC_SanitizeContentId(file.Name)
    info("inline_files").Add CStr(info("inline_files").Count), file.Path
    info("content_ids").Add CStr(info("content_ids").Count), cid
    info("html_body") = SC_RewriteInlineImageReferences(info("html_body"), templateBaseName, file.Name)
  Next

  info("inline_count") = info("inline_files").Count
  Set SC_PrepareInlineImages = info
End Function

Sub SC_AddAttachmentsToMail(mailItem, attachmentPaths, logPath)
  If Not IsObject(attachmentPaths) Then Exit Sub

  Dim key
  For Each key In attachmentPaths.Keys
    On Error Resume Next
    mailItem.Attachments.Add CStr(attachmentPaths(key))
    If Err.Number <> 0 Then
      SC_LogAssetWarning logPath, "Failed to add attachment: " & CStr(attachmentPaths(key)) & " - " & Err.Description
      Err.Clear
    End If
    On Error GoTo 0
  Next
End Sub

Sub SC_AddInlineImagesToMail(mailItem, inlineInfo, logPath)
  If Not IsObject(inlineInfo) Then Exit Sub
  If Not inlineInfo.Exists("inline_files") Then Exit Sub
  If Not inlineInfo.Exists("content_ids") Then Exit Sub

  Dim key
  For Each key In inlineInfo("inline_files").Keys
    On Error Resume Next
    Dim attachment
    Set attachment = mailItem.Attachments.Add(CStr(inlineInfo("inline_files")(key)))
    If Err.Number = 0 Then
      attachment.PropertyAccessor.SetProperty SC_PR_ATTACH_CONTENT_ID, CStr(inlineInfo("content_ids")(key))
    End If
    If Err.Number <> 0 Then
      SC_LogAssetWarning logPath, "Failed to add inline image: " & CStr(inlineInfo("inline_files")(key)) & " - " & Err.Description
      Err.Clear
    End If
    On Error GoTo 0
  Next
End Sub
