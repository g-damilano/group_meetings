' Attachment and inline image helpers

Function ResolveAttachments(rawValue, baseFolder)
  Dim cleaned: cleaned = Trim(rawValue)
  cleaned = Replace(cleaned, ",", ";")
  If Len(cleaned) = 0 Then
    ResolveAttachments = Array()
    Exit Function
  End If

  Dim parts: parts = Split(cleaned, ";")
  Dim resolvedList: Set resolvedList = CreateObject("System.Collections.ArrayList")
  Dim i, candidate, item
  For i = 0 To UBound(parts)
    item = Trim(parts(i))
    If Len(item) > 0 Then
      candidate = item
      If Not fso.FileExists(candidate) Then
        If Left(candidate, 2) = ".\" Or Left(candidate, 2) = "./" Then
          candidate = fso.BuildPath(baseFolder, Mid(candidate, 3))
        ElseIf Not IsAbsolutePath(candidate) Then
          candidate = fso.BuildPath(baseFolder, candidate)
        End If
      End If
      If fso.FileExists(candidate) Then
        resolvedList.Add candidate
        LogEvent "Attachment queued: " & candidate
      Else
        LogEvent "Attachment not found: " & item
      End If
    End If
  Next

  If resolvedList.Count = 0 Then
    ResolveAttachments = Array()
  Else
    ResolveAttachments = ArrayFromArrayList(resolvedList)
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
  Dim backslash: backslash = Chr(92)

  htmlText = Replace(htmlText, templateBaseName & ".files/" & fileName, target)
  htmlText = Replace(htmlText, templateBaseName & ".files" & backslash & fileName, target)
  htmlText = Replace(htmlText, encodedBase & ".files/" & encodedFile, target)
  htmlText = Replace(htmlText, encodedBase & ".files" & backslash & encodedFile, target)

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
