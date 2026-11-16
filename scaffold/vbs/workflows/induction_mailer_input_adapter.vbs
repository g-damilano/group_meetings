' Induction mailer input adapter.
' Converts anonymised recipient and template-manifest CSV data into workflow records.

Function SC_IM_Value(dictionary, key, defaultValue)
  SC_IM_Value = defaultValue
  On Error Resume Next
  If IsObject(dictionary) Then
    If dictionary.Exists(key) Then SC_IM_Value = dictionary(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_IM_SplitCsvLine(lineText)
  Dim fields
  Set fields = SC_CreateDictionary()

  Dim current
  current = ""
  Dim inQuotes
  inQuotes = False

  Dim i
  i = 1
  Do While i <= Len(CStr(lineText))
    Dim ch
    ch = Mid(CStr(lineText), i, 1)
    If ch = Chr(34) Then
      If inQuotes And i < Len(CStr(lineText)) And Mid(CStr(lineText), i + 1, 1) = Chr(34) Then
        current = current & Chr(34)
        i = i + 1
      Else
        inQuotes = Not inQuotes
      End If
    ElseIf ch = "," And Not inQuotes Then
      fields.Add CStr(fields.Count), current
      current = ""
    Else
      current = current & ch
    End If
    i = i + 1
  Loop

  fields.Add CStr(fields.Count), current
  Set SC_IM_SplitCsvLine = fields
End Function

Function SC_IM_NormalizeHeader(headerText)
  Dim raw
  raw = LCase(Trim(CStr(headerText)))
  Dim compact
  compact = raw
  compact = Replace(compact, " ", "")
  compact = Replace(compact, "_", "")
  compact = Replace(compact, "-", "")

  Select Case compact
    Case "firstname", "name", "givenname"
      SC_IM_NormalizeHeader = "first_name"
    Case "lastname", "surname", "familyname"
      SC_IM_NormalizeHeader = "last_name"
    Case "fullname"
      SC_IM_NormalizeHeader = "full_name"
    Case "displayname"
      If InStr(1, raw, "display_", vbTextCompare) > 0 Or InStr(1, raw, "display ", vbTextCompare) > 0 Or InStr(1, raw, "display-", vbTextCompare) > 0 Then
        SC_IM_NormalizeHeader = "display_name"
      Else
        SC_IM_NormalizeHeader = "full_name"
      End If
    Case "startdate", "arrivaldate"
      SC_IM_NormalizeHeader = "start_date"
    Case "templateid", "code"
      SC_IM_NormalizeHeader = "template_id"
    Case "templatepath", "path", "file"
      SC_IM_NormalizeHeader = "template_path"
    Case "inlineimagefolder", "inlinefolder"
      SC_IM_NormalizeHeader = "inline_image_folder"
    Case "targetkey", "target"
      SC_IM_NormalizeHeader = "target_key"
    Case "label"
      SC_IM_NormalizeHeader = "display_name"
    Case "attachmentbasefolder"
      SC_IM_NormalizeHeader = "attachment_base_folder"
    Case "deliveryaction"
      SC_IM_NormalizeHeader = "delivery_action"
    Case "deferminutes", "defermin"
      SC_IM_NormalizeHeader = "defer_minutes"
    Case "urgentverifyaftersubmit"
      SC_IM_NormalizeHeader = "urgent_verify_after_submit"
    Case Else
      SC_IM_NormalizeHeader = compact
  End Select
End Function

Function SC_IM_ParseBooleanText(value, defaultValue)
  Dim text
  text = LCase(Trim(CStr(value)))
  Select Case text
    Case "true", "yes", "y", "1", "on"
      SC_IM_ParseBooleanText = True
    Case "false", "no", "n", "0", "off"
      SC_IM_ParseBooleanText = False
    Case Else
      SC_IM_ParseBooleanText = CBool(defaultValue)
  End Select
End Function

Function SC_IM_IsValidTemplateMode(value)
  Dim text
  text = LCase(Trim(CStr(value)))
  SC_IM_IsValidTemplateMode = (text = "broadcast" Or text = "personal")
End Function

Function SC_IM_IsValidDeliveryAction(value)
  Dim text
  text = LCase(Trim(CStr(value)))
  SC_IM_IsValidDeliveryAction = (Len(text) = 0 Or text = "display" Or text = "draft" Or text = "send")
End Function

Function SC_IM_IsNonNegativeIntegerText(value)
  Dim text
  text = Trim(CStr(value))
  If Len(text) = 0 Then
    SC_IM_IsNonNegativeIntegerText = True
    Exit Function
  End If

  Dim i
  For i = 1 To Len(text)
    Dim ch
    ch = Mid(text, i, 1)
    If ch < "0" Or ch > "9" Then
      SC_IM_IsNonNegativeIntegerText = False
      Exit Function
    End If
  Next
  SC_IM_IsNonNegativeIntegerText = True
End Function

Sub SC_IM_AddInputIssue(issues, sourceLine, code, reason)
  Dim issue
  Set issue = SC_CreateDictionary()
  issue("source_line") = CStr(sourceLine)
  issue("code") = CStr(code)
  issue("reason") = CStr(reason)
  issues.Add CStr(issues.Count), issue
End Sub

Function SC_IM_ReadCsvRows(csvPath)
  Dim rows
  Set rows = SC_CreateDictionary()

  Dim text
  text = SC_ReadTextFile(csvPath)
  If Len(text) = 0 Then
    Set SC_IM_ReadCsvRows = rows
    Exit Function
  End If

  Dim lines
  lines = Split(SC_NormalizeLineEndings(text), vbLf)
  Dim headers
  Set headers = SC_CreateDictionary()
  Dim headerSeen
  headerSeen = False

  Dim i
  For i = LBound(lines) To UBound(lines)
    Dim line
    line = Trim(CStr(lines(i)))
    If Len(line) > 0 And Left(line, 1) <> "#" Then
      Dim fields
      Set fields = SC_IM_SplitCsvLine(line)
      If Not headerSeen Then
        Dim h
        For Each h In fields.Keys
          headers(CStr(h)) = SC_IM_NormalizeHeader(fields(h))
        Next
        headerSeen = True
      Else
        Dim row
        Set row = SC_CreateDictionary()
        row("source_line") = i + 1
        Dim key
        For Each key In headers.Keys
          If fields.Exists(CStr(key)) Then
            row(headers(key)) = Trim(CStr(fields(CStr(key))))
          Else
            row(headers(key)) = ""
          End If
        Next
        rows.Add CStr(rows.Count), row
      End If
    End If
  Next

  Set SC_IM_ReadCsvRows = rows
End Function

Function SC_IM_ResolvePath(baseFolder, pathText)
  Dim value
  value = Trim(CStr(pathText))
  If Len(value) = 0 Then
    SC_IM_ResolvePath = ""
    Exit Function
  End If

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If InStr(1, value, ":", vbTextCompare) > 0 Or Left(value, 2) = "\\" Then
    SC_IM_ResolvePath = value
  Else
  SC_IM_ResolvePath = fso.BuildPath(baseFolder, value)
  End If
End Function

Sub SC_IM_FillMissingRecipientNameParts(record, displayText, sourceName)
  If Not IsObject(record) Then Exit Sub
  If Len(Trim(CStr(displayText))) = 0 Then Exit Sub

  Dim parts
  Set parts = SC_ParseDisplayNameParts(displayText)
  If parts("name_parse_status") = "not_attempted" Then Exit Sub

  If Len(Trim(CStr(record("first_name")))) = 0 Then record("first_name") = parts("first_name")
  If Len(Trim(CStr(record("last_name")))) = 0 Then record("last_name") = parts("last_name")
  If Len(Trim(CStr(record("surname")))) = 0 Then record("surname") = parts("surname")
  If Len(Trim(CStr(record("full_name")))) = 0 Then record("full_name") = parts("full_name")
  If Len(Trim(CStr(record("name_parse_status")))) = 0 Or record("name_parse_status") = "not_attempted" Then record("name_parse_status") = parts("name_parse_status")
  If Len(Trim(CStr(record("name_parse_source")))) = 0 Then record("name_parse_source") = CStr(sourceName)
  If Len(Trim(CStr(record("name_parse_order")))) = 0 Then record("name_parse_order") = parts("name_parse_order")
End Sub

Function SC_IM_CreateRecipientRecord(row)
  Dim record
  Set record = SC_CreateDictionary()
  record("alias") = CStr(SC_IM_Value(row, "alias", ""))
  record("email") = CStr(SC_IM_Value(row, "email", ""))
  record("first_name") = CStr(SC_IM_Value(row, "first_name", ""))
  record("last_name") = CStr(SC_IM_Value(row, "last_name", ""))
  record("surname") = CStr(SC_IM_Value(row, "last_name", ""))
  record("display_name") = CStr(SC_IM_Value(row, "display_name", ""))
  record("full_name") = CStr(SC_IM_Value(row, "full_name", ""))
  record("name_parse_status") = "not_attempted"
  record("name_parse_source") = ""
  record("name_parse_order") = ""

  If Len(Trim(record("full_name"))) = 0 And Len(Trim(record("first_name") & record("last_name"))) > 0 Then record("full_name") = Trim(record("first_name") & " " & record("last_name"))
  If Len(Trim(record("full_name"))) = 0 And Len(Trim(record("display_name"))) > 0 Then record("full_name") = Trim(record("display_name"))
  If Len(Trim(record("first_name"))) = 0 Or Len(Trim(record("last_name"))) = 0 Then SC_IM_FillMissingRecipientNameParts record, record("full_name"), "full-name"
  If Len(Trim(record("first_name"))) = 0 Or Len(Trim(record("last_name"))) = 0 Then SC_IM_FillMissingRecipientNameParts record, record("display_name"), "display-name"
  If Len(Trim(record("first_name"))) = 0 Or Len(Trim(record("last_name"))) = 0 Then SC_IM_FillMissingRecipientNameParts record, SC_DisplayNameFromRecipientText(record("email")), "email-display"

  If Len(Trim(record("full_name"))) = 0 Then record("full_name") = Trim(record("first_name") & " " & record("last_name"))
  If Len(Trim(record("surname"))) = 0 Then record("surname") = record("last_name")
  record("start_date") = CStr(SC_IM_Value(row, "start_date", ""))
  record("cohort") = CStr(SC_IM_Value(row, "cohort", ""))
  record("active") = SC_IM_ParseBooleanText(SC_IM_Value(row, "active", "true"), True)
  Set SC_IM_CreateRecipientRecord = record
End Function

Function SC_IM_CreateTemplateRecord(row, manifestFolder)
  Dim record
  Set record = SC_CreateDictionary()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  record("template_id") = CStr(SC_IM_Value(row, "template_id", ""))
  record("mode") = LCase(CStr(SC_IM_Value(row, "mode", "personal")))
  record("target_key") = CStr(SC_IM_Value(row, "target_key", ""))
  If Len(Trim(record("target_key"))) = 0 Then record("target_key") = record("template_id")
  record("display_name") = CStr(SC_IM_Value(row, "display_name", record("template_id")))
  record("template_path") = SC_IM_ResolvePath(manifestFolder, SC_IM_Value(row, "template_path", ""))
  record("template_exists") = (Len(Trim(CStr(record("template_path")))) > 0 And fso.FileExists(CStr(record("template_path"))))
  record("inline_image_folder") = SC_IM_ResolvePath(manifestFolder, SC_IM_Value(row, "inline_image_folder", ""))
  record("attachment_base_folder") = SC_IM_ResolvePath(manifestFolder, SC_IM_Value(row, "attachment_base_folder", ""))
  record("attachments") = CStr(SC_IM_Value(row, "attachments", ""))
  record("cc") = CStr(SC_IM_Value(row, "cc", ""))
  record("bcc") = CStr(SC_IM_Value(row, "bcc", ""))
  record("delivery_action") = LCase(CStr(SC_IM_Value(row, "delivery_action", "")))
  record("defer_minutes") = CStr(SC_IM_Value(row, "defer_minutes", ""))
  record("urgent_verify_after_submit") = SC_IM_ParseBooleanText(SC_IM_Value(row, "urgent_verify_after_submit", "false"), False)
  record("enabled") = SC_IM_ParseBooleanText(SC_IM_Value(row, "enabled", "true"), True)
  Set SC_IM_CreateTemplateRecord = record
End Function

Function SC_IM_LoadRecipientCsv(csvPath)
  Dim result
  Set result = SC_CreateDictionary()
  result("state") = SC_STATE_FAILED_SAFE
  result("reason") = ""
  result("source") = CStr(csvPath)
  Set result("recipients") = SC_CreateDictionary()
  result("row_count") = 0
  result("active_count") = 0
  result("invalid_count") = 0
  Set result("invalid_rows") = SC_CreateDictionary()

  Dim rows
  Set rows = SC_IM_ReadCsvRows(csvPath)
  result("row_count") = rows.Count

  Dim key
  For Each key In rows.Keys
    Dim record
    Set record = SC_IM_CreateRecipientRecord(rows(key))
    result("recipients").Add CStr(result("recipients").Count), record
    If CBool(record("active")) Then
      result("active_count") = CLng(result("active_count")) + 1
      If Len(Trim(CStr(record("email")))) = 0 Then
        result("invalid_count") = CLng(result("invalid_count")) + 1
        SC_IM_AddInputIssue result("invalid_rows"), SC_IM_Value(rows(key), "source_line", ""), "ACTIVE_RECIPIENT_MISSING_EMAIL", "Active recipient rows require an email value."
      End If
    End If
  Next

  If rows.Count = 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Recipient source is empty or unreadable."
  ElseIf CLng(result("active_count")) = 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Recipient source has no active recipients."
  ElseIf CLng(result("invalid_count")) > 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Recipient source has invalid active recipient rows."
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Recipients loaded."
  End If

  Set SC_IM_LoadRecipientCsv = result
End Function

Function SC_IM_LoadTemplateManifestCsv(csvPath)
  Dim result
  Set result = SC_CreateDictionary()
  result("state") = SC_STATE_FAILED_SAFE
  result("reason") = ""
  result("source") = CStr(csvPath)
  Set result("templates") = SC_CreateDictionary()
  result("row_count") = 0
  result("enabled_count") = 0
  result("enabled_existing_template_count") = 0
  result("invalid_count") = 0
  Set result("invalid_rows") = SC_CreateDictionary()

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim manifestFolder
  manifestFolder = fso.GetParentFolderName(CStr(csvPath))

  Dim rows
  Set rows = SC_IM_ReadCsvRows(csvPath)
  result("row_count") = rows.Count

  Dim key
  For Each key In rows.Keys
    Dim record
    Set record = SC_IM_CreateTemplateRecord(rows(key), manifestFolder)
    result("templates").Add CStr(result("templates").Count), record
    If CBool(record("enabled")) Then result("enabled_count") = CLng(result("enabled_count")) + 1
    If CBool(record("enabled")) And CBool(record("template_exists")) Then result("enabled_existing_template_count") = CLng(result("enabled_existing_template_count")) + 1
    If CBool(record("enabled")) Then
      If Len(Trim(CStr(record("template_id")))) = 0 Then
        result("invalid_count") = CLng(result("invalid_count")) + 1
        SC_IM_AddInputIssue result("invalid_rows"), SC_IM_Value(rows(key), "source_line", ""), "TEMPLATE_ID_MISSING", "Enabled template rows require template_id."
      End If
      If Not SC_IM_IsValidTemplateMode(record("mode")) Then
        result("invalid_count") = CLng(result("invalid_count")) + 1
        SC_IM_AddInputIssue result("invalid_rows"), SC_IM_Value(rows(key), "source_line", ""), "TEMPLATE_MODE_INVALID", "Enabled template mode must be broadcast or personal."
      End If
      If Len(Trim(CStr(record("template_path")))) = 0 Then
        result("invalid_count") = CLng(result("invalid_count")) + 1
        SC_IM_AddInputIssue result("invalid_rows"), SC_IM_Value(rows(key), "source_line", ""), "TEMPLATE_PATH_MISSING", "Enabled template rows require template_path."
      End If
      If Not SC_IM_IsValidDeliveryAction(record("delivery_action")) Then
        result("invalid_count") = CLng(result("invalid_count")) + 1
        SC_IM_AddInputIssue result("invalid_rows"), SC_IM_Value(rows(key), "source_line", ""), "DELIVERY_ACTION_INVALID", "delivery_action must be blank, display, draft, or send."
      End If
      If Not SC_IM_IsNonNegativeIntegerText(record("defer_minutes")) Then
        result("invalid_count") = CLng(result("invalid_count")) + 1
        SC_IM_AddInputIssue result("invalid_rows"), SC_IM_Value(rows(key), "source_line", ""), "DEFER_MINUTES_INVALID", "defer_minutes must be a non-negative integer."
      End If
    End If
  Next

  If rows.Count = 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Template manifest is empty or unreadable."
  ElseIf CLng(result("enabled_count")) = 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Template manifest has no enabled templates."
  ElseIf CLng(result("invalid_count")) > 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Template manifest has invalid enabled rows."
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Template manifest loaded."
  End If

  Set SC_IM_LoadTemplateManifestCsv = result
End Function

Function SC_IM_BuildRecipientDirectory(recipients)
  Dim directory
  Set directory = SC_CreateDictionary()

  If IsObject(recipients) Then
    Dim key
    For Each key In recipients.Keys
      Dim recipient
      Set recipient = recipients(key)
      If Len(Trim(CStr(recipient("alias")))) > 0 And Len(Trim(CStr(recipient("email")))) > 0 Then
        directory(CStr(recipient("alias"))) = CStr(recipient("email"))
      End If
      If Len(Trim(CStr(recipient("full_name")))) > 0 And Len(Trim(CStr(recipient("email")))) > 0 Then
        directory(CStr(recipient("full_name"))) = CStr(recipient("email"))
      End If
    Next
  End If

  Set SC_IM_BuildRecipientDirectory = directory
End Function
