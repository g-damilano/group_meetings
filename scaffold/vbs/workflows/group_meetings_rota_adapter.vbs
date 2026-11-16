' Group meetings rota/workbook adapter.
' Converts anonymised workbook-shaped data into meetingRecord contracts. This module is
' deliberately file/contract focused and does not use Excel or Outlook COM.

Function SC_GM_RotaAdapterValue(dictionary, key, defaultValue)
  SC_GM_RotaAdapterValue = defaultValue
  On Error Resume Next
  If IsObject(dictionary) Then
    If dictionary.Exists(key) Then SC_GM_RotaAdapterValue = dictionary(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_GM_NormalizeRotaHeader(headerText)
  Dim compact
  compact = LCase(Trim(CStr(headerText)))
  compact = Replace(compact, " ", "")
  compact = Replace(compact, "_", "")
  compact = Replace(compact, "-", "")

  Select Case compact
    Case "date", "meetingdate", "friday", "weekdate"
      SC_GM_NormalizeRotaHeader = "meeting_date"
    Case "literature", "literaturepresenter", "lit"
      SC_GM_NormalizeRotaHeader = "literature_presenter"
    Case "results", "resultspresenter", "res"
      SC_GM_NormalizeRotaHeader = "results_presenter"
    Case "alternate", "alternatetext", "alt"
      SC_GM_NormalizeRotaHeader = "alternate_text"
    Case "room", "location"
      SC_GM_NormalizeRotaHeader = "room"
    Case "labmanager", "labmanagement"
      SC_GM_NormalizeRotaHeader = "lab_manager"
    Case "meetinglead", "chair", "timekeeper"
      SC_GM_NormalizeRotaHeader = "meeting_lead"
    Case "grouprecipient", "groupemail", "group"
      SC_GM_NormalizeRotaHeader = "group_recipient"
    Case "managerrecipient", "manageremail", "manager"
      SC_GM_NormalizeRotaHeader = "manager_recipient"
    Case "sender", "sendername", "coordinator"
      SC_GM_NormalizeRotaHeader = "sender_name"
    Case "includesamedayreminder", "samedayreminder"
      SC_GM_NormalizeRotaHeader = "include_same_day_reminder"
    Case "sendpresenterreminders"
      SC_GM_NormalizeRotaHeader = "send_presenter_reminders"
    Case "sendmeetingleadreminder"
      SC_GM_NormalizeRotaHeader = "send_meeting_lead_reminder"
    Case "sendlabmanagerreminder"
      SC_GM_NormalizeRotaHeader = "send_lab_manager_reminder"
    Case Else
      SC_GM_NormalizeRotaHeader = compact
  End Select
End Function

Function SC_GM_SplitCsvLine(lineText)
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
  Set SC_GM_SplitCsvLine = fields
End Function

Function SC_GM_ParseBooleanText(value, defaultValue)
  Dim text
  text = LCase(Trim(CStr(value)))
  Select Case text
    Case "true", "yes", "y", "1", "on"
      SC_GM_ParseBooleanText = True
    Case "false", "no", "n", "0", "off"
      SC_GM_ParseBooleanText = False
    Case Else
      SC_GM_ParseBooleanText = CBool(defaultValue)
  End Select
End Function

Function SC_GM_ParseDateText(value)
  Dim text
  text = Trim(CStr(value))
  If Len(text) = 0 Then
    SC_GM_ParseDateText = text
    Exit Function
  End If

  Dim parts
  parts = Split(text, "-")
  If UBound(parts) = 2 Then
    On Error Resume Next
    SC_GM_ParseDateText = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
    If Err.Number = 0 Then
      On Error GoTo 0
      Exit Function
    End If
    Err.Clear
    On Error GoTo 0
  End If

  If IsDate(value) Then
    SC_GM_ParseDateText = CDate(value)
  Else
    SC_GM_ParseDateText = text
  End If
End Function

Function SC_GM_RotaDateKey(value)
  SC_GM_RotaDateKey = SC_FormatKeyDate(SC_GM_ParseDateText(value))
End Function

Function SC_GM_ReadRotaCsv(csvPath)
  Dim rows
  Set rows = SC_CreateDictionary()

  Dim text
  text = SC_ReadTextFile(csvPath)
  If Len(text) = 0 Then
    Set SC_GM_ReadRotaCsv = rows
    Exit Function
  End If

  Dim lines
  lines = Split(SC_NormalizeLineEndings(text), vbLf)

  Dim headerFields
  Set headerFields = Nothing
  Dim headers
  Set headers = SC_CreateDictionary()

  Dim i
  For i = LBound(lines) To UBound(lines)
    Dim line
    line = Trim(CStr(lines(i)))
    If Len(line) > 0 And Left(line, 1) <> "#" Then
      If headerFields Is Nothing Then
        Set headerFields = SC_GM_SplitCsvLine(line)
        Dim h
        For Each h In headerFields.Keys
          headers(CStr(h)) = SC_GM_NormalizeRotaHeader(headerFields(h))
        Next
      Else
        Dim fields
        Set fields = SC_GM_SplitCsvLine(CStr(lines(i)))
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

  Set SC_GM_ReadRotaCsv = rows
End Function

Function SC_GM_RotaRowText(row, key)
  SC_GM_RotaRowText = Trim(CStr(SC_GM_RotaAdapterValue(row, key, "")))
End Function

Function SC_GM_RotaRowsDateDiff(row, targetDate)
  SC_GM_RotaRowsDateDiff = 0
  On Error Resume Next
  SC_GM_RotaRowsDateDiff = DateDiff("d", CDate(SC_GM_ParseDateText(SC_GM_RotaRowText(row, "meeting_date"))), CDate(SC_GM_ParseDateText(targetDate)))
  If Err.Number <> 0 Then
    Err.Clear
    SC_GM_RotaRowsDateDiff = 0
  End If
  On Error GoTo 0
End Function

Function SC_GM_FindRotaRowByDate(rows, targetDate)
  Set SC_GM_FindRotaRowByDate = Nothing
  If Not IsObject(rows) Then Exit Function

  Dim targetKey
  targetKey = SC_GM_RotaDateKey(targetDate)

  Dim key
  For Each key In rows.Keys
    Dim row
    Set row = rows(key)
    If SC_GM_RotaDateKey(SC_GM_RotaRowText(row, "meeting_date")) = targetKey Then
      Set SC_GM_FindRotaRowByDate = row
      Exit Function
    End If
  Next
End Function

Function SC_GM_FindPreviousRotaRow(rows, targetDate)
  Set SC_GM_FindPreviousRotaRow = Nothing
  If Not IsObject(rows) Then Exit Function

  Dim bestRow
  Set bestRow = Nothing
  Dim bestDistance
  bestDistance = 999999

  Dim key
  For Each key In rows.Keys
    Dim row
    Set row = rows(key)
    Dim distance
    distance = SC_GM_RotaRowsDateDiff(row, targetDate)
    If distance > 0 And distance < bestDistance Then
      bestDistance = distance
      Set bestRow = row
    End If
  Next

  Set SC_GM_FindPreviousRotaRow = bestRow
End Function

Function SC_GM_FindNextRotaRow(rows, targetDate)
  Set SC_GM_FindNextRotaRow = Nothing
  If Not IsObject(rows) Then Exit Function

  Dim bestRow
  Set bestRow = Nothing
  Dim bestDistance
  bestDistance = 999999

  Dim key
  For Each key In rows.Keys
    Dim row
    Set row = rows(key)
    Dim distance
    distance = SC_GM_RotaRowsDateDiff(row, targetDate)
    If distance < 0 And Abs(distance) < bestDistance Then
      bestDistance = Abs(distance)
      Set bestRow = row
    End If
  Next

  Set SC_GM_FindNextRotaRow = bestRow
End Function

Function SC_GM_DefaultRecordText(defaults, key, fallback)
  SC_GM_DefaultRecordText = CStr(SC_GM_RotaAdapterValue(defaults, key, fallback))
End Function

Function SC_GM_RowOrDefault(row, defaults, key, fallback)
  Dim value
  value = SC_GM_RotaRowText(row, key)
  If Len(value) = 0 Then value = SC_GM_DefaultRecordText(defaults, key, fallback)
  SC_GM_RowOrDefault = value
End Function

Function SC_GM_BuildMeetingRecordFromRotaRows(currentRow, previousRow, nextRow, defaults)
  Set SC_GM_BuildMeetingRecordFromRotaRows = Nothing
  If currentRow Is Nothing Then Exit Function

  Dim meetingDate
  meetingDate = SC_GM_ParseDateText(SC_GM_RotaRowText(currentRow, "meeting_date"))

  Dim record
  Set record = SC_GM_CreateMeetingRecord(meetingDate)
  record("literature_presenter") = SC_GM_RotaRowText(currentRow, "literature_presenter")
  record("results_presenter") = SC_GM_RotaRowText(currentRow, "results_presenter")
  record("alternate_text") = SC_GM_RotaRowText(currentRow, "alternate_text")
  record("room") = SC_GM_RotaRowText(currentRow, "room")
  record("lab_manager") = SC_GM_RotaRowText(currentRow, "lab_manager")

  Dim meetingLead
  meetingLead = SC_GM_RotaRowText(currentRow, "meeting_lead")
  If Len(meetingLead) = 0 Then
    If Not previousRow Is Nothing Then meetingLead = SC_GM_RotaRowText(previousRow, "lab_manager")
  End If
  If Len(meetingLead) = 0 Then meetingLead = SC_GM_DefaultRecordText(defaults, "default_meeting_lead", "")
  record("meeting_lead") = meetingLead

  record("group_recipient") = SC_GM_RowOrDefault(currentRow, defaults, "group_recipient", "group")
  record("manager_recipient") = SC_GM_RowOrDefault(currentRow, defaults, "manager_recipient", "manager")
  record("sender_name") = SC_GM_RowOrDefault(currentRow, defaults, "sender_name", "Coordinator")
  record("include_same_day_reminder") = SC_GM_ParseBooleanText(SC_GM_RowOrDefault(currentRow, defaults, "include_same_day_reminder", "false"), False)
  record("send_presenter_reminders") = SC_GM_ParseBooleanText(SC_GM_RowOrDefault(currentRow, defaults, "send_presenter_reminders", "true"), True)
  record("send_meeting_lead_reminder") = SC_GM_ParseBooleanText(SC_GM_RowOrDefault(currentRow, defaults, "send_meeting_lead_reminder", "true"), True)
  record("send_lab_manager_reminder") = SC_GM_ParseBooleanText(SC_GM_RowOrDefault(currentRow, defaults, "send_lab_manager_reminder", "true"), True)

  If Not nextRow Is Nothing Then
    record("next_meeting_date") = SC_GM_ParseDateText(SC_GM_RotaRowText(nextRow, "meeting_date"))
    record("next_literature_presenter") = SC_GM_RotaRowText(nextRow, "literature_presenter")
    record("next_results_presenter") = SC_GM_RotaRowText(nextRow, "results_presenter")
    record("next_lab_manager") = SC_GM_RotaRowText(nextRow, "lab_manager")
    record("next_meeting_lead") = SC_GM_RotaRowText(nextRow, "meeting_lead")
    If Len(record("next_meeting_lead")) = 0 Then record("next_meeting_lead") = record("lab_manager")
    record("next_alternate_text") = SC_GM_RotaRowText(nextRow, "alternate_text")
  End If

  Set SC_GM_BuildMeetingRecordFromRotaRows = record
End Function

Function SC_GM_LoadMeetingRecordFromCsv(csvPath, targetDate, defaults)
  Dim result
  Set result = SC_CreateDictionary()
  result("state") = SC_STATE_FAILED_SAFE
  result("reason") = ""
  result("target_date") = SC_GM_RotaDateKey(targetDate)
  result("row_count") = 0
  Set result("rows") = SC_CreateDictionary()

  Dim rows
  Set rows = SC_GM_ReadRotaCsv(csvPath)
  Set result("rows") = rows
  result("row_count") = rows.Count

  If rows.Count = 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Rota source is empty or unreadable."
    Set SC_GM_LoadMeetingRecordFromCsv = result
    Exit Function
  End If

  Dim currentRow
  Set currentRow = SC_GM_FindRotaRowByDate(rows, targetDate)
  If currentRow Is Nothing Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "No rota row found for target date."
    Set SC_GM_LoadMeetingRecordFromCsv = result
    Exit Function
  End If

  Dim previousRow
  Set previousRow = SC_GM_FindPreviousRotaRow(rows, targetDate)
  Dim nextRow
  Set nextRow = SC_GM_FindNextRotaRow(rows, targetDate)

  Dim record
  Set record = SC_GM_BuildMeetingRecordFromRotaRows(currentRow, previousRow, nextRow, defaults)
  result("state") = SC_STATE_COMPOSED
  result("reason") = "Meeting record loaded from rota source."
  Set result("record") = record
  Set SC_GM_LoadMeetingRecordFromCsv = result
End Function

Function SC_GM_CreateMessagePlanFromCsv(workflowContext, csvPath, targetDate, defaults)
  Dim loadResult
  Set loadResult = SC_GM_LoadMeetingRecordFromCsv(csvPath, targetDate, defaults)

  If loadResult("state") <> SC_STATE_COMPOSED Then
    Dim failedPlan
    Set failedPlan = SC_GM_CreatePlan()
    failedPlan("state") = SC_STATE_ACTION_REQUIRED
    failedPlan("reason") = CStr(loadResult("reason"))
    failedPlan("meeting_date") = SC_GM_RotaDateKey(targetDate)
    SC_GM_AddPlanSkip failedPlan, "ROTA_RECORD_UNAVAILABLE", CStr(loadResult("reason"))
    Set SC_GM_CreateMessagePlanFromCsv = failedPlan
    Exit Function
  End If

  Dim plan
  Set plan = SC_GM_CreateMessagePlan(workflowContext, loadResult("record"))
  plan("rota_source") = CStr(csvPath)
  Set SC_GM_CreateMessagePlanFromCsv = plan
End Function
