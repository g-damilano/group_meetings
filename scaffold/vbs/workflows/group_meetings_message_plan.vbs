' Group meetings message-plan adapter.
' This Stage 2 scaffold module translates anonymised meeting records into workflow
' message requests. It owns workflow decisions, but not Outlook COM or dispatch internals.

Const SC_GM_TYPE_REGULAR = "REGULAR"
Const SC_GM_TYPE_LITERATURE_REMINDER = "LITERATURE-REMINDER"
Const SC_GM_TYPE_RESULTS_REMINDER = "RESULTS-REMINDER"
Const SC_GM_TYPE_MEETING_LEAD_REMINDER = "MEETING-LEAD-REMINDER"
Const SC_GM_TYPE_LAB_MANAGER_REMINDER = "LAB-MANAGER-REMINDER"
Const SC_GM_TYPE_SAME_DAY_REMINDER = "SAME-DAY-REMINDER"
Const SC_GM_TYPE_MEETING_NOT_SET_ALERT = "MEETING-NOT-SET-ALERT"

Function SC_GM_RecordValue(record, key, defaultValue)
  SC_GM_RecordValue = defaultValue
  On Error Resume Next
  If IsObject(record) Then
    If record.Exists(key) Then SC_GM_RecordValue = record(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_GM_RecordText(record, key)
  SC_GM_RecordText = Trim(CStr(SC_GM_RecordValue(record, key, "")))
End Function

Function SC_GM_RecordBool(record, key, defaultValue)
  SC_GM_RecordBool = SC_WorkflowBoolValue(SC_GM_RecordValue(record, key, defaultValue), defaultValue)
End Function

Function SC_GM_DisplayDate(value)
  If IsDate(value) Then
    SC_GM_DisplayDate = SC_FormatKeyDate(value)
  Else
    SC_GM_DisplayDate = Trim(CStr(value))
  End If
End Function

Function SC_GM_FormatDayMonth(value)
  If IsDate(value) Then
    SC_GM_FormatDayMonth = Right("0" & Day(CDate(value)), 2) & "/" & Right("0" & Month(CDate(value)), 2)
  Else
    SC_GM_FormatDayMonth = Trim(CStr(value))
  End If
End Function

Function SC_GM_WeekStart(meetingDate)
  If IsDate(meetingDate) Then
    SC_GM_WeekStart = DateAdd("d", -4, CDate(meetingDate))
  Else
    SC_GM_WeekStart = meetingDate
  End If
End Function

Function SC_GM_WeekEnd(meetingDate)
  SC_GM_WeekEnd = meetingDate
End Function

Function SC_GM_IsCollegeClosureEvent(alternateText)
  SC_GM_IsCollegeClosureEvent = (InStr(1, LCase(Trim(CStr(alternateText))), "college closure", vbTextCompare) > 0)
End Function

Function SC_GM_IsNoMeetingEvent(alternateText)
  Dim normalized
  normalized = LCase(Trim(CStr(alternateText)))
  SC_GM_IsNoMeetingEvent = _
    (InStr(1, normalized, "college closure", vbTextCompare) > 0 Or _
     InStr(1, normalized, "no meeting", vbTextCompare) > 0 Or _
     InStr(1, normalized, "no-meeting", vbTextCompare) > 0)
End Function

Function SC_GM_NoMeetingReasonText(alternateText)
  If SC_GM_IsCollegeClosureEvent(alternateText) Then
    SC_GM_NoMeetingReasonText = "because of closure"
  Else
    SC_GM_NoMeetingReasonText = ""
  End If
End Function

Function SC_GM_IsMeetingNotSet(record)
  SC_GM_IsMeetingNotSet = _
    (Len(SC_GM_RecordText(record, "literature_presenter")) = 0 And _
     Len(SC_GM_RecordText(record, "results_presenter")) = 0 And _
     Len(SC_GM_RecordText(record, "alternate_text")) = 0)
End Function

Function SC_GM_CreateMeetingRecord(meetingDate)
  Dim record
  Set record = SC_CreateDictionary()
  record("meeting_date") = meetingDate
  record("literature_presenter") = ""
  record("results_presenter") = ""
  record("lab_manager") = ""
  record("meeting_lead") = ""
  record("alternate_text") = ""
  record("room") = ""
  record("group_recipient") = "group"
  record("manager_recipient") = "manager"
  record("sender_name") = "Coordinator"
  record("next_meeting_date") = ""
  record("next_literature_presenter") = ""
  record("next_results_presenter") = ""
  record("next_lab_manager") = ""
  record("next_meeting_lead") = ""
  record("next_alternate_text") = ""
  record("include_same_day_reminder") = False
  record("send_presenter_reminders") = True
  record("send_meeting_lead_reminder") = True
  record("send_lab_manager_reminder") = True
  Set SC_GM_CreateMeetingRecord = record
End Function

Function SC_GM_BuildRegularSubject(record)
  Dim meetingDate
  meetingDate = SC_GM_RecordValue(record, "meeting_date", Date)
  If SC_GM_IsNoMeetingEvent(SC_GM_RecordText(record, "alternate_text")) Then
    SC_GM_BuildRegularSubject = "No group meeting (" & SC_GM_DisplayDate(meetingDate) & ")"
  Else
    SC_GM_BuildRegularSubject = "Group meeting @ 11am (" & SC_GM_DisplayDate(meetingDate) & ")"
  End If
End Function

Function SC_GM_BuildPresenterReminderSubject(presentationType, meetingDate)
  Select Case LCase(Trim(CStr(presentationType)))
    Case "literature"
      SC_GM_BuildPresenterReminderSubject = "Literature presentation reminder (" & SC_GM_DisplayDate(meetingDate) & ")"
    Case "results"
      SC_GM_BuildPresenterReminderSubject = "Results presentation reminder (" & SC_GM_DisplayDate(meetingDate) & ")"
    Case Else
      SC_GM_BuildPresenterReminderSubject = "Presentation reminder (" & SC_GM_DisplayDate(meetingDate) & ")"
  End Select
End Function

Function SC_GM_BuildMeetingLeadReminderSubject(meetingDate)
  SC_GM_BuildMeetingLeadReminderSubject = "Meeting lead reminder (" & SC_GM_DisplayDate(meetingDate) & ")"
End Function

Function SC_GM_BuildSameDayReminderSubject(meetingDate, room)
  If Len(Trim(CStr(room))) > 0 Then
    SC_GM_BuildSameDayReminderSubject = "Reminder: group meeting " & SC_GM_DisplayDate(meetingDate) & " @ 11am in " & Trim(CStr(room))
  Else
    SC_GM_BuildSameDayReminderSubject = "Reminder: group meeting " & SC_GM_DisplayDate(meetingDate) & " @ 11am"
  End If
End Function

Function SC_GM_BuildLabManagerReminderSubject(meetingDate)
  SC_GM_BuildLabManagerReminderSubject = "Lab management reminder (" & _
    SC_GM_DisplayDate(SC_GM_WeekStart(meetingDate)) & " - " & SC_GM_DisplayDate(SC_GM_WeekEnd(meetingDate)) & ")"
End Function

Function SC_GM_BodyGreeting()
  SC_GM_BodyGreeting = "Hi all," & vbCrLf & vbCrLf
End Function

Function SC_GM_BodyClosing(record)
  SC_GM_BodyClosing = vbCrLf & "Best," & vbCrLf & SC_GM_RecordText(record, "sender_name")
End Function

Function SC_GM_BuildRegularBody(record)
  Dim meetingDate, alternate, room, body
  meetingDate = SC_GM_RecordValue(record, "meeting_date", Date)
  alternate = SC_GM_RecordText(record, "alternate_text")
  room = SC_GM_RecordText(record, "room")

  body = SC_GM_BodyGreeting()
  If SC_GM_IsNoMeetingEvent(alternate) Then
    body = body & "Please note that no group meeting will take place on " & SC_GM_DisplayDate(meetingDate) & "."
    If Len(SC_GM_NoMeetingReasonText(alternate)) > 0 Then body = body & " Reason: " & SC_GM_NoMeetingReasonText(alternate) & "."
    body = body & vbCrLf
  Else
    body = body & "We will have our group meeting on " & SC_GM_DisplayDate(meetingDate) & " at 11:00"
    If Len(room) > 0 Then body = body & " in " & room
    body = body & "." & vbCrLf & vbCrLf
    If Len(alternate) > 0 Then
      body = body & "- " & alternate & vbCrLf
    Else
      body = body & "- Results: " & SC_GM_RecordText(record, "results_presenter") & vbCrLf
      body = body & "- Literature: " & SC_GM_RecordText(record, "literature_presenter") & vbCrLf
    End If
    body = body & "- Lab manager: " & SC_GM_RecordText(record, "lab_manager") & vbCrLf
    body = body & "- Chair / timekeeper: " & SC_GM_RecordText(record, "meeting_lead") & vbCrLf
  End If

  SC_GM_BuildRegularBody = body & SC_GM_BodyClosing(record)
End Function

Function SC_GM_BuildPresenterReminderBody(record, presenterName, presentationType)
  SC_GM_BuildPresenterReminderBody = "Hi " & Trim(CStr(presenterName)) & "," & vbCrLf & vbCrLf & _
    "This is a reminder that you are scheduled for the " & LCase(CStr(presentationType)) & _
    " presentation at the group meeting on " & SC_GM_DisplayDate(SC_GM_RecordValue(record, "meeting_date", Date)) & "." & vbCrLf & _
    SC_GM_BodyClosing(record)
End Function

Function SC_GM_BuildMeetingLeadReminderBody(record)
  SC_GM_BuildMeetingLeadReminderBody = "Hi " & SC_GM_RecordText(record, "meeting_lead") & "," & vbCrLf & vbCrLf & _
    "This is a reminder that you will lead the group meeting on " & SC_GM_DisplayDate(SC_GM_RecordValue(record, "meeting_date", Date)) & "." & vbCrLf & _
    SC_GM_BodyClosing(record)
End Function

Function SC_GM_BuildLabManagerReminderBody(record)
  SC_GM_BuildLabManagerReminderBody = "Hi " & SC_GM_RecordText(record, "lab_manager") & "," & vbCrLf & vbCrLf & _
    "This is a reminder that you are the lab manager for the week ending " & SC_GM_DisplayDate(SC_GM_RecordValue(record, "meeting_date", Date)) & "." & vbCrLf & _
    SC_GM_BodyClosing(record)
End Function

Function SC_GM_BuildSameDayReminderBody(record)
  Dim room
  room = SC_GM_RecordText(record, "room")
  Dim body
  body = SC_GM_BodyGreeting() & "Just a quick reminder that our group meeting is today at 11:00"
  If Len(room) > 0 Then body = body & " in " & room
  body = body & "." & vbCrLf & SC_GM_BodyClosing(record)
  SC_GM_BuildSameDayReminderBody = body
End Function

Function SC_GM_BuildMeetingNotSetAlertBody(record)
  SC_GM_BuildMeetingNotSetAlertBody = "Meeting setup is incomplete for " & _
    SC_GM_DisplayDate(SC_GM_RecordValue(record, "meeting_date", Date)) & _
    ". No regular meeting email should be sent until the rota is completed."
End Function

Function SC_GM_CreatePlan()
  Dim plan
  Set plan = SC_CreateDictionary()
  plan("workflow_reference") = "group-meetings"
  plan("state") = SC_STATE_NOT_REQUESTED
  plan("reason") = ""
  Set plan("messages") = SC_CreateDictionary()
  Set plan("skips") = SC_CreateDictionary()
  Set SC_GM_CreatePlan = plan
End Function

Sub SC_GM_AddPlanMessage(plan, request, planCode)
  request("plan_code") = CStr(planCode)
  request("sequence") = plan("messages").Count + 1
  plan("messages").Add CStr(plan("messages").Count), request
End Sub

Sub SC_GM_AddPlanSkip(plan, skipCode, reason)
  Dim skipInfo
  Set skipInfo = SC_CreateDictionary()
  skipInfo("code") = CStr(skipCode)
  skipInfo("reason") = CStr(reason)
  plan("skips").Add CStr(plan("skips").Count), skipInfo
End Sub

Function SC_GM_RequestRecipientOrFallback(rawName, fallbackRecipient)
  If Len(Trim(CStr(rawName))) > 0 Then
    SC_GM_RequestRecipientOrFallback = Trim(CStr(rawName))
  Else
    SC_GM_RequestRecipientOrFallback = Trim(CStr(fallbackRecipient))
  End If
End Function

Function SC_GM_CreateRegularRequest(record)
  Dim request
  Set request = SC_CreateGroupMeetingsMessageRequest( _
    SC_GM_RecordValue(record, "meeting_date", Date), _
    SC_GM_TYPE_REGULAR, _
    "team", _
    SC_GM_RecordText(record, "group_recipient"), _
    SC_GM_BuildRegularSubject(record), _
    SC_GM_BuildRegularBody(record))
  SC_AddWorkflowRequestMetadata request, "PLAN_KIND", "regular-distribution"
  Set SC_GM_CreateRegularRequest = request
End Function

Function SC_GM_CreatePresenterReminderRequest(record, presentationType)
  Dim messageType, targetKey, presenterField, presenterName, request
  If LCase(CStr(presentationType)) = "literature" Then
    messageType = SC_GM_TYPE_LITERATURE_REMINDER
    targetKey = "literature"
    presenterField = "literature_presenter"
  Else
    messageType = SC_GM_TYPE_RESULTS_REMINDER
    targetKey = "results"
    presenterField = "results_presenter"
  End If

  presenterName = SC_GM_RecordText(record, presenterField)
  Set request = SC_CreateGroupMeetingsMessageRequest( _
    SC_GM_RecordValue(record, "meeting_date", Date), _
    messageType, _
    targetKey, _
    SC_GM_RequestRecipientOrFallback(presenterName, SC_GM_RecordText(record, "manager_recipient")), _
    SC_GM_BuildPresenterReminderSubject(presentationType, SC_GM_RecordValue(record, "meeting_date", Date)), _
    SC_GM_BuildPresenterReminderBody(record, presenterName, presentationType))
  SC_AddWorkflowRequestMetadata request, "PLAN_KIND", "presenter-reminder"
  If Len(presenterName) = 0 Then SC_AddWorkflowRequestMetadata request, "RECIPIENT_FALLBACK_USED", "true"
  Set SC_GM_CreatePresenterReminderRequest = request
End Function

Function SC_GM_CreateMeetingLeadReminderRequest(record)
  Dim request
  Set request = SC_CreateGroupMeetingsMessageRequest( _
    SC_GM_RecordValue(record, "meeting_date", Date), _
    SC_GM_TYPE_MEETING_LEAD_REMINDER, _
    "meeting-lead", _
    SC_GM_RequestRecipientOrFallback(SC_GM_RecordText(record, "meeting_lead"), SC_GM_RecordText(record, "manager_recipient")), _
    SC_GM_BuildMeetingLeadReminderSubject(SC_GM_RecordValue(record, "meeting_date", Date)), _
    SC_GM_BuildMeetingLeadReminderBody(record))
  SC_AddWorkflowRequestMetadata request, "PLAN_KIND", "meeting-lead-reminder"
  Set SC_GM_CreateMeetingLeadReminderRequest = request
End Function

Function SC_GM_CreateLabManagerReminderRequest(record)
  Dim request
  Set request = SC_CreateGroupMeetingsMessageRequest( _
    SC_GM_RecordValue(record, "meeting_date", Date), _
    SC_GM_TYPE_LAB_MANAGER_REMINDER, _
    "lab-manager", _
    SC_GM_RequestRecipientOrFallback(SC_GM_RecordText(record, "lab_manager"), SC_GM_RecordText(record, "manager_recipient")), _
    SC_GM_BuildLabManagerReminderSubject(SC_GM_RecordValue(record, "meeting_date", Date)), _
    SC_GM_BuildLabManagerReminderBody(record))
  SC_AddWorkflowRequestMetadata request, "PLAN_KIND", "lab-manager-reminder"
  Set SC_GM_CreateLabManagerReminderRequest = request
End Function

Function SC_GM_CreateSameDayReminderRequest(record)
  Dim request
  Set request = SC_CreateGroupMeetingsMessageRequest( _
    SC_GM_RecordValue(record, "meeting_date", Date), _
    SC_GM_TYPE_SAME_DAY_REMINDER, _
    "team", _
    SC_GM_RecordText(record, "group_recipient"), _
    SC_GM_BuildSameDayReminderSubject(SC_GM_RecordValue(record, "meeting_date", Date), SC_GM_RecordText(record, "room")), _
    SC_GM_BuildSameDayReminderBody(record))
  SC_AddWorkflowRequestMetadata request, "PLAN_KIND", "same-day-reminder"
  Set SC_GM_CreateSameDayReminderRequest = request
End Function

Function SC_GM_CreateMeetingNotSetAlertRequest(record)
  Dim request
  Set request = SC_CreateGroupMeetingsMessageRequest( _
    SC_GM_RecordValue(record, "meeting_date", Date), _
    SC_GM_TYPE_MEETING_NOT_SET_ALERT, _
    "manager", _
    SC_GM_RecordText(record, "manager_recipient"), _
    "ALERT: MeetingNotSet for " & SC_GM_DisplayDate(SC_GM_RecordValue(record, "meeting_date", Date)), _
    SC_GM_BuildMeetingNotSetAlertBody(record))
  SC_AddWorkflowRequestMetadata request, "PLAN_KIND", "manager-alert"
  Set SC_GM_CreateMeetingNotSetAlertRequest = request
End Function

Function SC_GM_CreateMessagePlan(workflowContext, record)
  Dim plan
  Set plan = SC_GM_CreatePlan()
  plan("meeting_date") = SC_GM_DisplayDate(SC_GM_RecordValue(record, "meeting_date", Date))

  Dim alternate
  alternate = SC_GM_RecordText(record, "alternate_text")

  If SC_GM_IsNoMeetingEvent(alternate) Then
    plan("state") = SC_STATE_COMPOSED
    plan("reason") = "No-meeting event: regular no-meeting notice only; reminders suppressed."
    SC_GM_AddPlanMessage plan, SC_GM_CreateRegularRequest(record), "no-meeting-notice"
    SC_GM_AddPlanSkip plan, "REMINDERS_SUPPRESSED_NO_MEETING", "No-meeting event suppresses presenter, lead, lab-manager, and same-day reminders."
    Set SC_GM_CreateMessagePlan = plan
    Exit Function
  End If

  If SC_GM_IsMeetingNotSet(record) Then
    plan("state") = SC_STATE_ACTION_REQUIRED
    plan("reason") = "Meeting is not configured: manager alert planned; regular distribution suppressed."
    SC_GM_AddPlanMessage plan, SC_GM_CreateMeetingNotSetAlertRequest(record), "meeting-not-set-alert"
    SC_GM_AddPlanSkip plan, "REGULAR_SUPPRESSED_MEETING_NOT_SET", "No presenters or alternate were provided."
    Set SC_GM_CreateMessagePlan = plan
    Exit Function
  End If

  If SC_GM_RecordBool(record, "send_presenter_reminders", True) Then
    SC_GM_AddPlanMessage plan, SC_GM_CreatePresenterReminderRequest(record, "Literature"), "literature-reminder"
    SC_GM_AddPlanMessage plan, SC_GM_CreatePresenterReminderRequest(record, "Results"), "results-reminder"
  Else
    SC_GM_AddPlanSkip plan, "PRESENTER_REMINDERS_DISABLED", "Presenter reminders disabled by workflow record."
  End If

  If SC_GM_RecordBool(record, "send_meeting_lead_reminder", True) Then
    SC_GM_AddPlanMessage plan, SC_GM_CreateMeetingLeadReminderRequest(record), "meeting-lead-reminder"
  Else
    SC_GM_AddPlanSkip plan, "MEETING_LEAD_REMINDER_DISABLED", "Meeting lead reminder disabled by workflow record."
  End If

  If SC_GM_RecordBool(record, "send_lab_manager_reminder", True) Then
    SC_GM_AddPlanMessage plan, SC_GM_CreateLabManagerReminderRequest(record), "lab-manager-reminder"
  Else
    SC_GM_AddPlanSkip plan, "LAB_MANAGER_REMINDER_DISABLED", "Lab manager reminder disabled by workflow record."
  End If

  If SC_GM_RecordBool(record, "include_same_day_reminder", False) Then
    SC_GM_AddPlanMessage plan, SC_GM_CreateSameDayReminderRequest(record), "same-day-reminder"
  Else
    SC_GM_AddPlanSkip plan, "SAME_DAY_NOT_REQUESTED", "Same-day reminder is controlled by orchestration timing and was not requested."
  End If

  SC_GM_AddPlanMessage plan, SC_GM_CreateRegularRequest(record), "regular-distribution"
  plan("state") = SC_STATE_COMPOSED
  plan("reason") = "Meeting configured: reminders and regular distribution planned."
  Set SC_GM_CreateMessagePlan = plan
End Function

Function SC_GM_RunMessagePlan(workflowContext, plan, recipientDirectory, templateContext)
  Dim results
  Set results = SC_CreateDictionary()
  results("workflow_reference") = "group-meetings"
  results("planned_count") = plan("messages").Count
  results("composed_count") = 0
  results("failed_count") = 0
  Set results("dispatch_results") = SC_CreateDictionary()

  Dim key
  For Each key In plan("messages").Keys
    Dim request
    Set request = plan("messages")(key)
    Dim result
    Set result = SC_RunGroupMeetingsMessage(workflowContext, request, recipientDirectory, templateContext)
    If result("state") = SC_STATE_COMPOSED Then
      results("composed_count") = CLng(results("composed_count")) + 1
    ElseIf result("state") = SC_STATE_FAILED_SAFE Then
      results("failed_count") = CLng(results("failed_count")) + 1
    End If
    results("dispatch_results").Add CStr(results("dispatch_results").Count), result
  Next

  Set SC_GM_RunMessagePlan = results
End Function
