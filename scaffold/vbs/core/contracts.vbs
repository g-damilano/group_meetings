' Core contract helpers for the Office communication automation scaffold.
' This module is loaded with ExecuteGlobal by test and entry scripts.

Const SC_STATE_NOT_REQUESTED = "NOT_REQUESTED"
Const SC_STATE_COMPOSED = "COMPOSED"
Const SC_STATE_DISPLAYED = "DISPLAYED"
Const SC_STATE_DRAFTED = "DRAFTED"
Const SC_STATE_SEND_REQUESTED = "SEND_REQUESTED"
Const SC_STATE_QUEUED_OUTBOX = "QUEUED_OUTBOX"
Const SC_STATE_SENT_CONFIRMED = "SENT_CONFIRMED"
Const SC_STATE_DUPLICATE_SUPPRESSED = "DUPLICATE_SUPPRESSED"
Const SC_STATE_WAITING_DEFERRED = "WAITING_DEFERRED"
Const SC_STATE_BLOCKED_OFFLINE = "BLOCKED_OFFLINE"
Const SC_STATE_BLOCKED_AUTH = "BLOCKED_AUTH"
Const SC_STATE_STALE_MANAGED_OUTBOX = "STALE_MANAGED_OUTBOX"
Const SC_STATE_REPLACED_MANAGED_OUTBOX = "REPLACED_MANAGED_OUTBOX"
Const SC_STATE_ACTION_REQUIRED = "ACTION_REQUIRED"
Const SC_STATE_FAILED_SAFE = "FAILED_SAFE"

Const SC_MANAGED_KEY_PROPERTY = "OfficeScaffoldManagedEmailKey"
Const SC_MANAGED_KEY_MARKER = "OfficeScaffoldManagedEmailKey"
Const SC_DEFAULT_OUTBOX_GRACE_MINUTES = 15

Function SC_CreateDictionary()
  Dim dictionary
  Set dictionary = CreateObject("Scripting.Dictionary")
  Set SC_CreateDictionary = dictionary
End Function

Function SC_CreateMessageSpec(messageKey, messageType, toAddress, subjectText, bodyText)
  Dim spec
  Set spec = SC_CreateDictionary()
  spec("message_key") = CStr(messageKey)
  spec("message_type") = CStr(messageType)
  spec("to") = CStr(toAddress)
  spec("cc") = ""
  spec("bcc") = ""
  spec("subject") = CStr(subjectText)
  spec("body_text") = CStr(bodyText)
  spec("body_html") = ""
  spec("inline_image_folder") = ""
  Set spec("attachments") = SC_CreateDictionary()
  Set spec("metadata") = SC_CreateDictionary()
  Set SC_CreateMessageSpec = spec
End Function

Function SC_CreateRecipientSpec(rawValue)
  Dim spec
  Set spec = SC_CreateDictionary()
  spec("raw_value") = CStr(rawValue)
  spec("display_name") = ""
  spec("first_name") = ""
  spec("last_name") = ""
  spec("surname") = ""
  spec("full_name") = ""
  spec("name_parse_status") = "not_attempted"
  spec("name_parse_source") = ""
  spec("name_parse_order") = ""
  spec("smtp_address") = ""
  spec("resolution_status") = "unresolved"
  spec("directory_source") = ""
  spec("fallback_used") = False
  Set SC_CreateRecipientSpec = spec
End Function

Function SC_CreateDeliveryPolicy(action)
  Dim policy
  Set policy = SC_CreateDictionary()
  policy("action") = LCase(Trim(CStr(action)))
  If Len(policy("action")) = 0 Then policy("action") = "display"
  policy("defer_minutes") = 0
  policy("urgent_verify_after_submit") = False
  policy("managed_dispatch_required") = True
  policy("enable_stale_outbox_replacement") = False
  Set SC_CreateDeliveryPolicy = policy
End Function

Function SC_CreateDispatchResult(stateText, reasonText)
  Dim result
  Set result = SC_CreateDictionary()
  result("state") = CStr(stateText)
  result("reason") = CStr(reasonText)
  result("managed_key") = ""
  result("folder") = ""
  result("outlook_entry_id") = ""
  result("action_required") = False
  result("created_new_mail") = False
  result("replaced_existing_mail") = False
  result("safe_to_retry") = False
  Set SC_CreateDispatchResult = result
End Function

Function SC_CreateWorkflowContext(workflowName)
  Dim context
  Set context = SC_CreateDictionary()
  context("workflow_name") = CStr(workflowName)
  context("run_id") = SC_TimestampId()
  context("source_workbook") = ""
  context("template_root") = ""
  context("log_path") = ""
  context("dry_run") = True
  context("operator") = ""
  Set context("settings") = SC_CreateDictionary()
  Set SC_CreateWorkflowContext = context
End Function

Function SC_CreateHealthReport()
  Dim report
  Set report = SC_CreateDictionary()
  report("offline") = ""
  report("exchange_connection_mode") = ""
  report("exchange_connection_mode_name") = ""
  report("blocked_or_auth_suspected") = False
  report("outbox_count") = 0
  report("managed_outbox_count") = 0
  report("unmanaged_outbox_count") = 0
  Set report("action_required_messages") = SC_CreateDictionary()
  Set SC_CreateHealthReport = report
End Function

Function SC_TimestampId()
  SC_TimestampId = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "-" & _
                   Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
End Function

Function SC_NormalizeKeyPart(value)
  Dim text
  text = UCase(Trim(CStr(value)))
  text = Replace(text, " ", "-")
  text = Replace(text, "_", "-")
  Do While InStr(text, "--") > 0
    text = Replace(text, "--", "-")
  Loop
  SC_NormalizeKeyPart = text
End Function

Function SC_FormatKeyDate(value)
  If IsDate(value) Then
    SC_FormatKeyDate = Year(CDate(value)) & "-" & Right("0" & Month(CDate(value)), 2) & "-" & Right("0" & Day(CDate(value)), 2)
  Else
    SC_FormatKeyDate = SC_NormalizeKeyPart(value)
  End If
End Function

Function SC_BuildManagedEmailKey(orgCode, workflowName, logicalDate, messageType, targetKey)
  SC_BuildManagedEmailKey = SC_NormalizeKeyPart(orgCode) & "|" & _
    SC_NormalizeKeyPart(workflowName) & "|" & _
    SC_FormatKeyDate(logicalDate) & "|" & _
    SC_NormalizeKeyPart(messageType) & "|" & _
    SC_NormalizeKeyPart(targetKey)
End Function

Function SC_IsKnownDispatchState(stateText)
  Select Case UCase(Trim(CStr(stateText)))
    Case SC_STATE_NOT_REQUESTED, SC_STATE_COMPOSED, SC_STATE_DISPLAYED, SC_STATE_DRAFTED, _
         SC_STATE_SEND_REQUESTED, SC_STATE_QUEUED_OUTBOX, SC_STATE_SENT_CONFIRMED, _
         SC_STATE_DUPLICATE_SUPPRESSED, SC_STATE_WAITING_DEFERRED, SC_STATE_BLOCKED_OFFLINE, _
         SC_STATE_BLOCKED_AUTH, SC_STATE_STALE_MANAGED_OUTBOX, SC_STATE_REPLACED_MANAGED_OUTBOX, _
         SC_STATE_ACTION_REQUIRED, SC_STATE_FAILED_SAFE
      SC_IsKnownDispatchState = True
    Case Else
      SC_IsKnownDispatchState = False
  End Select
End Function
