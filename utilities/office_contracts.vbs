'====================================================================
' office_contracts.vbs - Canonical contract helpers for Office workflows
'====================================================================
Option Explicit
SetLocale 1033

Const OFFICE_STATE_NOT_REQUESTED = "NOT_REQUESTED"
Const OFFICE_STATE_COMPOSED = "COMPOSED"
Const OFFICE_STATE_DISPLAYED = "DISPLAYED"
Const OFFICE_STATE_DRAFTED = "DRAFTED"
Const OFFICE_STATE_SEND_REQUESTED = "SEND_REQUESTED"
Const OFFICE_STATE_QUEUED_OUTBOX = "QUEUED_OUTBOX"
Const OFFICE_STATE_SENT_CONFIRMED = "SENT_CONFIRMED"
Const OFFICE_STATE_DUPLICATE_SUPPRESSED = "DUPLICATE_SUPPRESSED"
Const OFFICE_STATE_WAITING_DEFERRED = "WAITING_DEFERRED"
Const OFFICE_STATE_BLOCKED_OFFLINE = "BLOCKED_OFFLINE"
Const OFFICE_STATE_BLOCKED_AUTH = "BLOCKED_AUTH"
Const OFFICE_STATE_STALE_MANAGED_OUTBOX = "STALE_MANAGED_OUTBOX"
Const OFFICE_STATE_REPLACED_MANAGED_OUTBOX = "REPLACED_MANAGED_OUTBOX"
Const OFFICE_STATE_ACTION_REQUIRED = "ACTION_REQUIRED"
Const OFFICE_STATE_FAILED_SAFE = "FAILED_SAFE"

Const OFFICE_MANAGED_KEY_PROPERTY = "ManagedEmailKey"
Const OFFICE_MANAGED_KEY_MARKER = "ManagedEmailKey"

Function OfficeCreateDictionary()
  Set OfficeCreateDictionary = CreateObject("Scripting.Dictionary")
End Function

Function OfficeCreateMessageSpec(messageKey, messageType, toAddress, subjectText, bodyText)
  Dim spec: Set spec = OfficeCreateDictionary()
  spec("message_key") = CStr(messageKey)
  spec("message_type") = CStr(messageType)
  spec("to") = CStr(toAddress)
  spec("cc") = ""
  spec("bcc") = ""
  spec("subject") = CStr(subjectText)
  spec("body_text") = CStr(bodyText)
  spec("body_html") = ""
  spec("delivery_policy") = "display"
  spec("urgent_verify_after_submit") = False
  Set spec("attachments") = OfficeCreateDictionary()
  Set spec("metadata") = OfficeCreateDictionary()
  Set OfficeCreateMessageSpec = spec
End Function

Function OfficeCreateRecipientSpec(rawValue)
  Dim spec: Set spec = OfficeCreateDictionary()
  spec("raw_value") = CStr(rawValue)
  spec("display_name") = ""
  spec("smtp_address") = ""
  spec("resolution_status") = "unresolved"
  spec("directory_source") = ""
  spec("fallback_used") = False
  Set OfficeCreateRecipientSpec = spec
End Function

Function OfficeCreateDeliveryPolicy(action)
  Dim policy: Set policy = OfficeCreateDictionary()
  policy("action") = LCase(Trim(CStr(action)))
  If Len(policy("action")) = 0 Then policy("action") = "display"
  policy("defer_minutes") = 0
  policy("urgent_verify_after_submit") = False
  policy("managed_dispatch_required") = True
  Set OfficeCreateDeliveryPolicy = policy
End Function

Function OfficeCreateDispatchResult(state, reason)
  Dim result: Set result = OfficeCreateDictionary()
  result("state") = CStr(state)
  result("reason") = CStr(reason)
  result("managed_key") = ""
  result("folder") = ""
  result("outlook_entry_id") = ""
  result("action_required") = False
  result("created_new_mail") = False
  result("replaced_existing_mail") = False
  result("safe_to_retry") = False
  Set OfficeCreateDispatchResult = result
End Function

Function OfficeCreateWorkflowContext(workflowName)
  Dim context: Set context = OfficeCreateDictionary()
  context("workflow_name") = CStr(workflowName)
  context("run_id") = OfficeTimestampId()
  context("source_workbook") = ""
  context("template_root") = ""
  context("log_path") = ""
  context("dry_run") = False
  context("operator") = ""
  Set context("settings") = OfficeCreateDictionary()
  Set OfficeCreateWorkflowContext = context
End Function

Function OfficeCreateHealthReport()
  Dim report: Set report = OfficeCreateDictionary()
  report("offline") = ""
  report("exchange_connection_mode") = ""
  report("exchange_connection_mode_name") = ""
  report("blocked_or_auth_suspected") = False
  report("outbox_count") = 0
  report("managed_outbox_count") = 0
  report("unmanaged_outbox_count") = 0
  Set report("action_required_messages") = OfficeCreateDictionary()
  Set OfficeCreateHealthReport = report
End Function

Function OfficeTimestampId()
  OfficeTimestampId = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "-" & _
                      Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
End Function

Function OfficeNormalizeKeyPart(value)
  Dim text: text = UCase(Trim(CStr(value)))
  text = Replace(text, " ", "-")
  text = Replace(text, "_", "-")
  Do While InStr(text, "--") > 0
    text = Replace(text, "--", "-")
  Loop
  OfficeNormalizeKeyPart = text
End Function

Function OfficeBuildManagedEmailKey(orgCode, workflowName, logicalDate, messageType, targetKey)
  OfficeBuildManagedEmailKey = OfficeNormalizeKeyPart(orgCode) & "|" & _
    OfficeNormalizeKeyPart(workflowName) & "|" & _
    OfficeFormatKeyDate(logicalDate) & "|" & _
    OfficeNormalizeKeyPart(messageType) & "|" & _
    OfficeNormalizeKeyPart(targetKey)
End Function

Function OfficeFormatKeyDate(value)
  If IsDate(value) Then
    OfficeFormatKeyDate = Year(CDate(value)) & "-" & Right("0" & Month(CDate(value)), 2) & "-" & Right("0" & Day(CDate(value)), 2)
  Else
    OfficeFormatKeyDate = OfficeNormalizeKeyPart(value)
  End If
End Function

Function OfficeIsKnownDispatchState(stateText)
  Select Case UCase(Trim(CStr(stateText)))
    Case OFFICE_STATE_NOT_REQUESTED, OFFICE_STATE_COMPOSED, OFFICE_STATE_DISPLAYED, OFFICE_STATE_DRAFTED, _
         OFFICE_STATE_SEND_REQUESTED, OFFICE_STATE_QUEUED_OUTBOX, OFFICE_STATE_SENT_CONFIRMED, _
         OFFICE_STATE_DUPLICATE_SUPPRESSED, OFFICE_STATE_WAITING_DEFERRED, OFFICE_STATE_BLOCKED_OFFLINE, _
         OFFICE_STATE_BLOCKED_AUTH, OFFICE_STATE_STALE_MANAGED_OUTBOX, OFFICE_STATE_REPLACED_MANAGED_OUTBOX, _
         OFFICE_STATE_ACTION_REQUIRED, OFFICE_STATE_FAILED_SAFE
      OfficeIsKnownDispatchState = True
    Case Else
      OfficeIsKnownDispatchState = False
  End Select
End Function

