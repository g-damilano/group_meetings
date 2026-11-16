' Report-only Outlook watchdog sidecar.
' Stage 6 starts with diagnostics only. This module never repairs, deletes, sends,
' or replaces mail.

Function SC_WD_CreateSidecarResult()
  Dim result
  Set result = SC_CreateDictionary()
  result("workflow_reference") = "outlook-watchdog"
  result("state") = SC_STATE_NOT_REQUESTED
  result("reason") = ""
  result("report_only") = True
  result("repair_actions_taken") = 0
  result("action_required") = False
  result("total_count") = 0
  result("managed_count") = 0
  result("unmanaged_count") = 0
  result("queued_count") = 0
  result("waiting_deferred_count") = 0
  result("stale_count") = 0
  result("replace_eligible_count") = 0
  result("action_required_count") = 0
  result("failed_safe_count") = 0
  Set result("outbox_report") = SC_CreateOutboxInspectionReport()
  Set SC_WD_CreateSidecarResult = result
End Function

Sub SC_WD_ApplyReportSummary(result, report)
  Set result("outbox_report") = report
  result("total_count") = CLng(report("total_count"))
  result("managed_count") = CLng(report("managed_count"))
  result("unmanaged_count") = CLng(report("unmanaged_count"))
  result("queued_count") = CLng(report("queued_count"))
  result("waiting_deferred_count") = CLng(report("waiting_deferred_count"))
  result("stale_count") = CLng(report("stale_count"))
  result("replace_eligible_count") = CLng(report("replace_eligible_count"))
  result("action_required_count") = CLng(report("action_required_count"))
  result("failed_safe_count") = CLng(report("failed_safe_count"))
  result("repair_actions_taken") = CLng(report("repair_actions_taken"))

  If CLng(report("failed_safe_count")) > 0 Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "Watchdog report found uncertain managed Outbox state; no repair attempted."
    result("action_required") = True
  ElseIf CLng(report("action_required_count")) > 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Watchdog report found managed Outbox items requiring operator action."
    result("action_required") = True
  ElseIf CLng(report("stale_count")) > 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Watchdog report found stale scaffold-managed Outbox items; report-only mode did not repair."
    result("action_required") = True
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Watchdog report completed without repair."
    result("action_required") = False
  End If
End Sub

Sub SC_WD_ApplyHealthSummary(result, healthReport)
  result("total_count") = CLng(healthReport("outbox_count"))
  result("managed_count") = CLng(healthReport("managed_outbox_count"))
  result("unmanaged_count") = CLng(healthReport("unmanaged_outbox_count"))
  result("queued_count") = 0
  result("waiting_deferred_count") = 0
  result("stale_count") = 0
  result("replace_eligible_count") = 0
  result("action_required_count") = 0
  result("failed_safe_count") = 0
  result("repair_actions_taken") = 0

  If CBool(healthReport("blocked_or_auth_suspected")) Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Outlook transport requires action; report-only health inspection did not repair."
    result("action_required") = True
    result("action_required_count") = 1
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Live Outlook health inspection completed in report-only mode."
    result("action_required") = False
  End If
End Sub

Function SC_WD_RunReportOnly(outboxItems, transport, nowValue, graceMinutes, logPath)
  Dim result
  Set result = SC_WD_CreateSidecarResult()

  Dim report
  Set report = SC_InspectOutboxItems(outboxItems, transport, nowValue, graceMinutes, logPath)
  SC_WD_ApplyReportSummary result, report

  Set SC_WD_RunReportOnly = result
End Function

Function SC_WD_CreateLiveInspectionPlaceholder()
  Dim result
  Set result = SC_WD_CreateSidecarResult()
  result("state") = SC_STATE_COMPOSED
  result("reason") = "Report-only watchdog loaded; live Outlook inspection was not requested."
  Set SC_WD_CreateLiveInspectionPlaceholder = result
End Function
