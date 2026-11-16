' Group meetings dry-run runner.
' Stage 2 orchestration glue: loads workbook-shaped rota data, builds the message
' plan, and runs it through the core workflow facade without live Outlook actions.

Function SC_GM_CreateRunnerResult()
  Dim result
  Set result = SC_CreateDictionary()
  result("workflow_reference") = "group-meetings"
  result("state") = SC_STATE_NOT_REQUESTED
  result("reason") = ""
  result("action_required") = False
  result("dry_run") = True
  result("target_date") = ""
  result("rota_source") = ""
  result("row_count") = 0
  result("plan_state") = ""
  result("plan_reason") = ""
  result("planned_count") = 0
  result("composed_count") = 0
  result("failed_count") = 0
  result("skipped_count") = 0
  Set result("load_result") = SC_CreateDictionary()
  Set result("plan") = SC_GM_CreatePlan()
  Set result("dispatch_results") = SC_CreateDictionary()
  Set result("skips") = SC_CreateDictionary()
  Set SC_GM_CreateRunnerResult = result
End Function

Function SC_GM_RunnerDefaults(groupRecipient, managerRecipient, senderName)
  Dim defaults
  Set defaults = SC_CreateDictionary()
  defaults("group_recipient") = CStr(groupRecipient)
  defaults("manager_recipient") = CStr(managerRecipient)
  defaults("sender_name") = CStr(senderName)
  defaults("include_same_day_reminder") = "false"
  defaults("send_presenter_reminders") = "true"
  defaults("send_meeting_lead_reminder") = "true"
  defaults("send_lab_manager_reminder") = "true"
  Set SC_GM_RunnerDefaults = defaults
End Function

Sub SC_GM_ApplyRunnerPlanSummary(result, plan)
  result("plan_state") = CStr(plan("state"))
  result("plan_reason") = CStr(plan("reason"))
  result("planned_count") = plan("messages").Count
  result("skipped_count") = plan("skips").Count
  Set result("plan") = plan
  Set result("skips") = plan("skips")
End Sub

Sub SC_GM_ApplyRunnerDispatchSummary(result, planRun)
  result("planned_count") = CLng(planRun("planned_count"))
  result("composed_count") = CLng(planRun("composed_count"))
  result("failed_count") = CLng(planRun("failed_count"))
  Set result("dispatch_results") = planRun("dispatch_results")
End Sub

Function SC_GM_RunGroupMeetingsCsvDryRun(workflowContext, csvPath, targetDate, defaults, recipientDirectory, templateContext)
  Dim result
  Set result = SC_GM_CreateRunnerResult()
  result("target_date") = SC_GM_RotaDateKey(targetDate)
  result("rota_source") = CStr(csvPath)
  result("dry_run") = SC_WorkflowDryRun(workflowContext)

  If Not CBool(result("dry_run")) Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "Group meetings scaffold runner requires dry_run=True; live Outlook actions are not enabled in this runner."
    result("action_required") = True
    Set SC_GM_RunGroupMeetingsCsvDryRun = result
    Exit Function
  End If

  Dim loadResult
  Set loadResult = SC_GM_LoadMeetingRecordFromCsv(csvPath, targetDate, defaults)
  Set result("load_result") = loadResult
  result("row_count") = CLng(loadResult("row_count"))

  If loadResult("state") <> SC_STATE_COMPOSED Then
    Dim missingPlan
    Set missingPlan = SC_GM_CreatePlan()
    missingPlan("state") = SC_STATE_ACTION_REQUIRED
    missingPlan("reason") = CStr(loadResult("reason"))
    missingPlan("meeting_date") = CStr(loadResult("target_date"))
    SC_GM_AddPlanSkip missingPlan, "ROTA_RECORD_UNAVAILABLE", CStr(loadResult("reason"))
    SC_GM_ApplyRunnerPlanSummary result, missingPlan
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = CStr(loadResult("reason"))
    result("action_required") = True
    Set SC_GM_RunGroupMeetingsCsvDryRun = result
    Exit Function
  End If

  Dim plan
  Set plan = SC_GM_CreateMessagePlan(workflowContext, loadResult("record"))
  plan("rota_source") = CStr(csvPath)
  SC_GM_ApplyRunnerPlanSummary result, plan

  If plan("messages").Count > 0 Then
    Dim planRun
    Set planRun = SC_GM_RunMessagePlan(workflowContext, plan, recipientDirectory, templateContext)
    SC_GM_ApplyRunnerDispatchSummary result, planRun
  End If

  If CLng(result("failed_count")) > 0 Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "One or more planned group-meetings messages failed safe during dry-run composition."
    result("action_required") = True
  ElseIf plan("state") = SC_STATE_ACTION_REQUIRED Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = CStr(plan("reason"))
    result("action_required") = True
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Group meetings dry-run completed."
    result("action_required") = False
  End If

  Set SC_GM_RunGroupMeetingsCsvDryRun = result
End Function

Function SC_GM_RunGroupMeetingsCsvDryRunWithDefaults(workflowContext, csvPath, targetDate, recipientDirectory, templateContext)
  Dim defaults
  Set defaults = SC_GM_RunnerDefaults("group", "manager", "Coordinator")
  Set SC_GM_RunGroupMeetingsCsvDryRunWithDefaults = _
    SC_GM_RunGroupMeetingsCsvDryRun(workflowContext, csvPath, targetDate, defaults, recipientDirectory, templateContext)
End Function
