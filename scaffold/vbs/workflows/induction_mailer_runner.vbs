' Induction mailer dry-run runner.
' Coordinates template/recipient loading, planning, and core dry-run dispatch.

Function SC_IM_CreateRunnerResult()
  Dim result
  Set result = SC_CreateDictionary()
  result("workflow_reference") = "induction-mailer"
  result("state") = SC_STATE_NOT_REQUESTED
  result("reason") = ""
  result("action_required") = False
  result("dry_run") = True
  result("selected_mode") = SC_IM_MODE_ALL
  result("recipient_count") = 0
  result("active_recipient_count") = 0
  result("template_count") = 0
  result("enabled_template_count") = 0
  result("enabled_existing_template_count") = 0
  result("planned_count") = 0
  result("composed_count") = 0
  result("failed_count") = 0
  result("skipped_count") = 0
  Set result("recipient_load") = SC_CreateDictionary()
  Set result("template_load") = SC_CreateDictionary()
  Set result("plan") = SC_IM_CreatePlan()
  Set result("dispatch_results") = SC_CreateDictionary()
  Set SC_IM_CreateRunnerResult = result
End Function

Function SC_IM_RunInductionMailerCsvDryRun(workflowContext, templateManifestPath, recipientCsvPath, selectedMode, logicalDate)
  Dim result
  Set result = SC_IM_CreateRunnerResult()
  result("dry_run") = SC_WorkflowDryRun(workflowContext)
  result("selected_mode") = CStr(selectedMode)

  If Not CBool(result("dry_run")) Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "Induction mailer scaffold runner requires dry_run=True; live Outlook actions are not enabled in this runner."
    result("action_required") = True
    Set SC_IM_RunInductionMailerCsvDryRun = result
    Exit Function
  End If

  Dim recipientLoad
  Set recipientLoad = SC_IM_LoadRecipientCsv(recipientCsvPath)
  Set result("recipient_load") = recipientLoad
  result("recipient_count") = CLng(recipientLoad("row_count"))
  result("active_recipient_count") = CLng(recipientLoad("active_count"))
  If recipientLoad("state") <> SC_STATE_COMPOSED Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = CStr(recipientLoad("reason"))
    result("action_required") = True
    Set SC_IM_RunInductionMailerCsvDryRun = result
    Exit Function
  End If

  Dim templateLoad
  Set templateLoad = SC_IM_LoadTemplateManifestCsv(templateManifestPath)
  Set result("template_load") = templateLoad
  result("template_count") = CLng(templateLoad("row_count"))
  result("enabled_template_count") = CLng(templateLoad("enabled_count"))
  result("enabled_existing_template_count") = CLng(templateLoad("enabled_existing_template_count"))
  If templateLoad("state") <> SC_STATE_COMPOSED Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = CStr(templateLoad("reason"))
    result("action_required") = True
    Set SC_IM_RunInductionMailerCsvDryRun = result
    Exit Function
  End If

  Dim plan
  Set plan = SC_IM_CreateMessagePlan(workflowContext, templateLoad("templates"), recipientLoad("recipients"), selectedMode, logicalDate)
  Set result("plan") = plan
  result("planned_count") = plan("messages").Count
  result("skipped_count") = plan("skips").Count

  If plan("messages").Count > 0 Then
    Dim directory
    Set directory = SC_IM_BuildRecipientDirectory(recipientLoad("recipients"))
    Dim templateContext
    Set templateContext = SC_CreateDictionary()
    Dim planRun
    Set planRun = SC_IM_RunMessagePlan(workflowContext, plan, directory, templateContext)
    result("planned_count") = CLng(planRun("planned_count"))
    result("composed_count") = CLng(planRun("composed_count"))
    result("failed_count") = CLng(planRun("failed_count"))
    Set result("dispatch_results") = planRun("dispatch_results")
  End If

  If CLng(result("failed_count")) > 0 Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "One or more induction messages failed safe during dry-run composition."
    result("action_required") = True
  ElseIf plan("state") = SC_STATE_ACTION_REQUIRED Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = CStr(plan("reason"))
    result("action_required") = True
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Induction mailer dry-run completed."
    result("action_required") = False
  End If

  Set SC_IM_RunInductionMailerCsvDryRun = result
End Function
