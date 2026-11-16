Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

Sub LoadTextModule(modulePath)
  Dim ts
  Set ts = fso.OpenTextFile(modulePath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
SC_LoadCoreModules
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\workflow_helpers.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_workflow.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_input_adapter.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_message_plan.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_runner.vbs")

Dim failures
failures = 0

Sub AssertTrue(condition, message)
  If condition Then
    WScript.Echo "OK: " & message
  Else
    WScript.Echo "FAILED: " & message
    failures = failures + 1
  End If
End Sub

Function FileContains(path, token)
  FileContains = (InStr(1, SC_ReadTextFile(path), CStr(token), vbTextCompare) > 0)
End Function

Dim recipientsPath
recipientsPath = fso.BuildPath(scriptFolder, "fixtures\induction_recipients_fixture.csv")
Dim invalidRecipientsPath
invalidRecipientsPath = fso.BuildPath(scriptFolder, "fixtures\induction_recipients_invalid_fixture.csv")
Dim manifestPath
manifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_fixture.csv")
Dim missingManifestPath
missingManifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_missing_template_fixture.csv")
Dim invalidManifestPath
invalidManifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_invalid_fixture.csv")

Dim recipientLoad
Set recipientLoad = SC_IM_LoadRecipientCsv(recipientsPath)
AssertTrue recipientLoad("state") = SC_STATE_COMPOSED, "Recipient fixture loads"
AssertTrue recipientLoad("row_count") = 4, "Recipient fixture reports all rows"
AssertTrue recipientLoad("active_count") = 2, "Recipient fixture reports active recipients"
AssertTrue recipientLoad("invalid_count") = 0, "Recipient fixture reports no invalid rows"
AssertTrue recipientLoad("recipients")("0")("full_name") = "New Starter", "Recipient full name mapped"

Dim displayNameRow
Set displayNameRow = SC_CreateDictionary()
displayNameRow("alias") = "display-only"
displayNameRow("email") = "Example-Family, Casey <casey.family@example.invalid>"
displayNameRow("first_name") = ""
displayNameRow("last_name") = ""
displayNameRow("full_name") = ""
displayNameRow("active") = "true"
Dim displayNameRecord
Set displayNameRecord = SC_IM_CreateRecipientRecord(displayNameRow)
AssertTrue displayNameRecord("first_name") = "Casey", "Induction recipient display fallback fills first name"
AssertTrue displayNameRecord("surname") = "Example-Family", "Induction recipient display fallback fills surname"
AssertTrue displayNameRecord("full_name") = "Casey Example-Family", "Induction recipient display fallback fills full name"
AssertTrue displayNameRecord("name_parse_source") = "email-display", "Induction recipient records display parse source"

Dim structuredNameRow
Set structuredNameRow = SC_CreateDictionary()
structuredNameRow("alias") = "structured"
structuredNameRow("email") = "Display Fallback <structured.person@example.invalid>"
structuredNameRow("first_name") = "Structured"
structuredNameRow("last_name") = "Person"
structuredNameRow("display_name") = "Display Fallback"
structuredNameRow("full_name") = ""
structuredNameRow("active") = "true"
Dim structuredNameRecord
Set structuredNameRecord = SC_IM_CreateRecipientRecord(structuredNameRow)
AssertTrue structuredNameRecord("first_name") = "Structured", "Structured first name wins over display fallback"
AssertTrue structuredNameRecord("surname") = "Person", "Structured surname wins over display fallback"
AssertTrue structuredNameRecord("full_name") = "Structured Person", "Structured full name wins over display fallback"

Dim invalidRecipientLoad
Set invalidRecipientLoad = SC_IM_LoadRecipientCsv(invalidRecipientsPath)
AssertTrue invalidRecipientLoad("state") = SC_STATE_ACTION_REQUIRED, "Invalid recipient fixture requires action"
AssertTrue invalidRecipientLoad("invalid_count") = 1, "Invalid recipient fixture counts missing active email"
AssertTrue invalidRecipientLoad("invalid_rows")("0")("code") = "ACTIVE_RECIPIENT_MISSING_EMAIL", "Invalid recipient fixture reports missing active email"

Dim templateLoad
Set templateLoad = SC_IM_LoadTemplateManifestCsv(manifestPath)
AssertTrue templateLoad("state") = SC_STATE_COMPOSED, "Template manifest fixture loads"
AssertTrue templateLoad("row_count") = 3, "Template manifest reports all rows"
AssertTrue templateLoad("enabled_count") = 2, "Template manifest reports enabled templates"
AssertTrue templateLoad("enabled_existing_template_count") = 2, "Template manifest reports enabled templates with existing files"
AssertTrue templateLoad("invalid_count") = 0, "Template manifest reports no invalid rows"
AssertTrue fso.FileExists(templateLoad("templates")("0")("template_path")), "Template path resolves relative to manifest"
AssertTrue templateLoad("templates")("0")("display_name") = "Overview broadcast", "Template display name mapped"
AssertTrue fso.FolderExists(templateLoad("templates")("0")("attachment_base_folder")), "Attachment base folder resolves"
AssertTrue templateLoad("templates")("0")("attachments") = "overview.txt;setup checklist.txt", "Template attachments mapped"
AssertTrue fso.FolderExists(templateLoad("templates")("1")("inline_image_folder")), "Inline image folder resolves"

Dim directory
Set directory = SC_IM_BuildRecipientDirectory(recipientLoad("recipients"))
AssertTrue directory("new-starter") = "new.starter@example.invalid", "Recipient directory maps aliases"
AssertTrue directory("New Starter") = "new.starter@example.invalid", "Recipient directory maps display names"
AssertTrue directory("coordinator") = "coordinator@example.invalid", "Recipient directory maps inactive support aliases"

Dim context
Set context = SC_CreateInductionMailerWorkflowContext()
context("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_induction_stage3_smoke.log")

Dim plan
Set plan = SC_IM_CreateMessagePlan(context, templateLoad("templates"), recipientLoad("recipients"), SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue plan("state") = SC_STATE_COMPOSED, "Induction all-mode plan composes"
AssertTrue plan("messages").Count = 3, "All-mode plan has one broadcast and two personal messages"
AssertTrue plan("skips").Count = 1, "Disabled template is recorded as a skip"
AssertTrue plan("messages")("0")("message_type") = SC_IM_MODE_BROADCAST, "Broadcast template planned first"
AssertTrue plan("messages")("1")("message_type") = SC_IM_MODE_PERSONAL, "Personal template planned for active recipient"
AssertTrue plan("messages")("0")("cc") = "coordinator", "Broadcast template carries CC aliases"
AssertTrue plan("messages")("0")("defer_minutes") = "15", "Broadcast template carries defer minutes"
AssertTrue plan("messages")("1")("inline_image_folder") <> "", "Personal template carries inline image folder"

Dim missingTemplateLoad
Set missingTemplateLoad = SC_IM_LoadTemplateManifestCsv(missingManifestPath)
AssertTrue missingTemplateLoad("state") = SC_STATE_COMPOSED, "Missing-template manifest loads for planning diagnostics"
AssertTrue missingTemplateLoad("enabled_count") = 1, "Missing-template manifest reports enabled row"
AssertTrue missingTemplateLoad("enabled_existing_template_count") = 0, "Missing-template manifest detects no existing enabled templates"
AssertTrue Not CBool(missingTemplateLoad("templates")("0")("template_exists")), "Missing template record is marked missing"

Dim invalidTemplateLoad
Set invalidTemplateLoad = SC_IM_LoadTemplateManifestCsv(invalidManifestPath)
AssertTrue invalidTemplateLoad("state") = SC_STATE_ACTION_REQUIRED, "Invalid template manifest requires action"
AssertTrue invalidTemplateLoad("invalid_count") = 3, "Invalid template manifest counts invalid enabled rows"
AssertTrue invalidTemplateLoad("invalid_rows")("0")("code") = "TEMPLATE_MODE_INVALID", "Invalid template manifest reports bad mode"
AssertTrue invalidTemplateLoad("invalid_rows")("1")("code") = "DELIVERY_ACTION_INVALID", "Invalid template manifest reports bad delivery action"
AssertTrue invalidTemplateLoad("invalid_rows")("2")("code") = "DEFER_MINUTES_INVALID", "Invalid template manifest reports bad defer minutes"

Dim missingPlan
Set missingPlan = SC_IM_CreateMessagePlan(context, missingTemplateLoad("templates"), recipientLoad("recipients"), SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue missingPlan("state") = SC_STATE_ACTION_REQUIRED, "Missing template becomes action-required plan"
AssertTrue missingPlan("messages").Count = 0, "Missing template does not fabricate messages"
AssertTrue missingPlan("skips").Count = 1, "Missing template is recorded as a skip"
AssertTrue missingPlan("skips")("0")("code") = "TEMPLATE_MISSING", "Missing template skip code is explicit"

Dim templateContext
Set templateContext = SC_CreateDictionary()
Dim planRun
Set planRun = SC_IM_RunMessagePlan(context, plan, directory, templateContext)
AssertTrue planRun("planned_count") = 3, "Plan run sees all induction messages"
AssertTrue planRun("composed_count") = 3, "Plan run remains dry-run composed"
AssertTrue planRun("failed_count") = 0, "Plan run has no failed-safe result"
AssertTrue planRun("dispatch_results")("0")("prepared_to") = "new.starter@example.invalid; observer@example.invalid", "Broadcast resolves active recipients"
AssertTrue planRun("dispatch_results")("0")("prepared_cc") = "coordinator@example.invalid", "Broadcast resolves support CC recipient"
AssertTrue CLng(planRun("dispatch_results")("0")("prepared_attachment_count")) = 2, "Broadcast resolves attachments through core"
AssertTrue CLng(planRun("dispatch_results")("0")("prepared_defer_minutes")) = 15, "Broadcast applies per-template defer policy"
AssertTrue planRun("dispatch_results")("1")("prepared_to") = "new.starter@example.invalid", "Personal message resolves first active recipient"
AssertTrue InStr(1, planRun("dispatch_results")("1")("prepared_subject"), "Welcome New", vbTextCompare) > 0, "Personal subject receives recipient context"
AssertTrue CLng(planRun("dispatch_results")("1")("prepared_inline_count")) = 1, "Personal message resolves inline image through core"
AssertTrue Not CBool(planRun("dispatch_results")("0")("created_new_mail")), "Dry-run creates no Outlook mail"

Dim runnerResult
Set runnerResult = SC_IM_RunInductionMailerCsvDryRun(context, manifestPath, recipientsPath, SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue runnerResult("state") = SC_STATE_COMPOSED, "Induction runner composes"
AssertTrue runnerResult("enabled_existing_template_count") = 2, "Induction runner reports existing enabled template count"
AssertTrue runnerResult("planned_count") = 3, "Induction runner reports planned count"
AssertTrue runnerResult("composed_count") = 3, "Induction runner dry-runs all messages"
AssertTrue runnerResult("failed_count") = 0, "Induction runner reports no failed-safe messages"
AssertTrue runnerResult("skipped_count") = 1, "Induction runner reports disabled-template skip"

Dim broadcastOnly
Set broadcastOnly = SC_IM_RunInductionMailerCsvDryRun(context, manifestPath, recipientsPath, SC_IM_MODE_BROADCAST, #6/1/2026#)
AssertTrue broadcastOnly("planned_count") = 1, "Broadcast-only runner plans one message"
AssertTrue broadcastOnly("composed_count") = 1, "Broadcast-only runner composes one message"

Dim missingRunnerResult
Set missingRunnerResult = SC_IM_RunInductionMailerCsvDryRun(context, missingManifestPath, recipientsPath, SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue missingRunnerResult("state") = SC_STATE_ACTION_REQUIRED, "Missing-template runner is action-required"
AssertTrue CBool(missingRunnerResult("action_required")), "Missing-template runner flags action required"
AssertTrue missingRunnerResult("planned_count") = 0, "Missing-template runner plans no messages"
AssertTrue missingRunnerResult("skipped_count") = 1, "Missing-template runner reports skip"

Dim invalidRecipientRunner
Set invalidRecipientRunner = SC_IM_RunInductionMailerCsvDryRun(context, manifestPath, invalidRecipientsPath, SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue invalidRecipientRunner("state") = SC_STATE_ACTION_REQUIRED, "Invalid-recipient runner is action-required"
AssertTrue invalidRecipientRunner("planned_count") = 0, "Invalid-recipient runner does not plan messages"
AssertTrue invalidRecipientRunner("recipient_load")("invalid_count") = 1, "Invalid-recipient runner preserves invalid count"

Dim invalidTemplateRunner
Set invalidTemplateRunner = SC_IM_RunInductionMailerCsvDryRun(context, invalidManifestPath, recipientsPath, SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue invalidTemplateRunner("state") = SC_STATE_ACTION_REQUIRED, "Invalid-template runner is action-required"
AssertTrue invalidTemplateRunner("planned_count") = 0, "Invalid-template runner does not plan messages"
AssertTrue invalidTemplateRunner("template_load")("invalid_count") = 3, "Invalid-template runner preserves invalid count"

Dim liveContext
Set liveContext = SC_CreateInductionMailerWorkflowContext()
liveContext("dry_run") = False
Dim blockedRun
Set blockedRun = SC_IM_RunInductionMailerCsvDryRun(liveContext, manifestPath, recipientsPath, SC_IM_MODE_ALL, #6/1/2026#)
AssertTrue blockedRun("state") = SC_STATE_FAILED_SAFE, "Induction runner blocks non-dry-run context"
AssertTrue CBool(blockedRun("action_required")), "Blocked induction runner flags action required"
AssertTrue blockedRun("planned_count") = 0, "Blocked induction runner does not plan messages"

Dim adapterPath
adapterPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_input_adapter.vbs")
Dim plannerPath
plannerPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_message_plan.vbs")
Dim runnerPath
runnerPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_runner.vbs")
AssertTrue Not FileContains(adapterPath, "CreateObject(""Outlook.Application"")"), "Induction adapter owns no Outlook application creation"
AssertTrue Not FileContains(plannerPath, "CreateObject(""Outlook.Application"")"), "Induction planner owns no Outlook application creation"
AssertTrue Not FileContains(runnerPath, "CreateObject(""Outlook.Application"")"), "Induction runner owns no Outlook application creation"
AssertTrue Not FileContains(plannerPath, "GetDefaultFolder"), "Induction planner owns no Outlook folder access"
AssertTrue Not FileContains(runnerPath, ".Send"), "Induction runner owns no Outlook send call"
AssertTrue Not FileContains(runnerPath, ".Delete"), "Induction runner owns no Outlook delete call"
AssertTrue Not FileContains(runnerPath, "sharepoint.com"), "Induction runner contains no private schedule links"
AssertTrue Not FileContains(plannerPath, "Attachments.Add"), "Induction planner delegates attachment handling to core"

If failures > 0 Then
  WScript.Echo "Induction mailer Stage 3 smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Induction mailer Stage 3 smoke test passed."
