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
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_workflow.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_message_plan.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_rota_adapter.vbs")
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_runner.vbs")

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

Function CreateContext()
  Dim context
  Set context = SC_CreateGroupMeetingsWorkflowContext()
  context("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_group_runner_smoke.log")
  Set CreateContext = context
End Function

Function CreateDirectory()
  Dim directory
  Set directory = SC_CreateDictionary()
  directory("group") = "group@example.invalid"
  directory("manager") = "manager@example.invalid"
  directory("Literature Presenter") = "literature.presenter@example.invalid"
  directory("Results Presenter") = "results.presenter@example.invalid"
  directory("Previous Lab Manager") = "meeting.lead@example.invalid"
  directory("Lab Manager") = "lab.manager@example.invalid"
  directory("Future Lab Manager") = "future.lab.manager@example.invalid"
  Set CreateDirectory = directory
End Function

Dim fixturePath
fixturePath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_rota_fixture.csv")

Dim defaults
Set defaults = SC_GM_RunnerDefaults("group", "manager", "Coordinator")

Dim templateContext
Set templateContext = SC_CreateDictionary()

Dim context
Set context = CreateContext()
Dim directory
Set directory = CreateDirectory()

Dim configuredRun
Set configuredRun = SC_GM_RunGroupMeetingsCsvDryRun(context, fixturePath, #5/29/2026#, defaults, directory, templateContext)
AssertTrue configuredRun("state") = SC_STATE_COMPOSED, "Configured rota row dry-runs as composed"
AssertTrue configuredRun("row_count") = 4, "Runner reports rota row count"
AssertTrue configuredRun("planned_count") = 6, "Configured runner plans six messages"
AssertTrue configuredRun("composed_count") = 6, "Configured runner dry-runs all messages"
AssertTrue configuredRun("failed_count") = 0, "Configured runner has no failed-safe messages"
AssertTrue configuredRun("skipped_count") = 0, "Configured runner has no skipped planning branch"
AssertTrue configuredRun("dispatch_results")("5")("prepared_to") = "group@example.invalid", "Runner resolves group recipient"
AssertTrue Not CBool(configuredRun("dispatch_results")("5")("created_new_mail")), "Runner creates no Outlook mail"
AssertTrue configuredRun("plan")("messages")("5")("message_type") = SC_GM_TYPE_REGULAR, "Runner preserves plan message order"

Dim defaultRun
Set defaultRun = SC_GM_RunGroupMeetingsCsvDryRunWithDefaults(context, fixturePath, #5/29/2026#, directory, templateContext)
AssertTrue defaultRun("state") = SC_STATE_COMPOSED, "Default runner convenience path composes"
AssertTrue defaultRun("planned_count") = 6, "Default runner convenience path plans configured row"

Dim noMeetingRun
Set noMeetingRun = SC_GM_RunGroupMeetingsCsvDryRun(context, fixturePath, #6/5/2026#, defaults, directory, templateContext)
AssertTrue noMeetingRun("state") = SC_STATE_COMPOSED, "No-meeting rota row still composes"
AssertTrue noMeetingRun("planned_count") = 1, "No-meeting runner plans one notice"
AssertTrue noMeetingRun("composed_count") = 1, "No-meeting runner dry-runs one notice"
AssertTrue noMeetingRun("skipped_count") = 1, "No-meeting runner reports suppressed reminders"
AssertTrue InStr(1, noMeetingRun("plan_reason"), "No-meeting", vbTextCompare) > 0, "No-meeting runner keeps plan reason"

Dim incompleteRun
Set incompleteRun = SC_GM_RunGroupMeetingsCsvDryRun(context, fixturePath, #6/12/2026#, defaults, directory, templateContext)
AssertTrue incompleteRun("state") = SC_STATE_ACTION_REQUIRED, "Incomplete rota row becomes action-required"
AssertTrue CBool(incompleteRun("action_required")), "Incomplete rota row flags action required"
AssertTrue incompleteRun("planned_count") = 1, "Incomplete rota row plans one manager alert"
AssertTrue incompleteRun("composed_count") = 1, "Incomplete manager alert remains dry-run composed"
AssertTrue incompleteRun("dispatch_results")("0")("prepared_to") = "manager@example.invalid", "Incomplete rota alert resolves manager recipient"

Dim missingRun
Set missingRun = SC_GM_RunGroupMeetingsCsvDryRun(context, fixturePath, #7/3/2026#, defaults, directory, templateContext)
AssertTrue missingRun("state") = SC_STATE_ACTION_REQUIRED, "Missing target row becomes action-required"
AssertTrue CBool(missingRun("action_required")), "Missing target row flags action required"
AssertTrue missingRun("planned_count") = 0, "Missing target row plans no messages"
AssertTrue missingRun("composed_count") = 0, "Missing target row composes no messages"
AssertTrue missingRun("skipped_count") = 1, "Missing target row records skip"

Dim liveContext
Set liveContext = CreateContext()
liveContext("dry_run") = False
Dim liveBlockedRun
Set liveBlockedRun = SC_GM_RunGroupMeetingsCsvDryRun(liveContext, fixturePath, #5/29/2026#, defaults, directory, templateContext)
AssertTrue liveBlockedRun("state") = SC_STATE_FAILED_SAFE, "Runner blocks non-dry-run context"
AssertTrue CBool(liveBlockedRun("action_required")), "Blocked live context flags action required"
AssertTrue liveBlockedRun("planned_count") = 0, "Blocked live context does not plan messages"
AssertTrue liveBlockedRun("row_count") = 0, "Blocked live context does not read rota"

Dim runnerPath
runnerPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_runner.vbs")
AssertTrue Not FileContains(runnerPath, "CreateObject(""Excel.Application"")"), "Runner owns no Excel application creation"
AssertTrue Not FileContains(runnerPath, "CreateObject(""Outlook.Application"")"), "Runner owns no Outlook application creation"
AssertTrue Not FileContains(runnerPath, "GetDefaultFolder"), "Runner owns no Outlook folder access"
AssertTrue Not FileContains(runnerPath, ".Send"), "Runner owns no Outlook send call"
AssertTrue Not FileContains(runnerPath, ".Delete"), "Runner owns no Outlook delete call"
AssertTrue Not FileContains(runnerPath, "sharepoint.com"), "Runner contains no private schedule links"

If failures > 0 Then
  WScript.Echo "Group meetings runner smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Group meetings runner smoke test passed."
