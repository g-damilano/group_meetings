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

Dim fixturePath
fixturePath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_rota_fixture.csv")

Dim rows
Set rows = SC_GM_ReadRotaCsv(fixturePath)
AssertTrue rows.Count = 4, "CSV rota fixture loads four rows"
AssertTrue SC_GM_RotaRowText(rows("1"), "room") = "Room A, East", "CSV parser preserves quoted comma field"
AssertTrue SC_GM_NormalizeRotaHeader("Lab Manager") = "lab_manager", "Header normalization maps lab manager"

Dim defaults
Set defaults = SC_CreateDictionary()
defaults("group_recipient") = "group"
defaults("manager_recipient") = "manager"
defaults("sender_name") = "Coordinator"

Dim loadResult
Set loadResult = SC_GM_LoadMeetingRecordFromCsv(fixturePath, #5/29/2026#, defaults)
AssertTrue loadResult("state") = SC_STATE_COMPOSED, "Target meeting record loads"
AssertTrue loadResult("row_count") = 4, "Load result records row count"

Dim record
Set record = loadResult("record")
AssertTrue SC_GM_DisplayDate(record("meeting_date")) = "2026-05-29", "Meeting date normalized"
AssertTrue record("literature_presenter") = "Literature Presenter", "Literature presenter mapped"
AssertTrue record("results_presenter") = "Results Presenter", "Results presenter mapped"
AssertTrue record("room") = "Room A, East", "Room mapped"
AssertTrue record("meeting_lead") = "Previous Lab Manager", "Blank meeting lead derives from previous lab manager"
AssertTrue CBool(record("include_same_day_reminder")), "Same-day reminder flag parsed"
AssertTrue SC_GM_DisplayDate(record("next_meeting_date")) = "2026-06-05", "Next meeting date mapped"
AssertTrue record("next_alternate_text") = "no meeting", "Next alternate text mapped"

Dim context
Set context = SC_CreateGroupMeetingsWorkflowContext()
context("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_rota_adapter_smoke.log")

Dim plan
Set plan = SC_GM_CreateMessagePlanFromCsv(context, fixturePath, #5/29/2026#, defaults)
AssertTrue plan("state") = SC_STATE_COMPOSED, "CSV-derived message plan composes"
AssertTrue plan("messages").Count = 6, "CSV-derived message plan includes six messages"

Dim regularSpec
Set regularSpec = SC_BuildGroupMeetingsMessageSpec(context, plan("messages")("5"))
AssertTrue regularSpec("message_key") = "ORG|GROUP-MEETINGS|2026-05-29|REGULAR|TEAM", "CSV-derived regular message key is deterministic"

Dim directory
Set directory = SC_CreateDictionary()
directory("group") = "group@example.invalid"
directory("manager") = "manager@example.invalid"
directory("Literature Presenter") = "literature.presenter@example.invalid"
directory("Results Presenter") = "results.presenter@example.invalid"
directory("Previous Lab Manager") = "meeting.lead@example.invalid"
directory("Lab Manager") = "lab.manager@example.invalid"

Dim templateContext
Set templateContext = SC_CreateDictionary()
Dim runResult
Set runResult = SC_GM_RunMessagePlan(context, plan, directory, templateContext)
AssertTrue runResult("planned_count") = 6, "CSV-derived plan run sees all messages"
AssertTrue runResult("composed_count") = 6, "CSV-derived plan remains dry-run composed"
AssertTrue runResult("failed_count") = 0, "CSV-derived plan has no failed-safe result"
AssertTrue Not CBool(runResult("dispatch_results")("5")("created_new_mail")), "CSV-derived plan creates no Outlook mail"

Dim noMeetingPlan
Set noMeetingPlan = SC_GM_CreateMessagePlanFromCsv(context, fixturePath, #6/5/2026#, defaults)
AssertTrue noMeetingPlan("messages").Count = 1, "No-meeting row plans one notice"
AssertTrue InStr(1, noMeetingPlan("messages")("0")("subject"), "No group meeting", vbTextCompare) > 0, "No-meeting row subject is explicit"

Dim missingSetupPlan
Set missingSetupPlan = SC_GM_CreateMessagePlanFromCsv(context, fixturePath, #6/12/2026#, defaults)
AssertTrue missingSetupPlan("state") = SC_STATE_ACTION_REQUIRED, "Unconfigured rota row becomes action-required"
AssertTrue missingSetupPlan("messages")("0")("message_type") = SC_GM_TYPE_MEETING_NOT_SET_ALERT, "Unconfigured rota row plans manager alert"

Dim missingDatePlan
Set missingDatePlan = SC_GM_CreateMessagePlanFromCsv(context, fixturePath, #7/3/2026#, defaults)
AssertTrue missingDatePlan("state") = SC_STATE_ACTION_REQUIRED, "Missing target date becomes action-required"
AssertTrue missingDatePlan("messages").Count = 0, "Missing target date does not fabricate messages"
AssertTrue missingDatePlan("skips").Count = 1, "Missing target date records skip reason"

Dim adapterPath
adapterPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_rota_adapter.vbs")
AssertTrue Not FileContains(adapterPath, "CreateObject(""Excel.Application"")"), "Rota adapter owns no Excel application creation"
AssertTrue Not FileContains(adapterPath, "CreateObject(""Outlook.Application"")"), "Rota adapter owns no Outlook application creation"
AssertTrue Not FileContains(adapterPath, "GetDefaultFolder"), "Rota adapter owns no Outlook folder access"
AssertTrue Not FileContains(adapterPath, ".Send"), "Rota adapter owns no Outlook send call"
AssertTrue Not FileContains(adapterPath, ".Delete"), "Rota adapter owns no Outlook delete call"
AssertTrue Not FileContains(adapterPath, "sharepoint.com"), "Rota adapter contains no private schedule links"

If failures > 0 Then
  WScript.Echo "Group meetings rota adapter smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Group meetings rota adapter smoke test passed."
