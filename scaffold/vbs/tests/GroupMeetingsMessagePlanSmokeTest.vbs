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

Function MessageAt(plan, index)
  Set MessageAt = plan("messages")(CStr(index))
End Function

Dim context
Set context = SC_CreateGroupMeetingsWorkflowContext()
context("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_group_plan_smoke.log")

Dim directory
Set directory = SC_CreateDictionary()
directory("group") = "group@example.invalid"
directory("manager") = "manager@example.invalid"
directory("Literature Presenter") = "literature.presenter@example.invalid"
directory("Results Presenter") = "results.presenter@example.invalid"
directory("Lab Manager") = "lab.manager@example.invalid"
directory("Meeting Lead") = "meeting.lead@example.invalid"

Dim templateContext
Set templateContext = SC_CreateDictionary()

Dim record
Set record = SC_GM_CreateMeetingRecord(#5/29/2026#)
record("literature_presenter") = "Literature Presenter"
record("results_presenter") = "Results Presenter"
record("lab_manager") = "Lab Manager"
record("meeting_lead") = "Meeting Lead"
record("room") = "Room A"
record("include_same_day_reminder") = True

Dim plan
Set plan = SC_GM_CreateMessagePlan(context, record)
AssertTrue plan("state") = SC_STATE_COMPOSED, "Configured meeting plan composes"
AssertTrue plan("messages").Count = 6, "Configured meeting plans four reminders, same-day reminder, and regular distribution"
AssertTrue MessageAt(plan, 0)("message_type") = SC_GM_TYPE_LITERATURE_REMINDER, "Literature reminder first"
AssertTrue MessageAt(plan, 1)("message_type") = SC_GM_TYPE_RESULTS_REMINDER, "Results reminder second"
AssertTrue MessageAt(plan, 2)("message_type") = SC_GM_TYPE_MEETING_LEAD_REMINDER, "Meeting lead reminder third"
AssertTrue MessageAt(plan, 3)("message_type") = SC_GM_TYPE_LAB_MANAGER_REMINDER, "Lab manager reminder fourth"
AssertTrue MessageAt(plan, 4)("message_type") = SC_GM_TYPE_SAME_DAY_REMINDER, "Same-day reminder fifth"
AssertTrue MessageAt(plan, 5)("message_type") = SC_GM_TYPE_REGULAR, "Regular distribution last"

Dim regularSpec
Set regularSpec = SC_BuildGroupMeetingsMessageSpec(context, MessageAt(plan, 5))
AssertTrue regularSpec("message_key") = "ORG|GROUP-MEETINGS|2026-05-29|REGULAR|TEAM", "Regular message uses managed key"

Dim planRun
Set planRun = SC_GM_RunMessagePlan(context, plan, directory, templateContext)
AssertTrue planRun("planned_count") = 6, "Plan run sees all planned messages"
AssertTrue planRun("composed_count") = 6, "Plan run remains dry-run composed"
AssertTrue planRun("failed_count") = 0, "Plan run has no failed-safe result"
AssertTrue planRun("dispatch_results")("5")("prepared_to") = "group@example.invalid", "Regular distribution resolves group recipient"
AssertTrue Not CBool(planRun("dispatch_results")("5")("created_new_mail")), "Plan run creates no Outlook mail"

Dim noMeetingRecord
Set noMeetingRecord = SC_GM_CreateMeetingRecord(#6/5/2026#)
noMeetingRecord("alternate_text") = "no meeting"
Dim noMeetingPlan
Set noMeetingPlan = SC_GM_CreateMessagePlan(context, noMeetingRecord)
AssertTrue noMeetingPlan("messages").Count = 1, "No-meeting event plans only one notice"
AssertTrue MessageAt(noMeetingPlan, 0)("message_type") = SC_GM_TYPE_REGULAR, "No-meeting notice preserves regular logical key type"
AssertTrue InStr(1, MessageAt(noMeetingPlan, 0)("subject"), "No group meeting", vbTextCompare) > 0, "No-meeting notice subject is explicit"
AssertTrue noMeetingPlan("skips").Count = 1, "No-meeting event records suppressed reminders"

Dim missingRecord
Set missingRecord = SC_GM_CreateMeetingRecord(#6/12/2026#)
Dim missingPlan
Set missingPlan = SC_GM_CreateMessagePlan(context, missingRecord)
AssertTrue missingPlan("state") = SC_STATE_ACTION_REQUIRED, "Unconfigured meeting becomes action-required plan"
AssertTrue missingPlan("messages").Count = 1, "Unconfigured meeting plans only manager alert"
AssertTrue MessageAt(missingPlan, 0)("message_type") = SC_GM_TYPE_MEETING_NOT_SET_ALERT, "Unconfigured meeting produces manager alert"
AssertTrue missingPlan("skips").Count = 1, "Unconfigured meeting records regular suppression"

Dim alternateRecord
Set alternateRecord = SC_GM_CreateMeetingRecord(#6/19/2026#)
alternateRecord("alternate_text") = "journal club"
alternateRecord("include_same_day_reminder") = False
Dim alternatePlan
Set alternatePlan = SC_GM_CreateMessagePlan(context, alternateRecord)
AssertTrue alternatePlan("messages").Count = 5, "Alternate configured meeting omits same-day unless requested"
AssertTrue alternatePlan("skips").Count = 1, "Alternate configured meeting records same-day skip"

Dim planPath
planPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_message_plan.vbs")
AssertTrue Not FileContains(planPath, "CreateObject(""Outlook.Application"")"), "Message plan owns no Outlook application creation"
AssertTrue Not FileContains(planPath, "GetDefaultFolder"), "Message plan owns no Outlook folder access"
AssertTrue Not FileContains(planPath, ".Send"), "Message plan owns no Outlook send call"
AssertTrue Not FileContains(planPath, ".Delete"), "Message plan owns no Outlook delete call"
AssertTrue Not FileContains(planPath, "sharepoint.com"), "Message plan contains no private schedule links"

If failures > 0 Then
  WScript.Echo "Group meetings message-plan smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Group meetings message-plan smoke test passed."
