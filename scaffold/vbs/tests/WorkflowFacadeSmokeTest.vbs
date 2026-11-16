Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

Dim bootstrapPath
bootstrapPath = fso.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
Dim ts
Set ts = fso.OpenTextFile(bootstrapPath, 1)
ExecuteGlobal ts.ReadAll
ts.Close

SC_LoadCoreModules

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

Dim context
Set context = SC_CreateWorkflowContext("group meetings")
context("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_workflow_facade_smoke.log")
context("settings")("org_code") = "org"
context("settings")("defer_minutes") = 5
context("settings")("enable_subject_fallback") = True
context("settings")("managed_outbox_grace_minutes") = 12

AssertTrue CBool(context("dry_run")), "Workflow context defaults to dry-run"

Dim spec
Set spec = SC_CreateWorkflowManagedMessageSpec( _
  context, _
  #5/29/2026#, _
  "regular", _
  "team", _
  "team lead", _
  "Hello [NAME]", _
  "", _
  "<p>Hello [NAME]</p>")

AssertTrue spec("message_key") = "ORG|GROUP-MEETINGS|2026-05-29|REGULAR|TEAM", "Facade builds deterministic managed key"
AssertTrue spec("metadata")("WORKFLOW_NAME") = "group meetings", "Facade stamps workflow metadata"
AssertTrue spec("metadata")("LOGICAL_DATE") = "2026-05-29", "Facade stamps logical date metadata"

SC_SetWorkflowMessageRecipients spec, "team lead", "Alex Person <alex.person@example.invalid>", ""
SC_AddWorkflowMessageMetadata spec, "custom_token", "custom value"

Dim attachments
Set attachments = SC_CreateDictionary()
attachments("0") = "attachment.txt"
SC_SetWorkflowMessageAttachments spec, attachments, fso.BuildPath(scriptFolder, "fixtures")

Dim policy
Set policy = SC_CreateWorkflowDeliveryPolicy(context, "send")
AssertTrue policy("action") = "send", "Facade preserves requested delivery action"
AssertTrue CLng(policy("defer_minutes")) = 5, "Facade reads delivery settings"
AssertTrue Not CBool(policy("enable_stale_outbox_replacement")), "Stale replacement remains opt-in by default"

Dim directory
Set directory = SC_CreateDictionary()
directory("team lead") = "team.lead@example.invalid"

Dim templateContext
Set templateContext = SC_CreateDictionary()
templateContext("NAME") = "Team"

Dim prepared
Set prepared = SC_PrepareWorkflowMessage(spec, policy, context, directory, templateContext)
AssertTrue prepared("to") = "team.lead@example.invalid", "Facade preparation resolves To recipients"
AssertTrue prepared("cc") = "alex.person@example.invalid", "Facade preparation resolves Cc recipients"
AssertTrue prepared("subject") = "Hello Team", "Facade preparation renders subject placeholders"
AssertTrue InStr(1, prepared("body_html"), "Hello Team", vbTextCompare) > 0, "Facade preparation renders HTML placeholders"
AssertTrue prepared("attachment_paths").Count = 1, "Facade preparation resolves attachments"

Dim dispatchConfig
Set dispatchConfig = SC_CreateWorkflowDispatchConfig(context, directory, templateContext)
AssertTrue CBool(dispatchConfig("dry_run")), "Facade dispatch config remains dry-run"
AssertTrue CLng(dispatchConfig("managed_outbox_grace_minutes")) = 12, "Facade dispatch config carries grace setting"

Dim result
Set result = SC_RunWorkflowMessage(spec, policy, context, directory, templateContext)
AssertTrue result("state") = SC_STATE_COMPOSED, "Facade run uses dry-run dispatch"
AssertTrue result("prepared_to") = "team.lead@example.invalid", "Facade run returns prepared To"
AssertTrue result("prepared_subject") = "Hello Team", "Facade run returns prepared subject"
AssertTrue CLng(result("prepared_attachment_count")) = 1, "Facade run returns attachment count"
AssertTrue Not CBool(result("created_new_mail")), "Facade dry-run creates no Outlook mail"

If failures > 0 Then
  WScript.Echo "Workflow facade smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Workflow facade smoke test passed."
