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
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_workflow.vbs")

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
  FileContains = False
  Dim text
  text = SC_ReadTextFile(path)
  FileContains = (InStr(1, text, CStr(token), vbTextCompare) > 0)
End Function

Dim directory
Set directory = SC_CreateDictionary()
directory("team lead") = "team.lead@example.invalid"
directory("new starter") = "new.starter@example.invalid"

Dim groupTemplateContext
Set groupTemplateContext = SC_CreateDictionary()
groupTemplateContext("MEETING_DATE") = "2026-05-29"
groupTemplateContext("ROOM") = "Room A"

Dim groupContext
Set groupContext = SC_CreateGroupMeetingsWorkflowContext()
groupContext("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_group_workflow_smoke.log")

Dim groupRequest
Set groupRequest = SC_CreateGroupMeetingsMessageRequest( _
  #5/29/2026#, _
  "regular", _
  "team", _
  "team lead", _
  "Group meeting [MEETING_DATE]", _
  "<p>Meeting in [ROOM] on [MEETING_DATE].</p>")
SC_SetWorkflowRequestRecipients groupRequest, "team lead", "", ""
SC_SetWorkflowRequestAttachments groupRequest, "attachment.txt", fso.BuildPath(scriptFolder, "fixtures")

Dim groupSpec
Set groupSpec = SC_BuildGroupMeetingsMessageSpec(groupContext, groupRequest)
AssertTrue groupSpec("message_key") = "ORG|GROUP-MEETINGS|2026-05-29|REGULAR|TEAM", "Group workflow builds deterministic managed key"
AssertTrue groupSpec("workflow_reference") = "group-meetings", "Group workflow records reference name"

Dim groupResult
Set groupResult = SC_RunGroupMeetingsMessage(groupContext, groupRequest, directory, groupTemplateContext)
AssertTrue groupResult("state") = SC_STATE_COMPOSED, "Group workflow dry-run dispatch composes only"
AssertTrue groupResult("prepared_to") = "team.lead@example.invalid", "Group workflow resolves recipient through core"
AssertTrue groupResult("prepared_subject") = "Group meeting 2026-05-29", "Group workflow renders subject through core"
AssertTrue CLng(groupResult("prepared_attachment_count")) = 1, "Group workflow resolves attachment through core"
AssertTrue Not CBool(groupResult("created_new_mail")), "Group workflow dry-run creates no Outlook mail"

Dim inductionTemplateContext
Set inductionTemplateContext = SC_CreateDictionary()
inductionTemplateContext("NAME") = "New Starter"
inductionTemplateContext("START_DATE") = "2026-06-01"

Dim inductionContext
Set inductionContext = SC_CreateInductionMailerWorkflowContext()
inductionContext("log_path") = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_induction_workflow_smoke.log")

Dim inductionRequest
Set inductionRequest = SC_CreateInductionMailerMessageRequest( _
  #6/1/2026#, _
  "personal", _
  "new-starter", _
  "new starter", _
  "Welcome [NAME]", _
  "<p>Welcome [NAME]. Your start date is [START_DATE].</p>")

Dim inductionSpec
Set inductionSpec = SC_BuildInductionMailerMessageSpec(inductionContext, inductionRequest)
AssertTrue inductionSpec("message_key") = "ORG|INDUCTION-MAILER|2026-06-01|PERSONAL|NEW-STARTER", "Induction workflow builds deterministic managed key"
AssertTrue inductionSpec("workflow_reference") = "induction-mailer", "Induction workflow records reference name"

Dim inductionResult
Set inductionResult = SC_RunInductionMailerMessage(inductionContext, inductionRequest, directory, inductionTemplateContext)
AssertTrue inductionResult("state") = SC_STATE_COMPOSED, "Induction workflow dry-run dispatch composes only"
AssertTrue inductionResult("prepared_to") = "new.starter@example.invalid", "Induction workflow resolves recipient through core"
AssertTrue inductionResult("prepared_subject") = "Welcome New Starter", "Induction workflow renders subject through core"
AssertTrue Not CBool(inductionResult("created_new_mail")), "Induction workflow dry-run creates no Outlook mail"

Dim groupPath
groupPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_workflow.vbs")
Dim inductionPath
inductionPath = fso.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_workflow.vbs")
AssertTrue Not FileContains(groupPath, "CreateObject(""Outlook.Application"")"), "Group workflow owns no Outlook application creation"
AssertTrue Not FileContains(groupPath, "GetDefaultFolder"), "Group workflow owns no Outlook folder access"
AssertTrue Not FileContains(groupPath, ".Send"), "Group workflow owns no Outlook send call"
AssertTrue Not FileContains(inductionPath, "CreateObject(""Outlook.Application"")"), "Induction workflow owns no Outlook application creation"
AssertTrue Not FileContains(inductionPath, "GetDefaultFolder"), "Induction workflow owns no Outlook folder access"
AssertTrue Not FileContains(inductionPath, ".Send"), "Induction workflow owns no Outlook send call"

If failures > 0 Then
  WScript.Echo "Workflow skeleton smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Workflow skeleton smoke test passed."
