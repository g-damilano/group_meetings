Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell
Set shell = CreateObject("WScript.Shell")

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
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\watchdog\outlook_watchdog_sidecar.vbs")

Dim failures
failures = 0
Dim tempLogPath
tempLogPath = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_watchdog_sidecar_smoke.log")

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

Class FakeUserProperty
  Public Value
End Class

Class FakeUserProperties
  Private props

  Private Sub Class_Initialize()
    Set props = CreateObject("Scripting.Dictionary")
  End Sub

  Public Function Find(name)
    If props.Exists(name) Then
      Set Find = props(name)
    Else
      Set Find = Nothing
    End If
  End Function

  Public Function Add(name, propType, addToFolderFields)
    Dim prop
    Set prop = New FakeUserProperty
    props(name) = prop
    Set Add = prop
  End Function
End Class

Class FakeAttachments
  Public Count
End Class

Class FakePropertyAccessor
  Public Function GetProperty(name)
    Err.Raise 4601, "FakePropertyAccessor", "Property unavailable"
  End Function
End Class

Class FakeMail
  Public UserProperties
  Public Attachments
  Public PropertyAccessor
  Public HTMLBody
  Public Body
  Public Subject
  Public EntryID
  Public CreationTime
  Public DeferredDeliveryTime
  Public DeleteCalled

  Private Sub Class_Initialize()
    Set UserProperties = New FakeUserProperties
    Set Attachments = New FakeAttachments
    Set PropertyAccessor = New FakePropertyAccessor
    HTMLBody = "<html><body><p>Queued</p></body></html>"
    Body = "Queued"
    Subject = "Queued"
    EntryID = "queued-entry"
    CreationTime = DateAdd("n", -30, Now)
    DeferredDeliveryTime = DateAdd("n", -1, Now)
    DeleteCalled = False
  End Sub

  Public Sub Delete()
    DeleteCalled = True
  End Sub
End Class

Function MakeTransport(blocked)
  Dim report
  Set report = SC_CreateHealthReport()
  report("blocked_or_auth_suspected") = CBool(blocked)
  report("offline") = CBool(blocked)
  Set MakeTransport = report
End Function

Function NewMail(subjectText, managedKey)
  Dim mail
  Set mail = New FakeMail
  mail.Subject = subjectText
  If Len(managedKey) > 0 Then SC_StampManagedMail mail, managedKey
  Set NewMail = mail
End Function

Dim managedKey
managedKey = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "target")
Dim managedMail
Set managedMail = NewMail("stale managed", managedKey)
Dim userMail
Set userMail = NewMail("user authored", "")

Dim items
Set items = SC_CreateDictionary()
Set items("managed") = managedMail
Set items("user") = userMail

Dim result
Set result = SC_WD_RunReportOnly(items, MakeTransport(False), Now, 15, tempLogPath)
AssertTrue result("state") = SC_STATE_ACTION_REQUIRED, "Stale managed mail is reported action-required in report-only mode"
AssertTrue CBool(result("report_only")), "Sidecar result is report-only"
AssertTrue result("repair_actions_taken") = 0, "Sidecar takes no repair actions"
AssertTrue result("total_count") = 2, "Sidecar summarizes total inspected mail"
AssertTrue result("managed_count") = 1, "Sidecar summarizes managed mail"
AssertTrue result("unmanaged_count") = 1, "Sidecar summarizes unmanaged mail"
AssertTrue result("stale_count") = 1, "Sidecar summarizes stale managed mail"
AssertTrue result("replace_eligible_count") = 1, "Sidecar summarizes replace-eligible mail without repair"
AssertTrue result("failed_safe_count") = 0, "Sidecar summarizes failed-safe count"
AssertTrue result("outbox_report")("unmanaged_count") = 1, "Sidecar reports unmanaged mail"
AssertTrue userMail.DeleteCalled = False, "Sidecar does not delete unmarked user mail"
AssertTrue managedMail.DeleteCalled = False, "Sidecar does not delete managed mail in report-only mode"

Dim uncertain
Set uncertain = NewMail("uncertain", managedKey)
uncertain.CreationTime = "not-a-date"
Dim uncertainItems
Set uncertainItems = SC_CreateDictionary()
Set uncertainItems("uncertain") = uncertain
Set result = SC_WD_RunReportOnly(uncertainItems, MakeTransport(False), Now, 15, tempLogPath)
AssertTrue result("state") = SC_STATE_FAILED_SAFE, "Uncertain watchdog state fails safe"
AssertTrue CBool(result("action_required")), "Uncertain watchdog state requires action"
AssertTrue result("failed_safe_count") = 1, "Uncertain watchdog state is summarized as failed safe"
AssertTrue uncertain.DeleteCalled = False, "Uncertain item is not modified"

Dim health
Set health = SC_CreateHealthReport()
health("outbox_count") = 4
health("managed_outbox_count") = 2
health("unmanaged_outbox_count") = 2
health("blocked_or_auth_suspected") = True
Set result = SC_WD_CreateSidecarResult()
SC_WD_ApplyHealthSummary result, health
AssertTrue result("state") = SC_STATE_ACTION_REQUIRED, "Blocked health summary requires action"
AssertTrue result("total_count") = 4, "Health summary preserves Outbox count"
AssertTrue result("managed_count") = 2, "Health summary preserves managed count"
AssertTrue result("unmanaged_count") = 2, "Health summary preserves unmanaged count"
AssertTrue result("repair_actions_taken") = 0, "Health summary remains non-mutating"

Function QuoteArg(value)
  QuoteArg = Chr(34) & Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function RunCommand(commandText)
  Dim runResult
  Set runResult = SC_CreateDictionary()
  Dim exec
  Set exec = shell.Exec(commandText)
  Do While exec.Status = 0
    WScript.Sleep 50
  Loop
  runResult("exit_code") = exec.ExitCode
  runResult("output") = exec.StdOut.ReadAll & exec.StdErr.ReadAll
  Set RunCommand = runResult
End Function

Function OutputContains(runResult, token)
  OutputContains = (InStr(1, CStr(runResult("output")), CStr(token), vbTextCompare) > 0)
End Function

Dim entrypointPath
entrypointPath = fso.BuildPath(scaffoldRoot, "vbs\entrypoints\run_watchdog_report_only.vbs")
Dim dryRunEntry
Set dryRunEntry = RunCommand("cscript.exe //NoLogo " & QuoteArg(entrypointPath))
AssertTrue dryRunEntry("exit_code") = 0, "Watchdog entrypoint defaults to no-live report-only mode"
AssertTrue OutputContains(dryRunEntry, "report_only=true"), "Watchdog entrypoint reports report-only boundary"
AssertTrue OutputContains(dryRunEntry, "total_count=0"), "Watchdog entrypoint prints total count"
AssertTrue OutputContains(dryRunEntry, "managed_count=0"), "Watchdog entrypoint prints managed count"
AssertTrue OutputContains(dryRunEntry, "replace_eligible_count=0"), "Watchdog entrypoint prints replace eligibility count"

Dim blockedEntry
Set blockedEntry = RunCommand("cscript.exe //NoLogo " & QuoteArg(entrypointPath) & " /repair")
AssertTrue blockedEntry("exit_code") = 1, "Watchdog repair request is blocked"
AssertTrue OutputContains(blockedEntry, "state=FAILED_SAFE"), "Blocked watchdog repair request fails safe"
AssertTrue OutputContains(blockedEntry, "repair_actions_taken=0"), "Blocked watchdog repair request reports no repair"

Dim sidecarPath
sidecarPath = fso.BuildPath(scaffoldRoot, "vbs\watchdog\outlook_watchdog_sidecar.vbs")
AssertTrue Not FileContains(sidecarPath, ".Send"), "Watchdog sidecar owns no send call"
AssertTrue Not FileContains(sidecarPath, ".Delete"), "Watchdog sidecar owns no delete call"
AssertTrue Not FileContains(sidecarPath, "CreateObject(""Outlook.Application"")"), "Watchdog sidecar owns no Outlook creation"

Dim taskPath
taskPath = fso.BuildPath(scaffoldRoot, "examples\task_scheduler\watchdog_report_only_task.xml")
AssertTrue fso.FileExists(taskPath), "Task Scheduler report-only example exists"
AssertTrue FileContains(taskPath, "run_watchdog_report_only.vbs"), "Task Scheduler example runs watchdog entrypoint"
AssertTrue FileContains(taskPath, "OFFICE_SCAFFOLD_ROOT"), "Task Scheduler example uses environment placeholder"
AssertTrue FileContains(taskPath, "<Enabled>false</Enabled>"), "Task Scheduler example is disabled by default"
AssertTrue Not FileContains(taskPath, "inspect_live:true"), "Task Scheduler example does not request live Outlook inspection"
AssertTrue Not FileContains(taskPath, "/repair"), "Task Scheduler example does not request repair"
AssertTrue Not FileContains(taskPath, "/delete"), "Task Scheduler example does not request delete"
AssertTrue Not FileContains(taskPath, "/send"), "Task Scheduler example does not request send"
AssertTrue Not FileContains(taskPath, "/replace"), "Task Scheduler example does not request replacement"

Dim watchdogFixturePath
watchdogFixturePath = fso.BuildPath(scaffoldRoot, "fixtures\watchdog\report_only_health_fixture.json")
AssertTrue fso.FileExists(watchdogFixturePath), "Watchdog report-only health fixture exists"
AssertTrue FileContains(watchdogFixturePath, "office-communication-scaffold.watchdog-report-fixture.v1"), "Watchdog fixture names canonical schema"
AssertTrue FileContains(watchdogFixturePath, """report_only"": true"), "Watchdog fixture is report-only"
AssertTrue FileContains(watchdogFixturePath, """live_outlook_required"": false"), "Watchdog fixture does not require live Outlook"
AssertTrue FileContains(watchdogFixturePath, """mailbox_mutation_allowed"": false"), "Watchdog fixture forbids mailbox mutation"
AssertTrue FileContains(watchdogFixturePath, """repair_committed"": false"), "Watchdog fixture keeps repair uncommitted"
AssertTrue FileContains(watchdogFixturePath, """would_modify"": false"), "Watchdog fixture records non-mutating items"
AssertTrue Not FileContains(watchdogFixturePath, """would_modify"": true"), "Watchdog fixture contains no mutating item"
AssertTrue Not FileContains(watchdogFixturePath, """send_committed"": true"), "Watchdog fixture contains no committed send"
AssertTrue Not FileContains(watchdogFixturePath, """delete_committed"": true"), "Watchdog fixture contains no committed delete"
AssertTrue Not FileContains(watchdogFixturePath, """replace_committed"": true"), "Watchdog fixture contains no committed replacement"
AssertTrue FileContains(watchdogFixturePath, "STALE_MANAGED_OUTBOX"), "Watchdog fixture covers stale managed Outbox state"
AssertTrue FileContains(watchdogFixturePath, "ACTION_REQUIRED"), "Watchdog fixture covers action-required state"
AssertTrue FileContains(watchdogFixturePath, "FAILED_SAFE"), "Watchdog fixture covers failed-safe uncertainty"
AssertTrue FileContains(watchdogFixturePath, "unmarked-user-mail"), "Watchdog fixture covers unmarked user mail"

If failures > 0 Then
  WScript.Echo "Watchdog sidecar smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

On Error Resume Next
If fso.FileExists(tempLogPath) Then fso.DeleteFile tempLogPath, True
On Error GoTo 0

WScript.Echo "Watchdog sidecar smoke test passed."
