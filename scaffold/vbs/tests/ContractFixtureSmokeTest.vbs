Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

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

Function FileText(relativePath)
  Dim path
  path = fso.BuildPath(scaffoldRoot, relativePath)
  Dim stream
  Set stream = fso.OpenTextFile(path, 1, False)
  FileText = stream.ReadAll
  stream.Close
End Function

Function Contains(text, pattern)
  Contains = (InStr(1, text, pattern, vbTextCompare) > 0)
End Function

Dim schemaText
schemaText = FileText("contracts\workflow_schema.json")

Dim fixtureText
fixtureText = FileText("fixtures\contracts\stage4_shared_contract_fixture.json")

Dim reportText
reportText = FileText("python\contract_fixture_report.py")

AssertTrue Contains(schemaText, """message_spec"""), "Workflow schema defines MessageSpec"
AssertTrue Contains(schemaText, """recipient_spec"""), "Workflow schema defines RecipientSpec"
AssertTrue Contains(schemaText, """delivery_policy"""), "Workflow schema defines DeliveryPolicy"
AssertTrue Contains(schemaText, """managed_email_key"""), "Workflow schema defines ManagedEmailKey"
AssertTrue Contains(schemaText, """dispatch_result"""), "Workflow schema defines DispatchResult"
AssertTrue Contains(schemaText, """workflow_context"""), "Workflow schema defines WorkflowContext"
AssertTrue Contains(schemaText, """health_report"""), "Workflow schema defines HealthReport"

AssertTrue Contains(fixtureText, "office-communication-scaffold.shared-contract-fixture.v1"), "Shared contract fixture declares schema"
AssertTrue Contains(fixtureText, "ORG|GROUP-MEETINGS|2026-05-29|REGULAR|TEAM"), "Shared fixture includes group meetings managed key"
AssertTrue Contains(fixtureText, "ORG|INDUCTION-MAILER|2026-06-01|PERSONAL|NEW-STARTER"), "Shared fixture includes induction managed key"
AssertTrue Contains(fixtureText, """dispatch_results"""), "Shared fixture includes dispatch results"
AssertTrue Contains(fixtureText, """health_report"""), "Shared fixture includes health report"
AssertTrue Contains(fixtureText, "example.invalid"), "Shared fixture uses example addresses"
AssertTrue Contains(reportText, "contract-fixture-report"), "Python contract fixture report exists"
AssertTrue Contains(reportText, "report_only"), "Python contract fixture report declares report-only output"
AssertTrue Contains(reportText, "mutations_performed"), "Python contract fixture report declares mutation-free output"
AssertTrue Contains(reportText, "real_email_sending_performed"), "Python contract fixture report declares no real sending"

AssertTrue Not Contains(fixtureText, "CreateObject(""Outlook.Application"")"), "Shared fixture owns no Outlook creation"
AssertTrue Not Contains(fixtureText, "GetDefaultFolder"), "Shared fixture owns no Outlook folder access"
AssertTrue Not Contains(fixtureText, ".Send"), "Shared fixture owns no Outlook send call"
AssertTrue Not Contains(fixtureText, ".Delete"), "Shared fixture owns no Outlook delete call"
AssertTrue Not Contains(fixtureText, "One" & "Drive -"), "Shared fixture contains no private OneDrive marker"
AssertTrue Not Contains(fixtureText, "Imperial" & " College"), "Shared fixture contains no institution marker"
AssertTrue Not Contains(fixtureText, "sharepoint"), "Shared fixture contains no SharePoint link"

If failures > 0 Then
  WScript.Echo "Contract fixture smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Contract fixture smoke test passed."
