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
Dim tempLogPath
tempLogPath = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_mail_composer_smoke.log")

Sub AssertTrue(condition, message)
  If condition Then
    WScript.Echo "OK: " & message
  Else
    WScript.Echo "FAILED: " & message
    failures = failures + 1
  End If
End Sub

Dim fixtureRoot
fixtureRoot = fso.BuildPath(scaffoldRoot, "vbs\tests\fixtures")
Dim templatePath
templatePath = fso.BuildPath(fixtureRoot, "sample_template.html")

Dim directory
Set directory = SC_CreateDictionary()
directory("team lead") = "team.lead@example.invalid"

Dim context
Set context = SC_CreateDictionary()
context("FIRST_NAME") = "Alex"

Dim key
key = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "target")

Dim spec
Set spec = SC_CreateMessageSpec(key, "regular", "team lead", "Prepared Subject", "")
spec("cc") = "Alex Person <alex.person@example.invalid>"
spec("body_html") = "<html><body><p>Hello [FIRST_NAME]</p><img src=""sample_template.files/logo%20one.png""></body></html>"
spec("template_path") = templatePath
spec("attachment_base_folder") = fixtureRoot
spec("attachments")("first") = "attachment.txt"

Dim policy
Set policy = SC_CreateDeliveryPolicy("display")
policy("defer_minutes") = 5

Dim composerContext
Set composerContext = SC_CreateDictionary()
Set composerContext("recipient_directory") = directory
Set composerContext("template_context") = context

Dim prepared
Set prepared = SC_PrepareMessageForComposition(spec, policy, composerContext, "")

AssertTrue prepared("message_key") = key, "managed key preserved"
AssertTrue prepared("to") = "team.lead@example.invalid", "To recipient resolved"
AssertTrue prepared("cc") = "alex.person@example.invalid", "Cc recipient resolved"
AssertTrue prepared("subject") = "Prepared Subject", "explicit subject preserved"
AssertTrue prepared("attachment_paths").Count = 1, "attachment prepared"
AssertTrue prepared("inline_info")("inline_count") = 1, "inline image prepared"
AssertTrue InStr(1, prepared("body_html"), "cid:logo_one.png", vbTextCompare) > 0, "inline CID applied"
AssertTrue prepared("defer_minutes") = 5, "deferred delivery policy carried"

Dim dryRunConfig
Set dryRunConfig = SC_CreateDictionary()
dryRunConfig("dry_run") = True
dryRunConfig("log_path") = tempLogPath
Set dryRunConfig("recipient_directory") = directory
Set dryRunConfig("template_context") = context

Dim dispatchResult
Set dispatchResult = SC_DispatchManagedEmail(spec, policy, dryRunConfig)
AssertTrue dispatchResult("state") = SC_STATE_COMPOSED, "dry-run dispatch still avoids Outlook"

If failures > 0 Then
  WScript.Echo "Mail composer smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

On Error Resume Next
If fso.FileExists(tempLogPath) Then fso.DeleteFile tempLogPath, True
On Error GoTo 0

WScript.Echo "Mail composer smoke test passed."
