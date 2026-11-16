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

Function QuoteArg(value)
  QuoteArg = Chr(34) & Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function ReadText(path)
  Dim stream
  Set stream = fso.OpenTextFile(path, 1)
  ReadText = stream.ReadAll
  stream.Close
End Function

Function FileContains(path, token)
  FileContains = (InStr(1, ReadText(path), CStr(token), vbTextCompare) > 0)
End Function

Function RunCommand(commandText)
  Dim runResult
  Set runResult = CreateObject("Scripting.Dictionary")
  runResult("exit_code") = -1
  runResult("output") = ""

  Dim exec
  Set exec = shell.Exec(commandText)

  Dim output
  output = ""
  Do While exec.Status = 0
    Do While Not exec.StdOut.AtEndOfStream
      output = output & exec.StdOut.ReadLine & vbCrLf
    Loop
    Do While Not exec.StdErr.AtEndOfStream
      output = output & exec.StdErr.ReadLine & vbCrLf
    Loop
    WScript.Sleep 50
  Loop

  Do While Not exec.StdOut.AtEndOfStream
    output = output & exec.StdOut.ReadLine & vbCrLf
  Loop
  Do While Not exec.StdErr.AtEndOfStream
    output = output & exec.StdErr.ReadLine & vbCrLf
  Loop

  runResult("exit_code") = exec.ExitCode
  runResult("output") = output
  Set RunCommand = runResult
End Function

Function OutputContains(runResult, token)
  OutputContains = (InStr(1, CStr(runResult("output")), CStr(token), vbTextCompare) > 0)
End Function

Dim routerPath
routerPath = fso.BuildPath(scaffoldRoot, "Scaffold.vbs")

Dim helpRun
Set helpRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & " help")
AssertTrue helpRun("exit_code") = 0, "Router help exits successfully"
AssertTrue OutputContains(helpRun, "group-meetings --dry-run --input"), "Router help advertises group meetings"
AssertTrue OutputContains(helpRun, "induction --dry-run --recipients"), "Router help advertises induction"
AssertTrue OutputContains(helpRun, "watchdog --report-only"), "Router help advertises watchdog"
AssertTrue OutputContains(helpRun, "release-check"), "Router help advertises release check"
AssertTrue OutputContains(helpRun, "Exit codes"), "Router help documents exit codes"

Dim unknownRun
Set unknownRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & " unknown-command")
AssertTrue unknownRun("exit_code") = 2, "Router unknown command exits invalid-argument code"
AssertTrue OutputContains(unknownRun, "state=ACTION_REQUIRED"), "Router unknown command reports action required"

Dim validateRun
Set validateRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & " validate")
AssertTrue validateRun("exit_code") = 0, "Router validate exits successfully"
AssertTrue OutputContains(validateRun, "scaffold_command=validate"), "Router validate prints command summary"
AssertTrue OutputContains(validateRun, "Validation passed."), "Router validate delegates to scaffold validation"

Dim rotaPath
rotaPath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_rota_fixture.csv")
Dim directoryPath
directoryPath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_recipients_fixture.csv")
Dim groupRun
Set groupRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " group-meetings --dry-run --input " & QuoteArg(rotaPath) & _
  " --directory " & QuoteArg(directoryPath))
AssertTrue groupRun("exit_code") = 0, "Router group-meetings dry-run exits successfully"
AssertTrue OutputContains(groupRun, "scaffold_command=group-meetings"), "Router group-meetings prints command summary"
AssertTrue OutputContains(groupRun, "default_date_source=first meeting_date from input"), "Router group-meetings can infer target date"
AssertTrue OutputContains(groupRun, "state=COMPOSED"), "Router group-meetings reports composed state"
AssertTrue OutputContains(groupRun, "dry_run=true"), "Router group-meetings remains dry-run"

Dim exampleRotaPath
exampleRotaPath = fso.BuildPath(scaffoldRoot, "examples\workbooks\group_meetings_rota_example.csv")
Dim exampleGroupRun
Set exampleGroupRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " group-meetings --dry-run --input " & QuoteArg(exampleRotaPath))
AssertTrue exampleGroupRun("exit_code") = 0, "Router group-meetings example dry-run exits successfully"
AssertTrue OutputContains(exampleGroupRun, "state=COMPOSED"), "Router group-meetings example composes"
AssertTrue OutputContains(exampleGroupRun, "planned_count=6"), "Router group-meetings example plans expected messages"

Dim recipientsPath
recipientsPath = fso.BuildPath(scriptFolder, "fixtures\induction_recipients_fixture.csv")
Dim inductionRun
Set inductionRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " induction --dry-run --recipients " & QuoteArg(recipientsPath))
AssertTrue inductionRun("exit_code") = 0, "Router induction dry-run exits successfully"
AssertTrue OutputContains(inductionRun, "scaffold_command=induction"), "Router induction prints command summary"
AssertTrue OutputContains(inductionRun, "default_templates="), "Router induction uses default template manifest when omitted"
AssertTrue OutputContains(inductionRun, "state=COMPOSED"), "Router induction reports composed state"
AssertTrue OutputContains(inductionRun, "dry_run=true"), "Router induction remains dry-run"

Dim exampleRecipientsPath
exampleRecipientsPath = fso.BuildPath(scaffoldRoot, "examples\workbooks\induction_recipients_example.csv")
Dim exampleManifestPath
exampleManifestPath = fso.BuildPath(scaffoldRoot, "examples\workbooks\induction_template_manifest_example.csv")
Dim exampleInductionRun
Set exampleInductionRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " induction --dry-run --recipients " & QuoteArg(exampleRecipientsPath) & _
  " --templates " & QuoteArg(exampleManifestPath))
AssertTrue exampleInductionRun("exit_code") = 0, "Router induction example dry-run exits successfully"
AssertTrue OutputContains(exampleInductionRun, "state=COMPOSED"), "Router induction example composes"
AssertTrue OutputContains(exampleInductionRun, "planned_count=3"), "Router induction example plans expected messages"

Dim manifestPath
manifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_fixture.csv")
Dim invalidRecipientsPath
invalidRecipientsPath = fso.BuildPath(scriptFolder, "fixtures\induction_recipients_invalid_fixture.csv")
Dim invalidRecipientRun
Set invalidRecipientRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " induction --dry-run --recipients " & QuoteArg(invalidRecipientsPath) & _
  " --templates " & QuoteArg(manifestPath))
AssertTrue invalidRecipientRun("exit_code") = 1, "Router induction invalid recipients exits command-failure code"
AssertTrue OutputContains(invalidRecipientRun, "state=ACTION_REQUIRED"), "Router induction invalid recipients reports action required"
AssertTrue OutputContains(invalidRecipientRun, "Recipient source has invalid active recipient rows."), "Router induction invalid recipients reports source reason"

Dim invalidManifestPath
invalidManifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_invalid_fixture.csv")
Dim invalidTemplateRun
Set invalidTemplateRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " induction --dry-run --recipients " & QuoteArg(recipientsPath) & _
  " --templates " & QuoteArg(invalidManifestPath))
AssertTrue invalidTemplateRun("exit_code") = 1, "Router induction invalid templates exits command-failure code"
AssertTrue OutputContains(invalidTemplateRun, "state=ACTION_REQUIRED"), "Router induction invalid templates reports action required"
AssertTrue OutputContains(invalidTemplateRun, "Template manifest has invalid enabled rows."), "Router induction invalid templates reports manifest reason"

Dim watchdogRun
Set watchdogRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & " watchdog --report-only")
AssertTrue watchdogRun("exit_code") = 0, "Router watchdog report-only exits successfully"
AssertTrue OutputContains(watchdogRun, "scaffold_command=watchdog"), "Router watchdog prints command summary"
AssertTrue OutputContains(watchdogRun, "report_only=true"), "Router watchdog stays report-only"
AssertTrue OutputContains(watchdogRun, "repair_actions_taken=0"), "Router watchdog takes no repair actions"

Dim releaseRun
Set releaseRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & " release-check")
AssertTrue releaseRun("exit_code") = 0, "Router release-check exits successfully"
AssertTrue OutputContains(releaseRun, "Release gate: PASS"), "Router release-check delegates to release gate"
AssertTrue OutputContains(releaseRun, "repo_boundary_findings=advisory"), "Router release-check keeps parent repo findings advisory"

Dim unsafeRun
Set unsafeRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & _
  " group-meetings --send --input " & QuoteArg(rotaPath))
AssertTrue unsafeRun("exit_code") = 3, "Router blocks unsafe send request"
AssertTrue OutputContains(unsafeRun, "state=FAILED_SAFE"), "Router unsafe request fails safe"

Dim missingRun
Set missingRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(routerPath) & " induction --dry-run")
AssertTrue missingRun("exit_code") = 4, "Router missing required input exits missing-input code"
AssertTrue OutputContains(missingRun, "state=ACTION_REQUIRED"), "Router missing input reports action required"

AssertTrue fso.FileExists(routerPath), "Scaffold.vbs exists"
AssertTrue Not FileContains(routerPath, "CreateObject(""Outlook.Application"")"), "Router owns no Outlook creation"
AssertTrue Not FileContains(routerPath, "CreateObject(""Excel.Application"")"), "Router owns no Excel creation"
AssertTrue Not FileContains(routerPath, "GetDefaultFolder"), "Router owns no Outlook folder access"
AssertTrue Not FileContains(routerPath, ".Send"), "Router owns no Outlook send call"
AssertTrue Not FileContains(routerPath, ".Delete"), "Router owns no Outlook delete call"

If failures > 0 Then
  WScript.Echo "Scaffold router smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Scaffold router smoke test passed."
