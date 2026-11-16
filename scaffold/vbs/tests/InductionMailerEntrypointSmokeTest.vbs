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

Function QuoteArg(value)
  QuoteArg = Chr(34) & Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function RunCommand(commandText)
  Dim runResult
  Set runResult = SC_CreateDictionary()
  runResult("exit_code") = -1
  runResult("output") = ""

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
entrypointPath = fso.BuildPath(scaffoldRoot, "vbs\entrypoints\run_induction_mailer_dry_run.vbs")
Dim manifestPath
manifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_fixture.csv")
Dim missingManifestPath
missingManifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_missing_template_fixture.csv")
Dim recipientsPath
recipientsPath = fso.BuildPath(scriptFolder, "fixtures\induction_recipients_fixture.csv")

Dim allCommand
allCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /templates:" & QuoteArg(manifestPath) & _
  " /recipients:" & QuoteArg(recipientsPath) & _
  " /mode:all /date:2026-06-01"
Dim allRun
Set allRun = RunCommand(allCommand)
AssertTrue allRun("exit_code") = 0, "All-mode induction entrypoint exits successfully"
AssertTrue OutputContains(allRun, "state=COMPOSED"), "All-mode induction entrypoint reports composed"
AssertTrue OutputContains(allRun, "planned_count=3"), "All-mode induction entrypoint reports three planned messages"
AssertTrue OutputContains(allRun, "composed_count=3"), "All-mode induction entrypoint reports three composed messages"
AssertTrue OutputContains(allRun, "failed_count=0"), "All-mode induction entrypoint reports no failures"

Dim personalCommand
personalCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /templates:" & QuoteArg(manifestPath) & _
  " /recipients:" & QuoteArg(recipientsPath) & _
  " /mode:personal /date:2026-06-01"
Dim personalRun
Set personalRun = RunCommand(personalCommand)
AssertTrue personalRun("exit_code") = 0, "Personal-mode induction entrypoint exits successfully"
AssertTrue OutputContains(personalRun, "selected_mode=personal"), "Personal-mode induction entrypoint reports selected mode"
AssertTrue OutputContains(personalRun, "planned_count=2"), "Personal-mode induction entrypoint plans two personal messages"

Dim missingTemplateCommand
missingTemplateCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /templates:" & QuoteArg(missingManifestPath) & _
  " /recipients:" & QuoteArg(recipientsPath) & _
  " /mode:all /date:2026-06-01"
Dim missingTemplateRun
Set missingTemplateRun = RunCommand(missingTemplateCommand)
AssertTrue missingTemplateRun("exit_code") = 2, "Missing-template induction entrypoint exits action-required"
AssertTrue OutputContains(missingTemplateRun, "state=ACTION_REQUIRED"), "Missing-template induction entrypoint reports action-required"
AssertTrue OutputContains(missingTemplateRun, "planned_count=0"), "Missing-template induction entrypoint plans no messages"
AssertTrue OutputContains(missingTemplateRun, "skipped_count=1"), "Missing-template induction entrypoint reports skipped template"

Dim blockedCommand
blockedCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /templates:" & QuoteArg(manifestPath) & _
  " /recipients:" & QuoteArg(recipientsPath) & _
  " /mode:all /dry_run:false"
Dim blockedRun
Set blockedRun = RunCommand(blockedCommand)
AssertTrue blockedRun("exit_code") = 1, "Non-dry-run induction entrypoint request exits failed-safe"
AssertTrue OutputContains(blockedRun, "state=FAILED_SAFE"), "Non-dry-run induction entrypoint reports failed-safe"
AssertTrue OutputContains(blockedRun, "planned_count=0"), "Non-dry-run induction entrypoint does not plan messages"

Dim usageCommand
usageCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath)
Dim usageRun
Set usageRun = RunCommand(usageCommand)
AssertTrue usageRun("exit_code") = 64, "Missing induction entrypoint arguments show usage"
AssertTrue OutputContains(usageRun, "dry_run_only=true"), "Induction usage advertises dry-run-only boundary"

AssertTrue Not FileContains(entrypointPath, "CreateObject(""Outlook.Application"")"), "Induction entrypoint owns no Outlook application creation"
AssertTrue Not FileContains(entrypointPath, "GetDefaultFolder"), "Induction entrypoint owns no Outlook folder access"
AssertTrue Not FileContains(entrypointPath, ".Send"), "Induction entrypoint owns no Outlook send call"
AssertTrue Not FileContains(entrypointPath, ".Delete"), "Induction entrypoint owns no Outlook delete call"
AssertTrue Not FileContains(entrypointPath, "sharepoint.com"), "Induction entrypoint contains no private links"

If failures > 0 Then
  WScript.Echo "Induction mailer entrypoint smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Induction mailer entrypoint smoke test passed."
