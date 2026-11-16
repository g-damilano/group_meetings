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
  runResult("command") = commandText
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
entrypointPath = fso.BuildPath(scaffoldRoot, "vbs\entrypoints\run_group_meetings_dry_run.vbs")
Dim rotaPath
rotaPath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_rota_fixture.csv")
Dim directoryPath
directoryPath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_recipients_fixture.csv")

Dim configuredCommand
configuredCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /rota:" & QuoteArg(rotaPath) & _
  " /date:2026-05-29" & _
  " /directory:" & QuoteArg(directoryPath)
Dim configuredRun
Set configuredRun = RunCommand(configuredCommand)
AssertTrue configuredRun("exit_code") = 0, "Configured entrypoint exits successfully"
AssertTrue OutputContains(configuredRun, "state=COMPOSED"), "Configured entrypoint reports composed state"
AssertTrue OutputContains(configuredRun, "dry_run=true"), "Configured entrypoint reports dry-run mode"
AssertTrue OutputContains(configuredRun, "planned_count=6"), "Configured entrypoint reports six planned messages"
AssertTrue OutputContains(configuredRun, "composed_count=6"), "Configured entrypoint reports six composed messages"
AssertTrue OutputContains(configuredRun, "failed_count=0"), "Configured entrypoint reports no failed messages"

Dim actionRequiredCommand
actionRequiredCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /rota:" & QuoteArg(rotaPath) & _
  " /date:2026-06-12" & _
  " /directory:" & QuoteArg(directoryPath)
Dim actionRequiredRun
Set actionRequiredRun = RunCommand(actionRequiredCommand)
AssertTrue actionRequiredRun("exit_code") = 2, "Incomplete rota entrypoint exits with action-required code"
AssertTrue OutputContains(actionRequiredRun, "state=ACTION_REQUIRED"), "Incomplete rota entrypoint reports action-required state"
AssertTrue OutputContains(actionRequiredRun, "planned_count=1"), "Incomplete rota entrypoint reports manager alert"

Dim blockedCommand
blockedCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath) & _
  " /rota:" & QuoteArg(rotaPath) & _
  " /date:2026-05-29" & _
  " /directory:" & QuoteArg(directoryPath) & _
  " /dry_run:false"
Dim blockedRun
Set blockedRun = RunCommand(blockedCommand)
AssertTrue blockedRun("exit_code") = 1, "Non-dry-run entrypoint request exits failed-safe"
AssertTrue OutputContains(blockedRun, "state=FAILED_SAFE"), "Non-dry-run entrypoint request reports failed-safe state"
AssertTrue OutputContains(blockedRun, "dry_run=false"), "Non-dry-run entrypoint request reports blocked mode"
AssertTrue OutputContains(blockedRun, "planned_count=0"), "Non-dry-run entrypoint request does not plan messages"

Dim usageCommand
usageCommand = "cscript.exe //NoLogo " & QuoteArg(entrypointPath)
Dim usageRun
Set usageRun = RunCommand(usageCommand)
AssertTrue usageRun("exit_code") = 64, "Missing arguments show usage"
AssertTrue OutputContains(usageRun, "dry_run_only=true"), "Usage advertises dry-run-only boundary"

AssertTrue Not FileContains(entrypointPath, "CreateObject(""Excel.Application"")"), "Entrypoint owns no Excel application creation"
AssertTrue Not FileContains(entrypointPath, "CreateObject(""Outlook.Application"")"), "Entrypoint owns no Outlook application creation"
AssertTrue Not FileContains(entrypointPath, "GetDefaultFolder"), "Entrypoint owns no Outlook folder access"
AssertTrue Not FileContains(entrypointPath, ".Send"), "Entrypoint owns no Outlook send call"
AssertTrue Not FileContains(entrypointPath, ".Delete"), "Entrypoint owns no Outlook delete call"
AssertTrue Not FileContains(entrypointPath, "sharepoint.com"), "Entrypoint contains no private schedule links"

If failures > 0 Then
  WScript.Echo "Group meetings entrypoint smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Group meetings entrypoint smoke test passed."
