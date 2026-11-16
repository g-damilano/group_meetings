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

Function FileContains(path, token)
  FileContains = (InStr(1, SC_ReadTextFile(path), CStr(token), vbTextCompare) > 0)
End Function

Sub AssertFileSafe(relativePath)
  Dim path
  path = fso.BuildPath(scaffoldRoot, relativePath)
  AssertTrue fso.FileExists(path), relativePath & " exists"
  AssertTrue Not FileContains(path, "CreateObject(""Outlook.Application"")"), relativePath & " owns no Outlook creation"
  AssertTrue Not FileContains(path, "CreateObject(""Excel.Application"")"), relativePath & " owns no Excel creation"
  AssertTrue Not FileContains(path, "GetDefaultFolder"), relativePath & " owns no Outlook folder access"
  AssertTrue Not FileContains(path, ".Send"), relativePath & " owns no Outlook send call"
  AssertTrue Not FileContains(path, ".Delete"), relativePath & " owns no Outlook delete call"
  AssertTrue Not FileContains(path, "sharepoint.com"), relativePath & " contains no private links"
End Sub

Dim rotaPath
rotaPath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_rota_fixture.csv")
Dim directoryPath
directoryPath = fso.BuildPath(scriptFolder, "fixtures\group_meetings_recipients_fixture.csv")
Dim manifestPath
manifestPath = fso.BuildPath(scriptFolder, "fixtures\induction_template_manifest_fixture.csv")
Dim recipientsPath
recipientsPath = fso.BuildPath(scriptFolder, "fixtures\induction_recipients_fixture.csv")

Dim groupRun
Set groupRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(fso.BuildPath(scaffoldRoot, "RunGroupMeetings.vbs")) & _
  " /rota:" & QuoteArg(rotaPath) & _
  " /date:2026-05-29" & _
  " /directory:" & QuoteArg(directoryPath))
AssertTrue groupRun("exit_code") = 0, "Scaffold RunGroupMeetings wrapper exits successfully"
AssertTrue OutputContains(groupRun, "Launching scaffold group meetings dry-run"), "Group wrapper routes through scaffold launcher"
AssertTrue OutputContains(groupRun, "state=COMPOSED"), "Group wrapper reports composed state"
AssertTrue OutputContains(groupRun, "dry_run=true"), "Group wrapper stays dry-run"

Dim inductionRun
Set inductionRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(fso.BuildPath(scaffoldRoot, "RunInductionMailer.vbs")) & _
  " /templates:" & QuoteArg(manifestPath) & _
  " /recipients:" & QuoteArg(recipientsPath) & _
  " /mode:all /date:2026-06-01")
AssertTrue inductionRun("exit_code") = 0, "Scaffold RunInductionMailer wrapper exits successfully"
AssertTrue OutputContains(inductionRun, "Launching scaffold induction mailer dry-run"), "Induction wrapper routes through scaffold launcher"
AssertTrue OutputContains(inductionRun, "state=COMPOSED"), "Induction wrapper reports composed state"
AssertTrue OutputContains(inductionRun, "dry_run=true"), "Induction wrapper stays dry-run"

Dim watchdogRun
Set watchdogRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(fso.BuildPath(scaffoldRoot, "OutlookWatchdog.vbs")))
AssertTrue watchdogRun("exit_code") = 0, "Scaffold OutlookWatchdog wrapper exits successfully"
AssertTrue OutputContains(watchdogRun, "Launching scaffold watchdog report-only entrypoint"), "Watchdog wrapper routes through scaffold launcher"
AssertTrue OutputContains(watchdogRun, "report_only=true"), "Watchdog wrapper stays report-only by default"

Dim blockedWatchdogRun
Set blockedWatchdogRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(fso.BuildPath(scaffoldRoot, "OutlookWatchdog.vbs")) & " /repair")
AssertTrue blockedWatchdogRun("exit_code") = 1, "Scaffold OutlookWatchdog wrapper blocks repair requests"
AssertTrue OutputContains(blockedWatchdogRun, "state=FAILED_SAFE"), "Blocked watchdog request fails safe"

Dim validationRun
Set validationRun = RunCommand("cscript.exe //NoLogo " & QuoteArg(fso.BuildPath(scaffoldRoot, "ValidateWorkflow.vbs")))
AssertTrue validationRun("exit_code") = 0, "Scaffold ValidateWorkflow wrapper exits successfully"
AssertTrue OutputContains(validationRun, "Launching scaffold validation workflow"), "Validation wrapper routes through scaffold launcher"
AssertTrue OutputContains(validationRun, "Validation passed."), "Validation wrapper reports scaffold validation"

AssertFileSafe "RunGroupMeetings.vbs"
AssertFileSafe "RunInductionMailer.vbs"
AssertFileSafe "OutlookWatchdog.vbs"
AssertFileSafe "ValidateWorkflow.vbs"
AssertFileSafe "vbs\entrypoints\scaffold_launcher.vbs"

If failures > 0 Then
  WScript.Echo "Scaffold alternate entrypoints smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Scaffold alternate entrypoints smoke test passed."
