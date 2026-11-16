Option Explicit
SetLocale 1033

' Preferred scaffold command router.
' This file is orchestration only: it parses safe command intent and delegates
' to existing dry-run, report-only, validation, or release-check entrypoints.

Const SC_ROUTER_EXIT_SUCCESS = 0
Const SC_ROUTER_EXIT_COMMAND_FAILURE = 1
Const SC_ROUTER_EXIT_INVALID_ARGS = 2
Const SC_ROUTER_EXIT_UNSAFE_BLOCKED = 3
Const SC_ROUTER_EXIT_MISSING_INPUT = 4
Const SC_ROUTER_EXIT_INTERNAL_ERROR = 5

Dim SC_ROUTER_FSO
Set SC_ROUTER_FSO = CreateObject("Scripting.FileSystemObject")
Dim SC_ROUTER_SHELL
Set SC_ROUTER_SHELL = CreateObject("WScript.Shell")

Function SC_RouterRoot()
  SC_RouterRoot = SC_ROUTER_FSO.GetParentFolderName(WScript.ScriptFullName)
End Function

Function SC_RouterRepoRoot()
  SC_RouterRepoRoot = SC_ROUTER_FSO.GetParentFolderName(SC_RouterRoot())
End Function

Function SC_RouterQuoteArg(value)
  SC_RouterQuoteArg = Chr(34) & Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function SC_RouterBoolFalseLike(value)
  Dim text
  text = LCase(Trim(CStr(value)))
  SC_RouterBoolFalseLike = (text = "false" Or text = "no" Or text = "n" Or text = "0" Or text = "off")
End Function

Function SC_RouterOptionNeedsValue(name)
  Select Case LCase(CStr(name))
    Case "input", "recipients", "templates", "date", "directory", "group", "manager", "sender", "mode", "log"
      SC_RouterOptionNeedsValue = True
    Case Else
      SC_RouterOptionNeedsValue = False
  End Select
End Function

Function SC_RouterIsFlag(name)
  Select Case LCase(CStr(name))
    Case "help", "dry-run", "report-only", "strict-repo-boundary"
      SC_RouterIsFlag = True
    Case Else
      SC_RouterIsFlag = False
  End Select
End Function

Function SC_RouterIsUnsafeOption(name)
  Select Case LCase(CStr(name))
    Case "send", "live", "repair", "delete", "replace", "inspect-live", "no-dry-run"
      SC_RouterIsUnsafeOption = True
    Case Else
      SC_RouterIsUnsafeOption = False
  End Select
End Function

Sub SC_RouterDrainExecOutput(exec)
  Do While Not exec.StdOut.AtEndOfStream
    WScript.Echo exec.StdOut.ReadLine
  Loop
  Do While Not exec.StdErr.AtEndOfStream
    WScript.Echo exec.StdErr.ReadLine
  Loop
End Sub

Function SC_RouterRunCommand(commandText)
  On Error Resume Next
  Dim exec
  Set exec = SC_ROUTER_SHELL.Exec(commandText)
  If Err.Number <> 0 Then
    WScript.Echo "state=FAILED_SAFE"
    WScript.Echo "reason=Unable to launch delegated scaffold command: " & Err.Description
    Err.Clear
    SC_RouterRunCommand = SC_ROUTER_EXIT_INTERNAL_ERROR
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Do While exec.Status = 0
    SC_RouterDrainExecOutput exec
    WScript.Sleep 50
  Loop
  SC_RouterDrainExecOutput exec

  SC_RouterRunCommand = exec.ExitCode
End Function

Function SC_RouterRunScript(relativePath, argsText, displayName)
  Dim targetPath
  targetPath = SC_ROUTER_FSO.BuildPath(SC_RouterRoot(), CStr(relativePath))
  If Not SC_ROUTER_FSO.FileExists(targetPath) Then
    WScript.Echo "state=FAILED_SAFE"
    WScript.Echo "reason=Scaffold entrypoint not found: " & CStr(relativePath)
    SC_RouterRunScript = SC_ROUTER_EXIT_INTERNAL_ERROR
    Exit Function
  End If

  WScript.Echo "scaffold_command=" & CStr(displayName)
  WScript.Echo "route=" & CStr(relativePath)
  SC_RouterRunScript = SC_RouterRunCommand("cscript.exe //NoLogo " & SC_RouterQuoteArg(targetPath) & CStr(argsText))
End Function

Function SC_RouterRunPython(relativePath, argsText, displayName)
  Dim targetPath
  targetPath = SC_ROUTER_FSO.BuildPath(SC_RouterRoot(), CStr(relativePath))
  If Not SC_ROUTER_FSO.FileExists(targetPath) Then
    WScript.Echo "state=FAILED_SAFE"
    WScript.Echo "reason=Scaffold Python tool not found: " & CStr(relativePath)
    SC_RouterRunPython = SC_ROUTER_EXIT_INTERNAL_ERROR
    Exit Function
  End If

  WScript.Echo "scaffold_command=" & CStr(displayName)
  WScript.Echo "route=" & CStr(relativePath)
  SC_RouterRunPython = SC_RouterRunCommand("python -B " & SC_RouterQuoteArg(targetPath) & CStr(argsText))
End Function

Function SC_RouterNormalizeChildExit(exitCode)
  If CLng(exitCode) = 0 Then
    SC_RouterNormalizeChildExit = SC_ROUTER_EXIT_SUCCESS
  ElseIf CLng(exitCode) = SC_ROUTER_EXIT_INTERNAL_ERROR Then
    SC_RouterNormalizeChildExit = SC_ROUTER_EXIT_INTERNAL_ERROR
  Else
    SC_RouterNormalizeChildExit = SC_ROUTER_EXIT_COMMAND_FAILURE
  End If
End Function

Sub SC_RouterPrintHelp()
  WScript.Echo "Office Communication Automation Scaffold"
  WScript.Echo ""
  WScript.Echo "Usage:"
  WScript.Echo "  cscript.exe //NoLogo Scaffold.vbs help"
  WScript.Echo "  cscript.exe //NoLogo Scaffold.vbs validate"
  WScript.Echo "  cscript.exe //NoLogo Scaffold.vbs group-meetings --dry-run --input <rota.csv> [--date <yyyy-mm-dd>] [--directory <recipients.csv>]"
  WScript.Echo "  cscript.exe //NoLogo Scaffold.vbs induction --dry-run --recipients <recipients.csv> [--templates <manifest.csv>]"
  WScript.Echo "  cscript.exe //NoLogo Scaffold.vbs watchdog --report-only"
  WScript.Echo "  cscript.exe //NoLogo Scaffold.vbs release-check [--strict-repo-boundary]"
  WScript.Echo ""
  WScript.Echo "Safety defaults:"
  WScript.Echo "  group-meetings and induction run dry-run only through existing workflow entrypoints."
  WScript.Echo "  watchdog is report-only by default."
  WScript.Echo "  release-check treats parent repository boundary findings as advisory unless strict mode is requested."
  WScript.Echo "  live, send, repair, delete, replace, and inspect-live requests are blocked by this router."
  WScript.Echo ""
  WScript.Echo "Exit codes:"
  WScript.Echo "  0 success"
  WScript.Echo "  1 validation or delegated command failure"
  WScript.Echo "  2 invalid arguments or unknown command"
  WScript.Echo "  3 unsafe requested operation blocked"
  WScript.Echo "  4 missing required input"
  WScript.Echo "  5 internal scaffold error"
End Sub

Function SC_RouterNewParseResult()
  Dim parsed
  Set parsed = CreateObject("Scripting.Dictionary")
  parsed.Add "command", "help"
  parsed.Add "invalid", False
  parsed.Add "unsafe", False
  parsed.Add "missing", False
  parsed.Add "reason", ""
  Dim options
  Set options = CreateObject("Scripting.Dictionary")
  parsed.Add "options", options
  Set SC_RouterNewParseResult = parsed
End Function

Function SC_RouterParseArgs()
  Dim parsed
  Set parsed = SC_RouterNewParseResult()
  Dim options
  Set options = parsed("options")

  If WScript.Arguments.Count = 0 Then
    Set SC_RouterParseArgs = parsed
    Exit Function
  End If

  Dim commandText
  commandText = LCase(Trim(CStr(WScript.Arguments(0))))
  If commandText = "--help" Or commandText = "/?" Or commandText = "-h" Then
    commandText = "help"
  End If
  parsed("command") = commandText

  Dim i
  i = 1
  Do While i <= WScript.Arguments.Count - 1
    Dim token
    token = CStr(WScript.Arguments(i))
    If Left(token, 2) <> "--" Then
      parsed("invalid") = True
      parsed("reason") = "Unknown positional argument: " & token
      Exit Do
    End If

    Dim rawName
    Dim value
    Dim eqPos
    rawName = Mid(token, 3)
    value = ""
    eqPos = InStr(1, rawName, "=", vbTextCompare)
    If eqPos > 0 Then
      value = Mid(rawName, eqPos + 1)
      rawName = Left(rawName, eqPos - 1)
    End If

    Dim name
    name = Replace(LCase(Trim(rawName)), "_", "-")
    If SC_RouterIsUnsafeOption(name) Then
      parsed("unsafe") = True
      parsed("reason") = "Unsafe option requested: --" & name
      If Not options.Exists(name) Then options.Add name, "true"
    ElseIf SC_RouterOptionNeedsValue(name) Then
      If Len(value) = 0 Then
        If i + 1 > WScript.Arguments.Count - 1 Or Left(CStr(WScript.Arguments(i + 1)), 2) = "--" Then
          parsed("missing") = True
          parsed("reason") = "Missing value for --" & name
          Exit Do
        End If
        i = i + 1
        value = CStr(WScript.Arguments(i))
      End If
      If options.Exists(name) Then
        options(name) = value
      Else
        options.Add name, value
      End If
    ElseIf SC_RouterIsFlag(name) Then
      If Len(value) = 0 Then value = "true"
      If options.Exists(name) Then
        options(name) = value
      Else
        options.Add name, value
      End If
    Else
      parsed("invalid") = True
      parsed("reason") = "Unknown argument: --" & name
      Exit Do
    End If
    i = i + 1
  Loop

  Set SC_RouterParseArgs = parsed
End Function

Function SC_RouterHasOption(options, name)
  SC_RouterHasOption = options.Exists(LCase(CStr(name)))
End Function

Function SC_RouterOptionValue(options, name, defaultValue)
  If options.Exists(LCase(CStr(name))) Then
    SC_RouterOptionValue = options(LCase(CStr(name)))
  Else
    SC_RouterOptionValue = defaultValue
  End If
End Function

Sub SC_RouterQuitUnsafe(reason)
  WScript.Echo "state=FAILED_SAFE"
  WScript.Echo "reason=" & CStr(reason)
  WScript.Quit SC_ROUTER_EXIT_UNSAFE_BLOCKED
End Sub

Sub SC_RouterQuitMissing(reason)
  WScript.Echo "state=ACTION_REQUIRED"
  WScript.Echo "reason=" & CStr(reason)
  WScript.Quit SC_ROUTER_EXIT_MISSING_INPUT
End Sub

Sub SC_RouterQuitInvalid(reason)
  WScript.Echo "state=ACTION_REQUIRED"
  WScript.Echo "reason=" & CStr(reason)
  WScript.Quit SC_ROUTER_EXIT_INVALID_ARGS
End Sub

Function SC_RouterFirstCsvValue(path, headerName)
  SC_RouterFirstCsvValue = ""
  If Not SC_ROUTER_FSO.FileExists(CStr(path)) Then Exit Function

  Dim text
  Dim stream
  Set stream = SC_ROUTER_FSO.OpenTextFile(CStr(path), 1, False)
  text = stream.ReadAll
  stream.Close

  Dim lines
  lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
  If UBound(lines) < 1 Then Exit Function

  Dim header
  header = Replace(CStr(lines(0)), ChrW(&HFEFF), "")
  Dim headers
  headers = Split(header, ",")

  Dim targetIndex
  targetIndex = -1
  Dim i
  For i = LBound(headers) To UBound(headers)
    If LCase(Trim(CStr(headers(i)))) = LCase(CStr(headerName)) Then
      targetIndex = i
      Exit For
    End If
  Next
  If targetIndex < 0 Then Exit Function

  Dim rowIndex
  For rowIndex = 1 To UBound(lines)
    Dim line
    line = Trim(CStr(lines(rowIndex)))
    If Len(line) > 0 Then
      Dim values
      values = Split(line, ",")
      If targetIndex <= UBound(values) Then
        SC_RouterFirstCsvValue = Trim(CStr(values(targetIndex)))
        Exit Function
      End If
    End If
  Next
End Function

Sub SC_RouterRunValidate(options)
  If SC_RouterHasOption(options, "help") Then
    SC_RouterPrintHelp
    WScript.Quit SC_ROUTER_EXIT_SUCCESS
  End If
  Dim exitCode
  exitCode = SC_RouterRunScript("ValidateWorkflow.vbs", "", "validate")
  WScript.Quit SC_RouterNormalizeChildExit(exitCode)
End Sub

Sub SC_RouterRunGroupMeetings(options)
  If SC_RouterHasOption(options, "help") Then
    SC_RouterPrintHelp
    WScript.Quit SC_ROUTER_EXIT_SUCCESS
  End If
  If SC_RouterBoolFalseLike(SC_RouterOptionValue(options, "dry-run", "true")) Then
    SC_RouterQuitUnsafe "group-meetings is dry-run only; non-dry-run mode is blocked."
  End If

  Dim inputPath
  inputPath = CStr(SC_RouterOptionValue(options, "input", ""))
  If Len(Trim(inputPath)) = 0 Then SC_RouterQuitMissing "group-meetings requires --input <rota.csv>."
  If Not SC_ROUTER_FSO.FileExists(inputPath) Then SC_RouterQuitMissing "group-meetings input file was not found: " & inputPath

  Dim targetDate
  targetDate = CStr(SC_RouterOptionValue(options, "date", ""))
  If Len(Trim(targetDate)) = 0 Then
    targetDate = SC_RouterFirstCsvValue(inputPath, "meeting_date")
    If Len(Trim(targetDate)) > 0 Then WScript.Echo "default_date_source=first meeting_date from input"
  End If
  If Len(Trim(targetDate)) = 0 Then SC_RouterQuitMissing "group-meetings requires --date <yyyy-mm-dd> or an input meeting_date column."

  Dim argsText
  argsText = " /rota:" & SC_RouterQuoteArg(inputPath) & " /date:" & SC_RouterQuoteArg(targetDate) & " /dry_run:true"
  If Len(Trim(CStr(SC_RouterOptionValue(options, "directory", "")))) > 0 Then argsText = argsText & " /directory:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "directory", ""))
  If Len(Trim(CStr(SC_RouterOptionValue(options, "group", "")))) > 0 Then argsText = argsText & " /group:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "group", ""))
  If Len(Trim(CStr(SC_RouterOptionValue(options, "manager", "")))) > 0 Then argsText = argsText & " /manager:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "manager", ""))
  If Len(Trim(CStr(SC_RouterOptionValue(options, "sender", "")))) > 0 Then argsText = argsText & " /sender:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "sender", ""))
  If Len(Trim(CStr(SC_RouterOptionValue(options, "log", "")))) > 0 Then argsText = argsText & " /log:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "log", ""))

  Dim exitCode
  exitCode = SC_RouterRunScript("vbs\entrypoints\run_group_meetings_dry_run.vbs", argsText, "group-meetings")
  WScript.Quit SC_RouterNormalizeChildExit(exitCode)
End Sub

Sub SC_RouterRunInduction(options)
  If SC_RouterHasOption(options, "help") Then
    SC_RouterPrintHelp
    WScript.Quit SC_ROUTER_EXIT_SUCCESS
  End If
  If SC_RouterBoolFalseLike(SC_RouterOptionValue(options, "dry-run", "true")) Then
    SC_RouterQuitUnsafe "induction is dry-run only; non-dry-run mode is blocked."
  End If

  Dim recipientsPath
  recipientsPath = CStr(SC_RouterOptionValue(options, "recipients", ""))
  If Len(Trim(recipientsPath)) = 0 Then SC_RouterQuitMissing "induction requires --recipients <recipients.csv>."
  If Not SC_ROUTER_FSO.FileExists(recipientsPath) Then SC_RouterQuitMissing "induction recipients file was not found: " & recipientsPath

  Dim templatesPath
  templatesPath = CStr(SC_RouterOptionValue(options, "templates", ""))
  If Len(Trim(templatesPath)) = 0 Then
    templatesPath = SC_ROUTER_FSO.BuildPath(SC_RouterRoot(), "examples\workbooks\induction_template_manifest_example.csv")
    WScript.Echo "default_templates=" & templatesPath
  End If
  If Not SC_ROUTER_FSO.FileExists(templatesPath) Then SC_RouterQuitMissing "induction templates manifest was not found: " & templatesPath

  Dim argsText
  argsText = " /templates:" & SC_RouterQuoteArg(templatesPath) & " /recipients:" & SC_RouterQuoteArg(recipientsPath) & " /dry_run:true"
  If Len(Trim(CStr(SC_RouterOptionValue(options, "mode", "")))) > 0 Then
    argsText = argsText & " /mode:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "mode", ""))
  Else
    argsText = argsText & " /mode:all"
  End If
  If Len(Trim(CStr(SC_RouterOptionValue(options, "date", "")))) > 0 Then argsText = argsText & " /date:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "date", ""))
  If Len(Trim(CStr(SC_RouterOptionValue(options, "log", "")))) > 0 Then argsText = argsText & " /log:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "log", ""))

  Dim exitCode
  exitCode = SC_RouterRunScript("vbs\entrypoints\run_induction_mailer_dry_run.vbs", argsText, "induction")
  WScript.Quit SC_RouterNormalizeChildExit(exitCode)
End Sub

Sub SC_RouterRunWatchdog(options)
  If SC_RouterHasOption(options, "help") Then
    SC_RouterPrintHelp
    WScript.Quit SC_ROUTER_EXIT_SUCCESS
  End If
  If SC_RouterBoolFalseLike(SC_RouterOptionValue(options, "report-only", "true")) Then
    SC_RouterQuitUnsafe "watchdog is report-only through this router; live or repair mode is blocked."
  End If

  Dim argsText
  argsText = ""
  If Len(Trim(CStr(SC_RouterOptionValue(options, "log", "")))) > 0 Then argsText = argsText & " /log:" & SC_RouterQuoteArg(SC_RouterOptionValue(options, "log", ""))

  Dim exitCode
  exitCode = SC_RouterRunScript("vbs\entrypoints\run_watchdog_report_only.vbs", argsText, "watchdog")
  WScript.Quit SC_RouterNormalizeChildExit(exitCode)
End Sub

Sub SC_RouterRunReleaseCheck(options)
  If SC_RouterHasOption(options, "help") Then
    SC_RouterPrintHelp
    WScript.Quit SC_ROUTER_EXIT_SUCCESS
  End If

  Dim gateArgs
  gateArgs = " --root " & SC_RouterQuoteArg(SC_RouterRoot()) & " --repo-root " & SC_RouterQuoteArg(SC_RouterRepoRoot()) & " --fail-if-blocked"
  Dim exitCode
  exitCode = SC_RouterRunPython("python\release_gate.py", gateArgs, "release-check")
  If CLng(exitCode) <> 0 Then WScript.Quit SC_ROUTER_EXIT_COMMAND_FAILURE

  If SC_RouterHasOption(options, "strict-repo-boundary") Then
    WScript.Echo "strict_repo_boundary=true"
    Dim strictArgs
    strictArgs = " --repo-root " & SC_RouterQuoteArg(SC_RouterRepoRoot()) & " --scaffold-root " & SC_RouterQuoteArg(SC_RouterRoot()) & " --fail-on-findings"
    exitCode = SC_RouterRunPython("python\repo_boundary_audit.py", strictArgs, "strict-repo-boundary")
    If CLng(exitCode) <> 0 Then WScript.Quit SC_ROUTER_EXIT_COMMAND_FAILURE
  Else
    WScript.Echo "strict_repo_boundary=false"
    WScript.Echo "repo_boundary_findings=advisory"
  End If

  WScript.Quit SC_ROUTER_EXIT_SUCCESS
End Sub

Sub SC_RouterMain()
  Dim parsed
  Set parsed = SC_RouterParseArgs()
  Dim options
  Set options = parsed("options")

  If CBool(parsed("invalid")) Then SC_RouterQuitInvalid CStr(parsed("reason"))
  If CBool(parsed("missing")) Then SC_RouterQuitMissing CStr(parsed("reason"))
  If CBool(parsed("unsafe")) Then SC_RouterQuitUnsafe CStr(parsed("reason"))

  Dim commandText
  commandText = LCase(CStr(parsed("command")))

  Select Case commandText
    Case "help"
      SC_RouterPrintHelp
      WScript.Quit SC_ROUTER_EXIT_SUCCESS
    Case "validate"
      SC_RouterRunValidate options
    Case "group-meetings", "group", "meetings"
      SC_RouterRunGroupMeetings options
    Case "induction", "induction-mailer"
      SC_RouterRunInduction options
    Case "watchdog"
      SC_RouterRunWatchdog options
    Case "release-check", "release"
      SC_RouterRunReleaseCheck options
    Case Else
      SC_RouterQuitInvalid "Unknown command: " & commandText
  End Select
End Sub

SC_RouterMain
