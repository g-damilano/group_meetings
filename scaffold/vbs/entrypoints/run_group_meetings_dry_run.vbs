Option Explicit
SetLocale 1033

' Dry-run group-meetings entrypoint.
' This orchestration script gathers command-line intent and calls the Stage 2
' workflow runner. It does not open Excel or Outlook and cannot send mail.

Dim SC_GM_EP_FSO
Set SC_GM_EP_FSO = CreateObject("Scripting.FileSystemObject")

Function SC_GM_EP_ScaffoldRoot()
  Dim scriptFolder
  scriptFolder = SC_GM_EP_FSO.GetParentFolderName(WScript.ScriptFullName)
  SC_GM_EP_ScaffoldRoot = SC_GM_EP_FSO.GetParentFolderName(SC_GM_EP_FSO.GetParentFolderName(scriptFolder))
End Function

Sub SC_GM_EP_LoadTextModule(modulePath)
  Dim ts
  Set ts = SC_GM_EP_FSO.OpenTextFile(modulePath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

Sub SC_GM_EP_LoadModules(scaffoldRoot)
  SC_GM_EP_LoadTextModule SC_GM_EP_FSO.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
  SC_LoadCoreModules
  SC_GM_EP_LoadTextModule SC_GM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\workflow_helpers.vbs")
  SC_GM_EP_LoadTextModule SC_GM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_workflow.vbs")
  SC_GM_EP_LoadTextModule SC_GM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_message_plan.vbs")
  SC_GM_EP_LoadTextModule SC_GM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_rota_adapter.vbs")
  SC_GM_EP_LoadTextModule SC_GM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\group_meetings_runner.vbs")
End Sub

Function SC_GM_EP_ArgValue(argName, defaultValue)
  SC_GM_EP_ArgValue = defaultValue
  On Error Resume Next
  If WScript.Arguments.Named.Exists(argName) Then SC_GM_EP_ArgValue = WScript.Arguments.Named(argName)
  Err.Clear
  On Error GoTo 0
End Function

Function SC_GM_EP_BoolText(value)
  If CBool(value) Then
    SC_GM_EP_BoolText = "true"
  Else
    SC_GM_EP_BoolText = "false"
  End If
End Function

Function SC_GM_EP_IsDryRunRequested()
  Dim dryRunText
  dryRunText = LCase(Trim(CStr(SC_GM_EP_ArgValue("dry_run", "true"))))
  SC_GM_EP_IsDryRunRequested = (dryRunText = "true" Or dryRunText = "yes" Or dryRunText = "1" Or dryRunText = "on")
End Function

Function SC_GM_EP_DefaultLogPath()
  SC_GM_EP_DefaultLogPath = SC_GM_EP_FSO.BuildPath(SC_GM_EP_FSO.GetSpecialFolder(2), "office_scaffold_group_meetings_dry_run.log")
End Function

Function SC_GM_EP_ReadRecipientDirectoryCsv(directoryPath)
  Dim directory
  Set directory = SC_CreateDictionary()

  If Len(Trim(CStr(directoryPath))) = 0 Then
    Set SC_GM_EP_ReadRecipientDirectoryCsv = directory
    Exit Function
  End If

  If Not SC_GM_EP_FSO.FileExists(CStr(directoryPath)) Then
    Set SC_GM_EP_ReadRecipientDirectoryCsv = directory
    Exit Function
  End If

  Dim text
  text = SC_ReadTextFile(directoryPath)
  Dim lines
  lines = Split(SC_NormalizeLineEndings(text), vbLf)

  Dim i
  For i = LBound(lines) To UBound(lines)
    Dim line
    line = Trim(CStr(lines(i)))
    If Len(line) > 0 And Left(line, 1) <> "#" Then
      Dim fields
      Set fields = SC_GM_SplitCsvLine(line)
      If fields.Count >= 2 Then
        If LCase(Trim(CStr(fields("0")))) <> "alias" Then
          If Len(Trim(CStr(fields("0")))) > 0 And Len(Trim(CStr(fields("1")))) > 0 Then
            directory(Trim(CStr(fields("0")))) = Trim(CStr(fields("1")))
          End If
        End If
      End If
    End If
  Next

  Set SC_GM_EP_ReadRecipientDirectoryCsv = directory
End Function

Sub SC_GM_EP_PrintUsage()
  WScript.Echo "usage=run_group_meetings_dry_run.vbs /rota:<csv> /date:<yyyy-mm-dd> [/directory:<csv>] [/group:<address>] [/manager:<address>] [/sender:<name>]"
  WScript.Echo "dry_run_only=true"
End Sub

Sub SC_GM_EP_PrintResult(result)
  WScript.Echo "workflow_reference=" & CStr(result("workflow_reference"))
  WScript.Echo "state=" & CStr(result("state"))
  WScript.Echo "reason=" & CStr(result("reason"))
  WScript.Echo "action_required=" & SC_GM_EP_BoolText(result("action_required"))
  WScript.Echo "dry_run=" & SC_GM_EP_BoolText(result("dry_run"))
  WScript.Echo "target_date=" & CStr(result("target_date"))
  WScript.Echo "rota_source=" & CStr(result("rota_source"))
  WScript.Echo "row_count=" & CStr(result("row_count"))
  WScript.Echo "plan_state=" & CStr(result("plan_state"))
  WScript.Echo "planned_count=" & CStr(result("planned_count"))
  WScript.Echo "composed_count=" & CStr(result("composed_count"))
  WScript.Echo "failed_count=" & CStr(result("failed_count"))
  WScript.Echo "skipped_count=" & CStr(result("skipped_count"))
End Sub

Sub SC_GM_EP_Main()
  Dim scaffoldRoot
  scaffoldRoot = SC_GM_EP_ScaffoldRoot()
  SC_GM_EP_LoadModules scaffoldRoot

  Dim rotaPath
  rotaPath = CStr(SC_GM_EP_ArgValue("rota", ""))
  Dim targetDateText
  targetDateText = CStr(SC_GM_EP_ArgValue("date", ""))
  If Len(Trim(rotaPath)) = 0 Or Len(Trim(targetDateText)) = 0 Or WScript.Arguments.Named.Exists("help") Then
    SC_GM_EP_PrintUsage
    WScript.Quit 64
  End If

  If WScript.Arguments.Named.Exists("live") Or WScript.Arguments.Named.Exists("send") Or Not SC_GM_EP_IsDryRunRequested() Then
    Dim blocked
    Set blocked = SC_GM_CreateRunnerResult()
    blocked("state") = SC_STATE_FAILED_SAFE
    blocked("reason") = "This entrypoint is dry-run only; live, send, or dry_run=false requests are blocked."
    blocked("action_required") = True
    blocked("dry_run") = False
    blocked("target_date") = targetDateText
    blocked("rota_source") = rotaPath
    SC_GM_EP_PrintResult blocked
    WScript.Quit 1
  End If

  Dim context
  Set context = SC_CreateGroupMeetingsWorkflowContext()
  context("dry_run") = True
  context("log_path") = CStr(SC_GM_EP_ArgValue("log", SC_GM_EP_DefaultLogPath()))

  Dim groupAddress
  groupAddress = CStr(SC_GM_EP_ArgValue("group", ""))
  Dim managerAddress
  managerAddress = CStr(SC_GM_EP_ArgValue("manager", ""))
  Dim senderName
  senderName = CStr(SC_GM_EP_ArgValue("sender", "Coordinator"))

  Dim defaults
  Set defaults = SC_GM_RunnerDefaults("group", "manager", senderName)

  Dim directory
  Set directory = SC_GM_EP_ReadRecipientDirectoryCsv(CStr(SC_GM_EP_ArgValue("directory", "")))
  If Len(Trim(groupAddress)) > 0 Then directory("group") = groupAddress
  If Len(Trim(managerAddress)) > 0 Then directory("manager") = managerAddress

  Dim templateContext
  Set templateContext = SC_CreateDictionary()

  Dim result
  Set result = SC_GM_RunGroupMeetingsCsvDryRun(context, rotaPath, SC_GM_ParseDateText(targetDateText), defaults, directory, templateContext)
  SC_GM_EP_PrintResult result

  If CStr(result("state")) = SC_STATE_FAILED_SAFE Then
    WScript.Quit 1
  ElseIf CBool(result("action_required")) Then
    WScript.Quit 2
  Else
    WScript.Quit 0
  End If
End Sub

SC_GM_EP_Main
