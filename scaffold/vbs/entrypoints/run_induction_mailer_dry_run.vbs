Option Explicit
SetLocale 1033

' Dry-run induction mailer entrypoint.
' This orchestration script gathers command-line intent and calls the Stage 3
' workflow runner. It does not open Outlook and cannot send mail.

Dim SC_IM_EP_FSO
Set SC_IM_EP_FSO = CreateObject("Scripting.FileSystemObject")

Function SC_IM_EP_ScaffoldRoot()
  Dim scriptFolder
  scriptFolder = SC_IM_EP_FSO.GetParentFolderName(WScript.ScriptFullName)
  SC_IM_EP_ScaffoldRoot = SC_IM_EP_FSO.GetParentFolderName(SC_IM_EP_FSO.GetParentFolderName(scriptFolder))
End Function

Sub SC_IM_EP_LoadTextModule(modulePath)
  Dim ts
  Set ts = SC_IM_EP_FSO.OpenTextFile(modulePath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

Sub SC_IM_EP_LoadModules(scaffoldRoot)
  SC_IM_EP_LoadTextModule SC_IM_EP_FSO.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
  SC_LoadCoreModules
  SC_IM_EP_LoadTextModule SC_IM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\workflow_helpers.vbs")
  SC_IM_EP_LoadTextModule SC_IM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_workflow.vbs")
  SC_IM_EP_LoadTextModule SC_IM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_input_adapter.vbs")
  SC_IM_EP_LoadTextModule SC_IM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_message_plan.vbs")
  SC_IM_EP_LoadTextModule SC_IM_EP_FSO.BuildPath(scaffoldRoot, "vbs\workflows\induction_mailer_runner.vbs")
End Sub

Function SC_IM_EP_ArgValue(argName, defaultValue)
  SC_IM_EP_ArgValue = defaultValue
  On Error Resume Next
  If WScript.Arguments.Named.Exists(argName) Then SC_IM_EP_ArgValue = WScript.Arguments.Named(argName)
  Err.Clear
  On Error GoTo 0
End Function

Function SC_IM_EP_IsDryRunRequested()
  Dim dryRunText
  dryRunText = LCase(Trim(CStr(SC_IM_EP_ArgValue("dry_run", "true"))))
  SC_IM_EP_IsDryRunRequested = (dryRunText = "true" Or dryRunText = "yes" Or dryRunText = "1" Or dryRunText = "on")
End Function

Function SC_IM_EP_DefaultLogPath()
  SC_IM_EP_DefaultLogPath = SC_IM_EP_FSO.BuildPath(SC_IM_EP_FSO.GetSpecialFolder(2), "office_scaffold_induction_mailer_dry_run.log")
End Function

Function SC_IM_EP_BoolText(value)
  If CBool(value) Then
    SC_IM_EP_BoolText = "true"
  Else
    SC_IM_EP_BoolText = "false"
  End If
End Function

Sub SC_IM_EP_PrintUsage()
  WScript.Echo "usage=run_induction_mailer_dry_run.vbs /templates:<manifest.csv> /recipients:<recipients.csv> [/mode:all|broadcast|personal] [/date:<yyyy-mm-dd>]"
  WScript.Echo "dry_run_only=true"
End Sub

Sub SC_IM_EP_PrintResult(result)
  WScript.Echo "workflow_reference=" & CStr(result("workflow_reference"))
  WScript.Echo "state=" & CStr(result("state"))
  WScript.Echo "reason=" & CStr(result("reason"))
  WScript.Echo "action_required=" & SC_IM_EP_BoolText(result("action_required"))
  WScript.Echo "dry_run=" & SC_IM_EP_BoolText(result("dry_run"))
  WScript.Echo "selected_mode=" & CStr(result("selected_mode"))
  WScript.Echo "recipient_count=" & CStr(result("recipient_count"))
  WScript.Echo "active_recipient_count=" & CStr(result("active_recipient_count"))
  WScript.Echo "template_count=" & CStr(result("template_count"))
  WScript.Echo "enabled_template_count=" & CStr(result("enabled_template_count"))
  WScript.Echo "enabled_existing_template_count=" & CStr(result("enabled_existing_template_count"))
  WScript.Echo "planned_count=" & CStr(result("planned_count"))
  WScript.Echo "composed_count=" & CStr(result("composed_count"))
  WScript.Echo "failed_count=" & CStr(result("failed_count"))
  WScript.Echo "skipped_count=" & CStr(result("skipped_count"))
End Sub

Sub SC_IM_EP_Main()
  Dim scaffoldRoot
  scaffoldRoot = SC_IM_EP_ScaffoldRoot()
  SC_IM_EP_LoadModules scaffoldRoot

  Dim templateManifestPath
  templateManifestPath = CStr(SC_IM_EP_ArgValue("templates", ""))
  Dim recipientPath
  recipientPath = CStr(SC_IM_EP_ArgValue("recipients", ""))
  If Len(Trim(templateManifestPath)) = 0 Or Len(Trim(recipientPath)) = 0 Or WScript.Arguments.Named.Exists("help") Then
    SC_IM_EP_PrintUsage
    WScript.Quit 64
  End If

  If WScript.Arguments.Named.Exists("live") Or WScript.Arguments.Named.Exists("send") Or Not SC_IM_EP_IsDryRunRequested() Then
    Dim blocked
    Set blocked = SC_IM_CreateRunnerResult()
    blocked("state") = SC_STATE_FAILED_SAFE
    blocked("reason") = "This entrypoint is dry-run only; live, send, or dry_run=false requests are blocked."
    blocked("action_required") = True
    blocked("dry_run") = False
    SC_IM_EP_PrintResult blocked
    WScript.Quit 1
  End If

  Dim context
  Set context = SC_CreateInductionMailerWorkflowContext()
  context("dry_run") = True
  context("log_path") = CStr(SC_IM_EP_ArgValue("log", SC_IM_EP_DefaultLogPath()))

  Dim selectedMode
  selectedMode = LCase(CStr(SC_IM_EP_ArgValue("mode", SC_IM_MODE_ALL)))
  Dim logicalDate
  logicalDate = SC_IM_EP_ArgValue("date", Date)

  Dim result
  Set result = SC_IM_RunInductionMailerCsvDryRun(context, templateManifestPath, recipientPath, selectedMode, logicalDate)
  SC_IM_EP_PrintResult result

  If CStr(result("state")) = SC_STATE_FAILED_SAFE Then
    WScript.Quit 1
  ElseIf CBool(result("action_required")) Then
    WScript.Quit 2
  Else
    WScript.Quit 0
  End If
End Sub

SC_IM_EP_Main
