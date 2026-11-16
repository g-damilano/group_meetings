Option Explicit
SetLocale 1033

' Report-only watchdog entrypoint. Defaults to no live Outlook inspection so it
' is safe for validation. Scheduled live inspection can be explicitly enabled by
' passing /inspect_live:true in an operational environment.

Dim SC_WD_EP_FSO
Set SC_WD_EP_FSO = CreateObject("Scripting.FileSystemObject")

Function SC_WD_EP_ScaffoldRoot()
  Dim scriptFolder
  scriptFolder = SC_WD_EP_FSO.GetParentFolderName(WScript.ScriptFullName)
  SC_WD_EP_ScaffoldRoot = SC_WD_EP_FSO.GetParentFolderName(SC_WD_EP_FSO.GetParentFolderName(scriptFolder))
End Function

Sub SC_WD_EP_LoadTextModule(modulePath)
  Dim ts
  Set ts = SC_WD_EP_FSO.OpenTextFile(modulePath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

Sub SC_WD_EP_LoadModules(scaffoldRoot)
  SC_WD_EP_LoadTextModule SC_WD_EP_FSO.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
  SC_LoadCoreModules
  SC_WD_EP_LoadTextModule SC_WD_EP_FSO.BuildPath(scaffoldRoot, "vbs\watchdog\outlook_watchdog_sidecar.vbs")
End Sub

Function SC_WD_EP_ArgValue(argName, defaultValue)
  SC_WD_EP_ArgValue = defaultValue
  On Error Resume Next
  If WScript.Arguments.Named.Exists(argName) Then SC_WD_EP_ArgValue = WScript.Arguments.Named(argName)
  Err.Clear
  On Error GoTo 0
End Function

Function SC_WD_EP_BoolArg(argName, defaultValue)
  Dim text
  text = LCase(Trim(CStr(SC_WD_EP_ArgValue(argName, defaultValue))))
  SC_WD_EP_BoolArg = (text = "true" Or text = "yes" Or text = "1" Or text = "on")
End Function

Function SC_WD_EP_BoolText(value)
  If CBool(value) Then
    SC_WD_EP_BoolText = "true"
  Else
    SC_WD_EP_BoolText = "false"
  End If
End Function

Function SC_WD_EP_DefaultLogPath()
  SC_WD_EP_DefaultLogPath = SC_WD_EP_FSO.BuildPath(SC_WD_EP_FSO.GetSpecialFolder(2), "office_scaffold_watchdog_report_only.log")
End Function

Sub SC_WD_EP_PrintResult(result)
  WScript.Echo "workflow_reference=" & CStr(result("workflow_reference"))
  WScript.Echo "state=" & CStr(result("state"))
  WScript.Echo "reason=" & CStr(result("reason"))
  WScript.Echo "report_only=" & SC_WD_EP_BoolText(result("report_only"))
  WScript.Echo "action_required=" & SC_WD_EP_BoolText(result("action_required"))
  WScript.Echo "total_count=" & CStr(result("total_count"))
  WScript.Echo "managed_count=" & CStr(result("managed_count"))
  WScript.Echo "unmanaged_count=" & CStr(result("unmanaged_count"))
  WScript.Echo "queued_count=" & CStr(result("queued_count"))
  WScript.Echo "waiting_deferred_count=" & CStr(result("waiting_deferred_count"))
  WScript.Echo "stale_count=" & CStr(result("stale_count"))
  WScript.Echo "replace_eligible_count=" & CStr(result("replace_eligible_count"))
  WScript.Echo "action_required_count=" & CStr(result("action_required_count"))
  WScript.Echo "failed_safe_count=" & CStr(result("failed_safe_count"))
  WScript.Echo "repair_actions_taken=" & CStr(result("repair_actions_taken"))
End Sub

Sub SC_WD_EP_Main()
  Dim scaffoldRoot
  scaffoldRoot = SC_WD_EP_ScaffoldRoot()
  SC_WD_EP_LoadModules scaffoldRoot

  If WScript.Arguments.Named.Exists("repair") Or WScript.Arguments.Named.Exists("delete") Or _
     WScript.Arguments.Named.Exists("send") Or WScript.Arguments.Named.Exists("replace") Then
    Dim blocked
    Set blocked = SC_WD_CreateSidecarResult()
    blocked("state") = SC_STATE_FAILED_SAFE
    blocked("reason") = "This watchdog entrypoint is report-only; repair/delete/send/replace requests are blocked."
    blocked("action_required") = True
    SC_WD_EP_PrintResult blocked
    WScript.Quit 1
  End If

  Dim result
  If SC_WD_EP_BoolArg("inspect_live", "false") Then
    Dim report
    Set report = SC_BuildOutlookHealthReport(CStr(SC_WD_EP_ArgValue("log", SC_WD_EP_DefaultLogPath())), True)
    Set result = SC_WD_CreateSidecarResult()
    SC_WD_ApplyHealthSummary result, report
  Else
    Set result = SC_WD_CreateLiveInspectionPlaceholder()
  End If

  SC_WD_EP_PrintResult result
  If CStr(result("state")) = SC_STATE_FAILED_SAFE Then
    WScript.Quit 1
  ElseIf CBool(result("action_required")) Then
    WScript.Quit 2
  Else
    WScript.Quit 0
  End If
End Sub

SC_WD_EP_Main
