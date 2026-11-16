Option Explicit
SetLocale 1033

' Scheduled Outlook/O365 watchdog.
' Safety doctrine: inspect broadly, report clearly, modify narrowly, and never modify unmarked user mail.

Call BootstrapIncludeEntry()

Sub BootstrapIncludeEntry()
  Dim entryFso: Set entryFso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = entryFso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
  Dim entryPath: entryPath = entryFso.BuildPath(scriptFolder, "utilities\bootstrap_entry.vbs")
  If Not entryFso.FileExists(entryPath) Then
    WScript.Echo "Fatal: unable to locate utilities\bootstrap_entry.vbs"
    WScript.Quit 1
  End If
  Dim ts: Set ts = entryFso.OpenTextFile(entryPath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

LoadBootstrapModules Array( _
  "utilities\office_contracts.vbs", _
  "utilities\office_logger.vbs", _
  "utilities\office_session.vbs", _
  "utilities\office_managed_dispatch.vbs", _
  "utilities\office_health_diagnostics.vbs")

Dim logPath: logPath = OfficeDefaultLogPath("OutlookWatchdog.log")
Dim report: Set report = OfficeBuildOutlookHealthReport(logPath)

WScript.Echo "Outlook watchdog report"
WScript.Echo "Offline: " & CStr(report("offline"))
WScript.Echo "Exchange mode: " & CStr(report("exchange_connection_mode")) & " (" & report("exchange_connection_mode_name") & ")"
WScript.Echo "Blocked/auth suspected: " & CStr(report("blocked_or_auth_suspected"))
WScript.Echo "Outbox: " & CStr(report("outbox_count")) & " total, " & CStr(report("managed_outbox_count")) & " managed, " & CStr(report("unmanaged_outbox_count")) & " unmanaged"

If report("action_required_messages").Count > 0 Then
  Dim key
  For Each key In report("action_required_messages").Keys
    WScript.Echo "ACTION REQUIRED: " & report("action_required_messages")(key)
    OfficeLogActionRequired logPath, report("action_required_messages")(key)
  Next
End If

