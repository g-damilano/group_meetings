Option Explicit
SetLocale 1033

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

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim root: root = GetScriptFolder()
Dim failures: failures = 0

Sub RequireFile(relativePath)
  Dim path: path = fso.BuildPath(root, relativePath)
  If fso.FileExists(path) Then
    WScript.Echo "OK: " & relativePath
  Else
    WScript.Echo "MISSING: " & relativePath
    failures = failures + 1
  End If
End Sub

RequireFile "docs\ARCHITECTURE.md"
RequireFile "docs\CONTRACTS.md"
RequireFile "docs\DISPATCH_STATE_MACHINE.md"
RequireFile "docs\WATCHDOG_SAFETY_DOCTRINE.md"
RequireFile "docs\MIGRATION_SEQUENCE.md"
RequireFile "contracts\workflow_schema.json"
RequireFile "fixtures\dispatch_states\states.json"
RequireFile "RunGroupMeetings.vbs"
RequireFile "RunInductionMailer.vbs"
RequireFile "OutlookWatchdog.vbs"
RequireFile "utilities\office_contracts.vbs"
RequireFile "utilities\office_logger.vbs"
RequireFile "utilities\office_session.vbs"
RequireFile "utilities\office_managed_dispatch.vbs"
RequireFile "utilities\office_health_diagnostics.vbs"

If Not OfficeIsKnownDispatchState(OFFICE_STATE_SENT_CONFIRMED) Then
  WScript.Echo "FAILED: dispatch state constants are not available."
  failures = failures + 1
End If

If OfficeBuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "team") <> "ORG|WORKFLOW|2026-05-29|REGULAR|TEAM" Then
  WScript.Echo "FAILED: managed email key generation is not deterministic."
  failures = failures + 1
End If

If failures > 0 Then
  WScript.Echo "Validation failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Validation passed."
