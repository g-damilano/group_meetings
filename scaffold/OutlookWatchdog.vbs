Option Explicit
SetLocale 1033

' Scaffold-local alternate for the legacy OutlookWatchdog.vbs wrapper.
' Routes to the scaffold report-only watchdog entrypoint by default.

Sub LoadScaffoldLauncher()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  Dim scriptFolder
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)

  Dim launcherPath
  launcherPath = fso.BuildPath(scriptFolder, "vbs\entrypoints\scaffold_launcher.vbs")
  If Not fso.FileExists(launcherPath) Then
    WScript.Echo "Fatal: scaffold launcher not found."
    WScript.Quit 1
  End If

  Dim ts
  Set ts = fso.OpenTextFile(launcherPath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

LoadScaffoldLauncher
SC_Launcher_QuitWithEntrypoint "vbs\entrypoints\run_watchdog_report_only.vbs", "scaffold watchdog report-only entrypoint"
