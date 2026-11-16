Option Explicit
SetLocale 1033

' UI/orchestration wrapper for the induction mailer workflow.
' This script intentionally contains no Outlook COM logic.

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell: Set shell = CreateObject("WScript.Shell")
Dim scriptFolder: scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
If Len(scriptFolder) = 0 Then scriptFolder = "."

Dim workflowPath: workflowPath = fso.BuildPath(scriptFolder, "InductionMailer.vbs")
If Not fso.FileExists(workflowPath) Then
  WScript.Echo "Fatal: InductionMailer.vbs not found."
  WScript.Quit 1
End If

Dim command: command = "cscript.exe //nologo " & Chr(34) & workflowPath & Chr(34)
WScript.Echo "Launching induction mailer workflow..."
WScript.Quit shell.Run(command, 1, True)

