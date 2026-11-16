'====================================================================
' office_logger.vbs - Shared logging helpers for Office workflows
'====================================================================
Option Explicit
SetLocale 1033

Function OfficeDefaultLogPath(fileName)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  If Len(Trim(CStr(fileName))) = 0 Then fileName = "office_automation.log"
  OfficeDefaultLogPath = fso.BuildPath(folder, fileName)
End Function

Sub OfficeLog(logPath, levelText, message)
  On Error Resume Next
  Dim resolvedPath: resolvedPath = Trim(CStr(logPath))
  If Len(resolvedPath) = 0 Then resolvedPath = OfficeDefaultLogPath("office_automation.log")
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim ts: Set ts = fso.OpenTextFile(resolvedPath, 8, True)
  ts.WriteLine Now & " - [" & UCase(Trim(CStr(levelText))) & "] " & CStr(message)
  ts.Close
  On Error GoTo 0
End Sub

Sub OfficeLogInfo(logPath, message)
  OfficeLog logPath, "INFO", message
End Sub

Sub OfficeLogWarning(logPath, message)
  OfficeLog logPath, "WARNING", message
End Sub

Sub OfficeLogError(logPath, message)
  OfficeLog logPath, "ERROR", message
End Sub

Sub OfficeLogActionRequired(logPath, message)
  OfficeLog logPath, "ACTION REQUIRED", message
End Sub

