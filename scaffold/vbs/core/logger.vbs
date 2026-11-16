' Shared logging helpers for scaffold VBS modules.

Function SC_DefaultLogPath(fileName)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder
  folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  If Len(Trim(CStr(fileName))) = 0 Then fileName = "office_scaffold.log"
  SC_DefaultLogPath = fso.BuildPath(folder, fileName)
End Function

Sub SC_Log(logPath, levelText, message)
  On Error Resume Next
  Dim resolvedPath
  resolvedPath = Trim(CStr(logPath))
  If Len(resolvedPath) = 0 Then resolvedPath = SC_DefaultLogPath("office_scaffold.log")
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim ts
  Set ts = fso.OpenTextFile(resolvedPath, 8, True)
  ts.WriteLine Now & " - [" & UCase(Trim(CStr(levelText))) & "] " & CStr(message)
  ts.Close
  On Error GoTo 0
End Sub

Sub SC_LogInfo(logPath, message)
  SC_Log logPath, "INFO", message
End Sub

Sub SC_LogWarning(logPath, message)
  SC_Log logPath, "WARNING", message
End Sub

Sub SC_LogError(logPath, message)
  SC_Log logPath, "ERROR", message
End Sub

Sub SC_LogActionRequired(logPath, message)
  SC_Log logPath, "ACTION REQUIRED", message
End Sub

