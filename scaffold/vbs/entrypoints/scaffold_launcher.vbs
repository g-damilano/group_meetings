Option Explicit
SetLocale 1033

' Shared launcher for scaffold-local production-shaped wrappers.
' It forwards arguments to scaffold entrypoints without owning Office COM logic.

Function SC_Launcher_QuoteArg(value)
  SC_Launcher_QuoteArg = Chr(34) & Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function SC_Launcher_Root()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  SC_Launcher_Root = fso.GetParentFolderName(WScript.ScriptFullName)
End Function

Function SC_Launcher_ForwardedArgs()
  Dim text
  text = ""

  Dim i
  For i = 0 To WScript.Arguments.Count - 1
    text = text & " " & SC_Launcher_QuoteArg(WScript.Arguments(i))
  Next

  SC_Launcher_ForwardedArgs = text
End Function

Sub SC_Launcher_DrainExecOutput(exec)
  Do While Not exec.StdOut.AtEndOfStream
    WScript.Echo exec.StdOut.ReadLine
  Loop
  Do While Not exec.StdErr.AtEndOfStream
    WScript.Echo exec.StdErr.ReadLine
  Loop
End Sub

Function SC_Launcher_RunEntrypoint(relativePath, displayName)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  Dim shell
  Set shell = CreateObject("WScript.Shell")

  Dim targetPath
  targetPath = fso.BuildPath(SC_Launcher_Root(), CStr(relativePath))

  If Not fso.FileExists(targetPath) Then
    WScript.Echo "Fatal: scaffold entrypoint not found: " & CStr(relativePath)
    SC_Launcher_RunEntrypoint = 1
    Exit Function
  End If

  Dim command
  command = "cscript.exe //NoLogo " & SC_Launcher_QuoteArg(targetPath) & SC_Launcher_ForwardedArgs()

  WScript.Echo "Launching " & CStr(displayName) & "..."

  Dim exec
  Set exec = shell.Exec(command)
  Do While exec.Status = 0
    SC_Launcher_DrainExecOutput exec
    WScript.Sleep 50
  Loop
  SC_Launcher_DrainExecOutput exec

  SC_Launcher_RunEntrypoint = exec.ExitCode
End Function

Sub SC_Launcher_QuitWithEntrypoint(relativePath, displayName)
  WScript.Quit SC_Launcher_RunEntrypoint(relativePath, displayName)
End Sub
