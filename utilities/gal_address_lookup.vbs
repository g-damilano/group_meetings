Option Explicit
SetLocale 1033

Call BootstrapIncludeEntry()

Sub BootstrapIncludeEntry()
  Dim entryFso: Set entryFso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = entryFso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."

  Dim entryPath
  entryPath = entryFso.BuildPath(scriptFolder, "utilities\\bootstrap_entry.vbs")
  If Not entryFso.FileExists(entryPath) Then
    entryPath = entryFso.BuildPath(scriptFolder, "bootstrap_entry.vbs")
  End If

  If Not entryFso.FileExists(entryPath) Then
    WScript.Echo "Fatal: unable to locate utilities\\bootstrap_entry.vbs"
    WScript.Quit 1
  End If

  Dim entryStream
  On Error Resume Next
  Set entryStream = entryFso.OpenTextFile(entryPath, 1)
  If Err.Number <> 0 Then
    Dim entryError: entryError = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "Fatal: unable to read bootstrap entry: " & entryError
    WScript.Quit 1
  End If
  On Error GoTo 0

  Dim entryContents: entryContents = entryStream.ReadAll
  entryStream.Close

  ExecuteGlobal entryContents
  Set entryFso = Nothing
End Sub

Call Main()

Sub Main()
  Dim scriptFolder: scriptFolder = GetScriptFolder()
  Call RequireBootstrapModules(Array( _
    "common_console.vbs", _
    "mail_collections.vbs", _
    "mail_template_utils.vbs", _
    "common_outlook.vbs", _
    "mail_input_validation.vbs", _
    "MNS_loader.vbs"))
  Call LoadNameResolutionModules(scriptFolder)

  If Not EnsureConsoleHost() Then Exit Sub

  Dim count: count = WScript.Arguments.Count
  If count > 0 Then
    Dim i
    For i = 0 To count - 1
      Call DisplayResolution(Trim(WScript.Arguments(i)))
    Next
  Else
    Do
      Dim rawInput: rawInput = PromptNameResolutionLine("Enter a lookup value (prefix with name:/surname:/first:/last:/full:/email:, blank to quit): ")
      rawInput = Trim(rawInput)
      If Len(rawInput) = 0 Then Exit Do
      Call DisplayResolution(rawInput)
    Loop
  End If
End Sub

Sub RequireBootstrapModules(modulePaths)
  Dim loadError: loadError = ""

  On Error Resume Next
  Execute "Call LoadBootstrapModules(modulePaths)"
  If Err.Number <> 0 Then
    loadError = Err.Description
    Err.Clear
  End If
  On Error GoTo 0

  If Len(loadError) > 0 Then
    WScript.Echo "Fatal: unable to invoke LoadBootstrapModules: " & loadError
    WScript.Quit 1
  End If
End Sub

' Rely on shared name-resolution console helpers for parsing and display
