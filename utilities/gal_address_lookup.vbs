Option Explicit
SetLocale 1033

Call Main()

Sub Main()
  Dim scriptFolder: scriptFolder = GetScriptFolder()
  Call LoadModule(scriptFolder, "common_console.vbs")
  Call LoadModule(scriptFolder, "mail_collections.vbs")
  Call LoadModule(scriptFolder, "mail_template_utils.vbs")
  Call LoadModule(scriptFolder, "mail_input_validation.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_prompt.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_text.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_directory.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_core.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution_console.vbs")

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

Function GetScriptFolder()
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  GetScriptFolder = folder
End Function

Sub LoadModule(baseFolder, relativePath)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim modulePath: modulePath = fso.BuildPath(baseFolder, relativePath)
  If Not fso.FileExists(modulePath) Then
    WScript.Echo "Warning: utility module not found: " & relativePath
    Exit Sub
  End If

  Dim ts: Set ts = fso.OpenTextFile(modulePath, 1)
  Dim contents: contents = ts.ReadAll
  ts.Close

  contents = RemoveByteOrderMark(contents)

  On Error Resume Next
  ExecuteGlobal contents
  If Err.Number <> 0 Then
    WScript.Echo "Warning: failed to load module '" & relativePath & "': " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Function RemoveByteOrderMark(text)
  RemoveByteOrderMark = text
  If Len(text) = 0 Then Exit Function

  If Len(text) >= 3 Then
    Dim leading: leading = Mid(text, 1, 3)
    If StrComp(leading, Chr(239) & Chr(187) & Chr(191), vbBinaryCompare) = 0 Then
      RemoveByteOrderMark = Mid(text, 4)
      Exit Function
    End If
  End If

  Dim firstChar: firstChar = Mid(text, 1, 1)
  If Len(firstChar) = 0 Then Exit Function

  On Error Resume Next
  Dim codePoint: codePoint = AscW(firstChar)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If codePoint = &HFEFF Or codePoint = &HFFFE Then
    RemoveByteOrderMark = Mid(text, 2)
  End If
End Function


Function GetOutlookApp()
  Dim app
  On Error Resume Next
  Set app = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("Outlook.Application")
  End If
  On Error GoTo 0
  Set GetOutlookApp = app
End Function
