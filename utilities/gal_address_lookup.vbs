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

Function ParseLookupExpression(text)
  Dim result: Set result = CreateObject("Scripting.Dictionary")
  result("Field") = "auto"
  result("Value") = ""

  Dim trimmed: trimmed = Trim(text)
  If Len(trimmed) = 0 Then
    Set ParseLookupExpression = result
    Exit Function
  End If

  Dim separator: separator = InStr(trimmed, ":")
  If separator = 0 Then separator = InStr(trimmed, "=")
  If separator > 0 Then
    Dim potentialField: potentialField = Trim(Left(trimmed, separator - 1))
    Dim potentialValue: potentialValue = Trim(Mid(trimmed, separator + 1))
    If Len(potentialField) > 0 And Len(potentialValue) > 0 Then
      result("Field") = LCase(potentialField)
      result("Value") = potentialValue
      Set ParseLookupExpression = result
      Exit Function
    End If
  End If

  If InStr(trimmed, "@") > 0 Then
    result("Field") = "email"
  ElseIf InStr(trimmed, " ") > 0 Or InStr(trimmed, ",") > 0 Then
    result("Field") = "name"
  Else
    result("Field") = "auto"
  End If
  result("Value") = trimmed

  Set ParseLookupExpression = result
End Function

Function DescribeLookupField(fieldName, lookupValue)
  Dim lowered: lowered = LCase(Trim(fieldName))
  Select Case lowered
    Case "email", "mail", "address"
      DescribeLookupField = "email"
    Case "first", "given", "forename"
      DescribeLookupField = "given name"
    Case "surname", "last", "lastname", "family", "familyname"
      DescribeLookupField = "surname"
    Case "display", "name"
      DescribeLookupField = "display name"
    Case "fullname", "full", "full name", "full-name", "full_name"
      DescribeLookupField = "full name"
    Case Else
      Dim trimmedValue: trimmedValue = Trim(CStr(lookupValue))
      If InStr(trimmedValue, "@") > 0 Then
        DescribeLookupField = "email"
      ElseIf InStr(trimmedValue, " ") > 0 Or InStr(trimmedValue, ",") > 0 Then
        DescribeLookupField = "display name"
      Else
        DescribeLookupField = "auto-detected field"
      End If
  End Select
End Function

Sub DisplayResolution(rawInput)
  If Len(rawInput) = 0 Then Exit Sub

  Dim parsed: Set parsed = ParseLookupExpression(rawInput)
  Dim lookupField: lookupField = "auto"
  Dim lookupValue: lookupValue = ""
  If Not parsed Is Nothing Then
    If parsed.Exists("Field") Then lookupField = parsed("Field")
    If parsed.Exists("Value") Then lookupValue = parsed("Value")
  End If

  lookupValue = Trim(lookupValue)
  If Len(lookupValue) = 0 Then Exit Sub

  Dim fieldDescription: fieldDescription = DescribeLookupField(lookupField, lookupValue)

  WScript.Echo ""
  WScript.Echo "Resolving '" & lookupValue & "' by " & fieldDescription & "..."

  Dim info: Set info = ResolveNameLookup(lookupValue, lookupField)

  If LCase(lookupField) = "email" Then
    On Error Resume Next
    Call MaybeRefineNameUsingGivenName(lookupValue, info)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
  End If

  If info Is Nothing Then
    WScript.Echo "  No resolution information returned."
    Exit Sub
  End If

  Dim foundText: foundText = DescribeBoolean(info, "Found")
  WScript.Echo "  Resolution status: " & foundText
  WScript.Echo "  Resolved name for " & fieldDescription & " '" & lookupValue & "':"
  WScript.Echo "    Directory source: " & DescribeField(info, "DirectorySource", "(none)")
  WScript.Echo "    Address lists checked: " & DescribeField(info, "DirectoryAttempted", "(none)")
  WScript.Echo "    First name: " & DescribeField(info, "First", "(empty)")
  WScript.Echo "    Surname: " & DescribeField(info, "Last", "(empty)")
  WScript.Echo "    Email: " & DescribeField(info, "Email", "(empty)")
  WScript.Echo "    Display name: " & DescribeField(info, "Full", "(empty)")
End Sub
