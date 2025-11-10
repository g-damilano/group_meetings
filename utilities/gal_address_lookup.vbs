Option Explicit
SetLocale 1033

Call Main()

Sub Main()
  If Not EnsureConsoleHost() Then Exit Sub

  Dim scriptFolder: scriptFolder = GetScriptFolder()
  Call LoadModule(scriptFolder, "mail_collections.vbs")
  Call LoadModule(scriptFolder, "mail_template_utils.vbs")
  Call LoadModule(scriptFolder, "mail_input_validation.vbs")
  Call LoadModule(scriptFolder, "mail_name_resolution.vbs")

  Dim count: count = WScript.Arguments.Count
  If count > 0 Then
    Dim i
    For i = 0 To count - 1
      Call DisplayResolution(Trim(WScript.Arguments(i)))
    Next
  Else
    Do
      Dim rawInput: rawInput = PromptLine("Enter a lookup value (prefix with name:/surname:/first:/last:/full:/email:, blank to quit): ")
      rawInput = Trim(rawInput)
      If Len(rawInput) = 0 Then Exit Do
      Call DisplayResolution(rawInput)
    Loop
  End If
End Sub

Function EnsureConsoleHost()
  Dim host: host = LCase(WScript.FullName)
  If InStr(host, "cscript.exe") = 0 Then
    WScript.Echo "Please run this script with cscript.exe to use the console interface."
    EnsureConsoleHost = False
  Else
    EnsureConsoleHost = True
  End If
End Function

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

Function PromptLine(prompt)
  WScript.StdOut.Write prompt
  PromptLine = WScript.StdIn.ReadLine()
End Function

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

Function DescribeBoolean(info, key)
  Dim defaultText: defaultText = "Unknown"
  If info Is Nothing Then
    DescribeBoolean = defaultText
    Exit Function
  End If

  If info.Exists(key) Then
    If CBool(info(key)) Then
      DescribeBoolean = "Found"
    Else
      DescribeBoolean = "Not found"
    End If
  Else
    DescribeBoolean = defaultText
  End If
End Function

Function DescribeField(info, key, placeholder)
  Dim text: text = placeholder
  If Not info Is Nothing Then
    If info.Exists(key) Then
      Dim candidate: candidate = SafeText(info(key))
      If Len(candidate) > 0 Then text = candidate
    End If
  End If
  DescribeField = text
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
