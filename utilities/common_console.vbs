' Shared console helpers used across scripts.

Function EnsureConsoleHost()
  Dim host: host = LCase(WScript.FullName)
  If InStr(host, "cscript.exe") = 0 Then
    WScript.Echo "Please run this script with cscript.exe to use the console interface."
    EnsureConsoleHost = False
  Else
    EnsureConsoleHost = True
  End If
End Function

Function DescribeBoolean(info, key)
  Dim defaultText: defaultText = "Unknown"
  If info Is Nothing Then
    DescribeBoolean = defaultText
    Exit Function
  End If

  If info.Exists(key) Then
    On Error Resume Next
    Dim rawValue: rawValue = info(key)
    Dim boolValue: boolValue = False
    boolValue = CBool(rawValue)
    If Err.Number <> 0 Then
      Err.Clear
      DescribeBoolean = defaultText
    ElseIf boolValue Then
      DescribeBoolean = "Found"
    Else
      DescribeBoolean = "Not found"
    End If
    On Error GoTo 0
  Else
    DescribeBoolean = defaultText
  End If
End Function

Function DescribeField(info, key, placeholder)
  Dim text: text = placeholder
  If info Is Nothing Then
    DescribeField = text
    Exit Function
  End If

  If info.Exists(key) Then
    Dim candidate: candidate = ""
    On Error Resume Next
    candidate = SafeText(info(key))
    If Err.Number <> 0 Then
      Err.Clear
      candidate = CStr(info(key))
      If Err.Number <> 0 Then
        Err.Clear
        candidate = ""
      End If
    End If
    On Error GoTo 0

    If Len(candidate) > 0 Then text = candidate
  End If

  DescribeField = text
End Function
