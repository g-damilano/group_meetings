' Outlook, Excel, and transport-state helpers.

Function SC_GetOutlookApplication()
  Dim app
  On Error Resume Next
  Set app = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("Outlook.Application")
  End If
  If Err.Number <> 0 Then
    Err.Clear
    Set app = Nothing
  End If
  On Error GoTo 0
  Set SC_GetOutlookApplication = app
End Function

Function SC_GetMAPI(outlookApp)
  Dim app
  If outlookApp Is Nothing Then
    Set app = SC_GetOutlookApplication()
  Else
    Set app = outlookApp
  End If

  If app Is Nothing Then
    Set SC_GetMAPI = Nothing
    Exit Function
  End If

  On Error Resume Next
  Dim ns
  Set ns = app.GetNamespace("MAPI")
  If Err.Number <> 0 Then
    Err.Clear
    Set ns = Nothing
  End If
  On Error GoTo 0
  Set SC_GetMAPI = ns
End Function

Function SC_GetExcelApplication(visible)
  Dim app
  On Error Resume Next
  Set app = CreateObject("Excel.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set app = Nothing
  Else
    app.Visible = CBool(visible)
  End If
  On Error GoTo 0
  Set SC_GetExcelApplication = app
End Function

Function SC_ExchangeConnectionModeName(modeValue)
  If Not IsNumeric(modeValue) Then
    SC_ExchangeConnectionModeName = "Unavailable"
    Exit Function
  End If

  Select Case CLng(modeValue)
    Case 0: SC_ExchangeConnectionModeName = "olNoExchange"
    Case 100: SC_ExchangeConnectionModeName = "olOffline"
    Case 200: SC_ExchangeConnectionModeName = "olCachedOffline"
    Case 300: SC_ExchangeConnectionModeName = "olDisconnected"
    Case 400: SC_ExchangeConnectionModeName = "olCachedDisconnected"
    Case 500: SC_ExchangeConnectionModeName = "olCachedConnectedHeaders"
    Case 600: SC_ExchangeConnectionModeName = "olCachedConnectedDrizzle"
    Case 700: SC_ExchangeConnectionModeName = "olCachedConnectedFull"
    Case 800: SC_ExchangeConnectionModeName = "olOnline"
    Case Else: SC_ExchangeConnectionModeName = "Unknown"
  End Select
End Function

Function SC_IsBlockedExchangeMode(modeValue)
  If Not IsNumeric(modeValue) Then
    SC_IsBlockedExchangeMode = False
    Exit Function
  End If

  Select Case CLng(modeValue)
    Case 100, 200, 300, 400
      SC_IsBlockedExchangeMode = True
    Case Else
      SC_IsBlockedExchangeMode = False
  End Select
End Function

Function SC_GetTransportState(outlookApp)
  Dim report
  Set report = SC_CreateHealthReport()

  Dim ns
  Set ns = SC_GetMAPI(outlookApp)
  If ns Is Nothing Then
    report("blocked_or_auth_suspected") = True
    report("exchange_connection_mode_name") = "Unavailable"
    Set SC_GetTransportState = report
    Exit Function
  End If

  On Error Resume Next
  Dim offlineValue
  offlineValue = ns.Offline
  If Err.Number = 0 Then
    report("offline") = CBool(offlineValue)
  Else
    report("offline") = ""
    Err.Clear
  End If

  Dim modeValue
  modeValue = ns.ExchangeConnectionMode
  If Err.Number = 0 Then
    report("exchange_connection_mode") = modeValue
    report("exchange_connection_mode_name") = SC_ExchangeConnectionModeName(modeValue)
  Else
    report("exchange_connection_mode") = ""
    report("exchange_connection_mode_name") = "Unavailable"
    Err.Clear
  End If
  On Error GoTo 0

  report("blocked_or_auth_suspected") = False
  If report("offline") = True Then report("blocked_or_auth_suspected") = True
  If SC_IsBlockedExchangeMode(report("exchange_connection_mode")) Then report("blocked_or_auth_suspected") = True

  Set SC_GetTransportState = report
End Function

