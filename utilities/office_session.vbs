'====================================================================
' office_session.vbs - Shared Outlook/Excel session helpers
'====================================================================
Option Explicit
SetLocale 1033

Function OfficeGetOutlookApplication()
  Dim app
  On Error Resume Next
  Set app = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("Outlook.Application")
  End If
  On Error GoTo 0
  Set OfficeGetOutlookApplication = app
End Function

Function OfficeGetMAPI()
  Dim app: Set app = OfficeGetOutlookApplication()
  If app Is Nothing Then
    Set OfficeGetMAPI = Nothing
  Else
    On Error Resume Next
    Set OfficeGetMAPI = app.GetNamespace("MAPI")
    If Err.Number <> 0 Then
      Err.Clear
      Set OfficeGetMAPI = Nothing
    End If
    On Error GoTo 0
  End If
End Function

Function OfficeGetExcelApplication(visible)
  Dim xl
  On Error Resume Next
  Set xl = CreateObject("Excel.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set xl = Nothing
  Else
    xl.Visible = CBool(visible)
  End If
  On Error GoTo 0
  Set OfficeGetExcelApplication = xl
End Function

Function OfficeExchangeConnectionModeName(modeValue)
  If Not IsNumeric(modeValue) Then
    OfficeExchangeConnectionModeName = "Unavailable"
    Exit Function
  End If
  Select Case CLng(modeValue)
    Case 0: OfficeExchangeConnectionModeName = "olNoExchange"
    Case 100: OfficeExchangeConnectionModeName = "olOffline"
    Case 200: OfficeExchangeConnectionModeName = "olCachedOffline"
    Case 300: OfficeExchangeConnectionModeName = "olDisconnected"
    Case 400: OfficeExchangeConnectionModeName = "olCachedDisconnected"
    Case 500: OfficeExchangeConnectionModeName = "olCachedConnectedHeaders"
    Case 600: OfficeExchangeConnectionModeName = "olCachedConnectedDrizzle"
    Case 700: OfficeExchangeConnectionModeName = "olCachedConnectedFull"
    Case 800: OfficeExchangeConnectionModeName = "olOnline"
    Case Else: OfficeExchangeConnectionModeName = "Unknown"
  End Select
End Function

Function OfficeIsBlockedExchangeMode(modeValue)
  If Not IsNumeric(modeValue) Then
    OfficeIsBlockedExchangeMode = False
    Exit Function
  End If
  Select Case CLng(modeValue)
    Case 100, 200, 300, 400
      OfficeIsBlockedExchangeMode = True
    Case Else
      OfficeIsBlockedExchangeMode = False
  End Select
End Function

Function OfficeGetTransportState(outlookApp)
  Dim report: Set report = OfficeCreateHealthReport()
  Dim ns, offlineValue, modeValue

  On Error Resume Next
  Set ns = outlookApp.GetNamespace("MAPI")
  If Err.Number <> 0 Then
    Err.Clear
    report("blocked_or_auth_suspected") = True
    report("exchange_connection_mode_name") = "Unavailable"
    Set OfficeGetTransportState = report
    On Error GoTo 0
    Exit Function
  End If

  offlineValue = ns.Offline
  If Err.Number = 0 Then
    report("offline") = CBool(offlineValue)
  Else
    report("offline") = ""
    Err.Clear
  End If

  modeValue = ns.ExchangeConnectionMode
  If Err.Number = 0 Then
    report("exchange_connection_mode") = modeValue
    report("exchange_connection_mode_name") = OfficeExchangeConnectionModeName(modeValue)
  Else
    report("exchange_connection_mode") = ""
    report("exchange_connection_mode_name") = "Unavailable"
    Err.Clear
  End If
  On Error GoTo 0

  report("blocked_or_auth_suspected") = False
  If report("offline") = True Then report("blocked_or_auth_suspected") = True
  If OfficeIsBlockedExchangeMode(report("exchange_connection_mode")) Then report("blocked_or_auth_suspected") = True

  Set OfficeGetTransportState = report
End Function

