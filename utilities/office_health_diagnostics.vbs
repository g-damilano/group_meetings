'====================================================================
' office_health_diagnostics.vbs - Outlook health reporting helpers
'====================================================================
Option Explicit
SetLocale 1033

Sub OfficeAppendHealthAction(report, message)
  report("action_required_messages").Add CStr(report("action_required_messages").Count), CStr(message)
End Sub

Function OfficeBuildOutlookHealthReport(logPath)
  Dim report: Set report = OfficeCreateHealthReport()
  Dim app: Set app = OfficeGetOutlookApplication()
  If app Is Nothing Then
    report("blocked_or_auth_suspected") = True
    OfficeAppendHealthAction report, "Outlook.Application could not be created."
    Set OfficeBuildOutlookHealthReport = report
    Exit Function
  End If

  Dim transport: Set transport = OfficeGetTransportState(app)
  report("offline") = transport("offline")
  report("exchange_connection_mode") = transport("exchange_connection_mode")
  report("exchange_connection_mode_name") = transport("exchange_connection_mode_name")
  report("blocked_or_auth_suspected") = transport("blocked_or_auth_suspected")

  If report("blocked_or_auth_suspected") Then
    OfficeAppendHealthAction report, "Outlook appears offline/disconnected/auth-blocked. Reconnect Outlook, disable Work Offline, or complete O365 authentication."
  End If

  Dim ns: Set ns = app.GetNamespace("MAPI")
  Dim outbox, item, keyValue
  On Error Resume Next
  Set outbox = ns.GetDefaultFolder(OFFICE_OL_FOLDER_OUTBOX)
  If Err.Number <> 0 Then
    OfficeAppendHealthAction report, "Outbox could not be inspected: " & Err.Description
    Err.Clear
    Set OfficeBuildOutlookHealthReport = report
    On Error GoTo 0
    Exit Function
  End If
  report("outbox_count") = outbox.Items.Count
  On Error GoTo 0

  For Each item In outbox.Items
    If TypeName(item) = "MailItem" Then
      keyValue = OfficeGetManagedMailKey(item)
      If Len(keyValue) > 0 Then
        report("managed_outbox_count") = CLng(report("managed_outbox_count")) + 1
      Else
        report("unmanaged_outbox_count") = CLng(report("unmanaged_outbox_count")) + 1
      End If
    End If
  Next

  OfficeLogInfo logPath, "Outlook health report: offline=" & CStr(report("offline")) & _
    " mode=" & CStr(report("exchange_connection_mode")) & " (" & report("exchange_connection_mode_name") & ")" & _
    " outbox=" & CStr(report("outbox_count")) & _
    " managed_outbox=" & CStr(report("managed_outbox_count")) & _
    " unmanaged_outbox=" & CStr(report("unmanaged_outbox_count"))

  Set OfficeBuildOutlookHealthReport = report
End Function

