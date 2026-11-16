' Health diagnostics and watchdog-safe reporting helpers.

Sub SC_AppendHealthAction(report, message)
  report("action_required_messages").Add CStr(report("action_required_messages").Count), CStr(message)
End Sub

Function SC_BuildOutlookHealthReport(logPath, inspectOutbox)
  Dim report
  Set report = SC_CreateHealthReport()

  Dim app
  Set app = SC_GetOutlookApplication()
  If app Is Nothing Then
    report("blocked_or_auth_suspected") = True
    SC_AppendHealthAction report, "Outlook.Application could not be created."
    Set SC_BuildOutlookHealthReport = report
    Exit Function
  End If

  Dim transport
  Set transport = SC_GetTransportState(app)
  report("offline") = transport("offline")
  report("exchange_connection_mode") = transport("exchange_connection_mode")
  report("exchange_connection_mode_name") = transport("exchange_connection_mode_name")
  report("blocked_or_auth_suspected") = transport("blocked_or_auth_suspected")

  If report("blocked_or_auth_suspected") Then
    SC_AppendHealthAction report, "Outlook appears offline/disconnected/auth-blocked. Reconnect Outlook, disable Work Offline, or complete O365 authentication."
  End If

  If Not CBool(inspectOutbox) Then
    Set SC_BuildOutlookHealthReport = report
    Exit Function
  End If

  Dim ns
  Set ns = SC_GetMAPI(app)
  If ns Is Nothing Then
    SC_AppendHealthAction report, "Outlook MAPI namespace could not be inspected."
    Set SC_BuildOutlookHealthReport = report
    Exit Function
  End If

  Dim outbox
  Dim outboxAccess
  Set outboxAccess = SC_GetOutboxFolderSafe(ns, logPath)
  If Not CBool(outboxAccess("success")) Then
    SC_AppendHealthAction report, "Outbox could not be inspected: " & outboxAccess("reason")
    Set SC_BuildOutlookHealthReport = report
    Exit Function
  End If
  Set outbox = outboxAccess("folder")
  report("outbox_count") = outboxAccess("item_count")

  On Error Resume Next
  Dim items
  Set items = outbox.Items
  If Err.Number <> 0 Or items Is Nothing Then
    SC_AppendHealthAction report, "Outbox items could not be enumerated: " & Err.Description
    Err.Clear
    Set SC_BuildOutlookHealthReport = report
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Dim item
  For Each item In items
    If TypeName(item) = "MailItem" Then
      Dim keyValue
      keyValue = SC_GetManagedMailKey(item)
      If Len(keyValue) > 0 Then
        report("managed_outbox_count") = CLng(report("managed_outbox_count")) + 1
      Else
        report("unmanaged_outbox_count") = CLng(report("unmanaged_outbox_count")) + 1
      End If
    End If
  Next

  SC_LogInfo logPath, "Health report: offline=" & CStr(report("offline")) & _
    " mode=" & CStr(report("exchange_connection_mode")) & " (" & report("exchange_connection_mode_name") & ")" & _
    " outbox=" & CStr(report("outbox_count")) & _
    " managed_outbox=" & CStr(report("managed_outbox_count")) & _
    " unmanaged_outbox=" & CStr(report("unmanaged_outbox_count"))

  Set SC_BuildOutlookHealthReport = report
End Function
