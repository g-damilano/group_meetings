' Report-only Outbox inspection helpers.
' These functions inspect broadly and report clearly, but never modify mail items.

Function SC_CreateOutboxInspectionReport()
  Dim report
  Set report = SC_CreateDictionary()
  report("total_count") = 0
  report("managed_count") = 0
  report("unmanaged_count") = 0
  report("queued_count") = 0
  report("waiting_deferred_count") = 0
  report("stale_count") = 0
  report("replace_eligible_count") = 0
  report("action_required_count") = 0
  report("failed_safe_count") = 0
  report("repair_actions_taken") = 0
  Set report("items") = SC_CreateDictionary()
  Set SC_CreateOutboxInspectionReport = report
End Function

Function SC_SafeMailProperty(mailItem, propertyName)
  SC_SafeMailProperty = ""
  On Error Resume Next
  Select Case LCase(CStr(propertyName))
    Case "subject"
      SC_SafeMailProperty = CStr(mailItem.Subject)
    Case "entryid"
      SC_SafeMailProperty = CStr(mailItem.EntryID)
    Case "creationtime"
      SC_SafeMailProperty = CStr(mailItem.CreationTime)
    Case "deferreddeliverytime"
      SC_SafeMailProperty = CStr(mailItem.DeferredDeliveryTime)
  End Select
  If Err.Number <> 0 Then
    Err.Clear
    SC_SafeMailProperty = ""
  End If
  On Error GoTo 0
End Function

Function SC_CreateOutboxItemReport(mailItem, managedKey)
  Dim itemReport
  Set itemReport = SC_CreateDictionary()
  itemReport("managed_key") = CStr(managedKey)
  itemReport("is_managed") = (Len(Trim(CStr(managedKey))) > 0)
  itemReport("subject") = SC_SafeMailProperty(mailItem, "subject")
  itemReport("entry_id") = SC_SafeMailProperty(mailItem, "entryid")
  itemReport("creation_time") = SC_SafeMailProperty(mailItem, "creationtime")
  itemReport("deferred_delivery_time") = SC_SafeMailProperty(mailItem, "deferreddeliverytime")
  itemReport("decision_state") = SC_STATE_NOT_REQUESTED
  itemReport("decision_reason") = "Unmanaged or unmarked Outbox item; report-only inspection."
  itemReport("replace_eligible") = False
  itemReport("would_modify") = False
  itemReport("action_required") = False
  Set SC_CreateOutboxItemReport = itemReport
End Function

Function SC_InspectOutboxItem(mailItem, transport, nowValue, graceMinutes)
  Dim managedKey
  managedKey = SC_GetManagedMailKey(mailItem)

  Dim itemReport
  Set itemReport = SC_CreateOutboxItemReport(mailItem, managedKey)

  If Len(Trim(CStr(managedKey))) = 0 Then
    Set SC_InspectOutboxItem = itemReport
    Exit Function
  End If

  Dim decision
  Set decision = SC_GetQueuedManagedMailDecisionFromTransport(transport, mailItem, managedKey, nowValue, graceMinutes)
  itemReport("decision_state") = CStr(decision("state"))
  itemReport("decision_reason") = CStr(decision("reason"))
  itemReport("action_required") = CBool(decision("action_required"))
  itemReport("replace_eligible") = SC_CanReplaceQueuedManagedMail(transport, mailItem, managedKey, nowValue, graceMinutes)
  itemReport("would_modify") = False

  Set SC_InspectOutboxItem = itemReport
End Function

Sub SC_AppendOutboxInspectionItem(report, itemReport)
  report("total_count") = CLng(report("total_count")) + 1
  report("items").Add CStr(report("items").Count), itemReport

  If CBool(itemReport("is_managed")) Then
    report("managed_count") = CLng(report("managed_count")) + 1
  Else
    report("unmanaged_count") = CLng(report("unmanaged_count")) + 1
  End If

  Select Case CStr(itemReport("decision_state"))
    Case SC_STATE_QUEUED_OUTBOX
      report("queued_count") = CLng(report("queued_count")) + 1
    Case SC_STATE_WAITING_DEFERRED
      report("waiting_deferred_count") = CLng(report("waiting_deferred_count")) + 1
    Case SC_STATE_STALE_MANAGED_OUTBOX
      report("stale_count") = CLng(report("stale_count")) + 1
    Case SC_STATE_ACTION_REQUIRED
      report("action_required_count") = CLng(report("action_required_count")) + 1
    Case SC_STATE_FAILED_SAFE
      report("failed_safe_count") = CLng(report("failed_safe_count")) + 1
  End Select

  If CBool(itemReport("replace_eligible")) Then report("replace_eligible_count") = CLng(report("replace_eligible_count")) + 1
End Sub

Sub SC_InspectOutboxObjectItems(report, outboxItems, transport, nowValue, graceMinutes)
  Dim objectTypeName
  objectTypeName = TypeName(outboxItems)

  If StrComp(objectTypeName, "Dictionary", vbTextCompare) = 0 Then
    Dim key
    For Each key In outboxItems.Keys
      If IsObject(outboxItems(key)) Then
        Dim dictItem
        Set dictItem = outboxItems(key)
        SC_AppendOutboxInspectionItem report, SC_InspectOutboxItem(dictItem, transport, nowValue, graceMinutes)
      End If
    Next
  Else
    Dim item
    For Each item In outboxItems
      If IsObject(item) Then SC_AppendOutboxInspectionItem report, SC_InspectOutboxItem(item, transport, nowValue, graceMinutes)
    Next
  End If
End Sub

Sub SC_InspectOutboxArrayItems(report, outboxItems, transport, nowValue, graceMinutes)
  Dim item
  For Each item In outboxItems
    If IsObject(item) Then SC_AppendOutboxInspectionItem report, SC_InspectOutboxItem(item, transport, nowValue, graceMinutes)
  Next
End Sub

Function SC_InspectOutboxItems(outboxItems, transport, nowValue, graceMinutes, logPath)
  Dim report
  Set report = SC_CreateOutboxInspectionReport()

  If IsObject(outboxItems) Then
    SC_InspectOutboxObjectItems report, outboxItems, transport, nowValue, graceMinutes
  ElseIf IsArray(outboxItems) Then
    SC_InspectOutboxArrayItems report, outboxItems, transport, nowValue, graceMinutes
  End If

  SC_LogInfo logPath, "Outbox inspection report: total=" & CStr(report("total_count")) & _
    " managed=" & CStr(report("managed_count")) & _
    " unmanaged=" & CStr(report("unmanaged_count")) & _
    " stale=" & CStr(report("stale_count")) & _
    " replace_eligible=" & CStr(report("replace_eligible_count")) & _
    " action_required=" & CStr(report("action_required_count")) & _
    " failed_safe=" & CStr(report("failed_safe_count"))

  Set SC_InspectOutboxItems = report
End Function

Function SC_InspectOutboxFolder(outboxFolder, transport, nowValue, graceMinutes, logPath)
  Dim report
  Set report = SC_CreateOutboxInspectionReport()

  On Error Resume Next
  Dim items
  Set items = outboxFolder.Items
  If Err.Number <> 0 Then
    Err.Clear
    Set SC_InspectOutboxFolder = report
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Set SC_InspectOutboxFolder = SC_InspectOutboxItems(items, transport, nowValue, graceMinutes, logPath)
End Function

