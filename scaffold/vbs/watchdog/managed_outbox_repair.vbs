' Guarded managed Outbox repair primitives.
' This Stage 6 module contains the narrow modification path. Tests use fake mail
' items only; the live watchdog entrypoint remains report-only.

Function SC_WD_CreateRepairResult()
  Dim result
  Set result = SC_CreateDictionary()
  result("workflow_reference") = "outlook-watchdog"
  result("state") = SC_STATE_NOT_REQUESTED
  result("reason") = ""
  result("commit_repair") = False
  result("repair_actions_taken") = 0
  result("repair_available_count") = 0
  result("action_required") = False
  result("failed_safe_count") = 0
  result("unmanaged_count") = 0
  result("managed_count") = 0
  Set result("actions") = SC_CreateDictionary()
  Set SC_WD_CreateRepairResult = result
End Function

Function SC_WD_CreateRepairAction(itemReport)
  Dim action
  Set action = SC_CreateDictionary()
  action("managed_key") = CStr(itemReport("managed_key"))
  action("subject") = CStr(itemReport("subject"))
  action("decision_state") = CStr(itemReport("decision_state"))
  action("state") = SC_STATE_NOT_REQUESTED
  action("reason") = ""
  action("repair_available") = False
  action("replacement_created") = False
  action("stale_deleted") = False
  action("would_modify") = False
  Set SC_WD_CreateRepairAction = action
End Function

Sub SC_WD_AppendRepairAction(result, action)
  result("actions").Add CStr(result("actions").Count), action
  If CBool(action("repair_available")) Then result("repair_available_count") = CLng(result("repair_available_count")) + 1
  If CStr(action("state")) = SC_STATE_REPLACED_MANAGED_OUTBOX Then result("repair_actions_taken") = CLng(result("repair_actions_taken")) + 1
  If CStr(action("state")) = SC_STATE_FAILED_SAFE Then result("failed_safe_count") = CLng(result("failed_safe_count")) + 1
End Sub

Function SC_WD_CreateReplacementMail(replacementFactory, staleMail, itemReport)
  Set SC_WD_CreateReplacementMail = Nothing
  If Not IsObject(replacementFactory) Then Exit Function

  On Error Resume Next
  Set SC_WD_CreateReplacementMail = replacementFactory.CreateReplacement(staleMail, itemReport)
  If Err.Number <> 0 Then
    Err.Clear
    Set SC_WD_CreateReplacementMail = Nothing
  End If
  On Error GoTo 0
End Function

Function SC_WD_DeleteManagedMail(mailItem)
  SC_WD_DeleteManagedMail = False
  On Error Resume Next
  mailItem.Delete
  If Err.Number = 0 Then
    SC_WD_DeleteManagedMail = True
  Else
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SC_WD_RepairSingleOutboxItem(mailItem, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
  Dim itemReport
  Set itemReport = SC_InspectOutboxItem(mailItem, transport, nowValue, graceMinutes)

  Dim action
  Set action = SC_WD_CreateRepairAction(itemReport)

  If Not CBool(itemReport("is_managed")) Then
    action("state") = SC_STATE_NOT_REQUESTED
    action("reason") = "Unmarked Outbox item; repair is not allowed."
    Set SC_WD_RepairSingleOutboxItem = action
    Exit Function
  End If

  If Not CBool(itemReport("replace_eligible")) Then
    action("state") = CStr(itemReport("decision_state"))
    action("reason") = CStr(itemReport("decision_reason"))
    Set SC_WD_RepairSingleOutboxItem = action
    Exit Function
  End If

  action("repair_available") = True
  action("would_modify") = True

  If Not CBool(commitRepair) Then
    action("state") = SC_STATE_STALE_MANAGED_OUTBOX
    action("reason") = "Managed item is replace-eligible; repair was not committed."
    Set SC_WD_RepairSingleOutboxItem = action
    Exit Function
  End If

  Dim replacement
  Set replacement = SC_WD_CreateReplacementMail(replacementFactory, mailItem, itemReport)
  If replacement Is Nothing Then
    action("state") = SC_STATE_FAILED_SAFE
    action("reason") = "Replacement factory did not return a replacement mail item."
    Set SC_WD_RepairSingleOutboxItem = action
    Exit Function
  End If

  If SC_GetManagedMailKey(replacement) <> CStr(itemReport("managed_key")) Then
    action("state") = SC_STATE_FAILED_SAFE
    action("reason") = "Replacement mail item does not carry the expected managed key."
    Set SC_WD_RepairSingleOutboxItem = action
    Exit Function
  End If

  action("replacement_created") = True
  If SC_WD_DeleteManagedMail(mailItem) Then
    action("stale_deleted") = True
    action("state") = SC_STATE_REPLACED_MANAGED_OUTBOX
    action("reason") = "Stale scaffold-managed Outbox item replaced."
  Else
    action("state") = SC_STATE_FAILED_SAFE
    action("reason") = "Replacement was created but stale managed item could not be deleted."
  End If

  Set SC_WD_RepairSingleOutboxItem = action
End Function

Sub SC_WD_RepairObjectItems(result, outboxItems, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
  If StrComp(TypeName(outboxItems), "Dictionary", vbTextCompare) = 0 Then
    Dim key
    For Each key In outboxItems.Keys
      If IsObject(outboxItems(key)) Then
        Dim dictAction
        Set dictAction = SC_WD_RepairSingleOutboxItem(outboxItems(key), transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
        If Len(CStr(dictAction("managed_key"))) > 0 Then
          result("managed_count") = CLng(result("managed_count")) + 1
        Else
          result("unmanaged_count") = CLng(result("unmanaged_count")) + 1
        End If
        SC_WD_AppendRepairAction result, dictAction
      End If
    Next
  Else
    Dim item
    For Each item In outboxItems
      If IsObject(item) Then
        Dim objectAction
        Set objectAction = SC_WD_RepairSingleOutboxItem(item, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
        If Len(CStr(objectAction("managed_key"))) > 0 Then
          result("managed_count") = CLng(result("managed_count")) + 1
        Else
          result("unmanaged_count") = CLng(result("unmanaged_count")) + 1
        End If
        SC_WD_AppendRepairAction result, objectAction
      End If
    Next
  End If
End Sub

Sub SC_WD_RepairArrayItems(result, outboxItems, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
  Dim item
  For Each item In outboxItems
    If IsObject(item) Then
      Dim action
      Set action = SC_WD_RepairSingleOutboxItem(item, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
      If Len(CStr(action("managed_key"))) > 0 Then
        result("managed_count") = CLng(result("managed_count")) + 1
      Else
        result("unmanaged_count") = CLng(result("unmanaged_count")) + 1
      End If
      SC_WD_AppendRepairAction result, action
    End If
  Next
End Sub

Sub SC_WD_FinalizeRepairResult(result)
  If CLng(result("failed_safe_count")) > 0 Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "One or more managed repair actions failed safe."
    result("action_required") = True
  ElseIf CLng(result("repair_actions_taken")) > 0 Then
    result("state") = SC_STATE_REPLACED_MANAGED_OUTBOX
    result("reason") = "Stale scaffold-managed Outbox items were replaced."
    result("action_required") = False
  ElseIf CLng(result("repair_available_count")) > 0 Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Managed repair is available but was not committed."
    result("action_required") = True
  Else
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Managed repair scan completed; no repair actions were eligible."
    result("action_required") = False
  End If
End Sub

Function SC_WD_RunManagedRepair(outboxItems, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath)
  Dim result
  Set result = SC_WD_CreateRepairResult()
  result("commit_repair") = CBool(commitRepair)

  If IsObject(outboxItems) Then
    SC_WD_RepairObjectItems result, outboxItems, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath
  ElseIf IsArray(outboxItems) Then
    SC_WD_RepairArrayItems result, outboxItems, transport, nowValue, graceMinutes, replacementFactory, commitRepair, logPath
  End If

  SC_WD_FinalizeRepairResult result
  SC_LogInfo logPath, "Managed repair scan: managed=" & CStr(result("managed_count")) & _
    " unmanaged=" & CStr(result("unmanaged_count")) & _
    " available=" & CStr(result("repair_available_count")) & _
    " taken=" & CStr(result("repair_actions_taken")) & _
    " failed_safe=" & CStr(result("failed_safe_count"))

  Set SC_WD_RunManagedRepair = result
End Function
