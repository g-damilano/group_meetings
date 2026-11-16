' Induction mailer message-plan adapter.
' Builds broadcast and personal message requests while delegating composition and
' dispatch to core workflow facade helpers.

Const SC_IM_MODE_BROADCAST = "broadcast"
Const SC_IM_MODE_PERSONAL = "personal"
Const SC_IM_MODE_ALL = "all"

Function SC_IM_CreatePlan()
  Dim plan
  Set plan = SC_CreateDictionary()
  plan("workflow_reference") = "induction-mailer"
  plan("state") = SC_STATE_NOT_REQUESTED
  plan("reason") = ""
  Set plan("messages") = SC_CreateDictionary()
  Set plan("skips") = SC_CreateDictionary()
  Set SC_IM_CreatePlan = plan
End Function

Sub SC_IM_AddPlanMessage(plan, request, planCode)
  request("plan_code") = CStr(planCode)
  request("sequence") = plan("messages").Count + 1
  plan("messages").Add CStr(plan("messages").Count), request
End Sub

Sub SC_IM_AddPlanSkip(plan, skipCode, reason)
  Dim skipInfo
  Set skipInfo = SC_CreateDictionary()
  skipInfo("code") = CStr(skipCode)
  skipInfo("reason") = CStr(reason)
  plan("skips").Add CStr(plan("skips").Count), skipInfo
End Sub

Function SC_IM_TemplateModeIncluded(templateMode, selectedMode)
  Dim templateText
  templateText = LCase(Trim(CStr(templateMode)))
  Dim selectedText
  selectedText = LCase(Trim(CStr(selectedMode)))
  If Len(selectedText) = 0 Then selectedText = SC_IM_MODE_ALL
  SC_IM_TemplateModeIncluded = (selectedText = SC_IM_MODE_ALL Or selectedText = templateText)
End Function

Function SC_IM_ActiveRecipients(recipients)
  Dim active
  Set active = SC_CreateDictionary()

  If IsObject(recipients) Then
    Dim key
    For Each key In recipients.Keys
      Dim recipient
      Set recipient = recipients(key)
      If CBool(SC_IM_Value(recipient, "active", True)) Then active.Add CStr(active.Count), recipient
    Next
  End If

  Set SC_IM_ActiveRecipients = active
End Function

Function SC_IM_RecipientAlias(recipient)
  Dim value
  value = Trim(CStr(SC_IM_Value(recipient, "alias", "")))
  If Len(value) = 0 Then value = Trim(CStr(SC_IM_Value(recipient, "email", "")))
  SC_IM_RecipientAlias = value
End Function

Function SC_IM_RecipientContext(recipient, recipientCount)
  Dim context
  Set context = SC_CreateDictionary()
  context("ALIAS") = CStr(SC_IM_Value(recipient, "alias", ""))
  context("EMAIL") = CStr(SC_IM_Value(recipient, "email", ""))
  context("NAME") = CStr(SC_IM_Value(recipient, "first_name", ""))
  context("FIRST_NAME") = CStr(SC_IM_Value(recipient, "first_name", ""))
  context("SURNAME") = CStr(SC_IM_Value(recipient, "surname", SC_IM_Value(recipient, "last_name", "")))
  context("LAST_NAME") = CStr(SC_IM_Value(recipient, "last_name", ""))
  context("FULL_NAME") = CStr(SC_IM_Value(recipient, "full_name", ""))
  context("START_DATE") = CStr(SC_IM_Value(recipient, "start_date", ""))
  context("COHORT") = CStr(SC_IM_Value(recipient, "cohort", ""))
  context("RECIPIENT_COUNT") = CStr(recipientCount)
  Set SC_IM_RecipientContext = context
End Function

Function SC_IM_BroadcastContext(recipients)
  Dim context
  Set context = SC_CreateDictionary()
  context("RECIPIENT_COUNT") = CStr(recipients.Count)
  context("COHORT") = ""

  Dim key
  For Each key In recipients.Keys
    Dim recipient
    Set recipient = recipients(key)
    If Len(context("COHORT")) = 0 Then context("COHORT") = CStr(SC_IM_Value(recipient, "cohort", ""))
  Next

  Set SC_IM_BroadcastContext = context
End Function

Function SC_IM_JoinRecipientAliases(recipients)
  Dim list
  Set list = CreateObject("System.Collections.ArrayList")
  Dim key
  For Each key In recipients.Keys
    list.Add SC_IM_RecipientAlias(recipients(key))
  Next

  If list.Count = 0 Then
    SC_IM_JoinRecipientAliases = ""
  Else
    Dim values()
    ReDim values(list.Count - 1)
    Dim i
    For i = 0 To list.Count - 1
      values(i) = list(i)
    Next
    SC_IM_JoinRecipientAliases = Join(values, "; ")
  End If
End Function

Function SC_IM_TemplateSubject(templateRecord)
  Dim metadata
  Set metadata = SC_ParseTemplateMetadata(SC_ReadTextFile(CStr(templateRecord("template_path"))))
  If metadata.Exists("SUBJECT") Then
    SC_IM_TemplateSubject = CStr(metadata("SUBJECT"))
  Else
    SC_IM_TemplateSubject = CStr(SC_IM_Value(templateRecord, "display_name", "Induction message"))
  End If
End Function

Function SC_IM_CreateTemplateRequest(logicalDate, templateRecord, targetKey, toValue, context)
  Dim request
  Set request = SC_CreateInductionMailerMessageRequest( _
    logicalDate, _
    CStr(templateRecord("mode")), _
    targetKey, _
    toValue, _
    SC_IM_TemplateSubject(templateRecord), _
    "")
  SC_SetWorkflowRequestRecipients request, toValue, CStr(templateRecord("cc")), CStr(templateRecord("bcc"))
  SC_SetWorkflowRequestTemplate request, CStr(templateRecord("template_path")), CStr(templateRecord("inline_image_folder"))
  SC_SetWorkflowRequestAttachments request, CStr(templateRecord("attachments")), CStr(templateRecord("attachment_base_folder"))
  request("delivery_action") = CStr(templateRecord("delivery_action"))
  request("defer_minutes") = CStr(templateRecord("defer_minutes"))
  request("urgent_verify_after_submit") = CBool(templateRecord("urgent_verify_after_submit"))
  SC_AddWorkflowRequestMetadata request, "TEMPLATE_ID", CStr(templateRecord("template_id"))
  SC_AddWorkflowRequestMetadata request, "TEMPLATE_MODE", CStr(templateRecord("mode"))
  If Len(CStr(templateRecord("delivery_action"))) > 0 Then SC_AddWorkflowRequestMetadata request, "DELIVERY_ACTION", CStr(templateRecord("delivery_action"))
  Set request("template_context") = context
  Set SC_IM_CreateTemplateRequest = request
End Function

Function SC_IM_CreateMessagePlan(workflowContext, templates, recipients, selectedMode, logicalDate)
  Dim plan
  Set plan = SC_IM_CreatePlan()

  Dim activeRecipients
  Set activeRecipients = SC_IM_ActiveRecipients(recipients)
  If activeRecipients.Count = 0 Then
    plan("state") = SC_STATE_ACTION_REQUIRED
    plan("reason") = "No active recipients available."
    SC_IM_AddPlanSkip plan, "NO_ACTIVE_RECIPIENTS", "No active recipients available."
    Set SC_IM_CreateMessagePlan = plan
    Exit Function
  End If

  Dim key
  For Each key In templates.Keys
    Dim templateRecord
    Set templateRecord = templates(key)
    If Not CBool(templateRecord("enabled")) Then
      SC_IM_AddPlanSkip plan, "TEMPLATE_DISABLED", CStr(templateRecord("template_id"))
    ElseIf Not SC_IM_TemplateModeIncluded(templateRecord("mode"), selectedMode) Then
      SC_IM_AddPlanSkip plan, "TEMPLATE_MODE_NOT_SELECTED", CStr(templateRecord("template_id"))
    ElseIf Not CBool(SC_IM_Value(templateRecord, "template_exists", False)) Then
      SC_IM_AddPlanSkip plan, "TEMPLATE_MISSING", CStr(templateRecord("template_id"))
    ElseIf LCase(CStr(templateRecord("mode"))) = SC_IM_MODE_BROADCAST Then
      Dim broadcastContext
      Set broadcastContext = SC_IM_BroadcastContext(activeRecipients)
      Dim broadcastRequest
      Set broadcastRequest = SC_IM_CreateTemplateRequest( _
        logicalDate, _
        templateRecord, _
        CStr(templateRecord("target_key")), _
        SC_IM_JoinRecipientAliases(activeRecipients), _
        broadcastContext)
      SC_IM_AddPlanMessage plan, broadcastRequest, "broadcast:" & CStr(templateRecord("template_id"))
    Else
      Dim recipientKey
      For Each recipientKey In activeRecipients.Keys
        Dim recipient
        Set recipient = activeRecipients(recipientKey)
        Dim personalContext
        Set personalContext = SC_IM_RecipientContext(recipient, activeRecipients.Count)
        Dim target
        target = CStr(templateRecord("target_key")) & "-" & SC_IM_RecipientAlias(recipient)
        Dim personalRequest
        Set personalRequest = SC_IM_CreateTemplateRequest( _
          logicalDate, _
          templateRecord, _
          target, _
          SC_IM_RecipientAlias(recipient), _
          personalContext)
        SC_IM_AddPlanMessage plan, personalRequest, "personal:" & CStr(templateRecord("template_id"))
      Next
    End If
  Next

  If plan("messages").Count = 0 Then
    plan("state") = SC_STATE_ACTION_REQUIRED
    plan("reason") = "No induction messages were planned."
  Else
    plan("state") = SC_STATE_COMPOSED
    plan("reason") = "Induction mailer messages planned."
  End If

  Set SC_IM_CreateMessagePlan = plan
End Function

Function SC_IM_RunMessagePlan(workflowContext, plan, recipientDirectory, baseTemplateContext)
  Dim results
  Set results = SC_CreateDictionary()
  results("workflow_reference") = "induction-mailer"
  results("planned_count") = plan("messages").Count
  results("composed_count") = 0
  results("failed_count") = 0
  Set results("dispatch_results") = SC_CreateDictionary()

  Dim key
  For Each key In plan("messages").Keys
    Dim request
    Set request = plan("messages")(key)
    Dim context
    Set context = SC_CreateDictionary()
    If IsObject(baseTemplateContext) Then
      Dim baseKey
      For Each baseKey In baseTemplateContext.Keys
        context(baseKey) = baseTemplateContext(baseKey)
      Next
    End If
    If request.Exists("template_context") Then
      If IsObject(request("template_context")) Then
        Dim requestKey
        For Each requestKey In request("template_context").Keys
          context(requestKey) = request("template_context")(requestKey)
        Next
      End If
    End If

    Dim result
    Set result = SC_RunInductionMailerMessage(workflowContext, request, recipientDirectory, context)
    If result("state") = SC_STATE_COMPOSED Then
      results("composed_count") = CLng(results("composed_count")) + 1
    ElseIf result("state") = SC_STATE_FAILED_SAFE Then
      results("failed_count") = CLng(results("failed_count")) + 1
    End If
    results("dispatch_results").Add CStr(results("dispatch_results").Count), result
  Next

  Set SC_IM_RunMessagePlan = results
End Function
