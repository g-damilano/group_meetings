' Induction mailer workflow skeleton.
' This module proves the reference workflow can build core contracts without owning
' low-level Outlook COM behaviour. Template migration belongs to Stage 3.

Const SC_INDUCTION_MAILER_WORKFLOW_NAME = "induction mailer"

Function SC_CreateInductionMailerWorkflowContext()
  Dim context
  Set context = SC_CreateWorkflowContext(SC_INDUCTION_MAILER_WORKFLOW_NAME)
  context("settings")("org_code") = "ORG"
  context("settings")("workflow_scope") = "induction-mailer"
  context("settings")("default_delivery_action") = "display"
  context("settings")("managed_outbox_grace_minutes") = SC_DEFAULT_OUTBOX_GRACE_MINUTES
  context("settings")("enable_stale_outbox_replacement") = False
  Set SC_CreateInductionMailerWorkflowContext = context
End Function

Function SC_CreateInductionMailerMessageRequest(logicalDate, messageMode, targetKey, toValue, subjectText, bodyHtml)
  Dim request
  Set request = SC_CreateWorkflowMessageRequest(logicalDate, messageMode, targetKey, toValue, subjectText, "", bodyHtml)
  request("workflow_reference") = "induction-mailer"
  SC_AddWorkflowRequestMetadata request, "REFERENCE_WORKFLOW", "induction-mailer"
  SC_AddWorkflowRequestMetadata request, "LOGICAL_DATE", SC_FormatKeyDate(logicalDate)
  SC_AddWorkflowRequestMetadata request, "MESSAGE_MODE", CStr(messageMode)
  Set SC_CreateInductionMailerMessageRequest = request
End Function

Function SC_BuildInductionMailerMessageSpec(workflowContext, request)
  Dim spec
  Set spec = SC_CreateWorkflowManagedMessageSpec( _
    workflowContext, _
    SC_WorkflowRequestValue(request, "logical_date", Date), _
    SC_WorkflowRequestValue(request, "message_type", "personal"), _
    SC_WorkflowRequestValue(request, "target_key", "recipient"), _
    SC_WorkflowRequestValue(request, "to", ""), _
    SC_WorkflowRequestValue(request, "subject", ""), _
    SC_WorkflowRequestValue(request, "body_text", ""), _
    SC_WorkflowRequestValue(request, "body_html", ""))

  spec("workflow_reference") = "induction-mailer"
  SC_ApplyWorkflowRequestToMessageSpec spec, request
  Set SC_BuildInductionMailerMessageSpec = spec
End Function

Function SC_PrepareInductionMailerMessage(workflowContext, request, recipientDirectory, templateContext)
  Dim spec
  Set spec = SC_BuildInductionMailerMessageSpec(workflowContext, request)

  Dim policy
  Set policy = SC_CreateInductionMailerDeliveryPolicy(workflowContext, request)

  Set SC_PrepareInductionMailerMessage = SC_PrepareWorkflowMessage(spec, policy, workflowContext, recipientDirectory, templateContext)
End Function

Function SC_CreateInductionMailerDeliveryPolicy(workflowContext, request)
  Dim policy
  Set policy = SC_CreateWorkflowDeliveryPolicy(workflowContext, SC_WorkflowDeliveryAction(workflowContext, request))

  Dim deferText
  deferText = CStr(SC_WorkflowRequestValue(request, "defer_minutes", ""))
  If IsNumeric(deferText) Then policy("defer_minutes") = CLng(deferText)

  Dim verifyValue
  verifyValue = SC_WorkflowRequestValue(request, "urgent_verify_after_submit", "")
  If Len(CStr(verifyValue)) > 0 Then policy("urgent_verify_after_submit") = CBool(verifyValue)

  Set SC_CreateInductionMailerDeliveryPolicy = policy
End Function

Function SC_RunInductionMailerMessage(workflowContext, request, recipientDirectory, templateContext)
  Dim spec
  Set spec = SC_BuildInductionMailerMessageSpec(workflowContext, request)

  Dim policy
  Set policy = SC_CreateInductionMailerDeliveryPolicy(workflowContext, request)

  Dim result
  Set result = SC_RunWorkflowMessage(spec, policy, workflowContext, recipientDirectory, templateContext)
  result("workflow_reference") = "induction-mailer"
  Set SC_RunInductionMailerMessage = result
End Function
