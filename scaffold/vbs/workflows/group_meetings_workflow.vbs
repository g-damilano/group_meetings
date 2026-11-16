' Group meetings workflow adapter.
' This module builds core contracts without owning low-level Outlook COM behaviour.
' Rota parsing and production workbook compatibility are migrated incrementally in Stage 2.

Const SC_GROUP_MEETINGS_WORKFLOW_NAME = "group meetings"

Function SC_CreateGroupMeetingsWorkflowContext()
  Dim context
  Set context = SC_CreateWorkflowContext(SC_GROUP_MEETINGS_WORKFLOW_NAME)
  context("settings")("org_code") = "ORG"
  context("settings")("workflow_scope") = "group-meetings"
  context("settings")("default_delivery_action") = "display"
  context("settings")("managed_outbox_grace_minutes") = SC_DEFAULT_OUTBOX_GRACE_MINUTES
  context("settings")("enable_stale_outbox_replacement") = False
  Set SC_CreateGroupMeetingsWorkflowContext = context
End Function

Function SC_CreateGroupMeetingsMessageRequest(logicalDate, messageType, targetKey, toValue, subjectText, bodyHtml)
  Dim request
  Set request = SC_CreateWorkflowMessageRequest(logicalDate, messageType, targetKey, toValue, subjectText, "", bodyHtml)
  request("workflow_reference") = "group-meetings"
  SC_AddWorkflowRequestMetadata request, "REFERENCE_WORKFLOW", "group-meetings"
  SC_AddWorkflowRequestMetadata request, "LOGICAL_DATE", SC_FormatKeyDate(logicalDate)
  SC_AddWorkflowRequestMetadata request, "MESSAGE_TYPE", CStr(messageType)
  Set SC_CreateGroupMeetingsMessageRequest = request
End Function

Function SC_BuildGroupMeetingsMessageSpec(workflowContext, request)
  Dim spec
  Set spec = SC_CreateWorkflowManagedMessageSpec( _
    workflowContext, _
    SC_WorkflowRequestValue(request, "logical_date", Date), _
    SC_WorkflowRequestValue(request, "message_type", "regular"), _
    SC_WorkflowRequestValue(request, "target_key", "team"), _
    SC_WorkflowRequestValue(request, "to", ""), _
    SC_WorkflowRequestValue(request, "subject", ""), _
    SC_WorkflowRequestValue(request, "body_text", ""), _
    SC_WorkflowRequestValue(request, "body_html", ""))

  spec("workflow_reference") = "group-meetings"
  SC_ApplyWorkflowRequestToMessageSpec spec, request
  Set SC_BuildGroupMeetingsMessageSpec = spec
End Function

Function SC_PrepareGroupMeetingsMessage(workflowContext, request, recipientDirectory, templateContext)
  Dim spec
  Set spec = SC_BuildGroupMeetingsMessageSpec(workflowContext, request)

  Dim policy
  Set policy = SC_CreateWorkflowDeliveryPolicy(workflowContext, SC_WorkflowDeliveryAction(workflowContext, request))

  Set SC_PrepareGroupMeetingsMessage = SC_PrepareWorkflowMessage(spec, policy, workflowContext, recipientDirectory, templateContext)
End Function

Function SC_RunGroupMeetingsMessage(workflowContext, request, recipientDirectory, templateContext)
  Dim spec
  Set spec = SC_BuildGroupMeetingsMessageSpec(workflowContext, request)

  Dim policy
  Set policy = SC_CreateWorkflowDeliveryPolicy(workflowContext, SC_WorkflowDeliveryAction(workflowContext, request))

  Dim result
  Set result = SC_RunWorkflowMessage(spec, policy, workflowContext, recipientDirectory, templateContext)
  result("workflow_reference") = "group-meetings"
  Set SC_RunGroupMeetingsMessage = result
End Function
