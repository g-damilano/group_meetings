' Workflow-facing core facade.
' Workflows use these helpers to build canonical contracts and dry-run prepared messages
' without owning low-level Outlook COM, template, recipient, attachment, or dispatch details.

Function SC_WorkflowFacadeValue(dictionary, key, defaultValue)
  SC_WorkflowFacadeValue = defaultValue
  On Error Resume Next
  If IsObject(dictionary) Then
    If dictionary.Exists(key) Then SC_WorkflowFacadeValue = dictionary(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_WorkflowSetting(workflowContext, key, defaultValue)
  SC_WorkflowSetting = defaultValue
  On Error Resume Next
  If IsObject(workflowContext) Then
    If workflowContext.Exists("settings") Then
      If IsObject(workflowContext("settings")) Then
        If workflowContext("settings").Exists(key) Then SC_WorkflowSetting = workflowContext("settings")(key)
      End If
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_WorkflowBoolValue(value, defaultValue)
  On Error Resume Next
  If IsEmpty(value) Or IsNull(value) Then
    SC_WorkflowBoolValue = CBool(defaultValue)
  Else
    SC_WorkflowBoolValue = CBool(value)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    SC_WorkflowBoolValue = CBool(defaultValue)
  End If
  On Error GoTo 0
End Function

Function SC_WorkflowNumericValue(value, defaultValue)
  If IsNumeric(value) Then
    SC_WorkflowNumericValue = CLng(value)
  Else
    SC_WorkflowNumericValue = CLng(defaultValue)
  End If
End Function

Function SC_WorkflowName(workflowContext)
  SC_WorkflowName = CStr(SC_WorkflowFacadeValue(workflowContext, "workflow_name", "workflow"))
End Function

Function SC_WorkflowOrgCode(workflowContext)
  SC_WorkflowOrgCode = CStr(SC_WorkflowSetting(workflowContext, "org_code", "ORG"))
End Function

Function SC_WorkflowLogPath(workflowContext)
  SC_WorkflowLogPath = CStr(SC_WorkflowFacadeValue(workflowContext, "log_path", ""))
End Function

Function SC_WorkflowDryRun(workflowContext)
  SC_WorkflowDryRun = SC_WorkflowBoolValue(SC_WorkflowFacadeValue(workflowContext, "dry_run", True), True)
End Function

Function SC_CreateWorkflowManagedMessageSpec(workflowContext, logicalDate, messageType, targetKey, toValue, subjectText, bodyText, bodyHtml)
  Dim managedKey
  managedKey = SC_BuildManagedEmailKey( _
    SC_WorkflowOrgCode(workflowContext), _
    SC_WorkflowName(workflowContext), _
    logicalDate, _
    messageType, _
    targetKey)

  Dim spec
  Set spec = SC_CreateMessageSpec(managedKey, messageType, toValue, subjectText, bodyText)
  spec("body_html") = CStr(bodyHtml)
  spec("workflow_name") = SC_WorkflowName(workflowContext)
  spec("logical_date") = SC_FormatKeyDate(logicalDate)
  spec("target_key") = CStr(targetKey)
  spec("managed_key_format") = "ORG|WORKFLOW|YYYY-MM-DD|MESSAGE-TYPE|TARGET"
  spec("metadata")("WORKFLOW_NAME") = SC_WorkflowName(workflowContext)
  spec("metadata")("LOGICAL_DATE") = SC_FormatKeyDate(logicalDate)
  spec("metadata")("MESSAGE_TYPE") = CStr(messageType)
  spec("metadata")("TARGET_KEY") = CStr(targetKey)

  Set SC_CreateWorkflowManagedMessageSpec = spec
End Function

Sub SC_SetWorkflowMessageRecipients(messageSpec, toValue, ccValue, bccValue)
  messageSpec("to") = CStr(toValue)
  messageSpec("cc") = CStr(ccValue)
  messageSpec("bcc") = CStr(bccValue)
End Sub

Sub SC_SetWorkflowMessageBody(messageSpec, bodyText, bodyHtml)
  messageSpec("body_text") = CStr(bodyText)
  messageSpec("body_html") = CStr(bodyHtml)
End Sub

Sub SC_SetWorkflowMessageTemplate(messageSpec, templatePath, inlineImageFolder)
  messageSpec("template_path") = CStr(templatePath)
  messageSpec("inline_image_folder") = CStr(inlineImageFolder)
End Sub

Sub SC_SetWorkflowMessageAttachments(messageSpec, attachments, attachmentBaseFolder)
  If IsObject(attachments) Then
    Set messageSpec("attachments") = attachments
  Else
    messageSpec("attachments") = CStr(attachments)
  End If
  messageSpec("attachment_base_folder") = CStr(attachmentBaseFolder)
End Sub

Sub SC_AddWorkflowMessageMetadata(messageSpec, key, value)
  If Not messageSpec.Exists("metadata") Then Set messageSpec("metadata") = SC_CreateDictionary()
  messageSpec("metadata")(UCase(CStr(key))) = CStr(value)
End Sub

Function SC_CreateWorkflowDeliveryPolicy(workflowContext, actionText)
  Dim resolvedAction
  resolvedAction = CStr(actionText)
  If Len(Trim(resolvedAction)) = 0 Then resolvedAction = CStr(SC_WorkflowSetting(workflowContext, "default_delivery_action", "display"))

  Dim policy
  Set policy = SC_CreateDeliveryPolicy(resolvedAction)
  policy("defer_minutes") = SC_WorkflowNumericValue(SC_WorkflowSetting(workflowContext, "defer_minutes", 0), 0)
  policy("urgent_verify_after_submit") = SC_WorkflowBoolValue(SC_WorkflowSetting(workflowContext, "urgent_verify_after_submit", False), False)
  policy("managed_dispatch_required") = SC_WorkflowBoolValue(SC_WorkflowSetting(workflowContext, "managed_dispatch_required", True), True)
  policy("enable_stale_outbox_replacement") = SC_WorkflowBoolValue(SC_WorkflowSetting(workflowContext, "enable_stale_outbox_replacement", False), False)

  Set SC_CreateWorkflowDeliveryPolicy = policy
End Function

Function SC_CreateWorkflowComposerContext(workflowContext, recipientDirectory, templateContext)
  Dim composerContext
  Set composerContext = SC_CreateDictionary()
  composerContext("workflow_name") = SC_WorkflowName(workflowContext)
  composerContext("run_id") = CStr(SC_WorkflowFacadeValue(workflowContext, "run_id", ""))

  If IsObject(recipientDirectory) Then
    Set composerContext("recipient_directory") = recipientDirectory
  Else
    Set composerContext("recipient_directory") = SC_CreateDictionary()
  End If

  If IsObject(templateContext) Then
    Set composerContext("template_context") = templateContext
  Else
    Set composerContext("template_context") = SC_CreateDictionary()
  End If

  Set SC_CreateWorkflowComposerContext = composerContext
End Function

Function SC_CreateWorkflowDispatchConfig(workflowContext, recipientDirectory, templateContext)
  Dim dispatchConfig
  Set dispatchConfig = SC_CreateWorkflowComposerContext(workflowContext, recipientDirectory, templateContext)
  dispatchConfig("log_path") = SC_WorkflowLogPath(workflowContext)
  dispatchConfig("dry_run") = SC_WorkflowDryRun(workflowContext)
  dispatchConfig("enable_subject_fallback") = SC_WorkflowBoolValue(SC_WorkflowSetting(workflowContext, "enable_subject_fallback", True), True)
  dispatchConfig("managed_outbox_grace_minutes") = SC_WorkflowNumericValue(SC_WorkflowSetting(workflowContext, "managed_outbox_grace_minutes", SC_DEFAULT_OUTBOX_GRACE_MINUTES), SC_DEFAULT_OUTBOX_GRACE_MINUTES)
  Set SC_CreateWorkflowDispatchConfig = dispatchConfig
End Function

Function SC_PrepareWorkflowMessage(messageSpec, deliveryPolicy, workflowContext, recipientDirectory, templateContext)
  Dim composerContext
  Set composerContext = SC_CreateWorkflowComposerContext(workflowContext, recipientDirectory, templateContext)
  Set SC_PrepareWorkflowMessage = SC_PrepareMessageForComposition(messageSpec, deliveryPolicy, composerContext, SC_WorkflowLogPath(workflowContext))
End Function

Function SC_RunWorkflowMessage(messageSpec, deliveryPolicy, workflowContext, recipientDirectory, templateContext)
  Dim prepared
  Set prepared = SC_PrepareWorkflowMessage(messageSpec, deliveryPolicy, workflowContext, recipientDirectory, templateContext)

  Dim dispatchConfig
  Set dispatchConfig = SC_CreateWorkflowDispatchConfig(workflowContext, recipientDirectory, templateContext)

  Dim result
  Set result = SC_DispatchManagedEmail(messageSpec, deliveryPolicy, dispatchConfig)
  result("workflow_name") = SC_WorkflowName(workflowContext)
  result("run_id") = CStr(SC_WorkflowFacadeValue(workflowContext, "run_id", ""))
  result("prepared_to") = CStr(SC_WorkflowFacadeValue(prepared, "to", ""))
  result("prepared_cc") = CStr(SC_WorkflowFacadeValue(prepared, "cc", ""))
  result("prepared_bcc") = CStr(SC_WorkflowFacadeValue(prepared, "bcc", ""))
  result("prepared_subject") = CStr(SC_WorkflowFacadeValue(prepared, "subject", ""))
  result("prepared_defer_minutes") = SC_WorkflowFacadeValue(prepared, "defer_minutes", 0)
  result("prepared_attachment_count") = 0
  result("prepared_inline_count") = 0

  On Error Resume Next
  If prepared.Exists("attachment_paths") Then
    If IsObject(prepared("attachment_paths")) Then result("prepared_attachment_count") = prepared("attachment_paths").Count
  End If
  If prepared.Exists("inline_info") Then
    If IsObject(prepared("inline_info")) Then result("prepared_inline_count") = prepared("inline_info")("inline_count")
  End If
  Err.Clear
  On Error GoTo 0

  Set SC_RunWorkflowMessage = result
End Function
