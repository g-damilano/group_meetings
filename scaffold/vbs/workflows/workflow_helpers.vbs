' Shared workflow-layer helpers.
' These helpers shape workflow requests into core contracts. They deliberately avoid
' Outlook COM, folder access, or mailbox side effects.

Function SC_WorkflowRequestValue(request, key, defaultValue)
  SC_WorkflowRequestValue = defaultValue
  On Error Resume Next
  If IsObject(request) Then
    If request.Exists(key) Then SC_WorkflowRequestValue = request(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_WorkflowRequestObject(request, key)
  Set SC_WorkflowRequestObject = Nothing
  On Error Resume Next
  If IsObject(request) Then
    If request.Exists(key) Then
      If IsObject(request(key)) Then Set SC_WorkflowRequestObject = request(key)
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_CreateWorkflowMessageRequest(logicalDate, messageType, targetKey, toValue, subjectText, bodyText, bodyHtml)
  Dim request
  Set request = SC_CreateDictionary()
  request("logical_date") = logicalDate
  request("message_type") = CStr(messageType)
  request("target_key") = CStr(targetKey)
  request("to") = CStr(toValue)
  request("cc") = ""
  request("bcc") = ""
  request("subject") = CStr(subjectText)
  request("body_text") = CStr(bodyText)
  request("body_html") = CStr(bodyHtml)
  request("template_path") = ""
  request("inline_image_folder") = ""
  request("attachment_base_folder") = ""
  request("delivery_action") = ""
  Set request("attachments") = SC_CreateDictionary()
  Set request("metadata") = SC_CreateDictionary()
  Set SC_CreateWorkflowMessageRequest = request
End Function

Sub SC_SetWorkflowRequestRecipients(request, toValue, ccValue, bccValue)
  request("to") = CStr(toValue)
  request("cc") = CStr(ccValue)
  request("bcc") = CStr(bccValue)
End Sub

Sub SC_SetWorkflowRequestTemplate(request, templatePath, inlineImageFolder)
  request("template_path") = CStr(templatePath)
  request("inline_image_folder") = CStr(inlineImageFolder)
End Sub

Sub SC_SetWorkflowRequestAttachments(request, attachments, attachmentBaseFolder)
  If IsObject(attachments) Then
    Set request("attachments") = attachments
  Else
    request("attachments") = CStr(attachments)
  End If
  request("attachment_base_folder") = CStr(attachmentBaseFolder)
End Sub

Sub SC_AddWorkflowRequestMetadata(request, key, value)
  If Not request.Exists("metadata") Then Set request("metadata") = SC_CreateDictionary()
  request("metadata")(UCase(CStr(key))) = CStr(value)
End Sub

Sub SC_ApplyWorkflowRequestToMessageSpec(messageSpec, request)
  SC_SetWorkflowMessageRecipients messageSpec, _
    SC_WorkflowRequestValue(request, "to", ""), _
    SC_WorkflowRequestValue(request, "cc", ""), _
    SC_WorkflowRequestValue(request, "bcc", "")

  SC_SetWorkflowMessageBody messageSpec, _
    SC_WorkflowRequestValue(request, "body_text", ""), _
    SC_WorkflowRequestValue(request, "body_html", "")

  SC_SetWorkflowMessageTemplate messageSpec, _
    SC_WorkflowRequestValue(request, "template_path", ""), _
    SC_WorkflowRequestValue(request, "inline_image_folder", "")

  Dim attachments
  Set attachments = SC_WorkflowRequestObject(request, "attachments")
  If attachments Is Nothing Then
    SC_SetWorkflowMessageAttachments messageSpec, _
      SC_WorkflowRequestValue(request, "attachments", ""), _
      SC_WorkflowRequestValue(request, "attachment_base_folder", "")
  Else
    SC_SetWorkflowMessageAttachments messageSpec, attachments, SC_WorkflowRequestValue(request, "attachment_base_folder", "")
  End If

  Dim metadata
  Set metadata = SC_WorkflowRequestObject(request, "metadata")
  If Not metadata Is Nothing Then
    Dim key
    For Each key In metadata.Keys
      SC_AddWorkflowMessageMetadata messageSpec, key, metadata(key)
    Next
  End If
End Sub

Function SC_WorkflowDeliveryAction(workflowContext, request)
  Dim actionText
  actionText = CStr(SC_WorkflowRequestValue(request, "delivery_action", ""))
  If Len(Trim(actionText)) = 0 Then actionText = CStr(SC_WorkflowSetting(workflowContext, "default_delivery_action", "display"))
  SC_WorkflowDeliveryAction = actionText
End Function
