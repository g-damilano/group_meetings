' Mail composition facade for scaffold-managed Outlook messages.
' This module prepares MessageSpec content through recipient, template, attachment, and
' inline-image helpers. It does not decide workflow policy or duplicate-prevention state.

Function SC_ComposerValue(dictionary, key, defaultValue)
  SC_ComposerValue = defaultValue
  On Error Resume Next
  If IsObject(dictionary) Then
    If dictionary.Exists(key) Then SC_ComposerValue = dictionary(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_ComposerObject(dictionary, key)
  Set SC_ComposerObject = Nothing
  On Error Resume Next
  If IsObject(dictionary) Then
    If dictionary.Exists(key) Then
      If IsObject(dictionary(key)) Then Set SC_ComposerObject = dictionary(key)
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_ComposerRecipientField(messageSpec, fieldName, recipientDirectory)
  Dim rawValue
  rawValue = CStr(SC_ComposerValue(messageSpec, fieldName, ""))
  If Len(Trim(rawValue)) = 0 Then
    SC_ComposerRecipientField = ""
    Exit Function
  End If

  Dim resolved
  Set resolved = SC_ResolveRecipientList(rawValue, recipientDirectory)
  SC_ComposerRecipientField = SC_JoinResolvedRecipientAddresses(resolved)
End Function

Function SC_ComposerRenderBodies(messageSpec, templateContext)
  Dim result
  Set result = SC_CreateDictionary()
  result("body_html") = CStr(SC_ComposerValue(messageSpec, "body_html", ""))
  result("body_text") = CStr(SC_ComposerValue(messageSpec, "body_text", ""))
  Set result("metadata") = SC_CreateDictionary()
  Set result("placeholders") = SC_CreateDictionary()

  Dim templatePath
  templatePath = CStr(SC_ComposerValue(messageSpec, "template_path", ""))
  If Len(Trim(templatePath)) > 0 Then
    Dim rendered
    Set rendered = SC_RenderTemplateFile(templatePath, templateContext)
    result("body_html") = CStr(rendered("body"))
    result("body_text") = ""
    Set result("metadata") = rendered("metadata")
    Set result("placeholders") = rendered("placeholders")
    Set SC_ComposerRenderBodies = result
    Exit Function
  End If

  If IsObject(templateContext) Then
    result("body_html") = SC_ApplyContextToText(result("body_html"), templateContext)
    result("body_text") = SC_ApplyContextToText(result("body_text"), templateContext)
    If Len(result("body_html")) > 0 Then
      Set result("placeholders") = SC_ExtractPlaceholders(result("body_html"))
    Else
      Set result("placeholders") = SC_ExtractPlaceholders(result("body_text"))
    End If
  End If

  Set SC_ComposerRenderBodies = result
End Function

Function SC_PrepareMessageForComposition(messageSpec, deliveryPolicy, composerContext, logPath)
  Dim prepared
  Set prepared = SC_CreateDictionary()

  Dim recipientDirectory
  Set recipientDirectory = SC_ComposerObject(composerContext, "recipient_directory")

  Dim templateContext
  Set templateContext = SC_ComposerObject(composerContext, "template_context")
  If templateContext Is Nothing Then Set templateContext = SC_ComposerObject(messageSpec, "metadata")

  Dim rendered
  Set rendered = SC_ComposerRenderBodies(messageSpec, templateContext)

  Dim subjectText
  subjectText = CStr(SC_ComposerValue(messageSpec, "subject", ""))
  If Len(Trim(subjectText)) = 0 Then
    If rendered("metadata").Exists("SUBJECT") Then subjectText = CStr(rendered("metadata")("SUBJECT"))
  End If
  If IsObject(templateContext) Then subjectText = SC_ApplyContextToText(subjectText, templateContext)

  prepared("message_key") = CStr(SC_ComposerValue(messageSpec, "message_key", ""))
  prepared("message_type") = CStr(SC_ComposerValue(messageSpec, "message_type", ""))
  prepared("to") = SC_ComposerRecipientField(messageSpec, "to", recipientDirectory)
  prepared("cc") = SC_ComposerRecipientField(messageSpec, "cc", recipientDirectory)
  prepared("bcc") = SC_ComposerRecipientField(messageSpec, "bcc", recipientDirectory)
  prepared("subject") = subjectText
  prepared("body_html") = CStr(rendered("body_html"))
  prepared("body_text") = CStr(rendered("body_text"))
  prepared("defer_minutes") = SC_ComposerValue(deliveryPolicy, "defer_minutes", 0)
  Set prepared("metadata") = rendered("metadata")
  Set prepared("placeholders") = rendered("placeholders")
  Set prepared("attachment_paths") = SC_ResolveMessageAttachments(messageSpec, logPath)

  Dim templatePath
  templatePath = CStr(SC_ComposerValue(messageSpec, "template_path", ""))
  Dim inlineFolder
  inlineFolder = CStr(SC_ComposerValue(messageSpec, "inline_image_folder", ""))
  Dim inlineInfo
  Set inlineInfo = SC_PrepareInlineImages(prepared("body_html"), templatePath, inlineFolder, logPath)
  prepared("body_html") = CStr(inlineInfo("html_body"))
  Set prepared("inline_info") = inlineInfo

  Set SC_PrepareMessageForComposition = prepared
End Function

Function SC_ComposeOutlookMail(outlookApp, messageSpec, deliveryPolicy, composerContext, logPath)
  Dim prepared
  Set prepared = SC_PrepareMessageForComposition(messageSpec, deliveryPolicy, composerContext, logPath)

  Dim mailItem
  Set mailItem = outlookApp.CreateItem(0)
  mailItem.To = CStr(prepared("to"))
  mailItem.CC = CStr(prepared("cc"))
  mailItem.BCC = CStr(prepared("bcc"))
  mailItem.Subject = CStr(prepared("subject"))

  If Len(CStr(prepared("body_html"))) > 0 Then
    mailItem.BodyFormat = 2
    mailItem.HTMLBody = CStr(prepared("body_html"))
  Else
    mailItem.Body = CStr(prepared("body_text"))
  End If

  SC_AddAttachmentsToMail mailItem, prepared("attachment_paths"), logPath
  SC_AddInlineImagesToMail mailItem, prepared("inline_info"), logPath

  If IsNumeric(prepared("defer_minutes")) Then
    If CDbl(prepared("defer_minutes")) > 0 Then mailItem.DeferredDeliveryTime = Now + (CDbl(prepared("defer_minutes")) / (24 * 60))
  End If

  SC_StampManagedMail mailItem, CStr(prepared("message_key"))
  Set SC_ComposeOutlookMail = mailItem
End Function
