'====================================================================
' office_managed_dispatch.vbs - Managed email identity and dispatch helpers
'====================================================================
Option Explicit
SetLocale 1033

Const OFFICE_OL_FOLDER_OUTBOX = 4
Const OFFICE_OL_FOLDER_SENT = 5
Const OFFICE_MSGFLAG_SUBMIT = 4
Const OFFICE_MSGFLAG_UNSENT = 8
Const OFFICE_PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
Const OFFICE_PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"

Function OfficeManagedMarker(managedKey)
  OfficeManagedMarker = "<!-- " & OFFICE_MANAGED_KEY_MARKER & ": " & managedKey & " -->"
End Function

Sub OfficeStampManagedMail(mailItem, managedKey)
  On Error Resume Next
  Dim prop: Set prop = mailItem.UserProperties.Find(OFFICE_MANAGED_KEY_PROPERTY)
  If prop Is Nothing Then Set prop = mailItem.UserProperties.Add(OFFICE_MANAGED_KEY_PROPERTY, 1, True)
  prop.Value = managedKey
  Err.Clear

  Dim htmlBody: htmlBody = mailItem.HTMLBody
  If Err.Number = 0 Then
    If InStr(1, htmlBody, OFFICE_MANAGED_KEY_MARKER & ":", vbTextCompare) = 0 Then
      Dim marker: marker = OfficeManagedMarker(managedKey)
      Dim closePos: closePos = InStr(1, htmlBody, "</body>", vbTextCompare)
      If closePos > 0 Then
        mailItem.HTMLBody = Left(htmlBody, closePos - 1) & vbCrLf & marker & vbCrLf & Mid(htmlBody, closePos)
      Else
        mailItem.HTMLBody = htmlBody & vbCrLf & marker
      End If
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Function OfficeExtractManagedKeyFromText(textValue)
  OfficeExtractManagedKeyFromText = ""
  Dim pos: pos = InStr(1, CStr(textValue), OFFICE_MANAGED_KEY_MARKER & ":", vbTextCompare)
  If pos = 0 Then Exit Function
  Dim restText: restText = Trim(Mid(CStr(textValue), pos + Len(OFFICE_MANAGED_KEY_MARKER & ":")))
  Dim endPos: endPos = InStr(1, restText, "-->", vbTextCompare)
  If endPos = 0 Then endPos = InStr(1, restText, vbCrLf, vbTextCompare)
  If endPos = 0 Then endPos = InStr(1, restText, "<", vbTextCompare)
  If endPos > 0 Then
    OfficeExtractManagedKeyFromText = Trim(Left(restText, endPos - 1))
  Else
    OfficeExtractManagedKeyFromText = Trim(restText)
  End If
End Function

Function OfficeGetManagedMailKey(mailItem)
  OfficeGetManagedMailKey = ""
  On Error Resume Next
  Dim prop: Set prop = mailItem.UserProperties.Find(OFFICE_MANAGED_KEY_PROPERTY)
  If Err.Number = 0 Then
    If Not prop Is Nothing Then
      If Len(Trim(CStr(prop.Value))) > 0 Then
        OfficeGetManagedMailKey = Trim(CStr(prop.Value))
        On Error GoTo 0
        Exit Function
      End If
    End If
  End If
  Err.Clear

  Dim htmlKey: htmlKey = OfficeExtractManagedKeyFromText(mailItem.HTMLBody)
  If Len(htmlKey) > 0 Then
    OfficeGetManagedMailKey = htmlKey
    On Error GoTo 0
    Exit Function
  End If
  Err.Clear

  Dim bodyKey: bodyKey = OfficeExtractManagedKeyFromText(mailItem.Body)
  If Len(bodyKey) > 0 Then OfficeGetManagedMailKey = bodyKey
  Err.Clear
  On Error GoTo 0
End Function

Function OfficeFindManagedMailInFolder(ns, folderId, managedKey, subjectText, allowSubjectFallback, logPath)
  Dim matches: Set matches = OfficeCreateDictionary()
  Dim folder, items, item, itemKey

  On Error Resume Next
  Set folder = ns.GetDefaultFolder(CLng(folderId))
  If Err.Number <> 0 Then
    OfficeLogWarning logPath, "Unable to access Outlook folder " & folderId & " while searching for managed key " & managedKey & ": " & Err.Description
    Err.Clear
    Set OfficeFindManagedMailInFolder = matches
    On Error GoTo 0
    Exit Function
  End If
  Set items = folder.Items
  If CLng(folderId) = OFFICE_OL_FOLDER_SENT Then
    items.Sort "SentOn", True
  ElseIf CLng(folderId) = OFFICE_OL_FOLDER_OUTBOX Then
    items.Sort "CreationTime", True
  End If
  On Error GoTo 0

  For Each item In items
    If TypeName(item) = "MailItem" Then
      itemKey = OfficeGetManagedMailKey(item)
      If itemKey = managedKey Then
        matches.Add CStr(matches.Count), item
      ElseIf CBool(allowSubjectFallback) And Len(itemKey) = 0 Then
        On Error Resume Next
        If CStr(item.Subject) = CStr(subjectText) Then
          matches.Add CStr(matches.Count), item
          OfficeLogWarning logPath, "Legacy subject-only fallback matched subject: " & subjectText
        End If
        Err.Clear
        On Error GoTo 0
      End If
    End If
  Next

  Set OfficeFindManagedMailInFolder = matches
End Function

Function OfficeFindManagedMailInSentItems(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  Set OfficeFindManagedMailInSentItems = OfficeFindManagedMailInFolder(ns, OFFICE_OL_FOLDER_SENT, managedKey, subjectText, allowSubjectFallback, logPath)
End Function

Function OfficeFindManagedMailInOutbox(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  Set OfficeFindManagedMailInOutbox = OfficeFindManagedMailInFolder(ns, OFFICE_OL_FOLDER_OUTBOX, managedKey, subjectText, allowSubjectFallback, logPath)
End Function

Function OfficeFindManagedEmailState(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  Dim stateInfo: Set stateInfo = OfficeCreateDictionary()
  Dim sentMatches: Set sentMatches = OfficeFindManagedMailInSentItems(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  Dim outboxMatches: Set outboxMatches = OfficeFindManagedMailInOutbox(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  stateInfo("sent_count") = sentMatches.Count
  stateInfo("outbox_count") = outboxMatches.Count
  Set stateInfo("sent_matches") = sentMatches
  Set stateInfo("outbox_matches") = outboxMatches
  Set OfficeFindManagedEmailState = stateInfo
End Function

Function OfficeOutboxDiagnostic(mailItem, managedKey, logPath)
  Dim diag: Set diag = OfficeCreateDictionary()
  diag("managed_key") = managedKey
  On Error Resume Next
  diag("subject") = CStr(mailItem.Subject)
  diag("to") = CStr(mailItem.To)
  diag("entry_id") = CStr(mailItem.EntryID)
  diag("creation_time") = CStr(mailItem.CreationTime)
  diag("last_modification_time") = CStr(mailItem.LastModificationTime)
  diag("attachment_count") = CStr(mailItem.Attachments.Count)
  diag("deferred_delivery_time") = CStr(mailItem.DeferredDeliveryTime)
  diag("pr_message_flags") = "[unavailable]"
  diag("msgflag_submit") = "[unavailable]"
  diag("msgflag_unsent") = "[unavailable]"
  diag("pr_client_submit_time") = "[unavailable]"
  Err.Clear

  Dim flags: flags = ""
  flags = mailItem.PropertyAccessor.GetProperty(OFFICE_PR_MESSAGE_FLAGS)
  If Err.Number = 0 Then
    diag("pr_message_flags") = flags
    diag("msgflag_submit") = CBool((CLng(flags) And OFFICE_MSGFLAG_SUBMIT) <> 0)
    diag("msgflag_unsent") = CBool((CLng(flags) And OFFICE_MSGFLAG_UNSENT) <> 0)
  Else
    OfficeLogWarning logPath, "MAPI property lookup failed for PR_MESSAGE_FLAGS: " & Err.Description
    Err.Clear
  End If

  Dim submitTime: submitTime = mailItem.PropertyAccessor.GetProperty(OFFICE_PR_CLIENT_SUBMIT_TIME)
  If Err.Number = 0 Then
    diag("pr_client_submit_time") = CStr(submitTime)
  Else
    OfficeLogWarning logPath, "MAPI property lookup failed for PR_CLIENT_SUBMIT_TIME: " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0

  Set OfficeOutboxDiagnostic = diag
End Function

Sub OfficeLogOutboxDiagnostic(diag, logPath)
  On Error Resume Next
  OfficeLogInfo logPath, "Outbox diagnostic: key=" & diag("managed_key") & _
    " subject=" & diag("subject") & _
    " to=" & diag("to") & _
    " entry_id=" & diag("entry_id") & _
    " creation_time=" & diag("creation_time") & _
    " last_modification_time=" & diag("last_modification_time") & _
    " deferred_delivery_time=" & diag("deferred_delivery_time") & _
    " attachments=" & diag("attachment_count") & _
    " pr_message_flags=" & diag("pr_message_flags") & _
    " msgflag_submit=" & CStr(diag("msgflag_submit")) & _
    " msgflag_unsent=" & CStr(diag("msgflag_unsent")) & _
    " pr_client_submit_time=" & diag("pr_client_submit_time")
  On Error GoTo 0
End Sub

Function OfficeClassifyQueuedManagedMail(outlookApp, mailItem, nowValue, graceMinutes)
  Dim transport: Set transport = OfficeGetTransportState(outlookApp)
  If transport("blocked_or_auth_suspected") Then
    Set OfficeClassifyQueuedManagedMail = OfficeCreateDispatchResult(OFFICE_STATE_ACTION_REQUIRED, "Outlook appears offline/disconnected/auth-blocked.")
    Exit Function
  End If

  On Error Resume Next
  Dim deferredValue: deferredValue = mailItem.DeferredDeliveryTime
  If Err.Number = 0 Then
    If IsDate(deferredValue) Then
      If CDate(deferredValue) > CDate(nowValue) Then
        Set OfficeClassifyQueuedManagedMail = OfficeCreateDispatchResult(OFFICE_STATE_WAITING_DEFERRED, "Deferred delivery time has not been reached.")
        On Error GoTo 0
        Exit Function
      End If
    End If
  End If
  Err.Clear

  Dim ageMinutes: ageMinutes = DateDiff("n", mailItem.CreationTime, CDate(nowValue))
  If Err.Number <> 0 Then
    Err.Clear
    Set OfficeClassifyQueuedManagedMail = OfficeCreateDispatchResult(OFFICE_STATE_FAILED_SAFE, "Unable to determine queued item age.")
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If CLng(ageMinutes) < CLng(graceMinutes) Then
    Set OfficeClassifyQueuedManagedMail = OfficeCreateDispatchResult(OFFICE_STATE_QUEUED_OUTBOX, "Queued item is still within grace period.")
  Else
    Set OfficeClassifyQueuedManagedMail = OfficeCreateDispatchResult(OFFICE_STATE_STALE_MANAGED_OUTBOX, "Managed item appears stale in Outbox.")
  End If
End Function

Function OfficeDispatchConfigValue(config, key, defaultValue)
  OfficeDispatchConfigValue = defaultValue
  On Error Resume Next
  If IsObject(config) Then
    If config.Exists(key) Then OfficeDispatchConfigValue = config(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function OfficeDispatchConfigObject(config, key)
  Set OfficeDispatchConfigObject = Nothing
  On Error Resume Next
  If IsObject(config) Then
    If config.Exists(key) Then
      If IsObject(config(key)) Then Set OfficeDispatchConfigObject = config(key)
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function OfficeMessageSpecValue(spec, key, defaultValue)
  OfficeMessageSpecValue = defaultValue
  On Error Resume Next
  If IsObject(spec) Then
    If spec.Exists(key) Then OfficeMessageSpecValue = spec(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function OfficeMessageSpecObject(spec, key)
  Set OfficeMessageSpecObject = Nothing
  On Error Resume Next
  If IsObject(spec) Then
    If spec.Exists(key) Then
      If IsObject(spec(key)) Then Set OfficeMessageSpecObject = spec(key)
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function OfficePolicyValue(policy, key, defaultValue)
  OfficePolicyValue = defaultValue
  On Error Resume Next
  If IsObject(policy) Then
    If policy.Exists(key) Then OfficePolicyValue = policy(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function OfficeBool(value, defaultValue)
  On Error Resume Next
  If IsEmpty(value) Or IsNull(value) Then
    OfficeBool = CBool(defaultValue)
  Else
    OfficeBool = CBool(value)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    OfficeBool = CBool(defaultValue)
  End If
  On Error GoTo 0
End Function

Function OfficeCreateManagedMail(outlookApp, messageSpec, deliveryPolicy)
  Dim mailItem: Set mailItem = outlookApp.CreateItem(0)
  Dim htmlBody: htmlBody = CStr(OfficeMessageSpecValue(messageSpec, "body_html", ""))
  Dim textBody: textBody = CStr(OfficeMessageSpecValue(messageSpec, "body_text", ""))

  mailItem.To = CStr(OfficeMessageSpecValue(messageSpec, "to", ""))
  mailItem.CC = CStr(OfficeMessageSpecValue(messageSpec, "cc", ""))
  mailItem.BCC = CStr(OfficeMessageSpecValue(messageSpec, "bcc", ""))
  mailItem.Subject = CStr(OfficeMessageSpecValue(messageSpec, "subject", ""))
  If Len(htmlBody) > 0 Then
    mailItem.BodyFormat = 2
    mailItem.HTMLBody = htmlBody
  Else
    mailItem.Body = textBody
  End If

  Dim deferMinutes: deferMinutes = OfficePolicyValue(deliveryPolicy, "defer_minutes", 0)
  If IsNumeric(deferMinutes) Then
    If CDbl(deferMinutes) > 0 Then mailItem.DeferredDeliveryTime = Now + (CDbl(deferMinutes) / (24 * 60))
  End If

  Dim attachmentSpec: Set attachmentSpec = OfficeMessageSpecObject(messageSpec, "attachments")
  If Not attachmentSpec Is Nothing Then OfficeAddSpecAttachments mailItem, attachmentSpec
  OfficeStampManagedMail mailItem, CStr(OfficeMessageSpecValue(messageSpec, "message_key", ""))
  Set OfficeCreateManagedMail = mailItem
End Function

Sub OfficeAddSpecAttachments(mailItem, attachments)
  If IsEmpty(attachments) Then Exit Sub
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

  If IsArray(attachments) Then
    Dim i
    For i = LBound(attachments) To UBound(attachments)
      OfficeAddAttachmentIfPresent mailItem, attachments(i), fso
    Next
  ElseIf IsObject(attachments) Then
    Dim key
    For Each key In attachments.Keys
      OfficeAddAttachmentIfPresent mailItem, attachments(key), fso
    Next
  Else
    OfficeAddAttachmentIfPresent mailItem, attachments, fso
  End If
End Sub

Sub OfficeAddAttachmentIfPresent(mailItem, attachmentPath, fso)
  If IsObject(attachmentPath) Then Exit Sub
  If IsEmpty(attachmentPath) Or IsNull(attachmentPath) Then Exit Sub
  Dim resolvedPath: resolvedPath = Trim(CStr(attachmentPath))
  If Len(resolvedPath) = 0 Then Exit Sub
  If fso.FileExists(resolvedPath) Then mailItem.Attachments.Add resolvedPath
End Sub

Function OfficeSubmitManagedMail(outlookApp, messageSpec, deliveryPolicy, logPath, createdState)
  Dim result: Set result = OfficeCreateDispatchResult(createdState, "")
  result("managed_key") = CStr(OfficeMessageSpecValue(messageSpec, "message_key", ""))
  result("created_new_mail") = True

  On Error Resume Next
  Dim mailItem: Set mailItem = OfficeCreateManagedMail(outlookApp, messageSpec, deliveryPolicy)
  If Err.Number <> 0 Then
    result("state") = OFFICE_STATE_FAILED_SAFE
    result("reason") = "Failed to compose managed email: " & Err.Description
    OfficeLogError logPath, result("reason")
    Err.Clear
    Set OfficeSubmitManagedMail = result
    On Error GoTo 0
    Exit Function
  End If

  Dim action: action = LCase(Trim(CStr(OfficePolicyValue(deliveryPolicy, "action", "display"))))
  Select Case action
    Case "send"
      mailItem.Send
      If Err.Number = 0 Then
        result("state") = OFFICE_STATE_SEND_REQUESTED
        result("reason") = "Managed email submitted to Outlook."
        OfficeLogInfo logPath, "Created managed email and requested send: " & result("managed_key")
      Else
        result("state") = OFFICE_STATE_FAILED_SAFE
        result("reason") = "Outlook Send failed: " & Err.Description
        OfficeLogError logPath, result("reason")
        Err.Clear
      End If
    Case "draft", "save"
      mailItem.Save
      If Err.Number = 0 Then
        result("state") = OFFICE_STATE_DRAFTED
        result("reason") = "Managed email saved as draft."
        OfficeLogInfo logPath, "Created managed draft: " & result("managed_key")
      Else
        result("state") = OFFICE_STATE_FAILED_SAFE
        result("reason") = "Outlook Save failed: " & Err.Description
        OfficeLogError logPath, result("reason")
        Err.Clear
      End If
    Case Else
      mailItem.Display
      If Err.Number = 0 Then
        result("state") = OFFICE_STATE_DISPLAYED
        result("reason") = "Managed email displayed."
        OfficeLogInfo logPath, "Displayed managed email: " & result("managed_key")
      Else
        result("state") = OFFICE_STATE_FAILED_SAFE
        result("reason") = "Outlook Display failed: " & Err.Description
        OfficeLogError logPath, result("reason")
        Err.Clear
      End If
  End Select

  If Not mailItem Is Nothing Then result("outlook_entry_id") = CStr(mailItem.EntryID)
  On Error GoTo 0
  Set OfficeSubmitManagedMail = result
End Function

Function OfficeCanReplaceManagedOutboxMail(outlookApp, mailItem, managedKey, nowValue, graceMinutes)
  OfficeCanReplaceManagedOutboxMail = False
  If OfficeGetManagedMailKey(mailItem) <> managedKey Then Exit Function

  Dim transport: Set transport = OfficeGetTransportState(outlookApp)
  If transport("blocked_or_auth_suspected") Then Exit Function

  On Error Resume Next
  Dim deferredValue: deferredValue = mailItem.DeferredDeliveryTime
  If Err.Number = 0 Then
    If IsDate(deferredValue) Then
      If CDate(deferredValue) > CDate(nowValue) Then
        On Error GoTo 0
        Exit Function
      End If
    End If
  End If
  Err.Clear

  Dim ageMinutes: ageMinutes = DateDiff("n", mailItem.CreationTime, CDate(nowValue))
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  OfficeCanReplaceManagedOutboxMail = (CLng(ageMinutes) >= CLng(graceMinutes))
End Function

Function OfficeDispatchManagedEmail(messageSpec, deliveryPolicy, dispatchConfig)
  Dim logPath: logPath = CStr(OfficeDispatchConfigValue(dispatchConfig, "log_path", ""))
  Dim managedKey: managedKey = CStr(OfficeMessageSpecValue(messageSpec, "message_key", ""))
  Dim subjectText: subjectText = CStr(OfficeMessageSpecValue(messageSpec, "subject", ""))
  Dim graceMinutes: graceMinutes = CLng(OfficeDispatchConfigValue(dispatchConfig, "grace_minutes", 15))
  Dim enableReplacement: enableReplacement = OfficeBool(OfficeDispatchConfigValue(dispatchConfig, "enable_stale_outbox_replacement", False), False)
  Dim allowFallback: allowFallback = OfficeBool(OfficeDispatchConfigValue(dispatchConfig, "enable_subject_fallback", True), True)
  Dim nowValue: nowValue = OfficeDispatchConfigValue(dispatchConfig, "now", Now)

  Dim result: Set result = OfficeCreateDispatchResult(OFFICE_STATE_NOT_REQUESTED, "")
  result("managed_key") = managedKey

  If Len(managedKey) = 0 Then
    result("state") = OFFICE_STATE_FAILED_SAFE
    result("reason") = "MessageSpec has no managed email key."
    OfficeLogError logPath, result("reason")
    Set OfficeDispatchManagedEmail = result
    Exit Function
  End If

  OfficeLogInfo logPath, "Managed email dispatch requested. key=" & managedKey & " subject=" & subjectText

  Dim outlookApp: Set outlookApp = OfficeDispatchConfigObject(dispatchConfig, "outlook_app")
  If outlookApp Is Nothing Then Set outlookApp = OfficeGetOutlookApplication()
  If outlookApp Is Nothing Then
    result("state") = OFFICE_STATE_FAILED_SAFE
    result("reason") = "Outlook.Application is unavailable."
    result("action_required") = True
    OfficeLogActionRequired logPath, result("reason")
    Set OfficeDispatchManagedEmail = result
    Exit Function
  End If

  Dim ns
  On Error Resume Next
  Set ns = outlookApp.GetNamespace("MAPI")
  If Err.Number <> 0 Then
    result("state") = OFFICE_STATE_ACTION_REQUIRED
    result("reason") = "Outlook MAPI namespace is unavailable: " & Err.Description
    result("action_required") = True
    OfficeLogActionRequired logPath, result("reason")
    Err.Clear
    Set OfficeDispatchManagedEmail = result
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Dim stateInfo: Set stateInfo = OfficeFindManagedEmailState(ns, managedKey, subjectText, allowFallback, logPath)
  OfficeLogInfo logPath, "Sent Items search result for " & managedKey & ": " & CStr(stateInfo("sent_count"))
  If CLng(stateInfo("sent_count")) > 0 Then
    result("state") = OFFICE_STATE_DUPLICATE_SUPPRESSED
    result("reason") = "Managed email already exists in Sent Items."
    result("folder") = "Sent Items"
    OfficeLogInfo logPath, "Queue decision: DUPLICATE_SUPPRESSED. " & result("reason")
    Set OfficeDispatchManagedEmail = result
    Exit Function
  End If

  OfficeLogInfo logPath, "Outbox search result for " & managedKey & ": " & CStr(stateInfo("outbox_count"))
  If CLng(stateInfo("outbox_count")) = 0 Then
    Set OfficeDispatchManagedEmail = OfficeSubmitManagedMail(outlookApp, messageSpec, deliveryPolicy, logPath, OFFICE_STATE_COMPOSED)
    Exit Function
  End If

  If CLng(stateInfo("outbox_count")) > 1 Then
    result("state") = OFFICE_STATE_FAILED_SAFE
    result("reason") = "Multiple managed Outbox items matched; manual review required before repair."
    result("folder") = "Outbox"
    result("action_required") = True
    OfficeLogActionRequired logPath, result("reason") & " key=" & managedKey
    Set OfficeDispatchManagedEmail = result
    Exit Function
  End If

  Dim outboxItem: Set outboxItem = stateInfo("outbox_matches")("0")
  OfficeLogOutboxDiagnostic OfficeOutboxDiagnostic(outboxItem, managedKey, logPath), logPath

  Dim queuedDecision: Set queuedDecision = OfficeClassifyQueuedManagedMail(outlookApp, outboxItem, nowValue, graceMinutes)
  queuedDecision("managed_key") = managedKey
  queuedDecision("folder") = "Outbox"
  OfficeLogInfo logPath, "Queue decision: " & queuedDecision("state") & ". Reason: " & queuedDecision("reason")

  Select Case queuedDecision("state")
    Case OFFICE_STATE_ACTION_REQUIRED
      queuedDecision("action_required") = True
      OfficeLogActionRequired logPath, "Managed email is queued because Outlook appears offline/disconnected/auth-blocked. Keeping existing Outbox item."
      Set OfficeDispatchManagedEmail = queuedDecision
      Exit Function
    Case OFFICE_STATE_WAITING_DEFERRED, OFFICE_STATE_QUEUED_OUTBOX, OFFICE_STATE_FAILED_SAFE
      Set OfficeDispatchManagedEmail = queuedDecision
      Exit Function
    Case OFFICE_STATE_STALE_MANAGED_OUTBOX
      If Not enableReplacement Then
        queuedDecision("state") = OFFICE_STATE_ACTION_REQUIRED
        queuedDecision("reason") = "Managed email appears stale, but stale Outbox replacement is disabled."
        queuedDecision("action_required") = True
        OfficeLogActionRequired logPath, queuedDecision("reason") & " key=" & managedKey
        Set OfficeDispatchManagedEmail = queuedDecision
        Exit Function
      End If

      Dim finalSent: Set finalSent = OfficeFindManagedMailInSentItems(ns, managedKey, subjectText, allowFallback, logPath)
      If finalSent.Count > 0 Then
        result("state") = OFFICE_STATE_DUPLICATE_SUPPRESSED
        result("reason") = "Managed email reached Sent Items during final recheck."
        result("folder") = "Sent Items"
        OfficeLogInfo logPath, "Queue decision: DUPLICATE_SUPPRESSED. " & result("reason")
        Set OfficeDispatchManagedEmail = result
        Exit Function
      End If

      If Not OfficeCanReplaceManagedOutboxMail(outlookApp, outboxItem, managedKey, nowValue, graceMinutes) Then
        queuedDecision("state") = OFFICE_STATE_FAILED_SAFE
        queuedDecision("reason") = "Replacement guardrails failed; keeping existing Outbox item."
        OfficeLogWarning logPath, queuedDecision("reason") & " key=" & managedKey
        Set OfficeDispatchManagedEmail = queuedDecision
        Exit Function
      End If

      OfficeLogWarning logPath, "Managed email appears stale in Outbox while Outlook is connected and delivery time has passed. Replacing stale managed Outbox item. key=" & managedKey
      On Error Resume Next
      outboxItem.Delete
      If Err.Number <> 0 Then
        queuedDecision("state") = OFFICE_STATE_FAILED_SAFE
        queuedDecision("reason") = "Failed to delete stale managed Outbox item: " & Err.Description
        OfficeLogError logPath, queuedDecision("reason")
        Err.Clear
        Set OfficeDispatchManagedEmail = queuedDecision
        On Error GoTo 0
        Exit Function
      End If
      On Error GoTo 0

      Dim replacement: Set replacement = OfficeSubmitManagedMail(outlookApp, messageSpec, deliveryPolicy, logPath, OFFICE_STATE_REPLACED_MANAGED_OUTBOX)
      replacement("replaced_existing_mail") = True
      Set OfficeDispatchManagedEmail = replacement
      Exit Function
  End Select

  Set OfficeDispatchManagedEmail = queuedDecision
End Function
