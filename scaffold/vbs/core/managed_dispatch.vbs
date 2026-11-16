' Managed email identity, lookup, diagnostics, and dispatch guardrails.

Const SC_OL_MAIL_ITEM = 0
Const SC_OL_FORMAT_HTML = 2
Const SC_MSGFLAG_SUBMIT = 4
Const SC_MSGFLAG_UNSENT = 8
Const SC_PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
Const SC_PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"

Function SC_ManagedMarker(managedKey)
  SC_ManagedMarker = "<!-- " & SC_MANAGED_KEY_MARKER & ": " & managedKey & " -->"
End Function

Sub SC_StampManagedMail(mailItem, managedKey)
  On Error Resume Next
  Dim prop
  Set prop = mailItem.UserProperties.Find(SC_MANAGED_KEY_PROPERTY)
  If prop Is Nothing Then Set prop = mailItem.UserProperties.Add(SC_MANAGED_KEY_PROPERTY, 1, True)
  prop.Value = managedKey
  Err.Clear

  Dim htmlBody
  htmlBody = mailItem.HTMLBody
  If Err.Number = 0 Then
    If InStr(1, htmlBody, SC_MANAGED_KEY_MARKER & ":", vbTextCompare) = 0 Then
      Dim marker
      marker = SC_ManagedMarker(managedKey)
      Dim closePos
      closePos = InStr(1, htmlBody, "</body>", vbTextCompare)
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

Function SC_ExtractManagedKeyFromText(textValue)
  SC_ExtractManagedKeyFromText = ""
  Dim pos
  pos = InStr(1, CStr(textValue), SC_MANAGED_KEY_MARKER & ":", vbTextCompare)
  If pos = 0 Then Exit Function

  Dim restText
  restText = Trim(Mid(CStr(textValue), pos + Len(SC_MANAGED_KEY_MARKER & ":")))
  Dim endPos
  endPos = InStr(1, restText, "-->", vbTextCompare)
  If endPos = 0 Then endPos = InStr(1, restText, vbCrLf, vbTextCompare)
  If endPos = 0 Then endPos = InStr(1, restText, "<", vbTextCompare)

  If endPos > 0 Then
    SC_ExtractManagedKeyFromText = Trim(Left(restText, endPos - 1))
  Else
    SC_ExtractManagedKeyFromText = Trim(restText)
  End If
End Function

Function SC_GetManagedMailKey(mailItem)
  SC_GetManagedMailKey = ""
  On Error Resume Next
  Dim prop
  Set prop = mailItem.UserProperties.Find(SC_MANAGED_KEY_PROPERTY)
  If Err.Number = 0 Then
    If Not prop Is Nothing Then
      If Len(Trim(CStr(prop.Value))) > 0 Then
        SC_GetManagedMailKey = Trim(CStr(prop.Value))
        On Error GoTo 0
        Exit Function
      End If
    End If
  End If
  Err.Clear

  Dim htmlKey
  htmlKey = SC_ExtractManagedKeyFromText(mailItem.HTMLBody)
  If Len(htmlKey) > 0 Then
    SC_GetManagedMailKey = htmlKey
    On Error GoTo 0
    Exit Function
  End If
  Err.Clear

  Dim bodyKey
  bodyKey = SC_ExtractManagedKeyFromText(mailItem.Body)
  If Len(bodyKey) > 0 Then SC_GetManagedMailKey = bodyKey
  Err.Clear
  On Error GoTo 0
End Function

Function SC_FindManagedMailInFolder(ns, folderId, managedKey, subjectText, allowSubjectFallback, logPath)
  Dim matches
  Set matches = SC_CreateDictionary()

  Dim folderAccess
  Set folderAccess = SC_GetDefaultFolderSafe(ns, folderId, SC_FolderNameForId(folderId), "")
  If Not CBool(folderAccess("success")) Then
    SC_LogWarning logPath, "Unable to access Outlook " & folderAccess("folder_name") & _
      " while searching key " & managedKey & ": " & folderAccess("reason")
    Set SC_FindManagedMailInFolder = matches
    Exit Function
  End If

  Dim folder
  Set folder = folderAccess("folder")

  On Error Resume Next
  Dim items
  Set items = folder.Items
  If CLng(folderId) = SC_OL_FOLDER_SENT Then
    items.Sort "SentOn", True
  ElseIf CLng(folderId) = SC_OL_FOLDER_OUTBOX Then
    items.Sort "CreationTime", True
  End If
  On Error GoTo 0

  Dim item
  For Each item In items
    If TypeName(item) = "MailItem" Then
      Dim itemKey
      itemKey = SC_GetManagedMailKey(item)
      If itemKey = managedKey Then
        matches.Add CStr(matches.Count), item
      ElseIf CBool(allowSubjectFallback) And Len(itemKey) = 0 Then
        On Error Resume Next
        If CStr(item.Subject) = CStr(subjectText) Then
          matches.Add CStr(matches.Count), item
          SC_LogWarning logPath, "Legacy subject-only fallback matched subject: " & subjectText
        End If
        Err.Clear
        On Error GoTo 0
      End If
    End If
  Next

  Set SC_FindManagedMailInFolder = matches
End Function

Function SC_FindManagedMailInSentItems(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  Set SC_FindManagedMailInSentItems = SC_FindManagedMailInFolder(ns, SC_OL_FOLDER_SENT, managedKey, subjectText, allowSubjectFallback, logPath)
End Function

Function SC_FindManagedMailInOutbox(ns, managedKey, subjectText, allowSubjectFallback, logPath)
  Set SC_FindManagedMailInOutbox = SC_FindManagedMailInFolder(ns, SC_OL_FOLDER_OUTBOX, managedKey, subjectText, allowSubjectFallback, logPath)
End Function

Function SC_OutboxDiagnostic(mailItem, managedKey, logPath)
  Dim diag
  Set diag = SC_CreateDictionary()
  diag("managed_key") = managedKey
  diag("subject") = ""
  diag("to") = ""
  diag("entry_id") = ""
  diag("creation_time") = ""
  diag("last_modification_time") = ""
  diag("attachment_count") = ""
  diag("deferred_delivery_time") = ""
  diag("pr_message_flags") = "[unavailable]"
  diag("msgflag_submit") = "[unavailable]"
  diag("msgflag_unsent") = "[unavailable]"
  diag("pr_client_submit_time") = "[unavailable]"

  On Error Resume Next
  diag("subject") = CStr(mailItem.Subject)
  diag("to") = CStr(mailItem.To)
  diag("entry_id") = CStr(mailItem.EntryID)
  diag("creation_time") = CStr(mailItem.CreationTime)
  diag("last_modification_time") = CStr(mailItem.LastModificationTime)
  diag("attachment_count") = CStr(mailItem.Attachments.Count)
  diag("deferred_delivery_time") = CStr(mailItem.DeferredDeliveryTime)
  Err.Clear

  Dim flags
  flags = mailItem.PropertyAccessor.GetProperty(SC_PR_MESSAGE_FLAGS)
  If Err.Number = 0 Then
    diag("pr_message_flags") = flags
    diag("msgflag_submit") = CBool((CLng(flags) And SC_MSGFLAG_SUBMIT) <> 0)
    diag("msgflag_unsent") = CBool((CLng(flags) And SC_MSGFLAG_UNSENT) <> 0)
  Else
    SC_LogWarning logPath, "MAPI lookup failed for PR_MESSAGE_FLAGS: " & Err.Description
    Err.Clear
  End If

  Dim submitTime
  submitTime = mailItem.PropertyAccessor.GetProperty(SC_PR_CLIENT_SUBMIT_TIME)
  If Err.Number = 0 Then
    diag("pr_client_submit_time") = CStr(submitTime)
  Else
    SC_LogWarning logPath, "MAPI lookup failed for PR_CLIENT_SUBMIT_TIME: " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0

  Set SC_OutboxDiagnostic = diag
End Function

Function SC_ClassifyQueuedManagedMail(outlookApp, mailItem, nowValue, graceMinutes)
  Dim transport
  Set transport = SC_GetTransportState(outlookApp)
  Set SC_ClassifyQueuedManagedMail = SC_GetQueuedManagedMailDecisionFromTransport(transport, mailItem, SC_GetManagedMailKey(mailItem), nowValue, graceMinutes)
End Function

Function SC_TransportBlockedOrSuspected(transport)
  SC_TransportBlockedOrSuspected = True
  If Not IsObject(transport) Then Exit Function

  On Error Resume Next
  If transport.Exists("blocked_or_auth_suspected") Then
    SC_TransportBlockedOrSuspected = CBool(transport("blocked_or_auth_suspected"))
  Else
    SC_TransportBlockedOrSuspected = True
  End If
  If Err.Number <> 0 Then
    Err.Clear
    SC_TransportBlockedOrSuspected = True
  End If
  On Error GoTo 0
End Function

Function SC_QueuedMailDeferredInFuture(mailItem, nowValue)
  SC_QueuedMailDeferredInFuture = False
  On Error Resume Next
  Dim deferredValue
  deferredValue = mailItem.DeferredDeliveryTime
  If Err.Number = 0 Then
    If IsDate(deferredValue) Then
      SC_QueuedMailDeferredInFuture = (CDate(deferredValue) > CDate(nowValue))
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_QueuedMailAgeMinutes(mailItem, nowValue)
  SC_QueuedMailAgeMinutes = ""
  On Error Resume Next
  Dim creationValue
  creationValue = mailItem.CreationTime
  If Err.Number = 0 Then
    If IsDate(creationValue) Then SC_QueuedMailAgeMinutes = DateDiff("n", CDate(creationValue), CDate(nowValue))
  End If
  If Err.Number <> 0 Then
    Err.Clear
    SC_QueuedMailAgeMinutes = ""
  End If
  On Error GoTo 0
End Function

Function SC_GetQueuedManagedMailDecisionFromTransport(transport, mailItem, managedKey, nowValue, graceMinutes)
  Dim result
  Set result = SC_CreateDispatchResult(SC_STATE_FAILED_SAFE, "")
  result("managed_key") = CStr(managedKey)
  result("folder") = "Outbox"

  If Len(Trim(CStr(managedKey))) = 0 Then
    result("reason") = "Managed key is empty; refusing to classify for repair."
    Set SC_GetQueuedManagedMailDecisionFromTransport = result
    Exit Function
  End If

  If SC_GetManagedMailKey(mailItem) <> CStr(managedKey) Then
    result("reason") = "Outbox item does not have the expected managed key."
    Set SC_GetQueuedManagedMailDecisionFromTransport = result
    Exit Function
  End If

  If SC_TransportBlockedOrSuspected(transport) Then
    Set result = SC_CreateDispatchResult(SC_STATE_ACTION_REQUIRED, "Outlook appears offline/disconnected/auth-blocked.")
    result("managed_key") = CStr(managedKey)
    result("folder") = "Outbox"
    result("action_required") = True
    Set SC_GetQueuedManagedMailDecisionFromTransport = result
    Exit Function
  End If

  If SC_QueuedMailDeferredInFuture(mailItem, nowValue) Then
    Set result = SC_CreateDispatchResult(SC_STATE_WAITING_DEFERRED, "Deferred delivery time has not been reached.")
    result("managed_key") = CStr(managedKey)
    result("folder") = "Outbox"
    Set SC_GetQueuedManagedMailDecisionFromTransport = result
    Exit Function
  End If

  Dim ageMinutes
  ageMinutes = SC_QueuedMailAgeMinutes(mailItem, nowValue)
  If Len(CStr(ageMinutes)) = 0 Then
    result("reason") = "Unable to determine queued item age."
    Set SC_GetQueuedManagedMailDecisionFromTransport = result
    Exit Function
  End If

  If CLng(ageMinutes) < CLng(graceMinutes) Then
    Set result = SC_CreateDispatchResult(SC_STATE_QUEUED_OUTBOX, "Queued item is still within grace period.")
  Else
    Set result = SC_CreateDispatchResult(SC_STATE_STALE_MANAGED_OUTBOX, "Managed item appears stale in Outbox.")
  End If
  result("managed_key") = CStr(managedKey)
  result("folder") = "Outbox"
  Set SC_GetQueuedManagedMailDecisionFromTransport = result
End Function

Function SC_CanReplaceQueuedManagedMail(transport, mailItem, managedKey, nowValue, graceMinutes)
  SC_CanReplaceQueuedManagedMail = False

  If Len(Trim(CStr(managedKey))) = 0 Then Exit Function
  If SC_GetManagedMailKey(mailItem) <> CStr(managedKey) Then Exit Function
  If SC_TransportBlockedOrSuspected(transport) Then Exit Function
  If SC_QueuedMailDeferredInFuture(mailItem, nowValue) Then Exit Function

  Dim ageMinutes
  ageMinutes = SC_QueuedMailAgeMinutes(mailItem, nowValue)
  If Len(CStr(ageMinutes)) = 0 Then Exit Function

  SC_CanReplaceQueuedManagedMail = (CLng(ageMinutes) >= CLng(graceMinutes))
End Function

Function SC_DictValue(dictionary, key, defaultValue)
  SC_DictValue = defaultValue
  On Error Resume Next
  If IsObject(dictionary) Then
    If dictionary.Exists(key) Then SC_DictValue = dictionary(key)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function SC_Bool(value, defaultValue)
  On Error Resume Next
  If IsEmpty(value) Or IsNull(value) Then
    SC_Bool = CBool(defaultValue)
  Else
    SC_Bool = CBool(value)
  End If
  If Err.Number <> 0 Then
    Err.Clear
    SC_Bool = CBool(defaultValue)
  End If
  On Error GoTo 0
End Function

Function SC_CreateManagedMail(outlookApp, messageSpec, deliveryPolicy)
  Set SC_CreateManagedMail = SC_ComposeOutlookMail(outlookApp, messageSpec, deliveryPolicy, Nothing, "")
End Function

Function SC_DispatchManagedEmail(messageSpec, deliveryPolicy, dispatchConfig)
  Dim logPath
  logPath = CStr(SC_DictValue(dispatchConfig, "log_path", ""))
  Dim managedKey
  managedKey = CStr(SC_DictValue(messageSpec, "message_key", ""))
  Dim subjectText
  subjectText = CStr(SC_DictValue(messageSpec, "subject", ""))

  Dim result
  Set result = SC_CreateDispatchResult(SC_STATE_NOT_REQUESTED, "")
  result("managed_key") = managedKey

  If Len(managedKey) = 0 Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "MessageSpec has no managed email key."
    SC_LogError logPath, result("reason")
    Set SC_DispatchManagedEmail = result
    Exit Function
  End If

  Dim dryRun
  dryRun = SC_Bool(SC_DictValue(dispatchConfig, "dry_run", True), True)
  If dryRun Then
    result("state") = SC_STATE_COMPOSED
    result("reason") = "Dry run: managed email spec accepted without Outlook submission."
    SC_LogInfo logPath, "Dry-run dispatch accepted key=" & managedKey & " subject=" & subjectText
    Set SC_DispatchManagedEmail = result
    Exit Function
  End If

  Dim outlookApp
  Set outlookApp = Nothing
  On Error Resume Next
  If IsObject(dispatchConfig) Then
    If dispatchConfig.Exists("outlook_app") Then
      If IsObject(dispatchConfig("outlook_app")) Then Set outlookApp = dispatchConfig("outlook_app")
    End If
  End If
  Err.Clear
  On Error GoTo 0
  If outlookApp Is Nothing Then Set outlookApp = SC_GetOutlookApplication()

  If outlookApp Is Nothing Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Outlook.Application is unavailable."
    result("action_required") = True
    SC_LogActionRequired logPath, result("reason")
    Set SC_DispatchManagedEmail = result
    Exit Function
  End If

  Dim ns
  Set ns = SC_GetMAPI(outlookApp)
  If ns Is Nothing Then
    result("state") = SC_STATE_ACTION_REQUIRED
    result("reason") = "Outlook MAPI namespace is unavailable."
    result("action_required") = True
    SC_LogActionRequired logPath, result("reason")
    Set SC_DispatchManagedEmail = result
    Exit Function
  End If

  Dim allowFallback
  allowFallback = SC_Bool(SC_DictValue(dispatchConfig, "enable_subject_fallback", True), True)
  Dim sentMatches
  Set sentMatches = SC_FindManagedMailInSentItems(ns, managedKey, subjectText, allowFallback, logPath)
  If sentMatches.Count > 0 Then
    result("state") = SC_STATE_DUPLICATE_SUPPRESSED
    result("reason") = "Managed email already exists in Sent Items."
    result("folder") = "Sent Items"
    Set SC_DispatchManagedEmail = result
    Exit Function
  End If

  Dim outboxMatches
  Set outboxMatches = SC_FindManagedMailInOutbox(ns, managedKey, subjectText, allowFallback, logPath)
  If outboxMatches.Count > 0 Then
    result("state") = SC_STATE_DUPLICATE_SUPPRESSED
    result("reason") = "Managed email already exists in Outbox."
    result("folder") = "Outbox"
    Set SC_DispatchManagedEmail = result
    Exit Function
  End If

  Dim action
  action = LCase(Trim(CStr(SC_DictValue(deliveryPolicy, "action", "display"))))
  Dim mailItem
  On Error Resume Next
  Set mailItem = SC_ComposeOutlookMail(outlookApp, messageSpec, deliveryPolicy, dispatchConfig, logPath)
  If Err.Number <> 0 Then
    result("state") = SC_STATE_FAILED_SAFE
    result("reason") = "Failed to compose managed email: " & Err.Description
    SC_LogError logPath, result("reason")
    Err.Clear
    Set SC_DispatchManagedEmail = result
    On Error GoTo 0
    Exit Function
  End If

  Select Case action
    Case "send"
      mailItem.Send
      If Err.Number = 0 Then
        result("state") = SC_STATE_SEND_REQUESTED
        result("reason") = "Managed email submitted to Outlook."
        result("created_new_mail") = True
      Else
        result("state") = SC_STATE_FAILED_SAFE
        result("reason") = "Outlook Send failed: " & Err.Description
        Err.Clear
      End If
    Case "draft", "save"
      mailItem.Save
      If Err.Number = 0 Then
        result("state") = SC_STATE_DRAFTED
        result("reason") = "Managed email saved as draft."
        result("created_new_mail") = True
      Else
        result("state") = SC_STATE_FAILED_SAFE
        result("reason") = "Outlook Save failed: " & Err.Description
        Err.Clear
      End If
    Case Else
      mailItem.Display
      If Err.Number = 0 Then
        result("state") = SC_STATE_DISPLAYED
        result("reason") = "Managed email displayed."
        result("created_new_mail") = True
      Else
        result("state") = SC_STATE_FAILED_SAFE
        result("reason") = "Outlook Display failed: " & Err.Description
        Err.Clear
      End If
  End Select

  If Not mailItem Is Nothing Then result("outlook_entry_id") = CStr(mailItem.EntryID)
  On Error GoTo 0

  Set SC_DispatchManagedEmail = result
End Function
