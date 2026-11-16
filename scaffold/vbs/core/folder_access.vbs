' Safe Outlook account and folder access helpers.

Const SC_OL_FOLDER_OUTBOX = 4
Const SC_OL_FOLDER_SENT = 5

Function SC_FolderNameForId(folderId)
  Select Case CLng(folderId)
    Case SC_OL_FOLDER_OUTBOX
      SC_FolderNameForId = "Outbox"
    Case SC_OL_FOLDER_SENT
      SC_FolderNameForId = "Sent Items"
    Case Else
      SC_FolderNameForId = "Folder " & CStr(folderId)
  End Select
End Function

Function SC_CreateFolderAccessResult(folderId, folderName)
  Dim result
  Set result = SC_CreateDictionary()
  result("success") = False
  result("reason") = ""
  result("folder_id") = CLng(folderId)
  result("folder_name") = CStr(folderName)
  result("item_count") = ""
  Set SC_CreateFolderAccessResult = result
End Function

Function SC_GetFolderItemCountSafe(folder)
  SC_GetFolderItemCountSafe = ""
  If folder Is Nothing Then Exit Function

  On Error Resume Next
  SC_GetFolderItemCountSafe = folder.Items.Count
  If Err.Number <> 0 Then
    SC_GetFolderItemCountSafe = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Sub SC_LogFolderAccessWarning(logPath, message)
  If Len(Trim(CStr(logPath))) > 0 Then SC_LogWarning logPath, message
End Sub

Function SC_GetDefaultFolderSafe(ns, folderId, folderName, logPath)
  Dim resolvedName
  resolvedName = CStr(folderName)
  If Len(Trim(resolvedName)) = 0 Then resolvedName = SC_FolderNameForId(folderId)

  Dim result
  Set result = SC_CreateFolderAccessResult(folderId, resolvedName)

  If ns Is Nothing Then
    result("reason") = "Outlook MAPI namespace is unavailable."
    SC_LogFolderAccessWarning logPath, "Unable to access " & resolvedName & ": " & result("reason")
    Set SC_GetDefaultFolderSafe = result
    Exit Function
  End If

  On Error Resume Next
  Dim folder
  Set folder = ns.GetDefaultFolder(CLng(folderId))
  If Err.Number <> 0 Or folder Is Nothing Then
    result("reason") = Err.Description
    If Len(Trim(CStr(result("reason")))) = 0 Then result("reason") = "Folder object was not returned."
    Err.Clear
    On Error GoTo 0
    SC_LogFolderAccessWarning logPath, "Unable to access " & resolvedName & ": " & result("reason")
    Set SC_GetDefaultFolderSafe = result
    Exit Function
  End If
  On Error GoTo 0

  result("success") = True
  result("reason") = "OK"
  result("item_count") = SC_GetFolderItemCountSafe(folder)
  result.Add "folder", folder

  Set SC_GetDefaultFolderSafe = result
End Function

Function SC_GetOutboxFolderSafe(ns, logPath)
  Set SC_GetOutboxFolderSafe = SC_GetDefaultFolderSafe(ns, SC_OL_FOLDER_OUTBOX, "Outbox", logPath)
End Function

Function SC_GetSentItemsFolderSafe(ns, logPath)
  Set SC_GetSentItemsFolderSafe = SC_GetDefaultFolderSafe(ns, SC_OL_FOLDER_SENT, "Sent Items", logPath)
End Function

Function SC_CreateAccountInfo(displayName, smtpAddress, accountType)
  Dim info
  Set info = SC_CreateDictionary()
  info("display_name") = CStr(displayName)
  info("smtp_address") = CStr(smtpAddress)
  info("account_type") = CStr(accountType)
  Set SC_CreateAccountInfo = info
End Function

Function SC_SafeAccountProperty(account, propertyName)
  SC_SafeAccountProperty = ""
  If account Is Nothing Then Exit Function

  On Error Resume Next
  Dim value
  Select Case LCase(Trim(CStr(propertyName)))
    Case "displayname"
      value = account.DisplayName
    Case "smtpaddress"
      value = account.SmtpAddress
    Case "accounttype"
      value = account.AccountType
    Case Else
      value = ""
  End Select

  If Err.Number = 0 Then
    SC_SafeAccountProperty = CStr(value)
  Else
    Err.Clear
    SC_SafeAccountProperty = ""
  End If
  On Error GoTo 0
End Function

Sub SC_AppendAccountInfo(accountList, account)
  Dim info
  Set info = SC_CreateAccountInfo( _
    SC_SafeAccountProperty(account, "DisplayName"), _
    SC_SafeAccountProperty(account, "SmtpAddress"), _
    SC_SafeAccountProperty(account, "AccountType"))
  accountList.Add CStr(accountList.Count), info
End Sub

Function SC_ListOutlookAccounts(ns, logPath)
  Dim accountList
  Set accountList = SC_CreateDictionary()

  If ns Is Nothing Then
    SC_LogFolderAccessWarning logPath, "Unable to list Outlook accounts: Outlook MAPI namespace is unavailable."
    Set SC_ListOutlookAccounts = accountList
    Exit Function
  End If

  On Error Resume Next
  Dim accounts
  Set accounts = ns.Accounts
  If Err.Number <> 0 Or accounts Is Nothing Then
    Dim reason
    reason = Err.Description
    If Len(Trim(CStr(reason))) = 0 Then reason = "Accounts collection was not returned."
    Err.Clear
    On Error GoTo 0
    SC_LogFolderAccessWarning logPath, "Unable to list Outlook accounts: " & reason
    Set SC_ListOutlookAccounts = accountList
    Exit Function
  End If
  On Error GoTo 0

  If TypeName(accounts) = "Dictionary" Then
    Dim key
    For Each key In accounts.Keys
      If IsObject(accounts(key)) Then SC_AppendAccountInfo accountList, accounts(key)
    Next
  Else
    Dim account
    For Each account In accounts
      If IsObject(account) Then SC_AppendAccountInfo accountList, account
    Next
  End If

  Set SC_ListOutlookAccounts = accountList
End Function

Function SC_GetNamespaceSummary(ns, logPath)
  Dim summary
  Set summary = SC_CreateDictionary()
  summary("account_count") = 0
  summary("outbox_accessible") = False
  summary("outbox_count") = ""
  summary("sent_items_accessible") = False
  summary("sent_items_count") = ""

  Dim accounts
  Set accounts = SC_ListOutlookAccounts(ns, logPath)
  summary("account_count") = accounts.Count

  Dim outboxAccess
  Set outboxAccess = SC_GetOutboxFolderSafe(ns, logPath)
  summary("outbox_accessible") = CBool(outboxAccess("success"))
  summary("outbox_count") = outboxAccess("item_count")

  Dim sentAccess
  Set sentAccess = SC_GetSentItemsFolderSafe(ns, logPath)
  summary("sent_items_accessible") = CBool(sentAccess("success"))
  summary("sent_items_count") = sentAccess("item_count")

  Set SC_GetNamespaceSummary = summary
End Function
