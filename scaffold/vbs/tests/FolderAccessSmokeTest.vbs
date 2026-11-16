Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

Dim bootstrapPath
bootstrapPath = fso.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
Dim ts
Set ts = fso.OpenTextFile(bootstrapPath, 1)
ExecuteGlobal ts.ReadAll
ts.Close

SC_LoadCoreModules

Dim failures
failures = 0

Sub AssertTrue(condition, message)
  If condition Then
    WScript.Echo "OK: " & message
  Else
    WScript.Echo "FAILED: " & message
    failures = failures + 1
  End If
End Sub

Class FakeItems
  Public Count

  Private Sub Class_Initialize()
    Count = 0
  End Sub
End Class

Class FakeFolder
  Public Name
  Public Items

  Private Sub Class_Initialize()
    Set Items = New FakeItems
  End Sub
End Class

Class FakeAccount
  Public DisplayName
  Public SmtpAddress
  Public AccountType
End Class

Class FakeNamespace
  Private folders
  Private accountMap

  Private Sub Class_Initialize()
    Set folders = CreateObject("Scripting.Dictionary")
    Set accountMap = CreateObject("Scripting.Dictionary")
  End Sub

  Public Sub AddFolder(folderId, folder)
    folders.Add CStr(folderId), folder
  End Sub

  Public Function GetDefaultFolder(folderId)
    If folders.Exists(CStr(folderId)) Then
      Set GetDefaultFolder = folders(CStr(folderId))
    Else
      Err.Raise 4520, "FakeNamespace", "Folder not available in fake namespace"
    End If
  End Function

  Public Sub AddAccount(account)
    accountMap.Add CStr(accountMap.Count), account
  End Sub

  Public Property Get Accounts()
    Set Accounts = accountMap
  End Property
End Class

Function MakeFolder(name, countValue)
  Dim folder
  Set folder = New FakeFolder
  folder.Name = CStr(name)
  folder.Items.Count = CLng(countValue)
  Set MakeFolder = folder
End Function

Function MakeAccount(displayName, smtpAddress, accountType)
  Dim account
  Set account = New FakeAccount
  account.DisplayName = CStr(displayName)
  account.SmtpAddress = CStr(smtpAddress)
  account.AccountType = CStr(accountType)
  Set MakeAccount = account
End Function

Dim ns
Set ns = New FakeNamespace
ns.AddFolder SC_OL_FOLDER_OUTBOX, MakeFolder("Outbox", 3)
ns.AddFolder SC_OL_FOLDER_SENT, MakeFolder("Sent Items", 7)
ns.AddAccount MakeAccount("Primary Account", "primary@example.invalid", 0)
ns.AddAccount MakeAccount("Shared Account", "shared@example.invalid", 1)

Dim outboxAccess
Set outboxAccess = SC_GetOutboxFolderSafe(ns, "")
AssertTrue CBool(outboxAccess("success")), "Outbox access succeeds"
AssertTrue CLng(outboxAccess("folder_id")) = SC_OL_FOLDER_OUTBOX, "Outbox folder id is preserved"
AssertTrue CStr(outboxAccess("item_count")) = "3", "Outbox item count is captured"

Dim sentAccess
Set sentAccess = SC_GetSentItemsFolderSafe(ns, "")
AssertTrue CBool(sentAccess("success")), "Sent Items access succeeds"
AssertTrue CStr(sentAccess("item_count")) = "7", "Sent Items item count is captured"

Dim missingAccess
Set missingAccess = SC_GetDefaultFolderSafe(ns, 999, "Missing Folder", "")
AssertTrue Not CBool(missingAccess("success")), "Missing folder access fails safely"
AssertTrue Len(CStr(missingAccess("reason"))) > 0, "Missing folder records a reason"

Dim accounts
Set accounts = SC_ListOutlookAccounts(ns, "")
AssertTrue accounts.Count = 2, "Account list captures fake accounts"
AssertTrue accounts("0")("display_name") = "Primary Account", "Account display name is captured"
AssertTrue accounts("1")("smtp_address") = "shared@example.invalid", "Account SMTP address is captured"

Dim summary
Set summary = SC_GetNamespaceSummary(ns, "")
AssertTrue summary("account_count") = 2, "Namespace summary captures account count"
AssertTrue CBool(summary("outbox_accessible")), "Namespace summary captures Outbox access"
AssertTrue CBool(summary("sent_items_accessible")), "Namespace summary captures Sent Items access"

Dim emptyAccounts
Set emptyAccounts = SC_ListOutlookAccounts(Nothing, "")
AssertTrue emptyAccounts.Count = 0, "Unavailable namespace returns empty account list"

Dim noNamespaceAccess
Set noNamespaceAccess = SC_GetOutboxFolderSafe(Nothing, "")
AssertTrue Not CBool(noNamespaceAccess("success")), "Unavailable namespace fails folder access safely"

If failures > 0 Then
  WScript.Echo "Folder access smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Folder access smoke test passed."
