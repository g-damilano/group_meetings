Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

Sub LoadTextModule(modulePath)
  Dim ts
  Set ts = fso.OpenTextFile(modulePath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
SC_LoadCoreModules
LoadTextModule fso.BuildPath(scaffoldRoot, "vbs\watchdog\managed_outbox_repair.vbs")

Dim failures
failures = 0
Dim tempLogPath
tempLogPath = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_managed_repair_smoke.log")

Sub AssertTrue(condition, message)
  If condition Then
    WScript.Echo "OK: " & message
  Else
    WScript.Echo "FAILED: " & message
    failures = failures + 1
  End If
End Sub

Function FileContains(path, token)
  FileContains = (InStr(1, SC_ReadTextFile(path), CStr(token), vbTextCompare) > 0)
End Function

Class FakeUserProperty
  Public Value
End Class

Class FakeUserProperties
  Private props

  Private Sub Class_Initialize()
    Set props = CreateObject("Scripting.Dictionary")
  End Sub

  Public Function Find(name)
    If props.Exists(name) Then
      Set Find = props(name)
    Else
      Set Find = Nothing
    End If
  End Function

  Public Function Add(name, propType, addToFolderFields)
    Dim prop
    Set prop = New FakeUserProperty
    props(name) = prop
    Set Add = prop
  End Function
End Class

Class FakeAttachments
  Public Count
End Class

Class FakePropertyAccessor
  Public Function GetProperty(name)
    Err.Raise 4701, "FakePropertyAccessor", "Property unavailable"
  End Function
End Class

Class FakeMail
  Public UserProperties
  Public Attachments
  Public PropertyAccessor
  Public HTMLBody
  Public Body
  Public Subject
  Public EntryID
  Public CreationTime
  Public DeferredDeliveryTime
  Public DeleteCalled

  Private Sub Class_Initialize()
    Set UserProperties = New FakeUserProperties
    Set Attachments = New FakeAttachments
    Set PropertyAccessor = New FakePropertyAccessor
    HTMLBody = "<html><body><p>Queued</p></body></html>"
    Body = "Queued"
    Subject = "Queued"
    EntryID = "queued-entry"
    CreationTime = DateAdd("n", -30, Now)
    DeferredDeliveryTime = DateAdd("n", -1, Now)
    DeleteCalled = False
  End Sub

  Public Sub Delete()
    DeleteCalled = True
  End Sub
End Class

Class ReplacementFactory
  Public CreateCount

  Private Sub Class_Initialize()
    CreateCount = 0
  End Sub

  Public Function CreateReplacement(staleMail, itemReport)
    Dim mail
    Set mail = New FakeMail
    mail.Subject = "replacement for " & staleMail.Subject
    SC_StampManagedMail mail, CStr(itemReport("managed_key"))
    CreateCount = CreateCount + 1
    Set CreateReplacement = mail
  End Function
End Class

Class BadReplacementFactory
  Public Function CreateReplacement(staleMail, itemReport)
    Dim mail
    Set mail = New FakeMail
    mail.Subject = "bad replacement"
    Set CreateReplacement = mail
  End Function
End Class

Function MakeTransport(blocked)
  Dim report
  Set report = SC_CreateHealthReport()
  report("blocked_or_auth_suspected") = CBool(blocked)
  report("offline") = CBool(blocked)
  Set MakeTransport = report
End Function

Function NewMail(subjectText, managedKey, ageMinutes, deferredOffsetMinutes)
  Dim mail
  Set mail = New FakeMail
  mail.Subject = subjectText
  mail.EntryID = subjectText & "-entry"
  mail.CreationTime = DateAdd("n", -CLng(ageMinutes), Now)
  mail.DeferredDeliveryTime = DateAdd("n", CLng(deferredOffsetMinutes), Now)
  If Len(managedKey) > 0 Then SC_StampManagedMail mail, managedKey
  Set NewMail = mail
End Function

Dim staleKey
staleKey = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "target")
Dim youngKey
youngKey = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "young", "target")
Dim futureKey
futureKey = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "future", "target")

Dim staleManaged
Set staleManaged = NewMail("stale managed", staleKey, 40, -1)
Dim youngManaged
Set youngManaged = NewMail("young managed", youngKey, 5, -1)
Dim futureManaged
Set futureManaged = NewMail("future managed", futureKey, 40, 30)
Dim unmanaged
Set unmanaged = NewMail("user authored", "", 90, -1)

Dim items
Set items = SC_CreateDictionary()
Set items("stale") = staleManaged
Set items("young") = youngManaged
Set items("future") = futureManaged
Set items("unmanaged") = unmanaged

Dim factory
Set factory = New ReplacementFactory

Dim preview
Set preview = SC_WD_RunManagedRepair(items, MakeTransport(False), Now, 15, factory, False, tempLogPath)
AssertTrue preview("state") = SC_STATE_ACTION_REQUIRED, "Preview reports available repair as action-required"
AssertTrue preview("repair_available_count") = 1, "Preview finds only stale managed repair"
AssertTrue preview("repair_actions_taken") = 0, "Preview takes no repair actions"
AssertTrue factory.CreateCount = 0, "Preview does not create replacement"
AssertTrue staleManaged.DeleteCalled = False, "Preview does not delete stale managed mail"
AssertTrue unmanaged.DeleteCalled = False, "Preview does not delete unmarked user mail"

Dim committed
Set committed = SC_WD_RunManagedRepair(items, MakeTransport(False), Now, 15, factory, True, tempLogPath)
AssertTrue committed("state") = SC_STATE_REPLACED_MANAGED_OUTBOX, "Committed repair reports replacement"
AssertTrue committed("repair_actions_taken") = 1, "Committed repair takes one action"
AssertTrue factory.CreateCount = 1, "Committed repair creates one replacement"
AssertTrue staleManaged.DeleteCalled = True, "Committed repair deletes only stale managed mail"
AssertTrue youngManaged.DeleteCalled = False, "Committed repair leaves young managed mail"
AssertTrue futureManaged.DeleteCalled = False, "Committed repair leaves deferred managed mail"
AssertTrue unmanaged.DeleteCalled = False, "Committed repair leaves unmarked user mail"

Dim blockedStale
Set blockedStale = NewMail("blocked stale", staleKey, 40, -1)
Dim blockedItems
Set blockedItems = SC_CreateDictionary()
Set blockedItems("blocked") = blockedStale
Set committed = SC_WD_RunManagedRepair(blockedItems, MakeTransport(True), Now, 15, factory, True, tempLogPath)
AssertTrue committed("state") = SC_STATE_COMPOSED, "Blocked transport produces no eligible repair action"
AssertTrue blockedStale.DeleteCalled = False, "Blocked transport does not delete managed mail"

Dim uncertain
Set uncertain = NewMail("uncertain", staleKey, 40, -1)
uncertain.CreationTime = "not-a-date"
Dim uncertainItems
Set uncertainItems = SC_CreateDictionary()
Set uncertainItems("uncertain") = uncertain
Set committed = SC_WD_RunManagedRepair(uncertainItems, MakeTransport(False), Now, 15, factory, True, tempLogPath)
AssertTrue committed("state") = SC_STATE_FAILED_SAFE, "Uncertain repair fails safe"
AssertTrue uncertain.DeleteCalled = False, "Uncertain repair does not delete"

Dim noFactoryStale
Set noFactoryStale = NewMail("no factory stale", staleKey, 40, -1)
Dim noFactoryItems
Set noFactoryItems = SC_CreateDictionary()
Set noFactoryItems("stale") = noFactoryStale
Set committed = SC_WD_RunManagedRepair(noFactoryItems, MakeTransport(False), Now, 15, Nothing, True, tempLogPath)
AssertTrue committed("state") = SC_STATE_FAILED_SAFE, "Missing replacement factory fails safe"
AssertTrue noFactoryStale.DeleteCalled = False, "Missing replacement factory does not delete"

Dim badFactoryStale
Set badFactoryStale = NewMail("bad factory stale", staleKey, 40, -1)
Dim badFactoryItems
Set badFactoryItems = SC_CreateDictionary()
Set badFactoryItems("stale") = badFactoryStale
Dim badFactory
Set badFactory = New BadReplacementFactory
Set committed = SC_WD_RunManagedRepair(badFactoryItems, MakeTransport(False), Now, 15, badFactory, True, tempLogPath)
AssertTrue committed("state") = SC_STATE_FAILED_SAFE, "Unmarked replacement fails safe"
AssertTrue badFactoryStale.DeleteCalled = False, "Unmarked replacement does not trigger stale delete"

Dim repairPath
repairPath = fso.BuildPath(scaffoldRoot, "vbs\watchdog\managed_outbox_repair.vbs")
AssertTrue Not FileContains(repairPath, ".Send"), "Managed repair owns no send call"
AssertTrue Not FileContains(repairPath, "CreateObject(""Outlook.Application"")"), "Managed repair owns no Outlook creation"
AssertTrue Not FileContains(repairPath, "GetDefaultFolder"), "Managed repair owns no folder access"
AssertTrue FileContains(repairPath, ".Delete"), "Managed repair delete exists only in guarded repair module"

Dim fixturePath
fixturePath = fso.BuildPath(scaffoldRoot, "fixtures\watchdog\managed_repair_guardrail_fixture.json")
AssertTrue fso.FileExists(fixturePath), "Managed repair guardrail fixture exists"
AssertTrue FileContains(fixturePath, "office-communication-scaffold.watchdog-managed-repair-guardrail-fixture.v1"), "Managed repair fixture names canonical schema"
AssertTrue FileContains(fixturePath, """live_outlook_required"": false"), "Managed repair fixture requires no live Outlook"
AssertTrue FileContains(fixturePath, """uses_fake_outbox_items"": true"), "Managed repair fixture is fake-object based"
AssertTrue FileContains(fixturePath, """real_mailbox_mutation_allowed"": false"), "Managed repair fixture forbids real mailbox mutation"
AssertTrue FileContains(fixturePath, """send_allowed"": false"), "Managed repair fixture forbids sending"
AssertTrue FileContains(fixturePath, """unmarked_mail_mutation_allowed"": false"), "Managed repair fixture forbids unmarked mail mutation"
AssertTrue FileContains(fixturePath, "preview-stale-managed"), "Managed repair fixture covers preview mode"
AssertTrue FileContains(fixturePath, "commit-stale-managed-fake-only"), "Managed repair fixture covers committed fake-object repair"
AssertTrue FileContains(fixturePath, "commit-blocked-transport"), "Managed repair fixture covers blocked transport"
AssertTrue FileContains(fixturePath, "commit-unmarked-mail"), "Managed repair fixture covers unmarked mail"
AssertTrue FileContains(fixturePath, "commit-uncertain-age"), "Managed repair fixture covers uncertain age"
AssertTrue Not FileContains(fixturePath, """send_committed"": true"), "Managed repair fixture contains no committed send"
AssertTrue Not FileContains(fixturePath, """folder_accessed"": true"), "Managed repair fixture contains no folder access"
AssertTrue Not FileContains(fixturePath, """outlook_created"": true"), "Managed repair fixture contains no Outlook creation"
AssertTrue Not FileContains(fixturePath, """unmarked_deleted"": true"), "Managed repair fixture contains no unmarked delete"

If failures > 0 Then
  WScript.Echo "Managed Outbox repair smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

On Error Resume Next
If fso.FileExists(tempLogPath) Then fso.DeleteFile tempLogPath, True
On Error GoTo 0

WScript.Echo "Managed Outbox repair smoke test passed."
