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

  Private Sub Class_Initialize()
    Count = 0
  End Sub
End Class

Class FakePropertyAccessor
  Public Function GetProperty(name)
    Err.Raise 4502, "FakePropertyAccessor", "Property unavailable in decision test"
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
  Public LastModificationTime
  Public DeferredDeliveryTime

  Private Sub Class_Initialize()
    Set UserProperties = New FakeUserProperties
    Set Attachments = New FakeAttachments
    Set PropertyAccessor = New FakePropertyAccessor
    HTMLBody = "<html><body><p>Queued</p></body></html>"
    Body = "Queued"
    Subject = "Queued Test"
    EntryID = "fake-queued-entry"
    CreationTime = DateAdd("n", -20, Now)
    LastModificationTime = CreationTime
    DeferredDeliveryTime = DateAdd("n", -1, Now)
  End Sub
End Class

Function MakeTransport(blocked)
  Dim report
  Set report = SC_CreateHealthReport()
  report("blocked_or_auth_suspected") = CBool(blocked)
  report("offline") = CBool(blocked)
  If CBool(blocked) Then
    report("exchange_connection_mode") = 300
    report("exchange_connection_mode_name") = "olDisconnected"
  Else
    report("exchange_connection_mode") = 800
    report("exchange_connection_mode_name") = "olOnline"
  End If
  Set MakeTransport = report
End Function

Dim nowValue
nowValue = Now
Dim graceMinutes
graceMinutes = 15
Dim key
key = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "target")

Dim mail
Set mail = New FakeMail
SC_StampManagedMail mail, key
mail.CreationTime = DateAdd("n", -30, nowValue)
mail.DeferredDeliveryTime = DateAdd("n", -1, nowValue)

Dim decision
Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(False), mail, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_STALE_MANAGED_OUTBOX, "stale connected managed mail is classified stale"
AssertTrue SC_CanReplaceQueuedManagedMail(MakeTransport(False), mail, key, nowValue, graceMinutes), "stale connected managed mail is replace-eligible"

Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(True), mail, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_ACTION_REQUIRED, "blocked transport becomes action required"
AssertTrue decision("action_required") = True, "blocked transport action flag set"
AssertTrue Not SC_CanReplaceQueuedManagedMail(MakeTransport(True), mail, key, nowValue, graceMinutes), "blocked transport is not replace-eligible"

mail.DeferredDeliveryTime = DateAdd("n", 30, nowValue)
Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(False), mail, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_WAITING_DEFERRED, "future deferred delivery waits"
AssertTrue Not SC_CanReplaceQueuedManagedMail(MakeTransport(False), mail, key, nowValue, graceMinutes), "future deferred mail is not replace-eligible"

mail.DeferredDeliveryTime = DateAdd("n", -1, nowValue)
mail.CreationTime = DateAdd("n", -5, nowValue)
Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(False), mail, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_QUEUED_OUTBOX, "young queued mail remains active attempt"
AssertTrue Not SC_CanReplaceQueuedManagedMail(MakeTransport(False), mail, key, nowValue, graceMinutes), "young queued mail is not replace-eligible"

Dim unmarked
Set unmarked = New FakeMail
unmarked.CreationTime = DateAdd("n", -30, nowValue)
Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(False), unmarked, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_FAILED_SAFE, "unmarked mail fails safe"
AssertTrue Not SC_CanReplaceQueuedManagedMail(MakeTransport(False), unmarked, key, nowValue, graceMinutes), "unmarked mail is never replace-eligible"

Dim wrongKeyMail
Set wrongKeyMail = New FakeMail
SC_StampManagedMail wrongKeyMail, "ORG|WORKFLOW|2026-05-29|OTHER|TARGET"
wrongKeyMail.CreationTime = DateAdd("n", -30, nowValue)
Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(False), wrongKeyMail, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_FAILED_SAFE, "wrong managed key fails safe"
AssertTrue Not SC_CanReplaceQueuedManagedMail(MakeTransport(False), wrongKeyMail, key, nowValue, graceMinutes), "wrong managed key is not replace-eligible"

Dim uncertain
Set uncertain = New FakeMail
SC_StampManagedMail uncertain, key
uncertain.CreationTime = "not-a-date"
Set decision = SC_GetQueuedManagedMailDecisionFromTransport(MakeTransport(False), uncertain, key, nowValue, graceMinutes)
AssertTrue decision("state") = SC_STATE_FAILED_SAFE, "uncertain age fails safe"
AssertTrue Not SC_CanReplaceQueuedManagedMail(MakeTransport(False), uncertain, key, nowValue, graceMinutes), "uncertain age is not replace-eligible"

If failures > 0 Then
  WScript.Echo "Managed Outbox decision smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Managed Outbox decision smoke test passed."

