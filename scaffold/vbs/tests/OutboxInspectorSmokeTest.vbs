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
Dim tempLogPath
tempLogPath = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_outbox_inspector_smoke.log")

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
    Err.Raise 4503, "FakePropertyAccessor", "Property unavailable in inspector test"
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
  Public DeleteCalled

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
    DeleteCalled = False
  End Sub

  Public Sub Delete()
    DeleteCalled = True
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

Function NewMail(subjectText, keyValue, ageMinutes, deferredOffsetMinutes)
  Dim mail
  Set mail = New FakeMail
  mail.Subject = subjectText
  mail.EntryID = subjectText & "-entry"
  mail.CreationTime = DateAdd("n", -CLng(ageMinutes), Now)
  mail.LastModificationTime = mail.CreationTime
  mail.DeferredDeliveryTime = DateAdd("n", CLng(deferredOffsetMinutes), Now)
  If Len(keyValue) > 0 Then SC_StampManagedMail mail, keyValue
  Set NewMail = mail
End Function

Dim key
key = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "target")
Dim futureKey
futureKey = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "reminder", "target")
Dim youngKey
youngKey = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "same-day", "target")

Dim staleManaged
Set staleManaged = NewMail("stale managed", key, 40, -1)
Dim futureManaged
Set futureManaged = NewMail("future managed", futureKey, 40, 30)
Dim youngManaged
Set youngManaged = NewMail("young managed", youngKey, 5, -1)
Dim unmanaged
Set unmanaged = NewMail("unmanaged user mail", "", 90, -1)

Dim items
Set items = SC_CreateDictionary()
Set items("stale") = staleManaged
Set items("future") = futureManaged
Set items("young") = youngManaged
Set items("unmanaged") = unmanaged

Dim report
Set report = SC_InspectOutboxItems(items, MakeTransport(False), Now, 15, tempLogPath)

AssertTrue report("total_count") = 4, "all Outbox items inspected"
AssertTrue report("managed_count") = 3, "managed Outbox count"
AssertTrue report("unmanaged_count") = 1, "unmanaged Outbox count"
AssertTrue report("stale_count") = 1, "stale managed count"
AssertTrue report("waiting_deferred_count") = 1, "waiting deferred count"
AssertTrue report("queued_count") = 1, "young queued count"
AssertTrue report("replace_eligible_count") = 1, "only stale connected managed mail is replace-eligible"
AssertTrue report("repair_actions_taken") = 0, "report-only inspection takes no repair action"
AssertTrue staleManaged.DeleteCalled = False, "stale managed item not deleted in report-only inspection"
AssertTrue unmanaged.DeleteCalled = False, "unmanaged user mail not deleted"

Dim itemReport
Set itemReport = report("items")("3")
AssertTrue itemReport("is_managed") = False, "unmanaged item explicitly reported as unmanaged"
AssertTrue itemReport("would_modify") = False, "unmanaged item would not be modified"

Set report = SC_InspectOutboxItems(items, MakeTransport(True), Now, 15, tempLogPath)
AssertTrue report("action_required_count") = 3, "blocked transport marks managed items action-required"
AssertTrue report("replace_eligible_count") = 0, "blocked transport has no replace eligibility"
AssertTrue report("unmanaged_count") = 1, "blocked transport still counts unmanaged mail without modifying it"
AssertTrue unmanaged.DeleteCalled = False, "unmanaged user mail remains untouched when blocked"

Dim uncertainItems
Set uncertainItems = SC_CreateDictionary()
Dim uncertain
Set uncertain = NewMail("uncertain managed", key, 40, -1)
uncertain.CreationTime = "not-a-date"
Set uncertainItems("uncertain") = uncertain
Set report = SC_InspectOutboxItems(uncertainItems, MakeTransport(False), Now, 15, tempLogPath)
AssertTrue report("failed_safe_count") = 1, "uncertain item fails safe"
AssertTrue report("replace_eligible_count") = 0, "uncertain item is not replace-eligible"
AssertTrue uncertain.DeleteCalled = False, "uncertain item not modified"

If failures > 0 Then
  WScript.Echo "Outbox inspector smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

On Error Resume Next
If fso.FileExists(tempLogPath) Then fso.DeleteFile tempLogPath, True
On Error GoTo 0

WScript.Echo "Outbox inspector smoke test passed."

