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
tempLogPath = fso.BuildPath(fso.GetSpecialFolder(2), "office_scaffold_managed_dispatch_smoke.log")

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
    Err.Raise 4501, "FakePropertyAccessor", "Property unavailable in smoke test"
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
    HTMLBody = "<html><body><p>Hello</p></body></html>"
    Body = "Hello"
    Subject = "Smoke Test"
    EntryID = "fake-entry"
    CreationTime = DateAdd("n", -20, Now)
    LastModificationTime = CreationTime
    DeferredDeliveryTime = DateAdd("n", -1, Now)
  End Sub
End Class

Dim key
key = SC_BuildManagedEmailKey("org", "workflow", #5/29/2026#, "regular", "target")

Dim mail
Set mail = New FakeMail
SC_StampManagedMail mail, key

AssertTrue SC_GetManagedMailKey(mail) = key, "managed key recovered from user property"
AssertTrue InStr(1, mail.HTMLBody, "<!-- " & SC_MANAGED_KEY_MARKER & ":", vbTextCompare) > 0, "hidden HTML marker added"
AssertTrue InStr(1, mail.HTMLBody, "<p>Hello</p>", vbTextCompare) > 0, "visible HTML body preserved"

Dim markerOnly
Set markerOnly = New FakeMail
Set markerOnly.UserProperties = New FakeUserProperties
markerOnly.HTMLBody = "<html><body><p>Hello</p>" & SC_ManagedMarker(key) & "</body></html>"
AssertTrue SC_GetManagedMailKey(markerOnly) = key, "managed key recovered from HTML marker"

Dim diag
Set diag = SC_OutboxDiagnostic(mail, key, tempLogPath)
AssertTrue diag("managed_key") = key, "outbox diagnostic keeps managed key"
AssertTrue diag("pr_message_flags") = "[unavailable]", "MAPI property failure is non-fatal"

Dim spec
Set spec = SC_CreateMessageSpec(key, "regular", "recipient@example.invalid", "Subject", "Body")
Dim policy
Set policy = SC_CreateDeliveryPolicy("send")
Dim config
Set config = SC_CreateDictionary()
config("dry_run") = True
config("log_path") = tempLogPath
Dim result
Set result = SC_DispatchManagedEmail(spec, policy, config)
AssertTrue result("state") = SC_STATE_COMPOSED, "dry-run dispatch avoids Outlook submission"

Dim missingKeySpec
Set missingKeySpec = SC_CreateMessageSpec("", "regular", "recipient@example.invalid", "Subject", "Body")
Set result = SC_DispatchManagedEmail(missingKeySpec, policy, config)
AssertTrue result("state") = SC_STATE_FAILED_SAFE, "missing managed key fails safe"

If failures > 0 Then
  WScript.Echo "Smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

On Error Resume Next
If fso.FileExists(tempLogPath) Then fso.DeleteFile tempLogPath, True
On Error GoTo 0

WScript.Echo "Smoke test passed."
