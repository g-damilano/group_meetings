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

Dim directory
Set directory = SC_CreateDictionary()
directory("team lead") = "team.lead@example.invalid"
directory("Example Visitor") = "Example Visitor <external.visitor@example.invalid>"
directory("exact external") = "exact.external@example.invalid"
directory("Exact External") = "Wrong Exact <wrong.external@example.invalid>"
directory("Ambiguous Person") = "ambiguous.one@example.invalid"
directory("ambiguous-person") = "ambiguous.two@example.invalid"

Dim direct
Set direct = SC_ResolveRecipient("Alex Person <alex.person@example.invalid>", directory)
AssertTrue direct("resolution_status") = "direct", "direct SMTP recipient resolved"
AssertTrue direct("smtp_address") = "alex.person@example.invalid", "direct SMTP address normalized"
AssertTrue direct("display_name") = "Alex Person", "display name extracted"
AssertTrue direct("first_name") = "Alex", "direct display first name parsed"
AssertTrue direct("surname") = "Person", "direct display surname parsed"
AssertTrue direct("name_parse_source") = "direct-display", "direct display parse source recorded"

Dim quotedDirect
Set quotedDirect = SC_ResolveRecipient("""External Visitor"" <external.visitor@example.invalid>", directory)
AssertTrue quotedDirect("resolution_status") = "direct", "quoted display SMTP recipient resolved"
AssertTrue quotedDirect("display_name") = "External Visitor", "quoted display name cleaned"
AssertTrue quotedDirect("first_name") = "External", "quoted display first name parsed"
AssertTrue quotedDirect("surname") = "Visitor", "quoted display surname parsed"

Dim familyGivenDirect
Set familyGivenDirect = SC_ResolveRecipient("Example-Family, Casey <casey.family@example.invalid>", directory)
AssertTrue familyGivenDirect("resolution_status") = "direct", "family-given display SMTP recipient resolved"
AssertTrue familyGivenDirect("first_name") = "Casey", "family-given display first name parsed"
AssertTrue familyGivenDirect("surname") = "Example-Family", "family-given display surname parsed"
AssertTrue familyGivenDirect("full_name") = "Casey Example-Family", "family-given display full name normalised"
AssertTrue familyGivenDirect("name_parse_order") = "family-given", "family-given parse order recorded"

Dim resolved
Set resolved = SC_ResolveRecipient("team lead", directory)
AssertTrue resolved("resolution_status") = "resolved", "directory recipient resolved"
AssertTrue resolved("smtp_address") = "team.lead@example.invalid", "directory SMTP address normalized"
AssertTrue resolved("name_parse_status") = "not_attempted", "exact directory lookup does not parse alias as person name"

Dim displayFallback
Set displayFallback = SC_ResolveRecipient("example visitor", directory)
AssertTrue displayFallback("resolution_status") = "resolved", "display fallback recipient resolved"
AssertTrue displayFallback("smtp_address") = "external.visitor@example.invalid", "display fallback SMTP address normalized"
AssertTrue displayFallback("display_name") = "Example Visitor", "directory display name preserved"
AssertTrue displayFallback("first_name") = "Example", "directory display first name parsed"
AssertTrue displayFallback("surname") = "Visitor", "directory display surname parsed"

Dim exactWins
Set exactWins = SC_ResolveRecipient("exact external", directory)
AssertTrue exactWins("smtp_address") = "exact.external@example.invalid", "exact directory key wins before display fallback"

Dim ambiguous
Set ambiguous = SC_ResolveRecipient("ambiguous person", directory)
AssertTrue ambiguous("resolution_status") = "unresolved", "ambiguous display fallback remains unresolved"
AssertTrue ambiguous("fallback_used") = True, "ambiguous display fallback is marked fallback"

Dim unresolved
Set unresolved = SC_ResolveRecipient("unknown person", directory)
AssertTrue unresolved("resolution_status") = "unresolved", "unknown recipient remains unresolved"
AssertTrue unresolved("fallback_used") = True, "unknown recipient marked as fallback"

Dim recipients
Set recipients = SC_ResolveRecipientList("Alex Person <alex.person@example.invalid>; team lead", directory)
AssertTrue recipients.Count = 2, "recipient list split"
AssertTrue SC_JoinResolvedRecipientAddresses(recipients) = "alex.person@example.invalid; team.lead@example.invalid", "recipient list joined"

Dim commaRecipients
Set commaRecipients = SC_ResolveRecipientList("Example-Family, Casey <casey.family@example.invalid>, team lead", directory)
AssertTrue commaRecipients.Count = 2, "recipient list preserves family-given display comma"
AssertTrue commaRecipients("0")("surname") = "Example-Family", "recipient list first item keeps family-given surname"
AssertTrue SC_JoinResolvedRecipientAddresses(commaRecipients) = "casey.family@example.invalid; team.lead@example.invalid", "recipient list comma split after completed recipient"

Dim context
Set context = SC_CreateDictionary()
context("FIRST_NAME") = "Alex"
context("PROJECT") = "Example Project"

Dim templateText
templateText = "<!-- Subject: Welcome -->" & vbCrLf & _
               "<p>Hello [FIRST_NAME], welcome to {{PROJECT}}.</p>"

Dim placeholders
Set placeholders = SC_ExtractPlaceholders(templateText)
AssertTrue placeholders.Exists("FIRST_NAME"), "square placeholder extracted"
AssertTrue placeholders.Exists("PROJECT"), "double-brace placeholder extracted"

Dim renderedText
renderedText = SC_ApplyContextToText(templateText, context)
AssertTrue InStr(1, renderedText, "Hello Alex", vbTextCompare) > 0, "square placeholder rendered"
AssertTrue InStr(1, renderedText, "Example Project", vbTextCompare) > 0, "double-brace placeholder rendered"

Dim rendered
Set rendered = SC_RenderTemplateString(templateText, context)
AssertTrue rendered("metadata").Exists("SUBJECT"), "metadata parsed"
AssertTrue rendered("metadata")("SUBJECT") = "Welcome", "metadata value preserved"
AssertTrue InStr(1, rendered("body"), "Hello Alex", vbTextCompare) > 0, "rendered body returned"

If failures > 0 Then
  WScript.Echo "Template/recipient smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Template/recipient smoke test passed."
