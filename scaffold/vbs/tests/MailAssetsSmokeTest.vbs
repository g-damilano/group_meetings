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

Dim fixtureRoot
fixtureRoot = fso.BuildPath(scaffoldRoot, "vbs\tests\fixtures")
Dim attachmentPath
attachmentPath = fso.BuildPath(fixtureRoot, "attachment.txt")
Dim templatePath
templatePath = fso.BuildPath(fixtureRoot, "sample_template.html")

Dim resolvedPath
resolvedPath = SC_ResolveExistingFilePath("attachment.txt", fixtureRoot)
AssertTrue StrComp(resolvedPath, fso.GetAbsolutePathName(attachmentPath), vbTextCompare) = 0, "relative attachment path resolved"

Dim attachmentList
Set attachmentList = SC_ResolveAttachmentList("attachment.txt", fixtureRoot, "")
AssertTrue attachmentList.Count = 1, "single attachment resolved"
AssertTrue StrComp(attachmentList("0"), fso.GetAbsolutePathName(attachmentPath), vbTextCompare) = 0, "attachment absolute path returned"

Dim dictAttachments
Set dictAttachments = SC_CreateDictionary()
dictAttachments("first") = "attachment.txt"
dictAttachments("duplicate") = attachmentPath
Set attachmentList = SC_ResolveAttachmentList(dictAttachments, fixtureRoot, "")
AssertTrue attachmentList.Count = 1, "duplicate attachment paths collapsed"

AssertTrue SC_SanitizeContentId("logo one.png") = "logo_one.png", "content id sanitized"

Dim htmlBody
htmlBody = SC_ReadTextFile(templatePath)
Dim inlineInfo
Set inlineInfo = SC_PrepareInlineImages(htmlBody, templatePath, "", "")
AssertTrue inlineInfo("inline_count") = 1, "inline image discovered"
AssertTrue InStr(1, inlineInfo("html_body"), "cid:logo_one.png", vbTextCompare) > 0, "inline image reference rewritten"
AssertTrue InStr(1, inlineInfo("html_body"), "sample_template.files/logo%20one.png", vbTextCompare) = 0, "encoded file reference removed"

Dim explicitRewrite
explicitRewrite = SC_RewriteInlineImageReferences("<img src=""logo one.png"">", "sample_template", "logo one.png")
AssertTrue InStr(1, explicitRewrite, "cid:logo_one.png", vbTextCompare) > 0, "direct file reference rewritten"

If failures > 0 Then
  WScript.Echo "Mail asset smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Mail asset smoke test passed."

