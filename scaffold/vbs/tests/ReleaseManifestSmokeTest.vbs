Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

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

Function FileText(relativePath)
  Dim path
  path = fso.BuildPath(scaffoldRoot, relativePath)
  Dim stream
  Set stream = fso.OpenTextFile(path, 1, False)
  FileText = stream.ReadAll
  stream.Close
End Function

Function FileContains(relativePath, token)
  FileContains = (InStr(1, FileText(relativePath), CStr(token), vbTextCompare) > 0)
End Function

AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\release_manifest.py")), "Release manifest generator exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\release_export_plan.py")), "Release export planner exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\tests\test_release_manifest.py")), "Release manifest Python tests exist"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\tests\test_release_export_plan.py")), "Release export plan Python tests exist"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_MANIFEST.md")), "Release manifest documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_EXPORT_PLAN.md")), "Release export plan documentation exists"

AssertTrue FileContains("python\release_manifest.py", "sha256_file"), "Manifest generator hashes files"
AssertTrue FileContains("python\release_manifest.py", "relative_to(root).as_posix()"), "Manifest generator uses relative POSIX paths"
AssertTrue FileContains("python\release_manifest.py", "dont_write_bytecode"), "Manifest generator avoids bytecode artifacts"
AssertTrue FileContains("python\release_manifest.py", "excluded_release_directories"), "Manifest generator excludes private mirror directories"
AssertTrue Not FileContains("python\release_manifest.py", "write_text("), "Manifest generator does not write output files"
AssertTrue FileContains("python\release_audit.py", "release_manifest"), "Release audit includes manifest validation"
AssertTrue FileContains("docs\PUBLIC_RELEASE_CHECKLIST.md", "release_manifest.py"), "Public checklist includes manifest command"
AssertTrue FileContains("docs\PUBLIC_RELEASE_CHECKLIST.md", "release_export_plan.py"), "Public checklist includes export plan command"
AssertTrue FileContains("python\release_export_plan.py", "legacy_paths_in_files"), "Export plan guards legacy paths"

If failures > 0 Then
  WScript.Echo "Release manifest smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Release manifest smoke test passed."
