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

Sub Fail(message)
  WScript.Echo "FAILED: " & message
  failures = failures + 1
End Sub

Function RelativePath(path)
  RelativePath = Replace(Mid(CStr(path), Len(scaffoldRoot) + 2), "\", "/")
End Function

Function IsReleaseExcludedPath(path)
  Dim rel
  rel = LCase(RelativePath(path))
  IsReleaseExcludedPath = (rel = "legacy" Or Left(rel, 7) = "legacy/")
End Function

Function FileText(path)
  Dim stream
  Set stream = fso.OpenTextFile(path, 1, False)
  FileText = stream.ReadAll
  stream.Close
End Function

Function HasTextExtension(path)
  Dim ext
  ext = LCase(fso.GetExtensionName(path))
  HasTextExtension = (ext = "md" Or ext = "txt" Or ext = "csv" Or ext = "json" Or ext = "vbs" Or ext = "html" Or ext = "py")
End Function

Function HasGeneratedExtension(path)
  Dim ext
  ext = "." & LCase(fso.GetExtensionName(path))
  HasGeneratedExtension = (ext = ".log" Or ext = ".pyc" Or ext = ".pyo" Or ext = ".zip" Or ext = ".msg")
End Function

Function RegexMatches(pattern, text)
  Dim re
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = pattern
  re.IgnoreCase = True
  re.Global = True
  Set RegexMatches = re.Execute(text)
End Function

Sub ScanGeneratedArtifacts(folder)
  Dim subFolder
  For Each subFolder In folder.SubFolders
    If IsReleaseExcludedPath(subFolder.Path) Then
      ' Local hard-copy mirror; never part of the public release surface.
    Else
    If LCase(subFolder.Name) = "__pycache__" Then
      Fail "Generated directory is present: " & RelativePath(subFolder.Path)
    End If
    ScanGeneratedArtifacts subFolder
    End If
  Next

  Dim file
  For Each file In folder.Files
    If HasGeneratedExtension(file.Path) Then
      Fail "Generated or release-blocked artifact is present: " & RelativePath(file.Path)
    End If
  Next
End Sub

Sub ScanTextFile(path)
  If Not HasTextExtension(path) Then Exit Sub

  Dim text
  text = FileText(path)

  Dim emailMatches
  Set emailMatches = RegexMatches("\b[A-Z0-9._%+-]+@([A-Z0-9.-]+\.[A-Z]{2,63})\b", text)
  Dim match
  For Each match In emailMatches
    If LCase(Split(match.Value, "@")(1)) <> "example.invalid" Then
      Fail RelativePath(path) & " contains non-example email address"
      Exit For
    End If
  Next

  If RegexMatches("OneDrive\s+-", text).Count > 0 Then
    Fail RelativePath(path) & " contains private OneDrive-style path marker"
  End If
  If RegexMatches("Imperial\s+College", text).Count > 0 Then
    Fail RelativePath(path) & " contains institution-specific marker"
  End If
  If RegexMatches("https?://[^\s" & Chr(34) & "']*sharepoint[^\s" & Chr(34) & "']*", text).Count > 0 Then
    Fail RelativePath(path) & " contains live SharePoint-style URL"
  End If
End Sub

Sub ScanTextFiles(folder)
  Dim subFolder
  For Each subFolder In folder.SubFolders
    If IsReleaseExcludedPath(subFolder.Path) Then
      ' Local hard-copy mirror; never part of the public release surface.
    Else
    ScanTextFiles subFolder
    End If
  Next

  Dim file
  For Each file In folder.Files
    ScanTextFile file.Path
  Next
End Sub

ScanGeneratedArtifacts fso.GetFolder(scaffoldRoot)
ScanTextFiles fso.GetFolder(scaffoldRoot)

AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, ".gitignore")), "Scaffold has local generated-artifact ignore guard"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\PUBLIC_RELEASE_CHECKLIST.md")), "Public release checklist exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\PUBLICATION_BOUNDARY.md")), "Publication boundary documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\LEGACY_MIRROR.md")), "Legacy mirror quarantine documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_AUDIT.md")), "Release audit documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_CANDIDATE_PLAN.md")), "Release candidate plan documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_EXPORT_PLAN.md")), "Release export plan documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_GATE.md")), "Release gate documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\REPO_BOUNDARY_AUDIT.md")), "Repository boundary audit documentation exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "docs\RELEASE_NOTES_TEMPLATE.md")), "Release notes template exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\release_audit.py")), "Python release audit exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\release_candidate_plan.py")), "Python release candidate planner exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\release_export_plan.py")), "Python release export planner exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\release_gate.py")), "Python release gate exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\legacy_mirror_audit.py")), "Python legacy mirror audit exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\legacy_sanitisation_plan.py")), "Python legacy sanitisation planner exists"
AssertTrue fso.FileExists(fso.BuildPath(scaffoldRoot, "python\repo_boundary_audit.py")), "Python repository boundary audit exists"
AssertTrue FileText(fso.BuildPath(scaffoldRoot, "python\release_gate.py")) <> "", "Release gate source is readable"
AssertTrue InStr(1, FileText(fso.BuildPath(scaffoldRoot, "python\release_gate.py")), "fail-if-blocked", vbTextCompare) > 0, "Release gate has explicit blocked exit option"
AssertTrue InStr(1, FileText(fso.BuildPath(scaffoldRoot, "python\release_export_plan.py")), "archive_created", vbTextCompare) > 0, "Export plan declares archive-free behaviour"

If failures > 0 Then
  WScript.Echo "Public release gate smoke test failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Public release gate smoke test passed."
