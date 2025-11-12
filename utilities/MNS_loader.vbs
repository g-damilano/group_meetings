' Loader bootstrap that ensures name resolution helpers are defined once

Dim __mnsLoaderAlreadyDefined: __mnsLoaderAlreadyDefined = False
Dim __mnsLoaderRef
Dim __mnsLoaderError
On Error Resume Next
Set __mnsLoaderRef = GetRef("LoadNameResolutionModules")
__mnsLoaderError = Err.Number
On Error GoTo 0
Set __mnsLoaderRef = Nothing
If __mnsLoaderError <> 0 Then Err.Clear
If __mnsLoaderError = 0 Then __mnsLoaderAlreadyDefined = True

If Not __mnsLoaderAlreadyDefined Then
  Dim __mnsLoaderLines
  __mnsLoaderLines = Array(_
  "' Loader for the MNS (mail name resolution) module set", _
  "'", _
  "' Provides LoadNameResolutionModules for console scripts so they can", _
  "' bring in the entire name-resolution helper suite by including a", _
  "' single file (`utilities/MNS_loader.vbs`).", _
  "", _
  "Sub LoadNameResolutionModules(baseFolder)", _
  "  Dim effectiveBase: effectiveBase = Trim(CStr(baseFolder))", _
  "  Dim fso: Set fso = CreateObject(""Scripting.FileSystemObject"")", _
  "", _
  "  If Len(effectiveBase) = 0 Then", _
  "    On Error Resume Next", _
  "    effectiveBase = GetScriptFolder()", _
  "    If Err.Number <> 0 Then", _
  "      Err.Clear", _
  "      effectiveBase = baseFolder", _
  "    End If", _
  "    On Error GoTo 0", _
  "  End If", _
  "", _
  "  Dim moduleRoot: moduleRoot = ResolveNameResolutionRoot(fso, effectiveBase)", _
  "  Dim resolvedFolder: resolvedFolder = ResolveFolderPath(fso, effectiveBase, moduleRoot)", _
  "  Dim modules", _
  "", _
  "  If Len(resolvedFolder) > 0 Then", _
  "    modules = CollectNameResolutionModules(fso, resolvedFolder)", _
  "    If IsArray(modules) Then", _
  "      Dim upper", _
  "      On Error Resume Next", _
  "      upper = UBound(modules)", _
  "      If Err.Number <> 0 Then", _
  "        Err.Clear", _
  "        upper = -1", _
  "      End If", _
  "      On Error GoTo 0", _
  "", _
  "      Dim i", _
  "      If upper >= 0 Then", _
  "        For i = 0 To upper", _
  "          Dim relativePath: relativePath = moduleRoot & ""\\"" & modules(i)", _
  "          Call LoadModule(effectiveBase, relativePath)", _
  "        Next", _
  "        Exit Sub", _
  "      End If", _
  "    End If", _
  "  End If", _
  "", _
  "  modules = Array( _", _
  "    ""MNS_collections.vbs"", _", _
  "    ""MNS_contacts.vbs"", _", _
  "    ""MNS_text.vbs"", _", _
  "    ""MNS_directory_recording.vbs"", _", _
  "    ""MNS_directory_lists.vbs"", _", _
  "    ""MNS_directory_population.vbs"", _", _
  "    ""MNS_directory_matching.vbs"", _", _
  "    ""MNS_core.vbs"", _", _
  "    ""MNS_console.vbs"", _", _
  "    ""MNS_prompt.vbs"", _", _
  "    ""MNS_directory.vbs"")", _
  "", _
  "  Dim i", _
  "  For i = LBound(modules) To UBound(modules)", _
  "    Dim relativePath: relativePath = modules(i)", _
  "    If Len(ResolveModulePath(fso, effectiveBase, relativePath)) = 0 Then", _
  "      relativePath = ""name_resolution\\"" & modules(i)", _
  "    End If", _
  "    If Len(ResolveModulePath(fso, effectiveBase, relativePath)) = 0 Then", _
  "      relativePath = ""utilities\\name_resolution\\"" & modules(i)", _
  "    End If", _
  "    Call LoadModule(effectiveBase, relativePath)", _
  "  Next", _
  "End Sub", _
  "", _
  "Function ResolveNameResolutionRoot(fso, baseFolder)", _
  "  ResolveNameResolutionRoot = """"", _
  "  Dim candidates: candidates = Array(""name_resolution"", ""utilities\\name_resolution"")", _
  "  Dim i", _
  "  For i = LBound(candidates) To UBound(candidates)", _
  "    Dim candidate: candidate = candidates(i)", _
  "    If Len(ResolveFolderPath(fso, baseFolder, candidate)) > 0 Then", _
  "      ResolveNameResolutionRoot = candidate", _
  "      Exit Function", _
  "    End If", _
  "  Next", _
  "End Function", _
  "", _
  "Function ResolveFolderPath(fso, baseFolder, relativeFolder)", _
  "  Dim folderPath: folderPath = """"", _
  "  If fso Is Nothing Then", _
  "    ResolveFolderPath = folderPath", _
  "    Exit Function", _
  "  End If", _
  "", _
  "  Dim currentFolder: currentFolder = baseFolder", _
  "  If Len(currentFolder) = 0 Then", _
  "    ResolveFolderPath = folderPath", _
  "    Exit Function", _
  "  End If", _
  "", _
  "  If Len(relativeFolder) = 0 Then", _
  "    ResolveFolderPath = folderPath", _
  "    Exit Function", _
  "  End If", _
  "", _
  "  Do While Len(currentFolder) > 0", _
  "    Dim candidate: candidate = fso.BuildPath(currentFolder, relativeFolder)", _
  "    If fso.FolderExists(candidate) Then", _
  "      folderPath = candidate", _
  "      Exit Do", _
  "    End If", _
  "", _
  "    Dim parentFolder: parentFolder = fso.GetParentFolderName(currentFolder)", _
  "    If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do", _
  "    currentFolder = parentFolder", _
  "  Loop", _
  "", _
  "  ResolveFolderPath = folderPath", _
  "End Function", _
  "", _
  "Function CollectNameResolutionModules(fso, folderPath)", _
  "  Dim emptyArray: emptyArray = Array()", _
  "  CollectNameResolutionModules = emptyArray", _
  "", _
  "  If fso Is Nothing Then Exit Function", _
  "  If Len(folderPath) = 0 Then Exit Function", _
  "", _
  "  Dim folder", _
  "  On Error Resume Next", _
  "  Set folder = fso.GetFolder(folderPath)", _
  "  If Err.Number <> 0 Then", _
  "    Err.Clear", _
  "    On Error GoTo 0", _
  "    Exit Function", _
  "  End If", _
  "  On Error GoTo 0", _
  "", _
  "  Dim entries: Set entries = CreateObject(""Scripting.Dictionary"")", _
  "  On Error Resume Next", _
  "  entries.CompareMode = 1", _
  "  On Error GoTo 0", _
  "", _
  "  Dim file", _
  "  For Each file In folder.Files", _
  "    If StrComp(fso.GetExtensionName(file.Name), ""vbs"", vbTextCompare) = 0 Then", _
  "      Dim key: key = CStr(entries.Count)", _
  "      If Not entries.Exists(key) Then entries.Add key, file.Name", _
  "    End If", _
  "  Next", _
  "", _
  "  If entries.Count = 0 Then Exit Function", _
  "", _
  "  Dim result()", _
  "  ReDim result(entries.Count - 1)", _
  "", _
  "  Dim idx, dictKey", _
  "  idx = 0", _
  "  For Each dictKey In entries.Keys", _
  "    result(idx) = entries(dictKey)", _
  "    idx = idx + 1", _
  "  Next", _
  "", _
  "  Call SortNameResolutionModules(result)", _
  "  CollectNameResolutionModules = result", _
  "End Function", _
  "", _
  "Sub SortNameResolutionModules(ByRef modules)", _
  "  If IsEmpty(modules) Then Exit Sub", _
  "  If Not IsArray(modules) Then Exit Sub", _
  "", _
  "  Dim upper", _
  "  On Error Resume Next", _
  "  upper = UBound(modules)", _
  "  If Err.Number <> 0 Then", _
  "    Err.Clear", _
  "    On Error GoTo 0", _
  "    Exit Sub", _
  "  End If", _
  "  On Error GoTo 0", _
  "", _
  "  Dim i, j", _
  "  For i = 0 To upper - 1", _
  "    For j = i + 1 To upper", _
  "      If ShouldSwapModules(modules(i), modules(j)) Then", _
  "        Dim temp: temp = modules(i)", _
  "        modules(i) = modules(j)", _
  "        modules(j) = temp", _
  "      End If", _
  "    Next", _
  "  Next", _
  "End Sub", _
  "", _
  "Function ShouldSwapModules(firstName, secondName)", _
  "  Dim firstPriority: firstPriority = ModulePriority(firstName)", _
  "  Dim secondPriority: secondPriority = ModulePriority(secondName)", _
  "", _
  "  If firstPriority = secondPriority Then", _
  "    ShouldSwapModules = (StrComp(LCase(CStr(firstName)), LCase(CStr(secondName)), vbTextCompare) > 0)", _
  "  Else", _
  "    ShouldSwapModules = (firstPriority > secondPriority)", _
  "  End If", _
  "End Function", _
  "", _
  "Function ModulePriority(fileName)", _
  "  Dim lowered: lowered = LCase(CStr(fileName))", _
  "  Select Case lowered", _
  "    Case ""mns_collections.vbs""", _
  "      ModulePriority = 10", _
  "    Case ""mns_contacts.vbs""", _
  "      ModulePriority = 20", _
  "    Case ""mns_text.vbs""", _
  "      ModulePriority = 30", _
  "    Case ""mns_directory_recording.vbs""", _
  "      ModulePriority = 40", _
  "    Case ""mns_directory_lists.vbs""", _
  "      ModulePriority = 50", _
  "    Case ""mns_directory_population.vbs""", _
  "      ModulePriority = 60", _
  "    Case ""mns_directory_matching.vbs""", _
  "      ModulePriority = 70", _
  "    Case ""mns_core.vbs""", _
  "      ModulePriority = 80", _
  "    Case ""mns_console.vbs""", _
  "      ModulePriority = 90", _
  "    Case ""mns_prompt.vbs""", _
  "      ModulePriority = 100", _
  "    Case ""mns_directory.vbs""", _
  "      ModulePriority = 110", _
  "    Case Else", _
  "      ModulePriority = 1000", _
  "  End Select", _
  "End Function"
  )
  ExecuteGlobal Join(__mnsLoaderLines, vbCrLf)
End If
