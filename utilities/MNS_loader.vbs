' Loader for the MNS (mail name resolution) module set
'
' Provides LoadNameResolutionModules for console scripts so they can
' bring in the entire name-resolution helper suite by including a
' single file (`utilities/MNS_loader.vbs`).

Sub LoadNameResolutionModules(baseFolder)
  Dim effectiveBase: effectiveBase = Trim(CStr(baseFolder))
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

  If Len(effectiveBase) = 0 Then
    On Error Resume Next
    effectiveBase = GetScriptFolder()
    If Err.Number <> 0 Then
      Err.Clear
      effectiveBase = baseFolder
    End If
    On Error GoTo 0
  End If

  Dim moduleRoot: moduleRoot = ResolveNameResolutionRoot(fso, effectiveBase)
  Dim resolvedFolder: resolvedFolder = ResolveFolderPath(fso, effectiveBase, moduleRoot)
  Dim modules

  If Len(resolvedFolder) > 0 Then
    modules = CollectNameResolutionModules(fso, resolvedFolder)
    If IsArray(modules) Then
      Dim upper
      On Error Resume Next
      upper = UBound(modules)
      If Err.Number <> 0 Then
        Err.Clear
        upper = -1
      End If
      On Error GoTo 0

      Dim i
      If upper >= 0 Then
        For i = 0 To upper
          Dim relativePath: relativePath = moduleRoot & "\\" & modules(i)
          Call LoadModule(effectiveBase, relativePath)
        Next
        Exit Sub
      End If
    End If
  End If

  modules = Array( _
    "MNS_collections.vbs", _
    "MNS_contacts.vbs", _
    "MNS_text.vbs", _
    "MNS_directory_recording.vbs", _
    "MNS_directory_lists.vbs", _
    "MNS_directory_population.vbs", _
    "MNS_directory_matching.vbs", _
    "MNS_core.vbs", _
    "MNS_console.vbs", _
    "MNS_prompt.vbs", _
    "MNS_directory.vbs")

  Dim i
  For i = LBound(modules) To UBound(modules)
    Dim relativePath: relativePath = modules(i)
    If Len(ResolveModulePath(fso, effectiveBase, relativePath)) = 0 Then
      relativePath = "name_resolution\\" & modules(i)
    End If
    If Len(ResolveModulePath(fso, effectiveBase, relativePath)) = 0 Then
      relativePath = "utilities\\name_resolution\\" & modules(i)
    End If
    Call LoadModule(effectiveBase, relativePath)
  Next
End Sub

Function ResolveNameResolutionRoot(fso, baseFolder)
  ResolveNameResolutionRoot = ""
  Dim candidates: candidates = Array("name_resolution", "utilities\\name_resolution")
  Dim i
  For i = LBound(candidates) To UBound(candidates)
    Dim candidate: candidate = candidates(i)
    If Len(ResolveFolderPath(fso, baseFolder, candidate)) > 0 Then
      ResolveNameResolutionRoot = candidate
      Exit Function
    End If
  Next
End Function

Function ResolveFolderPath(fso, baseFolder, relativeFolder)
  Dim folderPath: folderPath = ""
  If fso Is Nothing Then
    ResolveFolderPath = folderPath
    Exit Function
  End If

  Dim currentFolder: currentFolder = baseFolder
  If Len(currentFolder) = 0 Then
    ResolveFolderPath = folderPath
    Exit Function
  End If

  If Len(relativeFolder) = 0 Then
    ResolveFolderPath = folderPath
    Exit Function
  End If

  Do While Len(currentFolder) > 0
    Dim candidate: candidate = fso.BuildPath(currentFolder, relativeFolder)
    If fso.FolderExists(candidate) Then
      folderPath = candidate
      Exit Do
    End If

    Dim parentFolder: parentFolder = fso.GetParentFolderName(currentFolder)
    If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do
    currentFolder = parentFolder
  Loop

  ResolveFolderPath = folderPath
End Function

Function CollectNameResolutionModules(fso, folderPath)
  Dim emptyArray: emptyArray = Array()
  CollectNameResolutionModules = emptyArray

  If fso Is Nothing Then Exit Function
  If Len(folderPath) = 0 Then Exit Function

  Dim folder
  On Error Resume Next
  Set folder = fso.GetFolder(folderPath)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Dim entries: Set entries = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  entries.CompareMode = 1
  On Error GoTo 0

  Dim file
  For Each file In folder.Files
    If StrComp(fso.GetExtensionName(file.Name), "vbs", vbTextCompare) = 0 Then
      Dim key: key = CStr(entries.Count)
      If Not entries.Exists(key) Then entries.Add key, file.Name
    End If
  Next

  If entries.Count = 0 Then Exit Function

  Dim result()
  ReDim result(entries.Count - 1)

  Dim idx, dictKey
  idx = 0
  For Each dictKey In entries.Keys
    result(idx) = entries(dictKey)
    idx = idx + 1
  Next

  Call SortNameResolutionModules(result)
  CollectNameResolutionModules = result
End Function

Sub SortNameResolutionModules(ByRef modules)
  If IsEmpty(modules) Then Exit Sub
  If Not IsArray(modules) Then Exit Sub

  Dim upper
  On Error Resume Next
  upper = UBound(modules)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0

  Dim i, j
  For i = 0 To upper - 1
    For j = i + 1 To upper
      If ShouldSwapModules(modules(i), modules(j)) Then
        Dim temp: temp = modules(i)
        modules(i) = modules(j)
        modules(j) = temp
      End If
    Next
  Next
End Sub

Function ShouldSwapModules(firstName, secondName)
  Dim firstPriority: firstPriority = ModulePriority(firstName)
  Dim secondPriority: secondPriority = ModulePriority(secondName)

  If firstPriority = secondPriority Then
    ShouldSwapModules = (StrComp(LCase(CStr(firstName)), LCase(CStr(secondName)), vbTextCompare) > 0)
  Else
    ShouldSwapModules = (firstPriority > secondPriority)
  End If
End Function

Function ModulePriority(fileName)
  Dim lowered: lowered = LCase(CStr(fileName))
  Select Case lowered
    Case "mns_collections.vbs"
      ModulePriority = 10
    Case "mns_contacts.vbs"
      ModulePriority = 20
    Case "mns_text.vbs"
      ModulePriority = 30
    Case "mns_directory_recording.vbs"
      ModulePriority = 40
    Case "mns_directory_lists.vbs"
      ModulePriority = 50
    Case "mns_directory_population.vbs"
      ModulePriority = 60
    Case "mns_directory_matching.vbs"
      ModulePriority = 70
    Case "mns_core.vbs"
      ModulePriority = 80
    Case "mns_console.vbs"
      ModulePriority = 90
    Case "mns_prompt.vbs"
      ModulePriority = 100
    Case "mns_directory.vbs"
      ModulePriority = 110
    Case Else
      ModulePriority = 1000
  End Select
End Function
