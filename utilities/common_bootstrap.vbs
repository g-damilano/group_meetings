' Common bootstrap helpers shared across console scripts

Dim loadedModuleRegistry
Dim loadedModuleRegistryInitialized

Sub EnsureLoadedModuleRegistry()
  If IsEmpty(loadedModuleRegistryInitialized) Then loadedModuleRegistryInitialized = False
  If loadedModuleRegistryInitialized Then Exit Sub

  On Error Resume Next
  Set loadedModuleRegistry = CreateObject("Scripting.Dictionary")
  If Err.Number <> 0 Then
    Err.Clear
    Set loadedModuleRegistry = Nothing
  Else
    loadedModuleRegistry.CompareMode = 1
  End If
  On Error GoTo 0

  loadedModuleRegistryInitialized = True
End Sub

Function GetScriptFolder()
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."
  GetScriptFolder = folder
End Function

Sub LoadBootstrapModules(modulePaths)
  Dim baseFolder: baseFolder = GetScriptFolder()
  If IsEmpty(modulePaths) Then Exit Sub

  Dim idx, item, pathText

  If IsArray(modulePaths) Then
    For idx = LBound(modulePaths) To UBound(modulePaths)
      pathText = NormalizeModulePath(modulePaths(idx))
      If Len(pathText) > 0 Then Call LoadModule(baseFolder, pathText)
    Next
    Exit Sub
  End If

  If IsObject(modulePaths) Then
    For Each item In modulePaths
      pathText = NormalizeModulePath(item)
      If Len(pathText) > 0 Then Call LoadModule(baseFolder, pathText)
    Next
    Exit Sub
  End If

  pathText = NormalizeModulePath(modulePaths)
  If Len(pathText) > 0 Then Call LoadModule(baseFolder, pathText)
End Sub

Sub LoadModule(baseFolder, relativePath)
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim modulePath: modulePath = ResolveModulePath(fso, baseFolder, relativePath)
  If Len(modulePath) = 0 Then
    WScript.Echo "Warning: utility module not found: " & relativePath
    Exit Sub
  End If

  Call EnsureLoadedModuleRegistry()

  Dim canonicalPath: canonicalPath = modulePath
  If Not fso Is Nothing Then
    On Error Resume Next
    canonicalPath = fso.GetAbsolutePathName(modulePath)
    If Err.Number <> 0 Then
      Err.Clear
      canonicalPath = modulePath
    End If
    On Error GoTo 0
  End If

  Dim registryKey: registryKey = LCase(CStr(canonicalPath))
  If Not loadedModuleRegistry Is Nothing Then
    If loadedModuleRegistry.Exists(registryKey) Then Exit Sub
  End If

  Dim ts: Set ts = fso.OpenTextFile(modulePath, 1)
  Dim contents: contents = ts.ReadAll
  ts.Close

  contents = RemoveByteOrderMark(contents)

  On Error Resume Next
  ExecuteGlobal contents
  If Err.Number <> 0 Then
    WScript.Echo "Warning: failed to load module '" & relativePath & "': " & Err.Description
    Err.Clear
  ElseIf Not loadedModuleRegistry Is Nothing Then
    On Error Resume Next
    loadedModuleRegistry.Add registryKey, True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
  End If
  On Error GoTo 0
End Sub

Function NormalizeModulePath(value)
  NormalizeModulePath = ""
  If IsEmpty(value) Then Exit Function
  If IsNull(value) Then Exit Function

  Dim text
  On Error Resume Next
  text = CStr(value)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  NormalizeModulePath = Trim(text)
End Function

Function ResolveModulePath(fso, baseFolder, relativePath)
  Dim modulePath: modulePath = ""

  If Not fso Is Nothing Then
    Dim currentFolder: currentFolder = baseFolder
    Do While Len(currentFolder) > 0
      Dim candidate: candidate = fso.BuildPath(currentFolder, relativePath)
      If fso.FileExists(candidate) Then
        modulePath = candidate
        Exit Do
      End If

      Dim parentFolder: parentFolder = fso.GetParentFolderName(currentFolder)
      If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do
      currentFolder = parentFolder
    Loop
  End If

  ResolveModulePath = modulePath
End Function

Function RemoveByteOrderMark(text)
  RemoveByteOrderMark = text
  If Len(text) = 0 Then Exit Function

  If Len(text) >= 3 Then
    Dim leading: leading = Mid(text, 1, 3)
    If StrComp(leading, Chr(239) & Chr(187) & Chr(191), vbBinaryCompare) = 0 Then
      RemoveByteOrderMark = Mid(text, 4)
      Exit Function
    End If
  End If

  Dim firstChar: firstChar = Mid(text, 1, 1)
  If Len(firstChar) = 0 Then Exit Function

  On Error Resume Next
  Dim codePoint: codePoint = AscW(firstChar)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If codePoint = &HFEFF Or codePoint = &HFFFE Then
    RemoveByteOrderMark = Mid(text, 2)
  End If
End Function
