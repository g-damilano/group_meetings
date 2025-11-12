' Provides helpers for loading the shared common_bootstrap module.
Dim bootstrapCommonLoaded: bootstrapCommonLoaded = False

Function BootstrapLocateNearest(fso, scriptPath, relativeTargets)
  BootstrapLocateNearest = ""
  If fso Is Nothing Then Exit Function

  Dim startFolder: startFolder = fso.GetParentFolderName(scriptPath)
  If Len(startFolder) = 0 Then startFolder = "."

  If Not IsArray(relativeTargets) Then Exit Function

  Dim currentFolder: currentFolder = startFolder
  Do While Len(currentFolder) > 0
    Dim idx
    For idx = LBound(relativeTargets) To UBound(relativeTargets)
      Dim candidate: candidate = fso.BuildPath(currentFolder, relativeTargets(idx))
      If fso.FileExists(candidate) Then
        BootstrapLocateNearest = candidate
        Exit Function
      End If
    Next

    Dim parentFolder: parentFolder = fso.GetParentFolderName(currentFolder)
    If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do
    currentFolder = parentFolder
  Loop
End Function

Function BootstrapReadAllText(fso, filePath)
  Dim stream: Set stream = fso.OpenTextFile(filePath, 1)
  BootstrapReadAllText = stream.ReadAll
  stream.Close
End Function

Sub EnsureCommonBootstrapLoaded(fso)
  If bootstrapCommonLoaded Then Exit Sub

  Dim localFso
  If IsObject(fso) Then
    Set localFso = fso
  Else
    Set localFso = CreateObject("Scripting.FileSystemObject")
  End If

  Dim modulePath
  On Error Resume Next
  modulePath = BootstrapLocateNearest(localFso, WScript.ScriptFullName, _
    Array("utilities\common_bootstrap.vbs", "common_bootstrap.vbs"))
  If Err.Number <> 0 Then
    Dim locateError: locateError = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "Fatal: unable to resolve bootstrap locator helpers: " & locateError
    WScript.Quit 1
  End If
  On Error GoTo 0
  If Len(modulePath) = 0 Then
    WScript.Echo "Fatal: unable to locate utilities\\common_bootstrap.vbs"
    WScript.Quit 1
  End If

  Dim moduleContents
  On Error Resume Next
  moduleContents = BootstrapReadAllText(localFso, modulePath)
  If Err.Number <> 0 Then
    Dim readError: readError = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "Fatal: unable to read common bootstrap module: " & readError
    WScript.Quit 1
  End If
  On Error GoTo 0

  ExecuteGlobal moduleContents
  bootstrapCommonLoaded = True
End Sub
