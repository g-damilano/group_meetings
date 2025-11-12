' Central entry point to load the bootstrap loader and shared utilities.

Dim bootstrapEntryLoaded
If IsEmpty(bootstrapEntryLoaded) Then bootstrapEntryLoaded = False

Sub BootstrapEnsureLoaded()
  If bootstrapEntryLoaded Then Exit Sub

  Dim entryFso: Set entryFso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = entryFso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."

  Dim loaderPath: loaderPath = ""
  Dim currentFolder: currentFolder = scriptFolder
  Do While Len(currentFolder) > 0
    Dim candidate

    candidate = entryFso.BuildPath(currentFolder, "utilities\\bootstrap_loader.vbs")
    If entryFso.FileExists(candidate) Then
      loaderPath = candidate
      Exit Do
    End If

    candidate = entryFso.BuildPath(currentFolder, "bootstrap_loader.vbs")
    If entryFso.FileExists(candidate) Then
      loaderPath = candidate
      Exit Do
    End If

    Dim parentFolder: parentFolder = entryFso.GetParentFolderName(currentFolder)
    If StrComp(parentFolder, currentFolder, vbTextCompare) = 0 Then Exit Do
    currentFolder = parentFolder
  Loop

  If Len(loaderPath) = 0 Then
    WScript.Echo "Fatal: unable to locate utilities\\bootstrap_loader.vbs"
    WScript.Quit 1
  End If

  Dim loaderContents: loaderContents = BootstrapReadAllText(entryFso, loaderPath)
  If Len(loaderContents) = 0 Then
    WScript.Echo "Fatal: unable to read bootstrap loader: " & loaderPath
    WScript.Quit 1
  End If

  ExecuteGlobal loaderContents
  Call EnsureCommonBootstrapLoaded(entryFso)
  Set entryFso = Nothing

  bootstrapEntryLoaded = True
End Sub

Call BootstrapEnsureLoaded()

Function BootstrapReadAllText(fso, filePath)
  Dim text: text = ""

  If Not fso Is Nothing Then
    Dim stream
    On Error Resume Next
    Set stream = fso.OpenTextFile(filePath, 1, False, -2)
    If Err.Number <> 0 Then
      Err.Clear
      Set stream = Nothing
    End If
    On Error GoTo 0

    If Not stream Is Nothing Then
      text = stream.ReadAll
      stream.Close
      Set stream = Nothing
      BootstrapReadAllText = text
      Exit Function
    End If

    On Error Resume Next
    Set stream = fso.OpenTextFile(filePath, 1)
    If Err.Number <> 0 Then
      Err.Clear
      Set stream = Nothing
    End If
    On Error GoTo 0

    If Not stream Is Nothing Then
      text = stream.ReadAll
      stream.Close
      Set stream = Nothing
      BootstrapReadAllText = text
      Exit Function
    End If
  End If

  On Error Resume Next
  Dim ado: Set ado = CreateObject("ADODB.Stream")
  If Err.Number <> 0 Then
    Err.Clear
    Set ado = Nothing
  End If
  On Error GoTo 0

  If ado Is Nothing Then
    BootstrapReadAllText = text
    Exit Function
  End If

  On Error Resume Next
  ado.Type = 2 ' adTypeText
  ado.Mode = 3 ' adModeReadWrite
  ado.Charset = "utf-8"
  ado.Open
  ado.LoadFromFile filePath
  If Err.Number = 0 Then text = ado.ReadText
  ado.Close
  Set ado = Nothing
  On Error GoTo 0

  BootstrapReadAllText = text
End Function
