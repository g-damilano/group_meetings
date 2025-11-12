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

  Dim loaderStream
  On Error Resume Next
  Set loaderStream = entryFso.OpenTextFile(loaderPath, 1)
  If Err.Number <> 0 Then
    Dim loadError: loadError = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "Fatal: unable to read bootstrap loader: " & loadError
    WScript.Quit 1
  End If
  On Error GoTo 0

  Dim loaderContents: loaderContents = loaderStream.ReadAll
  loaderStream.Close

  ExecuteGlobal loaderContents
  Call EnsureCommonBootstrapLoaded(entryFso)
  Set entryFso = Nothing

  bootstrapEntryLoaded = True
End Sub

Call BootstrapEnsureLoaded()
