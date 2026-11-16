' Minimal module loader for scaffold VBS scripts.

Function SC_CoreRootFromScript()
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder
  folder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(folder) = 0 Then folder = "."

  Dim current
  current = folder
  Do While Len(current) > 0
    If fso.FolderExists(fso.BuildPath(current, "vbs\core")) Then
      SC_CoreRootFromScript = fso.BuildPath(current, "vbs\core")
      Exit Function
    End If
    If LCase(fso.GetFileName(current)) = "core" Then
      SC_CoreRootFromScript = current
      Exit Function
    End If
    Dim parent
    parent = fso.GetParentFolderName(current)
    If StrComp(parent, current, vbTextCompare) = 0 Then Exit Do
    current = parent
  Loop

  SC_CoreRootFromScript = ""
End Function

Sub SC_LoadModule(modulePath)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(modulePath) Then
    WScript.Echo "Fatal: module not found: " & modulePath
    WScript.Quit 1
  End If

  Dim ts
  Set ts = fso.OpenTextFile(modulePath, 1)
  ExecuteGlobal ts.ReadAll
  ts.Close
End Sub

Sub SC_LoadCoreModules()
  Dim coreRoot
  coreRoot = SC_CoreRootFromScript()
  If Len(coreRoot) = 0 Then
    WScript.Echo "Fatal: unable to locate scaffold vbs\core folder."
    WScript.Quit 1
  End If

  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  SC_LoadModule fso.BuildPath(coreRoot, "contracts.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "logger.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "session.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "folder_access.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "recipient_resolver.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "template_engine.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "mail_assets.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "mail_composer.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "managed_dispatch.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "outbox_inspector.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "health_diagnostics.vbs")
  SC_LoadModule fso.BuildPath(coreRoot, "workflow_facade.vbs")
End Sub
