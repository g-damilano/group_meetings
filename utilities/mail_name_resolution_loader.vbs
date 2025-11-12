' Loader for mail name resolution modules
'
' Provides LoadNameResolutionModules for console scripts that need the
' complete name resolution helper set.

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

  Dim modules
  modules = Array( _
    "mail_name_resolution_prompt.vbs", _
    "mail_name_resolution_text.vbs", _
    "mail_name_resolution_directory.vbs", _
    "mail_name_resolution_core.vbs", _
    "mail_name_resolution_console.vbs")

  Dim i
  For i = LBound(modules) To UBound(modules)
    Dim relativePath: relativePath = modules(i)
    If Len(ResolveModulePath(fso, effectiveBase, relativePath)) = 0 Then
      relativePath = "utilities\\" & modules(i)
    End If
    Call LoadModule(effectiveBase, relativePath)
  Next
End Sub
