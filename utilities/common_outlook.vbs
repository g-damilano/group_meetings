'====================================================================
' common_outlook.vbs — Shared Outlook automation helpers
'====================================================================

Function GetOutlookApp()
  Dim app
  On Error Resume Next
  Set app = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set app = CreateObject("Outlook.Application")
  End If
  On Error GoTo 0
  Set GetOutlookApp = app
End Function

Sub EmbedImagesFromFolder(olMail, folderPath)
  Dim fso, file, cid, att
  Set fso = CreateObject("Scripting.FileSystemObject")
  For Each file In fso.GetFolder(folderPath).Files
    cid = fso.GetFileName(file.Name)
    Set att = olMail.Attachments.Add(file.Path)
    att.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, cid
  Next
End Sub

Function DefaultLogPath()
  Dim fso, scriptFolder
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."
  DefaultLogPath = scriptFolder & "\\mail_util.log"
End Function

Sub LogAction(logPath, message)
  On Error Resume Next
  If Len(Trim(logPath)) = 0 Then logPath = DefaultLogPath()
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & message
  f.Close
End Sub
