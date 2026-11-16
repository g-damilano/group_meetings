Option Explicit
SetLocale 1033

'======================================================================
' Bootstrap shared utilities
'======================================================================
Call BootstrapIncludeEntry()

Sub BootstrapIncludeEntry()
  Dim entryFso: Set entryFso = CreateObject("Scripting.FileSystemObject")
  Dim scriptFolder: scriptFolder = entryFso.GetParentFolderName(WScript.ScriptFullName)
  If Len(scriptFolder) = 0 Then scriptFolder = "."

  Dim entryPath
  entryPath = entryFso.BuildPath(scriptFolder, "utilities\\bootstrap_entry.vbs")
  If Not entryFso.FileExists(entryPath) Then
    entryPath = entryFso.BuildPath(scriptFolder, "bootstrap_entry.vbs")
  End If

  If Not entryFso.FileExists(entryPath) Then
    WScript.Echo "Fatal: unable to locate utilities\\bootstrap_entry.vbs"
    WScript.Quit 1
  End If

  Dim entryStream
  On Error Resume Next
  Set entryStream = entryFso.OpenTextFile(entryPath, 1)
  If Err.Number <> 0 Then
    Dim entryError: entryError = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "Fatal: unable to read bootstrap entry: " & entryError
    WScript.Quit 1
  End If
  On Error GoTo 0

  Dim entryContents: entryContents = entryStream.ReadAll
  entryStream.Close

  ExecuteGlobal entryContents
  Set entryFso = Nothing
End Sub

'======================================================================
' InductionMailer.vbs — Console workflow for staged induction templates
'======================================================================

Const TEMPLATE_DIR_NAME = "templates"
Const LOG_FILE_NAME     = "InductionMailer.log"
Const DEFAULT_ACTION    = "display"   ' display | send | save
Const MAILER_LIB_FILE   = "mail_utils_mailer_dynamiclinks.vbs"
Const MAILER_LIB_EXT_SUBPATH = "mail_utils_utf8_plus_dynamiclinks_ext\mail_utils_mailer_dynamiclinks.vbs"
Const LINK_MAP_FILE = "links_map.txt"

Const HEADER_PREFIX = "@--"
Const HEADER_SUFFIX = "--"
Const HTML_HEADER_PREFIX = "<!--"
Const HTML_HEADER_SUFFIX = "-->"
Const DELIVERY_BROADCAST = "broadcast"
Const DELIVERY_PERSONAL  = "personal"
Const TASK_EMAILS = "emails"
Const TASK_ROTA_PEOPLE = "rota_people"

Const ROTA_RCLONE_EXE = "C:\rclone\rclone.exe"
Const ROTA_REMOTE_WORKBOOK = "ABTsharepoint:/Group Meetings/meeting_rota.xlsx"
Const ROTA_WORKBOOK_URL = "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx"
Const ROTA_LOCAL_WORKBOOK = "C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings\meeting_rota.xlsx"
Const ROTA_SHEET_EMAILS = "Emails"
Const ROTA_PEOPLE_TABLE = "Table1"
Const ROTA_RCLONE_DOWNLOAD_TIMEOUT_SECONDS = 120
Const ROTA_RCLONE_UPLOAD_TIMEOUT_SECONDS = 180
Const ROTA_RCLONE_STATUS_SECONDS = 10

Dim fso, scriptFolder, templateFolder, logFilePath, linkMapContext

Call LoadUtilityModules()
Call Main()

'======================================================================
' Support module bootstrap
'======================================================================
Sub LoadUtilityModules()
  Dim modules
  modules = Array( _
    "utilities\common_console.vbs", _
    "utilities\common_outlook.vbs", _
    "utilities\mail_collections.vbs", _
    "utilities\mail_template_utils.vbs", _
    "utilities\mail_delivery_utils.vbs", _
    "utilities\mail_input_validation.vbs", _
    "utilities\MNS_loader.vbs")

  Dim baseFolder: baseFolder = GetScriptFolder()

  On Error Resume Next
  Call LoadBootstrapModules(modules)
  If Err.Number <> 0 Then
    WScript.Echo "Fatal: unable to load utility modules: " & Err.Description
    Err.Clear
    WScript.Quit 1
  End If

  Call LoadNameResolutionModules(baseFolder)
  If Err.Number <> 0 Then
    WScript.Echo "Fatal: unable to load name resolution modules: " & Err.Description
    Err.Clear
    WScript.Quit 1
  End If
  On Error GoTo 0
End Sub

'======================================================================
' Entry point
'======================================================================
Sub Main()
  If Not EnsureConsoleHost() Then Exit Sub

  InitializeEnvironment

  Dim workflowTask: workflowTask = PromptWorkflowTask()
  If Len(workflowTask) = 0 Then
    WScript.Echo "Operation cancelled."
    Exit Sub
  End If

  If workflowTask = TASK_ROTA_PEOPLE Then
    Dim rotaPrep: rotaPrep = PrepareMeetingRotaWorkbookForEdit()
    If rotaPrep = "uploaded" Then
      WScript.Echo "Pending local meeting rota workbook uploaded."
      Exit Sub
    End If
    If rotaPrep <> "ready" Then
      WScript.Echo "Meeting rota workbook was not prepared for editing."
      Exit Sub
    End If

    Dim rotaRecipientPackage: Set rotaRecipientPackage = PromptRecipientResolutionPackage()
    If Not HasRecipientResolutionEntries(rotaRecipientPackage) Then
      WScript.Echo "No valid people supplied."
      Exit Sub
    End If

    If AddRecipientsToMeetingRotaPeople(rotaRecipientPackage) Then
      WScript.Echo "Meeting rota people sheet update completed."
    Else
      WScript.Echo "Meeting rota people sheet update did not complete. Check " & logFilePath & "."
    End If
    Exit Sub
  End If

  If Not LoadMailerLibrary() Then
    LogEvent "Unable to load mail_utils_mailer_dynamiclinks.vbs"
    WScript.Echo "Fatal: mail utility library missing."
    Exit Sub
  End If

  Set linkMapContext = LoadLinkMapContext()

  Dim templates: Set templates = LoadTemplates(templateFolder)
  If templates.Count = 0 Then
    WScript.Echo "No templates were found under '" & templateFolder & "'."
    Exit Sub
  End If

  Dim recipientPackage: Set recipientPackage = PromptRecipientResolutionPackage()
  If Not HasRecipientResolutionEntries(recipientPackage) Then
    WScript.Echo "No valid target email addresses supplied."
    Exit Sub
  End If

  Dim action: action = DetermineAction()

  Dim mode: mode = PromptOperationMode()
  If Len(mode) = 0 Then
    WScript.Echo "Operation cancelled."
    Exit Sub
  End If

  Dim selectedTemplates: Set selectedTemplates = Nothing
  If mode = "stage" Then
    Set selectedTemplates = SelectTemplatesByStage(templates)
  Else
    Set selectedTemplates = SelectTemplatesIndividually(templates)
  End If

  If selectedTemplates Is Nothing Or selectedTemplates.Count = 0 Then
    WScript.Echo "No templates selected."
    Exit Sub
  End If

  Dim key
  For Each key In selectedTemplates.Keys
    ProcessTemplate selectedTemplates(key), recipientPackage, action
  Next

  WScript.Echo "All requested templates processed."
End Sub

Function PromptLine(prompt)
  WScript.StdOut.Write prompt
  PromptLine = Trim(WScript.StdIn.ReadLine())
End Function

Function PromptWorkflowTask()
  WScript.Echo "Select the induction task:" & vbCrLf & _
               "  1) Prepare induction email(s)" & vbCrLf & _
               "  2) Add person to group meeting rota people sheet"
  Dim task: task = TASK_EMAILS
  Do
    Dim choice: choice = PromptLine("Enter choice (1/2, default 1) or press Enter for emails: ")
    If Len(choice) = 0 Then
      task = TASK_EMAILS
      Exit Do
    End If
    Select Case LCase(choice)
      Case "1", "email", "emails"
        task = TASK_EMAILS
        Exit Do
      Case "2", "rota", "people", "rota_people"
        task = TASK_ROTA_PEOPLE
        Exit Do
      Case "q", "quit", "cancel"
        task = ""
        Exit Do
      Case Else
        WScript.Echo "Invalid option. Please choose 1 or 2."
    End Select
  Loop
  PromptWorkflowTask = task
End Function

Function PromptOperationMode()
  WScript.Echo "Select the operation mode:"& vbCrLf & _
               "  1) Stage mode (send every email in a stage)" & vbCrLf & _
               "  2) Email mode (pick one or more templates)"
  Dim mode: mode = ""
  Do
    Dim choice: choice = PromptLine("Enter choice (1/2) or press Enter to cancel: ")
    If Len(choice) = 0 Then
      mode = ""
      Exit Do
    End If
    Select Case choice
      Case "1": mode = "stage": Exit Do
      Case "2": mode = "email": Exit Do
      Case Else
        WScript.Echo "Invalid option. Please choose 1 or 2."
    End Select
  Loop
  PromptOperationMode = mode
End Function

'======================================================================
' Environment bootstrap
'======================================================================
Sub InitializeEnvironment()
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptFolder = GetScriptFolder()
  templateFolder = fso.BuildPath(scriptFolder, TEMPLATE_DIR_NAME)
  logFilePath = fso.BuildPath(scriptFolder, LOG_FILE_NAME)
  LogEvent "=== Induction Mailer (console) started ==="
  LogEvent "Template folder: " & templateFolder
  Set linkMapContext = CreateCaseInsensitiveDictionary()
End Sub

Function LoadMailerLibrary()
  Dim candidates
  candidates = Array( _
    fso.BuildPath(scriptFolder, MAILER_LIB_FILE), _
    fso.BuildPath(scriptFolder, MAILER_LIB_EXT_SUBPATH), _
    fso.BuildPath(templateFolder, MAILER_LIB_FILE))

  Dim i
  For i = LBound(candidates) To UBound(candidates)
    If fso.FileExists(candidates(i)) Then
      Dim ts, contents
      Set ts = fso.OpenTextFile(candidates(i), 1)
      contents = ts.ReadAll
      ts.Close
      contents = RemoveByteOrderMark(contents)
      ExecuteGlobal contents
      LogEvent "Loaded mailer library from " & candidates(i)
      EnsureLinkMapAvailable candidates(i)
      LoadMailerLibrary = True
      Exit Function
    End If
  Next

  LoadMailerLibrary = False
End Function

Sub EnsureLinkMapAvailable(libraryPath)
  On Error Resume Next
  Dim libFolder: libFolder = fso.GetParentFolderName(libraryPath)
  Dim sourceMap: sourceMap = fso.BuildPath(libFolder, LINK_MAP_FILE)
  Dim targetMap: targetMap = fso.BuildPath(scriptFolder, LINK_MAP_FILE)
  If fso.FileExists(sourceMap) Then
    If Not fso.FileExists(targetMap) Then
      fso.CopyFile sourceMap, targetMap
      If Err.Number = 0 Then
        LogEvent "Copied " & LINK_MAP_FILE & " to script folder"
      Else
        LogEvent "Failed to copy " & LINK_MAP_FILE & ": " & Err.Description
        Err.Clear
      End If
    End If
  End If
  On Error GoTo 0
End Sub

'======================================================================
' Group meeting rota people-sheet update
'======================================================================
Function PrepareMeetingRotaWorkbookForEdit()
  PrepareMeetingRotaWorkbookForEdit = ""

  If fso.FileExists(ROTA_LOCAL_WORKBOOK) Then
    WScript.Echo "A local meeting_rota.xlsx already exists:"
    WScript.Echo "  " & ROTA_LOCAL_WORKBOOK
    WScript.Echo "It may contain a previous saved rota update that was not uploaded."
    WScript.Echo "  1) Sync this local workbook to SharePoint now"
    WScript.Echo "  2) Continue editing this local workbook"
    WScript.Echo "  3) Back it up, then replace it with the latest SharePoint copy"
    WScript.Echo "  4) Cancel"

    Do
      Dim choice: choice = PromptLine("Choose recovery option (1-4): ")
      Select Case LCase(Trim(choice))
        Case "1", "retry", "upload"
          If UploadMeetingRotaWorkbookAfterEdit() Then
            CleanupLocalRotaWorkbook True
            PrepareMeetingRotaWorkbookForEdit = "uploaded"
          Else
            WScript.Echo "SharePoint sync failed. The local workbook was kept for recovery."
          End If
          Exit Function
        Case "2", "continue", "local"
          LogEvent "Operator chose to continue editing existing local rota workbook."
          PrepareMeetingRotaWorkbookForEdit = "ready"
          Exit Function
        Case "3", "replace", "download"
          If Not BackupLocalRotaWorkbook() Then Exit Function
          If SyncMeetingRotaWorkbookForEdit() Then PrepareMeetingRotaWorkbookForEdit = "ready"
          Exit Function
        Case "4", "q", "quit", "cancel"
          LogEvent "Operator cancelled rota workbook recovery prompt."
          Exit Function
        Case Else
          WScript.Echo "Invalid option. Please choose 1, 2, 3, or 4."
      End Select
    Loop
  End If

  If SyncMeetingRotaWorkbookForEdit() Then PrepareMeetingRotaWorkbookForEdit = "ready"
End Function

Function BackupLocalRotaWorkbook()
  BackupLocalRotaWorkbook = False
  If Not fso.FileExists(ROTA_LOCAL_WORKBOOK) Then
    BackupLocalRotaWorkbook = True
    Exit Function
  End If

  Dim backupPath: backupPath = fso.BuildPath(scriptFolder, "meeting_rota_recovery_" & RotaTimestampForFile() & ".xlsx")
  On Error Resume Next
  fso.CopyFile ROTA_LOCAL_WORKBOOK, backupPath, True
  If Err.Number <> 0 Then
    WScript.Echo "Could not back up the existing local workbook:"
    WScript.Echo "  " & Err.Description
    LogEvent "ERROR: Failed to back up local rota workbook before replacement: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  WScript.Echo "Backed up existing local workbook to:"
  WScript.Echo "  " & backupPath
  LogEvent "Backed up local rota workbook before replacement: " & backupPath
  BackupLocalRotaWorkbook = True
End Function

Function RotaTimestampForFile()
  RotaTimestampForFile = _
    CStr(Year(Now)) & _
    Right("0" & Month(Now), 2) & _
    Right("0" & Day(Now), 2) & "_" & _
    Right("0" & Hour(Now), 2) & _
    Right("0" & Minute(Now), 2) & _
    Right("0" & Second(Now), 2)
End Function

Function AddRecipientsToMeetingRotaPeople(recipientPackage)
  AddRecipientsToMeetingRotaPeople = False

  Dim entries: entries = GetRecipientResolutionEntries(recipientPackage)
  Dim entryUpper: entryUpper = GetRecipientResolutionArrayUpperBound(entries)
  If entryUpper < 0 Then
    WScript.Echo "No recipient entries were available for the rota update."
    LogEvent "Rota people update aborted: no recipient entries."
    Exit Function
  End If

  Dim objExcel, objWorkbook, objSheetEmails
  Set objExcel = Nothing
  Set objWorkbook = Nothing
  Set objSheetEmails = Nothing

  If Not OpenMeetingRotaWorkbookForEdit(objExcel, objWorkbook, objSheetEmails) Then
    Exit Function
  End If

  Dim addedCount: addedCount = 0
  Dim skippedCount: skippedCount = 0
  Dim duplicateCount: duplicateCount = 0
  Dim i

  For i = 0 To entryUpper
    Dim entry
    Set entry = entries(i)
    If entry Is Nothing Then
      skippedCount = skippedCount + 1
    Else
      Dim person: Set person = BuildRotaPersonContext(entry)
      If person Is Nothing Then
        skippedCount = skippedCount + 1
      ElseIf Not ConfirmRotaPerson(person) Then
        skippedCount = skippedCount + 1
        LogEvent "Rota people update skipped by operator for " & GetDictionaryValue(person, "FULL_NAME")
      Else
        Dim status: status = AppendRotaPersonIfMissing(objSheetEmails, person)
        Select Case status
          Case "added"
            addedCount = addedCount + 1
          Case "duplicate"
            duplicateCount = duplicateCount + 1
          Case Else
            skippedCount = skippedCount + 1
        End Select
      End If
    End If
  Next

  If addedCount > 0 Then
    On Error Resume Next
    objWorkbook.Save
    If Err.Number <> 0 Then
      LogEvent "ERROR: Failed to save meeting rota workbook: " & Err.Description
      WScript.Echo "Could not save the local workbook copy. No upload was attempted."
      Err.Clear
      CloseMeetingRotaWorkbookForEdit objExcel, objWorkbook, objSheetEmails, False
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    LogEvent "Saved meeting rota workbook after adding " & addedCount & " people."
  Else
    LogEvent "No new rota people rows added. Duplicates: " & duplicateCount & " | Skipped: " & skippedCount
  End If

  CloseMeetingRotaWorkbookForEdit objExcel, objWorkbook, objSheetEmails, False

  If addedCount > 0 Then
    If Not UploadMeetingRotaWorkbookAfterEdit() Then
      WScript.Echo "The local workbook was saved, but SharePoint sync failed."
      WScript.Echo "Close any open copy of meeting_rota.xlsx, then rerun this option and choose recovery option 1."
      WScript.Echo "Local copy kept at: " & ROTA_LOCAL_WORKBOOK
      Exit Function
    End If
    CleanupLocalRotaWorkbook True
  End If

  WScript.Echo "Added: " & addedCount & " | Already present: " & duplicateCount & " | Skipped: " & skippedCount
  LogEvent "Rota people update completed. Added: " & addedCount & " | Duplicates: " & duplicateCount & " | Skipped: " & skippedCount
  AddRecipientsToMeetingRotaPeople = True
End Function

Function SyncMeetingRotaWorkbookForEdit()
  SyncMeetingRotaWorkbookForEdit = False

  If Not fso.FileExists(ROTA_RCLONE_EXE) Then
    WScript.Echo "rclone was not found at: " & ROTA_RCLONE_EXE
    LogEvent "ERROR: rclone not found for rota workbook sync: " & ROTA_RCLONE_EXE
    Exit Function
  End If

  Dim targetFolder: targetFolder = fso.GetParentFolderName(ROTA_LOCAL_WORKBOOK)
  If Len(targetFolder) > 0 Then
    If Not fso.FolderExists(targetFolder) Then fso.CreateFolder targetFolder
  End If

  Dim cmd: cmd = QuoteRotaArg(ROTA_RCLONE_EXE) & " copyto " & _
                 QuoteRotaArg(ROTA_REMOTE_WORKBOOK) & " " & _
                 QuoteRotaArg(ROTA_LOCAL_WORKBOOK) & " --ignore-times " & RcloneContextFlags(False)
  Dim shell: Set shell = CreateObject("WScript.Shell")
  LogEvent "Downloading meeting rota workbook for edit: " & cmd
  WScript.Echo "Downloading the latest meeting rota workbook..."

  Dim result: Set result = RunRcloneCommandWithTimeout(cmd, "download", ROTA_RCLONE_DOWNLOAD_TIMEOUT_SECONDS)
  LogRcloneResult "download", result
  If CBool(result("TimedOut")) Or CLng(result("ExitCode")) <> 0 Then
    WScript.Echo "Could not download the meeting rota workbook."
    PrintRcloneResultToConsole "download", result
    WScript.Echo "Full details were written to:"
    WScript.Echo "  " & logFilePath
    Exit Function
  End If

  If Not fso.FileExists(ROTA_LOCAL_WORKBOOK) Then
    WScript.Echo "Downloaded workbook was not found at: " & ROTA_LOCAL_WORKBOOK
    LogEvent "ERROR: Meeting rota workbook missing after download: " & ROTA_LOCAL_WORKBOOK
    Exit Function
  End If

  SyncMeetingRotaWorkbookForEdit = True
End Function

Function UploadMeetingRotaWorkbookAfterEdit()
  UploadMeetingRotaWorkbookAfterEdit = False

  If Not fso.FileExists(ROTA_LOCAL_WORKBOOK) Then
    LogEvent "ERROR: Cannot upload missing meeting rota workbook: " & ROTA_LOCAL_WORKBOOK
    Exit Function
  End If

  WScript.Echo "Syncing the edited meeting rota workbook through Excel first..."
  LogEvent "Attempting Excel COM sync before rclone replacement upload."
  If TryExcelComFallbackUploadFromLocalWorkbook() Then
    UploadMeetingRotaWorkbookAfterEdit = True
    Exit Function
  End If

  If Not PromptYesNo("Excel sync did not complete. Try rclone replacement upload instead? (y/N): ", False) Then
    WScript.Echo "rclone replacement upload skipped. Local workbook kept at: " & ROTA_LOCAL_WORKBOOK
    Exit Function
  End If

  Dim cmd: cmd = QuoteRotaArg(ROTA_RCLONE_EXE) & " copyto " & _
                 QuoteRotaArg(ROTA_LOCAL_WORKBOOK) & " " & _
                 QuoteRotaArg(ROTA_REMOTE_WORKBOOK) & " --ignore-times " & RcloneContextFlags(True)
  Dim shell: Set shell = CreateObject("WScript.Shell")
  LogEvent "Uploading edited meeting rota workbook: " & cmd
  WScript.Echo "Uploading the edited meeting rota workbook..."

  Dim result: Set result = RunRcloneCommandWithTimeout(cmd, "upload", ROTA_RCLONE_UPLOAD_TIMEOUT_SECONDS)
  LogRcloneResult "upload", result
  If CBool(result("TimedOut")) Then
    WScript.Echo "Upload timed out after " & ROTA_RCLONE_UPLOAD_TIMEOUT_SECONDS & " seconds."
    WScript.Echo "This usually means the workbook is open/locked or SharePoint is not accepting the replacement."
    PrintRcloneResultToConsole "upload", result
    WScript.Echo "Local workbook kept at: " & ROTA_LOCAL_WORKBOOK
    Exit Function
  End If
  If CLng(result("ExitCode")) <> 0 Then
    WScript.Echo "Upload failed with exit code " & result("ExitCode") & "."
    PrintRcloneResultToConsole "upload", result
    WScript.Echo "Full details were written to:"
    WScript.Echo "  " & logFilePath
    WScript.Echo "Local workbook kept at: " & ROTA_LOCAL_WORKBOOK
    Exit Function
  End If

  UploadMeetingRotaWorkbookAfterEdit = True
End Function

Function RcloneContextFlags(isUpload)
  Dim flags
  flags = "--stats " & ROTA_RCLONE_STATUS_SECONDS & "s --stats-log-level NOTICE " & _
          "--retries 1 --low-level-retries 1 --timeout 30s --contimeout 15s"
  If CBool(isUpload) Then
    flags = flags & " --onedrive-upload-cutoff 4M"
  End If
  RcloneContextFlags = flags
End Function

Function RunRcloneCommandWithTimeout(commandText, operationName, timeoutSeconds)
  Dim result: Set result = CreateCaseInsensitiveDictionary()
  result("ExitCode") = -1
  result("TimedOut") = False
  result("StdOut") = ""
  result("StdErr") = ""

  Dim shell: Set shell = CreateObject("WScript.Shell")
  Dim exec
  On Error Resume Next
  Set exec = shell.Exec(commandText)
  If Err.Number <> 0 Or exec Is Nothing Then
    result("ExitCode") = -9999
    result("StdErr") = "Unable to start rclone " & operationName & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    Set RunRcloneCommandWithTimeout = result
    Exit Function
  End If
  On Error GoTo 0

  Dim startedAt: startedAt = Now
  Dim lastStatusAt: lastStatusAt = startedAt
  WScript.Echo "rclone " & operationName & " started. Waiting for status..."
  Do While exec.Status = 0
    DrainRcloneStream exec.StdOut, "stdout", result
    DrainRcloneStream exec.StdErr, "stderr", result

    If DateDiff("s", lastStatusAt, Now) >= ROTA_RCLONE_STATUS_SECONDS Then
      WScript.Echo "rclone " & operationName & " still running... " & _
                   DateDiff("s", startedAt, Now) & "s elapsed, timeout at " & timeoutSeconds & "s."
      lastStatusAt = Now
    End If

    WScript.Sleep 500
    If DateDiff("s", startedAt, Now) >= CLng(timeoutSeconds) Then
      result("TimedOut") = True
      On Error Resume Next
      exec.Terminate
      If Err.Number <> 0 Then
        result("StdErr") = "Timed out and failed to terminate rclone: " & Err.Description
        Err.Clear
      End If
      On Error GoTo 0
      Exit Do
    End If
  Loop

  DrainRcloneStream exec.StdOut, "stdout", result
  DrainRcloneStream exec.StdErr, "stderr", result

  If Not CBool(result("TimedOut")) Then
    On Error Resume Next
    result("ExitCode") = CLng(exec.ExitCode)
    If Err.Number <> 0 Then
      result("ExitCode") = -1
      Err.Clear
    End If
    On Error GoTo 0
  End If

  On Error Resume Next
  If Not exec.StdOut.AtEndOfStream Then AppendRcloneOutput result, "StdOut", exec.StdOut.ReadAll, "stdout"
  If Err.Number <> 0 Then Err.Clear
  If Not exec.StdErr.AtEndOfStream Then AppendRcloneOutput result, "StdErr", exec.StdErr.ReadAll, "stderr"
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0

  Set RunRcloneCommandWithTimeout = result
End Function

Sub DrainRcloneStream(stream, streamName, result)
  On Error Resume Next
  Dim lineCount: lineCount = 0
  Do While Not stream.AtEndOfStream
    Dim line: line = stream.ReadLine
    If Err.Number <> 0 Then
      Err.Clear
      Exit Do
    End If
    AppendRcloneOutput result, ProperRcloneStreamKey(streamName), line & vbCrLf, streamName
    lineCount = lineCount + 1
    If lineCount >= 20 Then Exit Do
  Loop
  On Error GoTo 0
End Sub

Function ProperRcloneStreamKey(streamName)
  If LCase(streamName) = "stdout" Then
    ProperRcloneStreamKey = "StdOut"
  Else
    ProperRcloneStreamKey = "StdErr"
  End If
End Function

Sub AppendRcloneOutput(result, keyName, text, streamName)
  Dim value: value = SafeText(text)
  If Len(value) = 0 Then Exit Sub
  result(keyName) = SafeText(result(keyName)) & value

  Dim lines: lines = Split(value, vbCrLf)
  Dim i
  For i = 0 To UBound(lines)
    Dim line: line = Trim(lines(i))
    If Len(line) > 0 Then WScript.Echo "rclone " & streamName & ": " & line
  Next
End Sub

Sub LogRcloneResult(operationName, result)
  LogEvent "Meeting rota " & operationName & " exit code: " & result("ExitCode") & " | Timed out: " & CStr(result("TimedOut"))
  If Len(Trim(SafeText(result("StdOut")))) > 0 Then LogEvent "rclone " & operationName & " stdout: " & Trim(SafeText(result("StdOut")))
  If Len(Trim(SafeText(result("StdErr")))) > 0 Then LogEvent "rclone " & operationName & " stderr: " & Trim(SafeText(result("StdErr")))
End Sub

Sub PrintRcloneResultToConsole(operationName, result)
  Dim stdOut: stdOut = Trim(SafeText(result("StdOut")))
  Dim stdErr: stdErr = Trim(SafeText(result("StdErr")))

  WScript.Echo "rclone " & operationName & " result: exit code " & result("ExitCode") & ", timed out: " & CStr(result("TimedOut"))
  If Len(stdOut) > 0 Then
    WScript.Echo "rclone stdout:"
    WScript.Echo LastRcloneOutputChunk(stdOut)
  End If
  If Len(stdErr) > 0 Then
    WScript.Echo "rclone stderr:"
    WScript.Echo LastRcloneOutputChunk(stdErr)
  End If
  If Len(stdOut) = 0 And Len(stdErr) = 0 Then
    WScript.Echo "rclone did not return stdout/stderr before the operation ended."
  End If
End Sub

Function LastRcloneOutputChunk(text)
  Dim value: value = SafeText(text)
  Dim maxChars: maxChars = 2500
  If Len(value) <= maxChars Then
    LastRcloneOutputChunk = value
  Else
    LastRcloneOutputChunk = "... last " & maxChars & " characters ..." & vbCrLf & Right(value, maxChars)
  End If
End Function

Function PromptYesNo(promptText, defaultYes)
  Dim response: response = LCase(Trim(PromptLine(promptText)))
  If Len(response) = 0 Then
    PromptYesNo = CBool(defaultYes)
  Else
    PromptYesNo = (Left(response, 1) = "y")
  End If
End Function

Function TryExcelComFallbackUploadFromLocalWorkbook()
  TryExcelComFallbackUploadFromLocalWorkbook = False

  If Not fso.FileExists(ROTA_LOCAL_WORKBOOK) Then
    WScript.Echo "Excel sync cannot run because the local workbook is missing:"
    WScript.Echo "  " & ROTA_LOCAL_WORKBOOK
    LogEvent "ERROR: Excel sync missing local workbook: " & ROTA_LOCAL_WORKBOOK
    Exit Function
  End If

  WScript.Echo "Starting Excel sync. Excel may show sign-in, lock, or conflict prompts."
  WScript.Echo "Reading pending people from the local saved workbook..."
  LogEvent "Starting Excel COM sync for meeting rota upload."

  Dim objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet
  Set objExcel = Nothing
  Set localWorkbook = Nothing
  Set remoteWorkbook = Nothing
  Set localSheet = Nothing
  Set remoteSheet = Nothing

  On Error Resume Next
  Set objExcel = CreateObject("Excel.Application")
  If Err.Number <> 0 Or objExcel Is Nothing Then
    WScript.Echo "Excel sync could not start Excel: " & Err.Description
    LogEvent "ERROR: Excel sync could not start Excel: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  objExcel.Visible = False
  objExcel.DisplayAlerts = False

  Dim pendingPeople: Set pendingPeople = LoadPendingRotaPeopleFromLocalWorkbook(objExcel)
  If pendingPeople Is Nothing Then
    CloseExcelFallback objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet, False
    Exit Function
  End If
  WScript.Echo "Pending people loaded from local workbook: " & pendingPeople.Count

  WScript.Echo "Opening SharePoint workbook..."
  Set remoteWorkbook = OpenSharePointWorkbookWithOperatorFallback(objExcel)
  If remoteWorkbook Is Nothing Then
    CloseExcelFallback objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet, False
    Exit Function
  End If

  If Not EnsureSharePointWorkbookEditable(remoteWorkbook) Then
    WScript.Echo "Excel could not obtain edit access for the SharePoint workbook."
    LogEvent "ERROR: Excel sync could not obtain edit access for SharePoint workbook."
    CloseExcelFallback objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet, False
    On Error GoTo 0
    Exit Function
  End If

  Set remoteSheet = remoteWorkbook.Sheets(ROTA_SHEET_EMAILS)
  If Err.Number <> 0 Or remoteSheet Is Nothing Then
    WScript.Echo "Excel sync could not access the SharePoint '" & ROTA_SHEET_EMAILS & "' sheet: " & Err.Description
    LogEvent "ERROR: Excel sync could not access SharePoint Emails sheet: " & Err.Description
    Err.Clear
    CloseExcelFallback objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet, False
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Dim addedCount: addedCount = AppendPendingRotaPeopleToRemoteWorkbook(pendingPeople, remoteSheet)
  If addedCount > 0 Then
    On Error Resume Next
    WScript.Echo "Saving SharePoint workbook through Excel..."
    remoteWorkbook.Save
    If Err.Number <> 0 Then
      WScript.Echo "Excel sync could not save the SharePoint workbook: " & Err.Description
      LogEvent "ERROR: Excel sync save failed: " & Err.Description
      Err.Clear
      If Not WaitForExcelSaveResolution(objExcel, remoteWorkbook) Then
        CloseExcelFallback objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet, False
        On Error GoTo 0
        Exit Function
      End If
    End If
    On Error GoTo 0
  Else
    WScript.Echo "Excel sync found no missing people to append; SharePoint workbook already appears current."
  End If

  CloseExcelFallback objExcel, localWorkbook, remoteWorkbook, localSheet, remoteSheet, False

  WScript.Echo "Excel sync completed. Missing people appended: " & addedCount
  LogEvent "Excel COM sync completed. Missing people appended: " & addedCount
  TryExcelComFallbackUploadFromLocalWorkbook = True
End Function

Function OpenSharePointWorkbookWithOperatorFallback(objExcel)
  Set OpenSharePointWorkbookWithOperatorFallback = Nothing
  If objExcel Is Nothing Then Exit Function

  Dim remoteWorkbook
  Set remoteWorkbook = Nothing

  Do
    Err.Clear
    On Error Resume Next
    Set remoteWorkbook = objExcel.Workbooks.Open(ROTA_WORKBOOK_URL, 0, False)
    If Err.Number = 0 And Not (remoteWorkbook Is Nothing) Then
      On Error GoTo 0
      Set OpenSharePointWorkbookWithOperatorFallback = remoteWorkbook
      Exit Function
    End If

    Dim openError: openError = Err.Description
    Err.Clear
    On Error GoTo 0

    WScript.Echo "Excel sync could not open the SharePoint workbook: " & openError
    WScript.Echo "Excel will now be shown so you can resolve any sign-in, trust, or server prompt."
    LogEvent "ERROR: Excel sync could not open SharePoint workbook: " & openError
    ShowExcelForOperator objExcel

    Dim response: response = PromptLine("After resolving the Excel prompt, press Enter to retry opening it, or type cancel: ")
    If LCase(Trim(response)) = "cancel" Or LCase(Trim(response)) = "q" Then
      LogEvent "Operator cancelled SharePoint workbook open retry."
      Exit Function
    End If
  Loop
End Function

Function WaitForExcelSaveResolution(objExcel, remoteWorkbook)
  WaitForExcelSaveResolution = False
  If objExcel Is Nothing Or remoteWorkbook Is Nothing Then Exit Function

  ShowExcelForOperator objExcel
  Do
    WScript.Echo "Excel is visible. Resolve any save, conflict, read-only, or sign-in prompt."
    Dim response: response = PromptLine("After resolving it, press Enter to retry Save, or type cancel: ")
    If LCase(Trim(response)) = "cancel" Or LCase(Trim(response)) = "q" Then
      LogEvent "Operator cancelled Excel save retry."
      Exit Function
    End If

    On Error Resume Next
    remoteWorkbook.Save
    If Err.Number = 0 Then
      On Error GoTo 0
      WaitForExcelSaveResolution = True
      Exit Function
    End If
    WScript.Echo "Excel save still failed: " & Err.Description
    LogEvent "ERROR: Excel save retry failed: " & Err.Description
    Err.Clear
    On Error GoTo 0
  Loop
End Function

Sub ShowExcelForOperator(objExcel)
  On Error Resume Next
  If Not (objExcel Is Nothing) Then
    objExcel.Visible = True
    objExcel.DisplayAlerts = True
    objExcel.WindowState = -4137 ' xlMaximized
  End If
  On Error GoTo 0
End Sub

Function EnsureSharePointWorkbookEditable(remoteWorkbook)
  EnsureSharePointWorkbookEditable = False
  If remoteWorkbook Is Nothing Then Exit Function

  On Error Resume Next
  If Not CBool(remoteWorkbook.ReadOnly) Then
    EnsureSharePointWorkbookEditable = True
    On Error GoTo 0
    Exit Function
  End If
  If Err.Number <> 0 Then
    LogEvent "WARNING: Could not read SharePoint workbook ReadOnly state: " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0

  WScript.Echo "Excel opened the SharePoint workbook read-only."
  WScript.Echo "Trying to request edit access through Excel..."
  If TryRequestSharePointEditAccess(remoteWorkbook) Then
    EnsureSharePointWorkbookEditable = True
    Exit Function
  End If

  On Error Resume Next
  ShowExcelForOperator remoteWorkbook.Application
  On Error GoTo 0

  Do
    WScript.Echo "Excel still reports the workbook as read-only."
    WScript.Echo "In Excel, click the yellow-bar 'Edit Workbook' button."
    Dim response: response = PromptLine("After clicking it, press Enter here to continue, or type cancel: ")
    If LCase(Trim(response)) = "cancel" Or LCase(Trim(response)) = "q" Then
      LogEvent "Operator cancelled Excel edit-access wait."
      Exit Function
    End If

    On Error Resume Next
    If Not CBool(remoteWorkbook.ReadOnly) Then
      EnsureSharePointWorkbookEditable = True
      On Error GoTo 0
      Exit Function
    End If
    If Err.Number <> 0 Then
      LogEvent "WARNING: Could not re-check SharePoint workbook ReadOnly state: " & Err.Description
      Err.Clear
    End If
    On Error GoTo 0

    If TryRequestSharePointEditAccess(remoteWorkbook) Then
      EnsureSharePointWorkbookEditable = True
      Exit Function
    End If
  Loop
End Function

Function TryRequestSharePointEditAccess(remoteWorkbook)
  TryRequestSharePointEditAccess = False
  If remoteWorkbook Is Nothing Then Exit Function

  On Error Resume Next
  remoteWorkbook.LockServerFile
  If Err.Number <> 0 Then
    LogEvent "Excel LockServerFile did not grant edit access: " & Err.Description
    Err.Clear
  End If
  WScript.Sleep 1000
  If Not CBool(remoteWorkbook.ReadOnly) Then
    TryRequestSharePointEditAccess = True
    On Error GoTo 0
    Exit Function
  End If
  If Err.Number <> 0 Then Err.Clear

  ' XlFileAccess: xlReadWrite = 2. This may trigger the same server edit transition
  ' as the yellow-bar Edit Workbook button when Excel permits it.
  remoteWorkbook.ChangeFileAccess 2
  If Err.Number <> 0 Then
    LogEvent "Excel ChangeFileAccess(xlReadWrite) did not grant edit access: " & Err.Description
    Err.Clear
  End If
  WScript.Sleep 1000
  If Not CBool(remoteWorkbook.ReadOnly) Then TryRequestSharePointEditAccess = True
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Function

Function LoadPendingRotaPeopleFromLocalWorkbook(objExcel)
  Set LoadPendingRotaPeopleFromLocalWorkbook = Nothing
  If objExcel Is Nothing Then Exit Function

  Dim localWorkbook, localSheet
  Set localWorkbook = Nothing
  Set localSheet = Nothing

  On Error Resume Next
  Set localWorkbook = objExcel.Workbooks.Open(ROTA_LOCAL_WORKBOOK, 0, True)
  If Err.Number <> 0 Or localWorkbook Is Nothing Then
    WScript.Echo "Excel sync could not open the local workbook: " & Err.Description
    LogEvent "ERROR: Excel sync could not open local workbook: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If

  Set localSheet = localWorkbook.Sheets(ROTA_SHEET_EMAILS)
  If Err.Number <> 0 Or localSheet Is Nothing Then
    WScript.Echo "Excel sync could not access the local '" & ROTA_SHEET_EMAILS & "' sheet: " & Err.Description
    LogEvent "ERROR: Excel sync could not access local Emails sheet: " & Err.Description
    Err.Clear
    localWorkbook.Close False
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Dim people: Set people = CreateObject("System.Collections.ArrayList")
  Dim lastDataRow: lastDataRow = GetRotaPeopleLastDataRow(localSheet)
  Dim rowIndex
  For rowIndex = 2 To lastDataRow
    Dim person: Set person = BuildRotaPersonFromWorksheetRow(localSheet, rowIndex)
    If Not person Is Nothing Then
      people.Add person
    End If
  Next

  On Error Resume Next
  Set localSheet = Nothing
  localWorkbook.Close False
  Set localWorkbook = Nothing
  On Error GoTo 0

  Set LoadPendingRotaPeopleFromLocalWorkbook = people
End Function

Function AppendPendingRotaPeopleToRemoteWorkbook(people, remoteSheet)
  AppendPendingRotaPeopleToRemoteWorkbook = 0
  If people Is Nothing Then Exit Function

  Dim i
  For i = 0 To people.Count - 1
    Dim person: Set person = people(i)
    If Not person Is Nothing Then
      Dim status: status = AppendRotaPersonIfMissing(remoteSheet, person)
      If status = "added" Then AppendPendingRotaPeopleToRemoteWorkbook = AppendPendingRotaPeopleToRemoteWorkbook + 1
    End If
  Next
End Function

Function BuildRotaPersonFromWorksheetRow(objSheet, rowIndex)
  Dim person: Set person = CreateCaseInsensitiveDictionary()

  Dim fullName: fullName = Trim(SafeText(objSheet.Cells(rowIndex, FindWorksheetHeaderColumn(objSheet, "Name", 1)).Value))
  Dim emailValue: emailValue = NormalizeEmailText(objSheet.Cells(rowIndex, FindWorksheetHeaderColumn(objSheet, "Email", 2)).Value)
  If Len(fullName) = 0 Or Not IsValidEmail(emailValue) Then
    Set BuildRotaPersonFromWorksheetRow = Nothing
    Exit Function
  End If

  person("FULL_NAME") = fullName
  person("EMAIL") = emailValue
  person("NAME") = ""
  person("SURNAME") = ""

  Dim firstNameCol: firstNameCol = FindWorksheetHeaderColumn(objSheet, "First name", 0)
  If firstNameCol > 0 Then person("NAME") = Trim(SafeText(objSheet.Cells(rowIndex, firstNameCol).Value))

  Dim surnameCol: surnameCol = FindWorksheetHeaderColumn(objSheet, "Surname", 0)
  If surnameCol > 0 Then person("SURNAME") = Trim(SafeText(objSheet.Cells(rowIndex, surnameCol).Value))

  If Len(person("NAME")) = 0 Or Len(person("SURNAME")) = 0 Then ApplyRotaFullNameFallback person, fullName

  Set BuildRotaPersonFromWorksheetRow = person
End Function

Sub CloseExcelFallback(ByRef objExcel, ByRef localWorkbook, ByRef remoteWorkbook, ByRef localSheet, ByRef remoteSheet, saveRemote)
  On Error Resume Next
  Set localSheet = Nothing
  Set remoteSheet = Nothing
  If Not (remoteWorkbook Is Nothing) Then remoteWorkbook.Close CBool(saveRemote)
  If Not (localWorkbook Is Nothing) Then localWorkbook.Close False
  If Not (objExcel Is Nothing) Then objExcel.Quit
  Set remoteWorkbook = Nothing
  Set localWorkbook = Nothing
  Set objExcel = Nothing
  On Error GoTo 0
End Sub

Function OpenMeetingRotaWorkbookForEdit(ByRef objExcel, ByRef objWorkbook, ByRef objSheetEmails)
  OpenMeetingRotaWorkbookForEdit = False

  On Error Resume Next
  Set objExcel = CreateObject("Excel.Application")
  If Err.Number <> 0 Then
    LogEvent "ERROR: Failed to create Excel.Application for rota edit: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  objExcel.Visible = False
  objExcel.DisplayAlerts = False

  Set objWorkbook = objExcel.Workbooks.Open(ROTA_LOCAL_WORKBOOK, 0, False)
  If Err.Number <> 0 Or objWorkbook Is Nothing Then
    LogEvent "ERROR: Failed to open meeting rota workbook for edit: " & Err.Description
    WScript.Echo "Excel could not open the local meeting rota workbook."
    Err.Clear
    objExcel.Quit
    Set objExcel = Nothing
    On Error GoTo 0
    Exit Function
  End If

  If objWorkbook.ReadOnly Then
    LogEvent "ERROR: Meeting rota workbook opened read-only."
    WScript.Echo "The meeting rota workbook opened read-only. Please close it elsewhere and try again."
    objWorkbook.Close False
    objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0
    Exit Function
  End If

  Set objSheetEmails = objWorkbook.Sheets(ROTA_SHEET_EMAILS)
  If Err.Number <> 0 Or objSheetEmails Is Nothing Then
    LogEvent "ERROR: Failed to open '" & ROTA_SHEET_EMAILS & "' sheet: " & Err.Description
    Err.Clear
    objWorkbook.Close False
    objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  LogEvent "Meeting rota workbook opened for edit."
  OpenMeetingRotaWorkbookForEdit = True
End Function

Sub CloseMeetingRotaWorkbookForEdit(ByRef objExcel, ByRef objWorkbook, ByRef objSheetEmails, saveChanges)
  On Error Resume Next
  Set objSheetEmails = Nothing
  If Not (objWorkbook Is Nothing) Then objWorkbook.Close CBool(saveChanges)
  If Not (objExcel Is Nothing) Then objExcel.Quit
  Set objWorkbook = Nothing
  Set objExcel = Nothing
  On Error GoTo 0
End Sub

Function BuildRotaPersonContext(entry)
  Dim context: Set context = CreateCaseInsensitiveDictionary()
  context("NAME") = ""
  context("SURNAME") = ""
  context("FULL_NAME") = ""
  context("EMAIL") = ""
  context("DIRECTORY_SOURCE") = ""
  context("DIRECTORY_ATTEMPTED") = ""

  Dim rawValue: rawValue = SafeText(GetDictionaryValue(entry, "RawValue"))
  Dim address: address = SafeText(GetDictionaryValue(entry, "PreferredAddress"))
  If Len(address) = 0 Then address = rawValue

  If entry.Exists("ParsedDisplay") Then
    ApplyRotaFullNameFallback context, SafeText(entry("ParsedDisplay"))
  End If
  If entry.Exists("ParsedEmail") Then
    context("EMAIL") = NormalizeEmailText(entry("ParsedEmail"))
  End If

  Dim nameInfo: Set nameInfo = Nothing
  If entry.Exists("Resolution") Then Set nameInfo = entry("Resolution")
  If nameInfo Is Nothing Then Set nameInfo = CreateEmptyResolutionInfo()

  If nameInfo.Exists("First") Then context("NAME") = SafeText(nameInfo("First"))
  If nameInfo.Exists("Last") Then context("SURNAME") = SafeText(nameInfo("Last"))
  If nameInfo.Exists("Full") Then
    Dim resolvedFullName: resolvedFullName = SafeText(nameInfo("Full"))
    If IsUsableRotaFullName(resolvedFullName) Then context("FULL_NAME") = resolvedFullName
  End If
  If nameInfo.Exists("Email") Then
    If Len(NormalizeEmailText(nameInfo("Email"))) > 0 Then context("EMAIL") = NormalizeEmailText(nameInfo("Email"))
  End If
  If nameInfo.Exists("DirectorySource") Then context("DIRECTORY_SOURCE") = SafeText(nameInfo("DirectorySource"))
  If nameInfo.Exists("DirectoryAttempted") Then context("DIRECTORY_ATTEMPTED") = SafeText(nameInfo("DirectoryAttempted"))

  If Not IsUsableRotaFullName(context("FULL_NAME")) Then context("FULL_NAME") = ""
  If Len(SafeText(context("FULL_NAME"))) = 0 Then ApplyRotaFullNameFallback context, rawValue

  Call NormalizeNameParts(context)
  If Not IsUsableRotaFullName(context("FULL_NAME")) Then context("FULL_NAME") = ""
  If Len(SafeText(context("FULL_NAME"))) = 0 Then
    context("FULL_NAME") = Trim(SafeText(context("NAME")) & " " & SafeText(context("SURNAME")))
  End If

  Call ReviewResolvedName(address, context)
  Call NormalizeNameParts(context)
  If Not IsUsableRotaFullName(context("FULL_NAME")) Then context("FULL_NAME") = ""
  If Len(SafeText(context("FULL_NAME"))) = 0 Then
    context("FULL_NAME") = Trim(SafeText(context("NAME")) & " " & SafeText(context("SURNAME")))
  End If

  Dim emailValue: emailValue = NormalizeEmailText(context("EMAIL"))
  If Not IsValidEmail(emailValue) Then
    Dim promptName: promptName = SafeText(context("FULL_NAME"))
    If Len(promptName) = 0 Then promptName = rawValue
    emailValue = PromptLine("    Enter email address for " & promptName & " or press Enter to skip: ")
    emailValue = NormalizeEmailText(emailValue)
  End If

  If Not IsValidEmail(emailValue) Then
    WScript.Echo "  Skipping rota entry because no valid email address was supplied."
    LogEvent "Rota people update skipped: invalid email for raw value '" & rawValue & "'"
    Set BuildRotaPersonContext = Nothing
    Exit Function
  End If
  context("EMAIL") = emailValue

  If Len(SafeText(context("FULL_NAME"))) = 0 Then
    WScript.Echo "  Skipping rota entry because no name was supplied."
    LogEvent "Rota people update skipped: missing name for email '" & emailValue & "'"
    Set BuildRotaPersonContext = Nothing
    Exit Function
  End If

  Set BuildRotaPersonContext = context
End Function

Sub ApplyRotaFullNameFallback(context, value)
  Dim fullName: fullName = Trim(SafeText(value))
  If Len(fullName) = 0 Then Exit Sub
  If InStr(fullName, "@") > 0 Then Exit Sub

  If Len(SafeText(context("FULL_NAME"))) = 0 Then context("FULL_NAME") = fullName

  If Len(SafeText(context("NAME"))) > 0 And Len(SafeText(context("SURNAME"))) > 0 Then Exit Sub

  Dim firstName: firstName = ""
  Dim surname: surname = ""
  If InStr(fullName, ",") > 0 Then
    Dim commaParts: commaParts = Split(fullName, ",")
    surname = Trim(commaParts(0))
    If UBound(commaParts) >= 1 Then firstName = Trim(Split(Trim(commaParts(1)), " ")(0))
  Else
    Dim bits: bits = Split(fullName, " ")
    If UBound(bits) >= 0 Then firstName = Trim(bits(0))
    If UBound(bits) >= 1 Then surname = Trim(bits(UBound(bits)))
  End If

  If Len(SafeText(context("NAME"))) = 0 Then context("NAME") = firstName
  If Len(SafeText(context("SURNAME"))) = 0 Then context("SURNAME") = surname
End Sub

Function ConfirmRotaPerson(person)
  WScript.Echo "Rota people-sheet row:"
  WScript.Echo "  Name:  " & GetDictionaryValue(person, "FULL_NAME")
  WScript.Echo "  Email: " & GetDictionaryValue(person, "EMAIL")
  WScript.Echo "  Presenter: 1 | Lab Worker: 1 | Past member: 0"

  Dim response: response = PromptLine("Add this person to the meeting rota people sheet? (Y/n): ")
  response = LCase(Trim(response))
  ConfirmRotaPerson = (Len(response) = 0 Or Left(response, 1) = "y")
End Function

Function IsUsableRotaFullName(value)
  Dim text: text = Trim(SafeText(value))
  If Len(text) = 0 Then
    IsUsableRotaFullName = False
  ElseIf InStr(text, "@") > 0 Then
    IsUsableRotaFullName = False
  Else
    IsUsableRotaFullName = True
  End If
End Function

Function AppendRotaPersonIfMissing(objSheetEmails, person)
  AppendRotaPersonIfMissing = "skipped"

  Dim fullName: fullName = Trim(SafeText(GetDictionaryValue(person, "FULL_NAME")))
  Dim firstName: firstName = Trim(SafeText(GetDictionaryValue(person, "NAME")))
  Dim surname: surname = Trim(SafeText(GetDictionaryValue(person, "SURNAME")))
  Dim emailValue: emailValue = NormalizeEmailText(GetDictionaryValue(person, "EMAIL"))

  If Len(fullName) = 0 Or Not IsValidEmail(emailValue) Then
    LogEvent "Rota people row skipped because name or email was missing."
    Exit Function
  End If

  Dim duplicateRow: duplicateRow = FindRotaPeopleDuplicateRow(objSheetEmails, fullName, emailValue)
  If duplicateRow > 0 Then
    WScript.Echo "  Already present in the people sheet at row " & duplicateRow & ": " & fullName
    LogEvent "Rota people duplicate skipped at row " & duplicateRow & ": " & fullName & " <" & emailValue & ">"
    AppendRotaPersonIfMissing = "duplicate"
    Exit Function
  End If

  Dim table: Set table = GetRotaPeopleTable(objSheetEmails)
  If Not (table Is Nothing) Then
    Dim listRow: Set listRow = table.ListRows.Add
    SetRotaTableCell table, listRow.Range, "Name", 1, fullName
    SetRotaTableCell table, listRow.Range, "First name", 0, firstName
    SetRotaTableCell table, listRow.Range, "Surname", 0, surname
    SetRotaTableCell table, listRow.Range, "Email", 2, emailValue
    SetRotaTableCell table, listRow.Range, "Title", 3, ""
    SetRotaTableCell table, listRow.Range, "Degree pursued", 4, ""
    SetRotaTableCell table, listRow.Range, "Presenter", 5, 1
    SetRotaTableCell table, listRow.Range, "Lab Worker", 6, 1
    SetRotaTableCell table, listRow.Range, "Past member", 7, 0
  Else
    Dim insertRow: insertRow = DetermineRotaPeopleInsertRow(objSheetEmails)
    objSheetEmails.Rows(insertRow).Insert
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Name", 1, fullName
    SetWorksheetCellByHeader objSheetEmails, insertRow, "First name", 0, firstName
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Surname", 0, surname
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Email", 2, emailValue
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Title", 3, ""
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Degree pursued", 4, ""
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Presenter", 5, 1
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Lab Worker", 6, 1
    SetWorksheetCellByHeader objSheetEmails, insertRow, "Past member", 7, 0
  End If

  WScript.Echo "  Added to people sheet: " & fullName
  LogEvent "Added rota people row: " & fullName & " <" & emailValue & ">"
  AppendRotaPersonIfMissing = "added"
End Function

Function GetRotaPeopleTable(objSheetEmails)
  Dim table: Set table = Nothing
  On Error Resume Next
  Set table = objSheetEmails.ListObjects(ROTA_PEOPLE_TABLE)
  If Err.Number <> 0 Then
    Err.Clear
    Set table = Nothing
  End If
  If table Is Nothing Then
    If objSheetEmails.ListObjects.Count > 0 Then Set table = objSheetEmails.ListObjects(1)
  End If
  On Error GoTo 0
  Set GetRotaPeopleTable = table
End Function

Sub SetRotaTableCell(table, rowRange, headerText, defaultColumn, value)
  Dim col: col = FindRotaTableColumn(table, headerText, defaultColumn)
  If col <= 0 Then Exit Sub
  rowRange.Cells(1, col).Value = value
End Sub

Function FindRotaTableColumn(table, headerText, defaultColumn)
  FindRotaTableColumn = 0
  If table Is Nothing Then Exit Function

  Dim i
  For i = 1 To table.ListColumns.Count
    If LCase(Trim(CStr(table.ListColumns(i).Name))) = LCase(Trim(CStr(headerText))) Then
      FindRotaTableColumn = i
      Exit Function
    End If
  Next

  If IsNumeric(defaultColumn) Then
    If CLng(defaultColumn) > 0 And CLng(defaultColumn) <= table.ListColumns.Count Then FindRotaTableColumn = CLng(defaultColumn)
  End If
End Function

Sub SetWorksheetCellByHeader(objSheet, rowIndex, headerText, defaultColumn, value)
  Dim col: col = FindWorksheetHeaderColumn(objSheet, headerText, defaultColumn)
  If col <= 0 Then Exit Sub
  objSheet.Cells(rowIndex, col).Value = value
End Sub

Function FindWorksheetHeaderColumn(objSheet, headerText, defaultColumn)
  FindWorksheetHeaderColumn = 0
  Dim lastColumn: lastColumn = objSheet.Cells(1, objSheet.Columns.Count).End(-4159).Column
  Dim col
  For col = 1 To lastColumn
    If LCase(Trim(CStr(objSheet.Cells(1, col).Value))) = LCase(Trim(CStr(headerText))) Then
      FindWorksheetHeaderColumn = col
      Exit Function
    End If
  Next
  If IsNumeric(defaultColumn) Then FindWorksheetHeaderColumn = CLng(defaultColumn)
End Function

Function DetermineRotaPeopleInsertRow(objSheetEmails)
  Dim lastRow: lastRow = objSheetEmails.Cells(objSheetEmails.Rows.Count, "A").End(-4162).Row
  If lastRow < 2 Then
    DetermineRotaPeopleInsertRow = 2
    Exit Function
  End If

  Dim presenterCol: presenterCol = FindWorksheetHeaderColumn(objSheetEmails, "Presenter", 5)
  Dim formulaText: formulaText = ""
  On Error Resume Next
  formulaText = CStr(objSheetEmails.Cells(lastRow, presenterCol).Formula)
  If Err.Number <> 0 Then
    Err.Clear
    formulaText = ""
  End If
  On Error GoTo 0

  If InStr(1, formulaText, "SUBTOTAL", vbTextCompare) > 0 Then
    DetermineRotaPeopleInsertRow = lastRow
  Else
    DetermineRotaPeopleInsertRow = lastRow + 1
  End If
End Function

Function FindRotaPeopleDuplicateRow(objSheetEmails, fullName, emailValue)
  FindRotaPeopleDuplicateRow = 0

  Dim nameCol: nameCol = FindWorksheetHeaderColumn(objSheetEmails, "Name", 1)
  Dim emailCol: emailCol = FindWorksheetHeaderColumn(objSheetEmails, "Email", 2)
  Dim lastDataRow: lastDataRow = GetRotaPeopleLastDataRow(objSheetEmails)
  Dim rowIndex

  For rowIndex = 2 To lastDataRow
    Dim existingEmail: existingEmail = NormalizeEmailText(objSheetEmails.Cells(rowIndex, emailCol).Value)
    Dim existingName: existingName = Trim(SafeText(objSheetEmails.Cells(rowIndex, nameCol).Value))
    If Len(existingEmail) > 0 Then
      If LCase(existingEmail) = LCase(emailValue) Then
        FindRotaPeopleDuplicateRow = rowIndex
        Exit Function
      End If
    End If
    If Len(existingName) > 0 Then
      If LCase(existingName) = LCase(fullName) Then
        FindRotaPeopleDuplicateRow = rowIndex
        Exit Function
      End If
    End If
  Next
End Function

Function GetRotaPeopleLastDataRow(objSheetEmails)
  Dim table: Set table = GetRotaPeopleTable(objSheetEmails)
  If Not (table Is Nothing) Then
    If table.ListRows.Count > 0 Then
      GetRotaPeopleLastDataRow = table.DataBodyRange.Row + table.DataBodyRange.Rows.Count - 1
      Exit Function
    End If
  End If

  Dim lastRow: lastRow = objSheetEmails.Cells(objSheetEmails.Rows.Count, "A").End(-4162).Row
  Dim presenterCol: presenterCol = FindWorksheetHeaderColumn(objSheetEmails, "Presenter", 5)
  Dim formulaText: formulaText = ""
  On Error Resume Next
  formulaText = CStr(objSheetEmails.Cells(lastRow, presenterCol).Formula)
  If Err.Number <> 0 Then
    Err.Clear
    formulaText = ""
  End If
  On Error GoTo 0
  If InStr(1, formulaText, "SUBTOTAL", vbTextCompare) > 0 Then lastRow = lastRow - 1
  If lastRow < 1 Then lastRow = 1
  GetRotaPeopleLastDataRow = lastRow
End Function

Sub CleanupLocalRotaWorkbook(shouldCleanup)
  If Not shouldCleanup Then Exit Sub
  On Error Resume Next
  If fso.FileExists(ROTA_LOCAL_WORKBOOK) Then fso.DeleteFile ROTA_LOCAL_WORKBOOK, True
  If Err.Number <> 0 Then
    LogEvent "WARNING: Could not delete local rota workbook after upload: " & Err.Description
    Err.Clear
  Else
    LogEvent "Deleted local rota workbook after successful upload."
  End If
  On Error GoTo 0
End Sub

Function QuoteRotaArg(value)
  QuoteRotaArg = Chr(34) & CStr(value) & Chr(34)
End Function

'======================================================================
' Template discovery and parsing
'======================================================================
Function LoadTemplates(rootFolder)
  Dim items: Set items = CreateCaseInsensitiveDictionary()
  If Not fso.FolderExists(rootFolder) Then
    Set LoadTemplates = items
    Exit Function
  End If
  Dim counter: counter = 0
  Call CollectTemplates(rootFolder, items, counter)
  Set LoadTemplates = items
End Function

Sub CollectTemplates(folderPath, dict, ByRef counter)
  Dim folder: Set folder = fso.GetFolder(folderPath)
  Dim file
  For Each file In folder.Files
    Dim ext: ext = LCase(fso.GetExtensionName(file.Name))
    If ext = "htm" Or ext = "html" Or ext = "txt" Then
      Dim tmpl: Set tmpl = ParseTemplate(file.Path)
      If Not tmpl Is Nothing Then
        counter = counter + 1
        dict.Add CStr(counter), tmpl
      End If
    End If
  Next
  Dim subFolder
  For Each subFolder In folder.SubFolders
    CollectTemplates subFolder.Path, dict, counter
  Next
End Sub

Function ParseTemplate(filePath)
  Dim resultText: resultText = ReadFileSmart(filePath)
  If Len(resultText) = 0 Then
    LogEvent "Skipped empty template: " & filePath
    Set ParseTemplate = Nothing
    Exit Function
  End If

  Dim pieces: pieces = ExtractHeaderAndBody(resultText)
  Dim metaText, bodyText
  metaText = NormalizeLineEndings(pieces(0))
  bodyText = TrimLeadingBlankLines(pieces(1))

  Dim meta: Set meta = ParseHeaderBlock(metaText)

  Dim template: Set template = CreateObject("Scripting.Dictionary")
  template("Path") = filePath
  template("Folder") = fso.GetParentFolderName(filePath)
  template("FileName") = fso.GetFileName(filePath)
  Set template("Meta") = meta
  template("Body") = bodyText
  template("DisplayName") = fso.GetFileName(filePath)
  template("Stage") = LCase(GetDictionaryValue(meta, "stage"))
  template("Code") = GetDictionaryValue(meta, "code")
  Set ParseTemplate = template
End Function

Function ExtractHeaderAndBody(resultText)
  Dim headerText: headerText = ""
  Dim bodyText: bodyText = resultText

  Dim startPos, endPos

  startPos = InStr(1, resultText, HTML_HEADER_PREFIX, vbTextCompare)
  If startPos > 0 Then
    endPos = InStr(startPos + Len(HTML_HEADER_PREFIX), resultText, HTML_HEADER_SUFFIX)
    If endPos > 0 Then
      headerText = Mid(resultText, startPos + Len(HTML_HEADER_PREFIX), endPos - startPos - Len(HTML_HEADER_PREFIX))
      bodyText = Mid(resultText, endPos + Len(HTML_HEADER_SUFFIX))
      ExtractHeaderAndBody = Array(headerText, bodyText)
      Exit Function
    End If
  End If

  startPos = InStr(1, resultText, HEADER_PREFIX, vbTextCompare)
  If startPos > 0 Then
    endPos = InStr(startPos + Len(HEADER_PREFIX), resultText, HEADER_SUFFIX)
    If endPos > 0 Then
      headerText = Mid(resultText, startPos + Len(HEADER_PREFIX), endPos - startPos - Len(HEADER_PREFIX))
      bodyText = Mid(resultText, endPos + Len(HEADER_SUFFIX))
    End If
  End If

  ExtractHeaderAndBody = Array(headerText, bodyText)
End Function

Function ParseHeaderBlock(blockText)
  Dim dict: Set dict = CreateCaseInsensitiveDictionary()
  If Len(Trim(blockText)) = 0 Then
    Set ParseHeaderBlock = dict
    Exit Function
  End If

  Dim lines: lines = Split(blockText, vbLf)
  Dim line, pos, key, value
  For Each line In lines
    pos = InStr(line, ":")
    If pos > 0 Then
      key = Trim(Left(line, pos - 1))
      value = Trim(Mid(line, pos + 1))
      If Len(key) > 0 Then
        dict(LCase(key)) = value
      End If
    End If
  Next

  Set ParseHeaderBlock = dict
End Function
'======================================================================
' Template selection
'======================================================================
Function BuildTemplateLabel(tmpl)
  Dim stage: stage = tmpl("Stage")
  Dim code: code = tmpl("Code")
  Dim label: label = tmpl("FileName")
  Dim prefix: prefix = ""
  If Len(stage) > 0 Then prefix = Capitalize(stage)
  If Len(code) > 0 Then
    If Len(prefix) > 0 Then
      prefix = prefix & " - " & code
    Else
      prefix = code
    End If
  End If
  If Len(prefix) > 0 Then
    BuildTemplateLabel = prefix & " | " & label
  Else
    BuildTemplateLabel = label
  End If
End Function

Function BuildStageIndex(templates)
  Dim stageMap: Set stageMap = CreateCaseInsensitiveDictionary()
  Dim key, template, stageName
  For Each key In templates.Keys
    Set template = templates(key)
    stageName = template("Stage")
    If Len(stageName) = 0 Then stageName = "(unspecified)"
    If Not stageMap.Exists(stageName) Then
      Dim list: Set list = CreateObject("System.Collections.ArrayList")
      Set stageMap(stageName) = list
    End If
    stageMap(stageName).Add template
  Next
  Set BuildStageIndex = stageMap
End Function

Function SelectTemplatesByStage(templates)
  Dim stageMap: Set stageMap = BuildStageIndex(templates)
  Dim orderedStages: orderedStages = Array("onboarding", "scheduling", "authorisation", "finalisation")

  If stageMap.Count = 0 Then
    WScript.Echo "No stage information available in templates."
    Set SelectTemplatesByStage = Nothing
    Exit Function
  End If

  Dim displayList: Set displayList = CreateObject("System.Collections.ArrayList")
  Dim stageName
  For Each stageName In orderedStages
    If stageMap.Exists(stageName) Then displayList.Add stageName
  Next
  For Each stageName In stageMap.Keys
    If Not ContainsInArrayList(displayList, stageName) Then displayList.Add stageName
  Next

  WScript.Echo "Available stages:"
  Dim i
  For i = 0 To displayList.Count - 1
    stageName = displayList(i)
    WScript.Echo "  " & (i + 1) & ") " & Capitalize(stageName) & " (" & stageMap(stageName).Count & " template(s))"
  Next

  Dim selection: Set selection = Nothing
  Do
    Dim choice: choice = PromptLine("Select a stage by number or press Enter to cancel: ")
    If Len(choice) = 0 Then
      Exit Do
    End If
    If IsNumeric(choice) Then
      i = CInt(choice) - 1
      If i >= 0 And i < displayList.Count Then
        stageName = displayList(i)
        WScript.Echo "Stage '" & Capitalize(stageName) & "' selected."
        Set selection = BuildOrderedDictionary(stageMap(stageName))
        Exit Do
      End If
    End If
    WScript.Echo "Invalid selection."
  Loop

  Set SelectTemplatesByStage = selection
End Function

Function SelectTemplatesIndividually(templates)
  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  Dim key
  For Each key In templates.Keys
    list.Add templates(key)
  Next

  Dim ordered: Set ordered = SortTemplateList(list)
  Dim selectionResult: Set selectionResult = Nothing

  If ordered.Count = 0 Then
    Set SelectTemplatesIndividually = Nothing
    Exit Function
  End If

  WScript.Echo "Available templates:"
  Dim i
  For i = 0 To ordered.Count - 1
    WScript.Echo "  " & (i + 1) & ") " & BuildTemplateLabel(ordered(i))
  Next

  Do
    Dim response: response = PromptLine("Select template numbers (comma-separated) or press Enter to cancel: ")
    If Len(response) = 0 Then
      Exit Do
    End If

    Dim selections: Set selections = ParseNumberList(response)
    If selections.Count = 0 Then
      WScript.Echo "No valid selections detected."
    Else
      Dim dict: Set dict = CreateObject("Scripting.Dictionary")
      Dim invalid: invalid = False
      Dim number
      For Each number In selections.Keys
        i = CInt(number) - 1
        If i < 0 Or i >= ordered.Count Then
          invalid = True
          Exit For
        End If
      Next

      If invalid Then
        WScript.Echo "One or more selections were out of range."
      Else
        Dim orderedKeys: Set orderedKeys = CreateObject("System.Collections.ArrayList")
        For Each number In selections.Keys
          orderedKeys.Add CInt(number)
        Next
        orderedKeys.Sort

        Dim idx
        For i = 0 To orderedKeys.Count - 1
          idx = orderedKeys(i) - 1
          dict.Add CStr(dict.Count + 1), ordered(idx)
        Next

        Set selectionResult = dict
        Exit Do
      End If
    End If
  Loop

  Set SelectTemplatesIndividually = selectionResult
End Function

'======================================================================
' Delivery helpers
'======================================================================
Function ResolveDeliveryMode(tmpl, meta)
  Dim delivery: delivery = LCase(GetDictionaryValue(meta, "delivery"))
  If delivery = DELIVERY_BROADCAST Or delivery = DELIVERY_PERSONAL Then
    ResolveDeliveryMode = delivery
    Exit Function
  End If

  Dim inferred: inferred = DELIVERY_PERSONAL
  Dim toField: toField = LCase(GetDictionaryValue(meta, "to"))

  If InStr(toField, "[targets]") > 0 Then
    inferred = DELIVERY_BROADCAST
  ElseIf InStr(toField, "[email]") > 0 Then
    inferred = DELIVERY_PERSONAL
  ElseIf InStr(toField, "[name]") > 0 Or InStr(toField, "[surname]") > 0 Or InStr(toField, "[full_name]") > 0 Then
    inferred = DELIVERY_PERSONAL
  ElseIf Len(Trim(toField)) = 0 Then
    inferred = DELIVERY_PERSONAL
  Else
    inferred = DELIVERY_BROADCAST
  End If

  LogEvent "Delivery mode not specified for '" & tmpl("DisplayName") & "'. Defaulting to " & inferred
  WScript.Echo "  Delivery mode not specified. Defaulting to " & inferred & "."

  ResolveDeliveryMode = inferred
End Function

'======================================================================
' Template processing
'======================================================================
Sub ProcessTemplate(tmpl, recipientPackage, action)
  WScript.Echo "Preparing template: " & tmpl("DisplayName")
  LogEvent "Processing template: " & tmpl("DisplayName")

  Dim meta: Set meta = CloneDictionary(tmpl("Meta"))
  Call PromptForEmptyTags(meta, tmpl("DisplayName"))
  Call NormalizeNameParts(meta)

  Dim delivery: delivery = ResolveDeliveryMode(tmpl, meta)
  meta("delivery") = delivery

  Dim subjectText: subjectText = GetDictionaryValue(meta, "subject")
  Dim toText: toText = GetDictionaryValue(meta, "to")
  Dim ccText: ccText = GetDictionaryValue(meta, "cc")
  Dim bccText: bccText = GetDictionaryValue(meta, "bcc")
  Dim attachmentsValue: attachmentsValue = GetDictionaryValue(meta, "attachments")

  Dim attachments: attachments = ResolveAttachments(attachmentsValue, tmpl("Folder"))

  Dim inlineInfo: Set inlineInfo = PrepareInlineImageSupport(tmpl("Body"), tmpl("Path"))
  Dim htmlBody: htmlBody = inlineInfo("HtmlBody")
  Dim inlineFolder: inlineFolder = inlineInfo("InlineFolder")

  Dim targets: targets = GetRecipientResolutionAddresses(recipientPackage)
  Dim entries: entries = GetRecipientResolutionEntries(recipientPackage)
  Dim recipientCount: recipientCount = GetRecipientResolutionCount(recipientPackage)

  Dim baseContext: Set baseContext = BuildContextFromMeta(meta)
  baseContext("TEMPLATE_NAME") = tmpl("DisplayName")
  baseContext("RECIPIENT_COUNT") = CStr(recipientCount)
  baseContext("TARGETS") = Join(targets, "; ")
  Call NormalizeNameParts(baseContext)

  If delivery = DELIVERY_BROADCAST Then
    HandleBroadcast tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext
  Else
    HandlePersonal tmpl, entries, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext
  End If
End Sub

Sub HandleBroadcast(tmpl, targets, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext)
  Dim joinedTargets: joinedTargets = Join(targets, "; ")
  baseContext("TARGETS") = joinedTargets

  Dim context: Set context = CloneDictionary(baseContext)
  Call EnsurePlaceholderValues(Array(subjectText, htmlBody, toText, ccText, bccText), context, tmpl("DisplayName"))
  Call NormalizeNameParts(context)

  Dim finalSubject: finalSubject = ApplyContextToText(subjectText, context)
  Dim finalBody: finalBody = ApplyContextToText(htmlBody, context)
  Dim templatedTo: templatedTo = ApplyContextToText(toText, context)
  Dim finalTo: finalTo = MergeRecipients(templatedTo, targets)
  Dim finalCc: finalCc = ApplyContextToText(ccText, context)
  Dim finalBcc: finalBcc = ApplyContextToText(bccText, context)

  finalSubject = EnsureCapitalizedNameText(finalSubject, context)
  finalBody = EnsureCapitalizedNameText(finalBody, context)
  finalTo = EnsureCapitalizedNameText(finalTo, context)
  finalCc = EnsureCapitalizedNameText(finalCc, context)
  finalBcc = EnsureCapitalizedNameText(finalBcc, context)

  LogEvent "Broadcast delivery for " & tmpl("DisplayName") & " -> " & finalTo
  WScript.Echo "  Sending broadcast to " & finalTo

  Dim ok
  ok = CreateMail(finalTo, finalCc, finalBcc, finalSubject, finalBody, attachments, action, "", "", inlineFolder, 0, logFilePath)
  If Not ok Then
    WScript.Echo "  Failed to create broadcast email for " & tmpl("DisplayName")
    LogEvent "ERROR: Broadcast delivery failed for " & tmpl("DisplayName")
  End If
End Sub

Sub HandlePersonal(tmpl, entries, action, toText, ccText, bccText, subjectText, htmlBody, attachments, inlineFolder, baseContext)
  Dim entryUpper: entryUpper = GetRecipientResolutionArrayUpperBound(entries)
  If entryUpper < 0 Then Exit Sub

  Dim i
  For i = 0 To entryUpper
    Dim entry
    Set entry = entries(i)
    If entry Is Nothing Then
      ' Skip missing entries
    Else
      Dim address: address = SafeText(GetDictionaryValue(entry, "PreferredAddress"))
      If Len(address) = 0 Then address = SafeText(GetDictionaryValue(entry, "RawValue"))

      Dim context: Set context = CloneDictionary(baseContext)
      context("EMAIL") = address

      If Not context.Exists("NAME") Then context("NAME") = ""
      If Not context.Exists("SURNAME") Then context("SURNAME") = ""
      If Not context.Exists("FULL_NAME") Then context("FULL_NAME") = ""
      If Not context.Exists("DIRECTORY_SOURCE") Then context("DIRECTORY_SOURCE") = ""
      If Not context.Exists("DIRECTORY_ATTEMPTED") Then context("DIRECTORY_ATTEMPTED") = ""

      Dim nameInfo: Set nameInfo = Nothing
      If entry.Exists("Resolution") Then Set nameInfo = entry("Resolution")
      If nameInfo Is Nothing Then Set nameInfo = CreateEmptyResolutionInfo()

      If nameInfo("Found") Then
        Dim resolvedFirst: resolvedFirst = SafeText(nameInfo("First"))
        Dim resolvedLast: resolvedLast = SafeText(nameInfo("Last"))
        Dim resolvedFull: resolvedFull = SafeText(nameInfo("Full"))
        Dim resolvedEmail: resolvedEmail = ""
        If nameInfo.Exists("Email") Then resolvedEmail = SafeText(nameInfo("Email"))

        If Len(resolvedFirst) > 0 Then
          resolvedFirst = CapitalizeNamePart(resolvedFirst)
          nameInfo("First") = resolvedFirst
          context("NAME") = resolvedFirst
        End If

        If Len(resolvedLast) > 0 Then
          resolvedLast = CapitalizeNamePart(resolvedLast)
          nameInfo("Last") = resolvedLast
          context("SURNAME") = resolvedLast
        End If

        If Len(resolvedFull) > 0 Then
          context("FULL_NAME") = resolvedFull
        End If
        If Len(resolvedEmail) > 0 Then
          context("EMAIL") = resolvedEmail
        End If
      End If

      If nameInfo.Exists("DirectorySource") Then context("DIRECTORY_SOURCE") = nameInfo("DirectorySource")
      If nameInfo.Exists("DirectoryAttempted") Then context("DIRECTORY_ATTEMPTED") = nameInfo("DirectoryAttempted")

      Call NormalizeNameParts(context)
      Call ReviewResolvedName(address, context)

      Call NormalizeNameParts(context)

      If Len(context("FULL_NAME")) = 0 Then
        context("FULL_NAME") = Trim(context("NAME") & " " & context("SURNAME"))
      End If

      Call EnsurePlaceholderValues(Array(subjectText, htmlBody, toText, ccText, bccText), context, tmpl("DisplayName"))
      Call NormalizeNameParts(context)

      Dim finalSubject: finalSubject = ApplyContextToText(subjectText, context)
      Dim finalBody: finalBody = ApplyContextToText(htmlBody, context)
      Dim templatedTo: templatedTo = ApplyContextToText(toText, context)
      Dim finalTo: finalTo = MergeRecipients(templatedTo, Array(address))
      Dim finalCc: finalCc = ApplyContextToText(ccText, context)
      Dim finalBcc: finalBcc = ApplyContextToText(bccText, context)

      finalSubject = EnsureCapitalizedNameText(finalSubject, context)
      finalBody = EnsureCapitalizedNameText(finalBody, context)
      finalTo = EnsureCapitalizedNameText(finalTo, context)
      finalCc = EnsureCapitalizedNameText(finalCc, context)
      finalBcc = EnsureCapitalizedNameText(finalBcc, context)

      LogEvent "Personal delivery for " & tmpl("DisplayName") & " -> " & finalTo
      WScript.Echo "  Sending personal email to " & finalTo

      Dim ok
      ok = CreateMail(finalTo, finalCc, finalBcc, finalSubject, finalBody, attachments, action, "", "", inlineFolder, 0, logFilePath)
      If Not ok Then
        WScript.Echo "  Failed to create personal email for " & finalTo
        LogEvent "ERROR: Personal delivery failed for " & finalTo
      End If
    End If
  Next
End Sub

Function GetRecipientResolutionArrayUpperBound(values)
  GetRecipientResolutionArrayUpperBound = -1
  If Not IsArray(values) Then Exit Function

  On Error Resume Next
  Dim upper: upper = UBound(values)
  If Err.Number <> 0 Then
    Err.Clear
  Else
    GetRecipientResolutionArrayUpperBound = upper
  End If
  On Error GoTo 0
End Function

Sub NormalizeNameParts(dict)
  If dict Is Nothing Then Exit Sub

  Call NormalizeNameKey(dict, "NAME")
  Call NormalizeNameKey(dict, "SURNAME")
End Sub

Sub NormalizeNameKey(dict, canonicalKey)
  If dict Is Nothing Then Exit Sub

  Dim actualKey: actualKey = FindDictionaryKey(dict, canonicalKey)
  If Len(actualKey) = 0 Then
    If dict.Exists(canonicalKey) Then
      actualKey = canonicalKey
    ElseIf dict.Exists(LCase(canonicalKey)) Then
      actualKey = LCase(canonicalKey)
    Else
      Exit Sub
    End If
  End If

  Dim normalized: normalized = CapitalizeNamePart(dict(actualKey))
  dict(actualKey) = normalized

  If StrComp(actualKey, canonicalKey, vbTextCompare) <> 0 Then
    dict(canonicalKey) = normalized
  End If
End Sub

Function GetNormalizedNameValue(dict, canonicalKey)
  If dict Is Nothing Then
    GetNormalizedNameValue = ""
    Exit Function
  End If

  Call NormalizeNameKey(dict, canonicalKey)

  Dim actualKey: actualKey = FindDictionaryKey(dict, canonicalKey)
  If Len(actualKey) = 0 Then
    If dict.Exists(canonicalKey) Then
      actualKey = canonicalKey
    ElseIf dict.Exists(LCase(canonicalKey)) Then
      actualKey = LCase(canonicalKey)
    Else
      GetNormalizedNameValue = ""
      Exit Function
    End If
  End If

  GetNormalizedNameValue = SafeText(dict(actualKey))
End Function

Function EnsureCapitalizedNameText(text, context)
  Dim result: result = SafeText(text)
  If Len(result) = 0 Then
    EnsureCapitalizedNameText = result
    Exit Function
  End If

  Dim keys: keys = Array("NAME", "SURNAME")
  Dim i
  For i = 0 To UBound(keys)
    Dim key: key = keys(i)
    Dim normalized: normalized = GetNormalizedNameValue(context, key)
    If Len(normalized) > 0 Then
      Dim capitalized: capitalized = CapitalizeNamePart(normalized)
      If Len(capitalized) > 0 Then
        result = ReplaceWholeWordCaseInsensitive(result, normalized, capitalized)
      End If
    End If
  Next

  EnsureCapitalizedNameText = result
End Function

Function FindDictionaryKey(dict, target)
  Dim key
  For Each key In dict.Keys
    If StrComp(CStr(key), target, vbTextCompare) = 0 Then
      FindDictionaryKey = CStr(key)
      Exit Function
    End If
  Next
  FindDictionaryKey = ""
End Function

Function CapitalizeNamePart(value)
  Dim text: text = Trim(SafeText(value))
  If Len(text) = 0 Then
    CapitalizeNamePart = ""
  Else
    Dim firstChar: firstChar = UCase(Left(text, 1))
    If Len(text) > 1 Then
      CapitalizeNamePart = firstChar & Mid(text, 2)
    Else
      CapitalizeNamePart = firstChar
    End If
  End If
End Function

Function ReplaceWholeWordCaseInsensitive(text, findValue, replacement)
  Dim source: source = SafeText(text)
  Dim target: target = SafeText(findValue)

  If Len(target) = 0 Then
    ReplaceWholeWordCaseInsensitive = source
    Exit Function
  End If

  Dim result: result = source
  Dim startPos: startPos = 1
  Dim targetLen: targetLen = Len(target)

  Do While targetLen > 0
    Dim pos: pos = InStr(startPos, result, target, vbTextCompare)
    If pos = 0 Then Exit Do

    Dim beforeChar: beforeChar = ""
    Dim afterChar: afterChar = ""

    If pos > 1 Then beforeChar = Mid(result, pos - 1, 1)
    If pos + targetLen <= Len(result) Then afterChar = Mid(result, pos + targetLen, 1)

    If IsWordBoundary(beforeChar, afterChar) Then
      result = Left(result, pos - 1) & replacement & Mid(result, pos + targetLen)
      startPos = pos + Len(replacement)
    Else
      startPos = pos + targetLen
    End If
  Loop

  ReplaceWholeWordCaseInsensitive = result
End Function

Function IsWordBoundary(beforeChar, afterChar)
  If Len(beforeChar) = 0 Or Not IsWordCharacter(beforeChar) Then
    If Len(afterChar) = 0 Or Not IsWordCharacter(afterChar) Then
      IsWordBoundary = True
      Exit Function
    End If
  End If
  IsWordBoundary = False
End Function

Function IsWordCharacter(ch)
  If Len(ch) = 0 Then
    IsWordCharacter = False
  Else
    Dim code: code = AscW(ch)
    IsWordCharacter = _
      (code >= 48 And code <= 57) _
      Or (code >= 65 And code <= 90) _
      Or (code >= 97 And code <= 122) _
      Or (code >= &H00C0 And code <= &H02AF) _
      Or (code >= &H0370 And code <= &H1FFF)
  End If
End Function

Sub ReviewResolvedName(address, context)
  Call NormalizeNameParts(context)

  Dim summary: Set summary = CreateEmptyResolutionInfo()
  summary("DirectorySource") = SafeText(GetDictionaryValue(context, "DIRECTORY_SOURCE"))
  summary("DirectoryAttempted") = SafeText(GetDictionaryValue(context, "DIRECTORY_ATTEMPTED"))
  summary("First") = SafeText(GetDictionaryValue(context, "NAME"))
  summary("Last") = SafeText(GetDictionaryValue(context, "SURNAME"))
  summary("Full") = SafeText(GetDictionaryValue(context, "FULL_NAME"))
  summary("Email") = SafeText(GetDictionaryValue(context, "EMAIL"))

  Call DisplayResolutionSummary("Resolved name for " & address & ":", summary, "  ")

  Dim currentFirst: currentFirst = summary("First")
  Dim currentLast: currentLast = summary("Last")
  Dim currentFull: currentFull = summary("Full")

  Dim requireReview: requireReview = False
  If Len(currentFirst) = 0 Or Len(currentLast) = 0 Or Len(currentFull) = 0 Then
    requireReview = True
  End If

  If requireReview Then
    Dim input
    input = PromptLine("    Enter first name or press Enter to keep current: ")
    If Len(input) > 0 Then
      context("NAME") = CapitalizeNamePart(input)
    End If

    input = PromptLine("    Enter surname or press Enter to keep current: ")
    If Len(input) > 0 Then
      context("SURNAME") = CapitalizeNamePart(input)
    End If

    Call NormalizeNameParts(context)

    Dim computedFull: computedFull = Trim(SafeText(context("NAME")) & " " & SafeText(context("SURNAME")))
    If Len(computedFull) = 0 Then computedFull = currentFull

    input = PromptLine("    Enter display name or press Enter to keep current: ")
    If Len(input) > 0 Then
      context("FULL_NAME") = input
    ElseIf Len(SafeText(context("FULL_NAME"))) = 0 Then
      context("FULL_NAME") = computedFull
    End If
  End If
End Sub
Function LoadLinkMapContext()
  Dim dict: Set dict = CreateCaseInsensitiveDictionary()
  If fso Is Nothing Then
    Set LoadLinkMapContext = dict
    Exit Function
  End If

  Dim mapPath: mapPath = fso.BuildPath(scriptFolder, LINK_MAP_FILE)
  If Len(mapPath) = 0 Or Not fso.FileExists(mapPath) Then
    LogEvent "Link map not found at " & mapPath
    Set LoadLinkMapContext = dict
    Exit Function
  End If

  On Error Resume Next
  Dim ts: Set ts = fso.OpenTextFile(mapPath, 1)
  If Err.Number <> 0 Then
    LogEvent "Unable to open link map: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Set LoadLinkMapContext = dict
    Exit Function
  End If
  On Error GoTo 0

  Do
    If ts.AtEndOfStream Then Exit Do
    Dim rawLine: rawLine = Trim(ts.ReadLine)
    If Len(rawLine) > 0 Then
      If Left(rawLine, 1) <> "#" Then
        Dim pos: pos = InStr(rawLine, "=")
        If pos > 0 Then
          Dim key: key = Trim(Left(rawLine, pos - 1))
          Dim value: value = Trim(Mid(rawLine, pos + 1))
          If Len(key) > 0 Then dict(UCase(key)) = value
        End If
      End If
    End If
  Loop
  ts.Close

  If dict.Count > 0 Then
    LogEvent "Loaded " & dict.Count & " link map entries"
  Else
    LogEvent "Link map was empty: " & mapPath
  End If

  Set LoadLinkMapContext = dict
End Function
'======================================================================
' Action helpers
'======================================================================
Function DetermineAction()
  Dim action: action = DEFAULT_ACTION
  If WScript.Arguments.Count > 0 Then
    Dim arg: arg = LCase(Trim(WScript.Arguments(0)))
    Select Case arg
      Case "display", "send", "save"
        action = arg
      Case Else
        LogEvent "Unknown action argument: " & arg & ". Using default '" & DEFAULT_ACTION & "'"
    End Select
  End If
  DetermineAction = action
End Function

'======================================================================
' Logging
'======================================================================
Sub LogEvent(message)
  On Error Resume Next
  Dim ts: Set ts = fso.OpenTextFile(logFilePath, 8, True)
  ts.WriteLine Now & "  " & message
  ts.Close
  On Error GoTo 0
End Sub
