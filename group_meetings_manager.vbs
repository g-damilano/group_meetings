Option Explicit

'==============================================================
' Global Objects and State
'==============================================================
Dim objShell, fso
Dim SCRIPT_FOLDER, LOG_FILE_PATH
Dim OL, olNS

'==============================================================
' Configuration Constants
'==============================================================
Const DEBUG_MODE = False   ' Set to True to dry-run without sending emails
Const EXTENDED_LOGGING = True
Const LOG_VARIANT = "[Main] "
Const DISPLAY_DEBUG_POPUPS = False
Const MANAGER   = "gd2613@ic.ac.uk"
Const BUILDING_ADMIN_EMAIL   = "i.gomezvilchis@imperial.ac.uk"
Const SUPERVISOR_EMAIL       = "agi@imperial.ac.uk"
Const DEFAULT_PREVIOUS_LAB_MANAGER = "Agi"

Const EMAIL_TO = "brandt-talbot-group@imperial.ac.uk"
Const EXCEL_PATH = "C:\\Users\\giaco\\OneDrive - Imperial College London\\Work\\group_meetings\\meeting_rota.xlsx"
Const SHEET_ROTA = "Rota"
Const SHEET_EMAILS = "Emails"

'==============================================================
' Logging Helpers
'==============================================================
Sub LogEvent(message)
    On Error Resume Next
    Dim ts
    Set ts = fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    ts.WriteLine Now & " - " & LOG_VARIANT & message
    ts.Close
    Set ts = Nothing
    On Error GoTo 0
End Sub

Function ShouldResetLogForNewWeek(filePath)
    On Error Resume Next
    Dim fileObj, weekDelta

    If Not fso.FileExists(filePath) Then
        ShouldResetLogForNewWeek = False
        Exit Function
    End If

    Set fileObj = fso.GetFile(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        ShouldResetLogForNewWeek = False
        Exit Function
    End If

    weekDelta = DateDiff("ww", fileObj.DateLastModified, Now, vbMonday, vbFirstFourDays)
    If Err.Number <> 0 Then
        Err.Clear
        ShouldResetLogForNewWeek = False
        Exit Function
    End If

    ShouldResetLogForNewWeek = (weekDelta > 0)
End Function

Function ResetLogIfNewWeek(filePath)
    On Error Resume Next
    Dim ts

    ResetLogIfNewWeek = False

    If Not fso.FileExists(filePath) Then
        Exit Function
    End If

    If Not ShouldResetLogForNewWeek(filePath) Then
        Exit Function
    End If

    Set ts = fso.OpenTextFile(filePath, 2, True)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    ts.Write ""
    ts.Close
    Set ts = Nothing

    ResetLogIfNewWeek = True
End Function

Sub LogError(context)
    Dim errMsg
    errMsg = context & " | Err: " & Err.Number & " - " & Err.Description
    LogEvent errMsg
    Err.Clear
End Sub

Sub LogVerbose(message)
    If EXTENDED_LOGGING Then
        LogEvent "[VERBOSE] " & message
    End If
End Sub

'==============================================================
' Environment Bootstrap
'==============================================================
Sub InitializeFileSystem()
    Call LogVerbose("InitializeFileSystem invoked")
    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    SCRIPT_FOLDER = fso.GetParentFolderName(WScript.ScriptFullName)
    LOG_FILE_PATH = fso.BuildPath(SCRIPT_FOLDER, "group_meetings_manager.log")
End Sub

Sub InitializeLogging()
    Dim logExisted, resetPerformed

    logExisted = fso.FileExists(LOG_FILE_PATH)
    resetPerformed = ResetLogIfNewWeek(LOG_FILE_PATH)

    If resetPerformed Then
        LogEvent "Weekly log reset detected. Starting new log file: " & LOG_FILE_PATH
    ElseIf logExisted Then
        LogEvent "Continuing existing weekly log: " & LOG_FILE_PATH
    Else
        LogEvent "Creating new weekly log file: " & LOG_FILE_PATH
    End If

    Call LogVerbose("InitializeLogging invoked")
    LogEvent "=== Script execution started ==="
    LogEvent "Logging to file: " & LOG_FILE_PATH
End Sub

Function SyncMeetingWorkbook()
    Call LogVerbose("SyncMeetingWorkbook invoked")
    Dim cmd, exitCode
    cmd = "C:\rclone\rclone.exe copy ""ABTsharepoint:/Group Meetings/meeting_rota.xlsx"" ""C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings"""
    LogEvent "Executing rclone command: " & cmd
    exitCode = objShell.Run(cmd, 0, True)
    LogEvent "rclone command completed with exit code: " & exitCode
    If exitCode <> 0 Then
        LogEvent "WARNING: rclone returned non-zero exit code: " & exitCode
    End If
    SyncMeetingWorkbook = (exitCode = 0)
End Function

Function InitializeOutlook()
    Call LogVerbose("InitializeOutlook invoked")
    InitializeOutlook = False

    On Error Resume Next
    Set OL = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        LogError "Failed to create Outlook.Application"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    LogEvent "Outlook application created successfully."

    On Error Resume Next
    Set olNS = OL.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        LogError "Failed to obtain Outlook namespace"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    LogEvent "Obtained Outlook namespace for MAPI."

    InitializeOutlook = True
End Function

Sub CleanupSyncedWorkbook()
    Call LogVerbose("CleanupSyncedWorkbook invoked")
    If fso.FileExists(EXCEL_PATH) Then
        LogEvent "Deleting local Excel file after sync: " & EXCEL_PATH
        fso.DeleteFile EXCEL_PATH, True
    Else
        LogEvent "Local Excel file not found during cleanup: " & EXCEL_PATH
    End If
End Sub

'==============================================================
' Outlook Lookups
'==============================================================
Function GetExactEmail(searchName)
    Call LogVerbose("GetExactEmail invoked with searchName='" & searchName & "'")
    If Trim(searchName) = "" Then
        GetExactEmail = ""
        Exit Function
    End If

    Dim olApp, ns, recip, exchUser
    Set olApp = CreateObject("Outlook.Application")
    Set ns = olApp.GetNamespace("MAPI")

    Set recip = ns.CreateRecipient(searchName)
    recip.Resolve

    If recip.Resolved Then
        On Error Resume Next
        Set exchUser = recip.AddressEntry.GetExchangeUser()
        On Error GoTo 0

        If Not exchUser Is Nothing Then
            GetExactEmail = exchUser.PrimarySmtpAddress
        Else
            GetExactEmail = recip.AddressEntry.Address
        End If
    Else
        GetExactEmail = ""
    End If
End Function

Function EmailExists(header, folderConstant, checkOnly)
    Call LogVerbose("EmailExists invoked for subject='" & header & "'")
    Dim targetFolderId, folderLabel, sortField
    Dim targetFolder, items, mail
    Dim found, checkMode

    targetFolderId = 6 ' Inbox by default
    If Not IsEmpty(folderConstant) Then
        On Error Resume Next
        targetFolderId = CLng(folderConstant)
        On Error GoTo 0
        If targetFolderId = 0 Then targetFolderId = 6
    End If

    checkMode = False
    If Not IsEmpty(checkOnly) Then
        On Error Resume Next
        checkMode = CBool(checkOnly)
        If Err.Number <> 0 Then
            checkMode = False
        End If
        On Error GoTo 0
    End If

    Select Case targetFolderId
        Case 5
            folderLabel = "Sent Items"
            sortField = "SentOn"
        Case 6
            folderLabel = "Inbox"
            sortField = "ReceivedTime"
        Case Else
            folderLabel = "folder id " & targetFolderId
            sortField = "ReceivedTime"
    End Select

    LogEvent "Checking Outlook " & folderLabel & " for email with subject: " & header

    On Error Resume Next
    Set targetFolder = olNS.GetDefaultFolder(targetFolderId)
    If Err.Number <> 0 Then
        LogError "Unable to access Outlook folder id " & targetFolderId & " for subject '" & header & "'"
        On Error GoTo 0
        EmailExists = False
        Exit Function
    End If
    On Error GoTo 0

    Set items = targetFolder.Items
    On Error Resume Next
    items.Sort sortField, True
    On Error GoTo 0

    found = False
    For Each mail In items
        If TypeName(mail) = "MailItem" Then
            If mail.Subject = header Then
                found = True
                If checkMode Then
                    LogEvent "Confirmed email with subject '" & header & "' in " & folderLabel & "."
                Else
                    LogEvent "Email with subject '" & header & "' already exists in " & folderLabel & ". Skipping related actions."
                End If
                Exit For
            End If
        End If
    Next

    If Not found Then
        If checkMode Then
            LogEvent "Email with subject '" & header & "' not yet located in " & folderLabel & " during verification."
        Else
            LogEvent "Email with subject '" & header & "' not found in " & folderLabel & " prior to action."
        End If
    End If

    LogEvent "EmailExists('" & header & "', folder=" & folderLabel & ", checkOnly=" & checkMode & ") result: " & found
    EmailExists = found
End Function

'==============================================================
' Meeting Content Builders
'==============================================================
Function DeriveMeetingLeadName(objSheetRota, meetingDate)
    Call LogVerbose("DeriveMeetingLeadName invoked")
    Dim previousLabManager, previousFriday, previousFridayRow, searchRow, lastRow, candidateDate

    previousFriday = DateAdd("d", -7, meetingDate)
    previousFridayRow = -1
    lastRow = objSheetRota.Cells(objSheetRota.Rows.Count, "A").End(-4162).Row

    For searchRow = 2 To lastRow
        candidateDate = objSheetRota.Cells(searchRow, 1).Value
        If candidateDate = previousFriday Then
            previousFridayRow = searchRow
            Exit For
        End If
    Next

    If previousFridayRow <> -1 Then
        previousLabManager = Trim(objSheetRota.Cells(previousFridayRow, 6).Value)
    Else
        previousLabManager = ""
    End If

    If previousLabManager = "" Then
        previousLabManager = DEFAULT_PREVIOUS_LAB_MANAGER
    End If

    DeriveMeetingLeadName = previousLabManager
    LogEvent "Derived meeting lead as '" & previousLabManager & "' for meeting date " & FormatDateTime(meetingDate, 2)
End Function

Function BuildMeetingSubject(meetingDate)
    Call LogVerbose("BuildMeetingSubject invoked")
    BuildMeetingSubject = "ABT group meeting @ 11am (" & FormatDateTime(meetingDate, 2) & ")"
End Function

Function GenerateEmailContent(meetingDate, row, objSheetRota, nextFridayRow)
    Call LogVerbose("GenerateEmailContent invoked for row " & row)
    Dim literature, results, alternate, room, labManager, next_lit, next_res, next_alt, content, nameParts
    Dim meetingLeadName
    literature = Trim(objSheetRota.Cells(row, 2).Value)
    results    = Trim(objSheetRota.Cells(row, 3).Value)
    alternate  = Trim(objSheetRota.Cells(row, 4).Value)
    room       = Trim(objSheetRota.Cells(row, 5).Value)
    labManager = Trim(objSheetRota.Cells(row, 6).Value)
    LogEvent "Generating main email content for date " & meetingDate & _
              " | Row: " & row & " | Lit: " & literature & _
              " | Res: " & results & " | Alt: " & alternate & _
              " | Room: " & room

    If nextFridayRow <> -1 Then
        next_lit = objSheetRota.Cells(nextFridayRow, 2).Value
        next_res = objSheetRota.Cells(nextFridayRow, 3).Value
        next_alt = objSheetRota.Cells(nextFridayRow, 4).Value
        LogEvent "Next Friday presenters found | Row: " & nextFridayRow & _
                  " | Lit: " & next_lit & " | Res: " & next_res & _
                  " | Alt: " & next_alt
    Else
        next_lit = ""
        next_res = ""
        next_alt = ""
        LogEvent "No next Friday data found. Using empty placeholders."
    End If

    nameParts = Split(olNS.Accounts.Item(1).CurrentUser, ", ")
    Dim senderFirstName
    senderFirstName = ""
    If UBound(nameParts) >= 1 Then
        senderFirstName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        senderFirstName = Trim(nameParts(0))
    End If
    If senderFirstName = "" Then
        senderFirstName = Trim(olNS.Accounts.Item(1).CurrentUser)
    End If

    meetingLeadName = DeriveMeetingLeadName(objSheetRota, meetingDate)

    content = "Hi all," & vbCrLf
    If Trim(room) <> "" Then
        content = content & "We will have our group meeting on " & meetingDate & _
                  " at 11:00 am in " & room & " in person or via MS Teams." & vbCrLf
    Else
        content = content & "We will have our group meeting on " & meetingDate & _
                  " at 11:00 am in person or via MS Teams." & vbCrLf
    End If

    If Trim(alternate) <> "" Or Trim(results) <> "" Or Trim(literature) <> "" Then
        content = content & vbCrLf & "Meeting agenda:" & vbCrLf
        If Trim(alternate) <> "" Then
            content = content & "    " & Trim(alternate) & vbCrLf
        Else
            If Trim(results) <> "" Then
                content = content & "    " & Trim(results) & " will present results." & vbCrLf
            End If
            If Trim(literature) <> "" Then
                content = content & "    " & Trim(literature) & " will present literature." & vbCrLf
            End If
        End If
    Else
        content = content & vbCrLf & "No meeting has been set for the following week." & vbCrLf
    End If

    content = content & vbCrLf & Trim(meetingLeadName) & " will lead the meeting, keep presenters time, and do the inclusion activity." & vbCrLf

    If labManager = "" Then
        labManager = "TBD"
    End If

    content = content & vbCrLf 
    content = content & "The lab manager for this week is " & Trim(labManager) & "." & vbCrLf

    If next_lit <> "" Or next_res <> "" Or next_alt <> "" Then
        content = content & vbCrLf & "Next week:" & vbCrLf
        If Trim(next_alt) <> "" Then
            content = content & "    " & Trim(next_alt) & vbCrLf
        End If
        If Trim(next_res) <> "" Then
            content = content & "    " & Trim(next_res) & " will present results." & vbCrLf
        End If
        If Trim(next_lit) <> "" Then
            content = content & "    " & Trim(next_lit) & " will present literature." & vbCrLf
        End If
    Else
        content = content & vbCrLf & "No meeting has been set for the following week." & vbCrLf
    End If

    content = content & "Upcoming weeks schedule: " & _
              "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx?web=1" & vbCrLf

    content = content & "Best," & vbCrLf & Trim(senderFirstName)

    GenerateEmailContent = content
    LogEvent "Generated email body length: " & Len(content)
End Function

Function GeneratePresenterReminderEmail(presenterName, presentationType, meetingDate)
    Call LogVerbose("GeneratePresenterReminderEmail invoked for presenter='" & presenterName & "'")
    Dim content, nameParts, trimmedType, normalizedType, description, reminderSubject, meetingDateText, typeLabel
    Dim greetingName, tokens
    presenterName = Trim(presenterName)
    presentationType = Trim(presentationType)
    LogEvent "Generating presenter reminder email | Presenter: " & presenterName & _
              " | Type: " & presentationType

    trimmedType = presentationType
    normalizedType = LCase(trimmedType)
    Select Case normalizedType
        Case "literature"
            description = "the literature session"
            typeLabel = "Literature presentation"
        Case "results"
            description = "the results update"
            typeLabel = "Results presentation"
        Case Else
            If trimmedType <> "" Then
                description = "the " & presentationType & " presentation"
                typeLabel = trimmedType & " presentation"
            Else
                description = "your presentation"
                typeLabel = "Presentation"
            End If
    End Select

    reminderSubject = LCase(typeLabel)
    meetingDateText = FormatDateTime(meetingDate, 2)
    nameParts = Split(olNS.Accounts.Item(1).CurrentUser, ", ")
    Dim senderFirstName
    senderFirstName = ""
    If UBound(nameParts) >= 1 Then
        senderFirstName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        senderFirstName = Trim(nameParts(0))
    End If
    If senderFirstName = "" Then
        senderFirstName = Trim(olNS.Accounts.Item(1).CurrentUser)
    End If

    If presenterName = "" Then
        content = "Hi," & vbCrLf & vbCrLf
        content = content & "No presenter is currently assigned for the " & reminderSubject & _
                  " scheduled on " & meetingDateText & "." & vbCrLf & vbCrLf
        content = content & "Please update the rota or advise the group so we can keep the agenda on track." & vbCrLf & vbCrLf
        content = content & "Best," & vbCrLf & Trim(senderFirstName)
        GeneratePresenterReminderEmail = content
        LogEvent "Generated presenter reminder length: " & Len(content)
        Exit Function
    End If

    greetingName = presenterName
    If InStr(greetingName, ",") > 0 Then
        tokens = Split(greetingName, ",")
        If UBound(tokens) >= 1 Then
            greetingName = Trim(tokens(1))
        ElseIf UBound(tokens) = 0 Then
            greetingName = Trim(tokens(0))
        End If
    Else
        tokens = Split(greetingName, " ")
        If UBound(tokens) >= 0 Then
            greetingName = Trim(tokens(0))
        End If
    End If

    Dim enDash, apostrophe, bulletChar
    enDash = ChrW(&H2013)
    apostrophe = Chr(39)
    bulletChar = ChrW(&H2022)

    content = "Hi " & Trim(greetingName) & "," & vbCrLf & vbCrLf
    content = content & "Just a reminder that you are scheduled to present " & description & _
              " at this week's group meeting on " & meetingDateText & ". The slot is 15 minutes for your presentation plus 5 minutes for Q&A. " & _
              "As a rough guide, aim for about 1 minute per slide (so around 10" & enDash & "12 slides in total). " & _
              "The group meeting leader will keep time to ensure we don" & apostrophe & "t overrun, so please rehearse your talk beforehand to check that you are within the time limit." & vbCrLf & vbCrLf
    content = content & "A few pointers for preparing your slides:" & vbCrLf & vbCrLf
    content = content & bulletChar & " Aim for a balanced, well-planned presentation, rather than trying to include everything." & vbCrLf & vbCrLf
    content = content & bulletChar & " Use no more than ~3 concise bullet points per slide to support (not duplicate) what you say, so the audience focuses on you rather than reading the screen." & vbCrLf & vbCrLf
    content = content & bulletChar & " Let the slides support your performance, not be the performance" & enDash & "they shouldn" & apostrophe & "t steal your stage." & vbCrLf & vbCrLf
    content = content & bulletChar & " One clear, well-chosen image is often more effective than a lot of text, especially as many scientists are visual learners." & vbCrLf & vbCrLf
    content = content & bulletChar & " Assume a left-to-right, top-to-bottom reading pattern: place the most important information at the top left and less critical details toward the bottom right." & vbCrLf & vbCrLf
    content = content & "Please ensure your materials are ready and let me know if you need any support." & vbCrLf & vbCrLf
    content = content & "Best," & vbCrLf & Trim(senderFirstName)
    GeneratePresenterReminderEmail = content
    LogEvent "Generated presenter reminder length: " & Len(content)
End Function

Function GenerateMeetingLeadReminderEmail(meetingLeadName, meetingDate)
    Call LogVerbose("GenerateMeetingLeadReminderEmail invoked for lead='" & meetingLeadName & "'")
    Dim content, greetingName, tokens, meetingDateText
    Dim nameParts, senderFirstName

    meetingLeadName = Trim(meetingLeadName)
    meetingDateText = FormatDateTime(meetingDate, 2)

    nameParts = Split(olNS.Accounts.Item(1).CurrentUser, ", ")
    senderFirstName = ""
    If UBound(nameParts) >= 1 Then
        senderFirstName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        senderFirstName = Trim(nameParts(0))
    End If
    If senderFirstName = "" Then
        senderFirstName = Trim(olNS.Accounts.Item(1).CurrentUser)
    End If

    greetingName = meetingLeadName
    If InStr(greetingName, ",") > 0 Then
        tokens = Split(greetingName, ",")
        If UBound(tokens) >= 1 Then
            greetingName = Trim(tokens(1))
        ElseIf UBound(tokens) = 0 Then
            greetingName = Trim(tokens(0))
        End If
    Else
        tokens = Split(greetingName, " ")
        If UBound(tokens) >= 0 Then
            greetingName = Trim(tokens(0))
        End If
    End If

    If Trim(greetingName) = "" Then
        greetingName = "there"
    End If

    content = "Hi " & greetingName & "," & vbCrLf & vbCrLf
    content = content & "Just a reminder that you will lead this week's group meeting on " & meetingDateText & "." & vbCrLf & vbCrLf
    content = content & "Please arrive about 5 minutes early to set up the room so the computer is ready right at meeting time." & vbCrLf & vbCrLf
    content = content & "Bring the red/yellow timing cards (5 minutes/1 minute left) to signal time to each presenter, and keep presenters on schedule while running the inclusion activity." & vbCrLf & vbCrLf
    content = content & "Best," & vbCrLf & Trim(senderFirstName)

    GenerateMeetingLeadReminderEmail = content
    LogEvent "Generated meeting lead reminder length: " & Len(content)
End Function

Function GenerateLabManagerReminderEmail(labManagerName, meetingDate)
    Call LogVerbose("GenerateLabManagerReminderEmail invoked for manager='" & labManagerName & "'")
    Dim content, nameParts, manager
    Dim firstName, lastName, senderDisplayName, parcelAddressee
    Dim greetingName, tokens
    Dim startText, endText, weekRange
    manager = Trim(labManagerName)
    LogEvent "Generating lab manager reminder email | Manager: " & manager

    If manager = "" Then
        manager = "there"
    End If

    greetingName = manager
    If LCase(greetingName) <> "there" Then
        If InStr(greetingName, ",") > 0 Then
            tokens = Split(greetingName, ",")
            If UBound(tokens) >= 1 Then
                greetingName = Trim(tokens(1))
            ElseIf UBound(tokens) = 0 Then
                greetingName = Trim(tokens(0))
            End If
        Else
            tokens = Split(greetingName, " ")
            If UBound(tokens) >= 0 Then
                greetingName = Trim(tokens(0))
            End If
        End If

        If Trim(greetingName) = "" Then
            greetingName = manager
        End If
    End If

    nameParts = Split(olNS.Accounts.Item(1).CurrentUser, ", ")
    firstName = ""
    lastName = ""

    If UBound(nameParts) >= 1 Then
        lastName = Trim(nameParts(0))
        firstName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        firstName = Trim(nameParts(0))
    End If

    senderDisplayName = firstName
    If senderDisplayName = "" Then
        senderDisplayName = Trim(olNS.Accounts.Item(1).CurrentUser)
    End If

    parcelAddressee = senderDisplayName
    If lastName <> "" Then
        If senderDisplayName <> "" Then
            parcelAddressee = senderDisplayName & " " & lastName
        Else
            parcelAddressee = lastName
        End If
    End If

    If parcelAddressee = "" Then
        parcelAddressee = "me"
    End If

    If senderDisplayName = "" Then
        senderDisplayName = parcelAddressee
    End If

    startText = FormatDateTime(meetingDate, 2)
    endText = FormatDateTime(DateAdd("d", 7, meetingDate), 2)
    weekRange = startText & " " & ChrW(&H2013) & " " & endText

    Dim greetingDisplay
    If LCase(greetingName) = "there" Then
        greetingDisplay = greetingName
    Else
        greetingDisplay = Trim(greetingName)
    End If

    content = "Hi " & greetingDisplay & "," & vbCrLf & vbCrLf
    content = content & "This is a reminder that you are the lab manager for the week of " & weekRange & "." & vbCrLf & vbCrLf
    content = content & "Lab cleaning starts on Friday and must be completed by next Friday " & ChrW(&H2014) & _
              " please try to finish it by Friday EOD so the lab is ready for Monday." & vbCrLf & vbCrLf
    Dim parcelDisplay
    parcelDisplay = parcelAddressee
    If LCase(parcelAddressee) <> "me" Then
        parcelDisplay = Trim(parcelAddressee)
    End If

    content = content & "On friday, please pick up all parcels addressed to " & parcelDisplay & _
              "; most of them contain the consumables ordered the last week that need to be re-stocked in the lab." & vbCrLf & vbCrLf
    content = content & "Also, please procure the list of missing items on the whiteboard from the stores so that we can maintain supplies at all times." & vbCrLf & vbCrLf
    content = content & "Thank you for keeping the lab organized." & vbCrLf & vbCrLf
    content = content & "Best," & vbCrLf & Trim(senderDisplayName)
    GenerateLabManagerReminderEmail = content
    LogEvent "Generated lab manager reminder length: " & Len(content)
End Function

'==============================================================
' Reminder Workflow
'==============================================================
Function ResolveEmailOrManager(personName)
    Call LogVerbose("ResolveEmailOrManager invoked for person='" & personName & "'")
    Dim resolved
    resolved = GetExactEmail(personName)
    If Trim(resolved & "") <> "" Then
        LogEvent "Resolved exact email for '" & personName & "': " & resolved
        ResolveEmailOrManager = resolved
    Else
        LogEvent "Falling back to manager email for '" & personName & "'"
        ResolveEmailOrManager = MANAGER
    End If
End Function

Sub DispatchReminderEmail(toAddress, subjectText, bodyText)
    Call LogVerbose("DispatchReminderEmail invoked for subject='" & subjectText & "'")
    Dim mail
    LogEvent "Preparing reminder email with subject: " & subjectText & " to " & toAddress

    Set mail = OL.CreateItem(0)
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = bodyText
        If DEBUG_MODE Then
            LogEvent "DEBUG_MODE enabled - reminder email not sent. Subject: " & subjectText
            If DISPLAY_DEBUG_POPUPS Then
                MsgBox "[DEBUG] Would send reminder email to " & toAddress & ":" & vbCrLf & _
                       "Subject: " & subjectText & vbCrLf & vbCrLf & bodyText
            Else
                LogVerbose "Reminder preview suppressed for subject='" & subjectText & "'"
            End If
        Else
            LogEvent "Sending reminder email to " & toAddress & " with subject: " & subjectText
            .Send
            If EmailExists(subjectText, 5, True) Then
                LogEvent "Negative feedback loop confirmed for reminder subject: " & subjectText
            Else
                LogEvent "WARNING: Reminder subject not observed in Sent Items after sending: " & subjectText
            End If
        End If
    End With
End Sub

Function BuildPresenterReminderSubject(presentationType, meetingDate)
    Call LogVerbose("BuildPresenterReminderSubject invoked")
    Dim trimmedType, normalizedType, typeLabel
    trimmedType = Trim(presentationType)
    normalizedType = LCase(trimmedType)
    Select Case normalizedType
        Case "literature"
            typeLabel = "Literature presentation"
        Case "results"
            typeLabel = "Results presentation"
        Case Else
            If trimmedType = "" Then
                typeLabel = "Presentation"
            Else
                typeLabel = trimmedType & " presentation"
            End If
    End Select
    BuildPresenterReminderSubject = typeLabel & " reminder (" & FormatDateTime(meetingDate, 2) & ")"
End Function

Function BuildMeetingLeadReminderSubject(meetingDate)
    Call LogVerbose("BuildMeetingLeadReminderSubject invoked")
    BuildMeetingLeadReminderSubject = "Meeting lead reminder (" & FormatDateTime(meetingDate, 2) & ")"
End Function

Function BuildLabManagerReminderSubject(labManagerName, meetingDate)
    Call LogVerbose("BuildLabManagerReminderSubject invoked")
    Dim startText, endText
    startText = FormatDateTime(meetingDate, 2)
    endText = FormatDateTime(DateAdd("d", 7, meetingDate), 2)
    BuildLabManagerReminderSubject = "Lab management reminder (" & startText & " " & ChrW(&H2013) & " " & endText & ")"
End Function

Sub TriggerReminderEmails(literaturePresenter, resultsPresenter, labManagerName, meetingLeadName, meetingDate)
    Call LogVerbose("TriggerReminderEmails invoked")
    LogEvent "Triggering reminder workflow for meeting date " & FormatDateTime(meetingDate, 2)
    MaybeSendPresenterReminder literaturePresenter, "Literature", meetingDate
    MaybeSendPresenterReminder resultsPresenter, "Results", meetingDate
    MaybeSendMeetingLeadReminder meetingLeadName, meetingDate
    MaybeSendLabManagerReminder labManagerName, meetingDate
End Sub

Sub MaybeSendPresenterReminder(presenterName, presentationType, meetingDate)
    Call LogVerbose("MaybeSendPresenterReminder invoked for presenter='" & presenterName & "'")
    Dim trimmedName, subjectLine, targetEmail, body
    trimmedName = Trim(presenterName)

    targetEmail = ResolveEmailOrManager(trimmedName)
    subjectLine = BuildPresenterReminderSubject(presentationType, meetingDate)

    If EmailExists(subjectLine, 5, True) Then
        LogEvent "Skipping presenter reminder because Sent Items already contain subject '" & subjectLine & "'."
        Exit Sub
    End If

    body = GeneratePresenterReminderEmail(trimmedName, presentationType, meetingDate)
    DispatchReminderEmail targetEmail, subjectLine, body
End Sub

Sub MaybeSendMeetingLeadReminder(meetingLeadName, meetingDate)
    Call LogVerbose("MaybeSendMeetingLeadReminder invoked for meeting lead='" & meetingLeadName & "'")
    Dim trimmedName, subjectLine, targetEmail, body
    trimmedName = Trim(meetingLeadName)

    targetEmail = ResolveEmailOrManager(trimmedName)
    subjectLine = BuildMeetingLeadReminderSubject(meetingDate)

    If EmailExists(subjectLine, 5, True) Then
        LogEvent "Skipping meeting lead reminder because Sent Items already contain subject '" & subjectLine & "'."
        Exit Sub
    End If

    body = GenerateMeetingLeadReminderEmail(trimmedName, meetingDate)
    DispatchReminderEmail targetEmail, subjectLine, body
End Sub

Sub MaybeSendLabManagerReminder(labManagerName, meetingDate)
    Call LogVerbose("MaybeSendLabManagerReminder invoked for manager='" & labManagerName & "'")
    Dim trimmedName, subjectLine, targetEmail, body
    trimmedName = Trim(labManagerName)

    targetEmail = ResolveEmailOrManager(trimmedName)
    subjectLine = BuildLabManagerReminderSubject(trimmedName, meetingDate)

    If EmailExists(subjectLine, 5, True) Then
        LogEvent "Skipping lab manager reminder because Sent Items already contain subject '" & subjectLine & "'."
        Exit Sub
    End If

    body = GenerateLabManagerReminderEmail(trimmedName, meetingDate)
    DispatchReminderEmail targetEmail, subjectLine, body
End Sub

'==============================================================
' Standard Email Dispatch
'==============================================================
Sub DistributeEmail(header, body)
    Call LogVerbose("DistributeEmail invoked with subject='" & header & "'")
    Dim mail
    LogEvent "Preparing regular meeting email with subject: " & header

    If EmailExists(header, 5, True) Then
        LogEvent "Sent Items already contains the regular meeting email. Skipping send for subject: " & header
        Exit Sub
    End If

    Set mail = OL.CreateItem(0)
    With mail
        .To = EMAIL_TO
        .Subject = header
        .Body = body
        If DEBUG_MODE Then
            LogEvent "DEBUG_MODE enabled - email not sent. Subject: " & header
            If DISPLAY_DEBUG_POPUPS Then
                MsgBox "[DEBUG] Would send email:" & vbCrLf & vbCrLf & header & vbCrLf & body
            Else
                LogVerbose "Regular email preview suppressed for subject='" & header & "'"
            End If
        Else
            LogEvent "Sending email to " & EMAIL_TO & " with subject: " & header
            .Send
            If EmailExists(header, 5, True) Then
                LogEvent "Negative feedback loop confirmed for regular meeting subject: " & header
            Else
                LogEvent "WARNING: Regular meeting subject not observed in Sent Items after sending: " & header
            End If
        End If
    End With
End Sub

'==============================================================
' Alert Email Helpers
'==============================================================
Function GenerateEmailAlert(alertType, missingDate, recipientEmail, tentativeRota)
    Call LogVerbose("GenerateEmailAlert invoked for alertType='" & alertType & "'")
    Dim content
    content = "Hi," & vbCrLf & vbCrLf
    Select Case alertType
        Case "MissingDate"
            content = content & "There is no entry in the rota for " & missingDate & "." & vbCrLf & _
                      "Please update the rota accordingly." & vbCrLf & vbCrLf & tentativeRota
        Case "Unassigned"
            content = content & "The rota entry for " & missingDate & " has no assigned presenters." & vbCrLf & _
                      "Please review the tentative 6-month rota below and approve or adjust:" & vbCrLf & vbCrLf & tentativeRota
        Case "MissingRoom"
            content = content & "The rota entry for " & missingDate & " has no room reservation." & vbCrLf & _
                      "Please reserve a room for the meeting." & vbCrLf
            content = content & vbCrLf & "Send Email to Iliane to request room for the upcoming term: " & _
              "mailto:i.gomezvilchis@imperial.ac.uk?subject=Room%20Reservation%20Request%20for%20ABT%20Group%20Meetings%20%E2%80%93%20Upcoming%20Term&body=Hi%20Iliane,%0D%0A%0D%0AI%20hope%20you%20are%20well.%0D%0A%0D%0AI%20would%20like%20to%20request%20a%20recurring%20booking%20for%20our%20ABT%20group%20meetings%20for%20the%20upcoming%20academic%20term.%20Could%20you%20please%20reserve%20G25%20(MSRH)%20for%20Fridays%20from%2011:00%20to%2012:00,%20starting%20from%20the%20next%20available%20week?%0D%0A%0D%0ACould%20you%20also%20let%20me%20know%20the%20exact%20date%20the%20reservation%20ends?%0D%0A%0D%0AThank%20you%20very%20much%20for%20your%20help.%0D%0A%0D%0ABest%20regards,%0D%0AGiacomo" & vbCrLf & vbCrLf
        Case "MeetingNotSet"
            content = content & "The rota entry for " & missingDate & " has no presenters or alternate assigned, so the regular meeting email was not sent to the distribution list." & vbCrLf & _
                      "Please update the rota with presenters so the meeting can be scheduled." & vbCrLf
    End Select
    content = content & vbCrLf & "Contact directly via email: " & recipientEmail & vbCrLf & vbCrLf
    content = content & "Best," & vbCrLf & "Automated Meeting Scheduler"
    GenerateEmailAlert = content
End Function

Sub AlertEmail(toAddress, messageBody, subjectText)
    Call LogVerbose("AlertEmail invoked to='" & toAddress & "'")
    Dim mail
    LogEvent "Preparing alert email with subject: " & subjectText & " to " & toAddress
    Set mail = OL.CreateItem(0)
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = messageBody
        If DEBUG_MODE Then
            LogEvent "DEBUG_MODE enabled - alert email not sent. Subject: " & subjectText
            If DISPLAY_DEBUG_POPUPS Then
                MsgBox "[DEBUG] Would send ALERT email to " & toAddress & ":" & vbCrLf & _
                    "Subject: " & subjectText & vbCrLf & vbCrLf & messageBody
            Else
                LogVerbose "Alert email preview suppressed for subject='" & subjectText & "'"
            End If
        Else
            LogEvent "Sending alert email to " & toAddress & " with subject: " & subjectText
            .Send
        End If
    End With
End Sub

Sub AlertEmailWithAttachment(toAddress, messageBody, subjectText, attachmentPath)
    Call LogVerbose("AlertEmailWithAttachment invoked to='" & toAddress & "' with attachment='" & attachmentPath & "'")
    Dim mail
    LogEvent "Preparing alert email with attachment. Subject: " & subjectText & _
                  " | Recipient: " & toAddress & " | Attachment: " & attachmentPath
    Set mail = OL.CreateItem(0)
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = messageBody
        If attachmentPath <> "" Then
            .Attachments.Add attachmentPath
            LogEvent "Attached file to email: " & attachmentPath
        End If
        If DEBUG_MODE Then
            LogEvent "DEBUG_MODE enabled - alert email with attachment not sent. Subject: " & subjectText
            If DISPLAY_DEBUG_POPUPS Then
                MsgBox "[DEBUG] Would send ALERT email with attachment to " & toAddress & ":" & vbCrLf & _
                       "Subject: " & subjectText & vbCrLf & vbCrLf & messageBody & vbCrLf & _
                       "Attachment: " & attachmentPath
            Else
                LogVerbose "Alert email with attachment preview suppressed for subject='" & subjectText & "'"
            End If
        Else
            LogEvent "Sending alert email with attachment to " & toAddress & " with subject: " & subjectText
            .Send
        End If
    End With
End Sub

'==============================================================
' Tentative Rota Helpers
'==============================================================
Function GenerateTentativeRota(objSheetEmails)
    Call LogVerbose("GenerateTentativeRota invoked")
    Dim lastRow, i, numNames, names(), rndIndex, tmp
    Dim rotaDates(), rotaAssignments(), rotaBody, dt, totalWeeks
    Dim csvFile, txtStream
    LogEvent "Starting tentative rota generation"

    lastRow = objSheetEmails.Cells(objSheetEmails.Rows.Count, "A").End(-4162).Row
    numNames = 0
    For i = 2 To lastRow
        If objSheetEmails.Cells(i, 4).Value = 1 Then numNames = numNames + 1
    Next
    LogEvent "Eligible names counted: " & numNames
    If numNames = 0 Then
        LogEvent "No eligible names found for tentative rota."
        GenerateTentativeRota = "No eligible names for rota."
        Exit Function
    End If

    ReDim names(numNames - 1)
    numNames = 0
    For i = 2 To lastRow
        If objSheetEmails.Cells(i, 4).Value = 1 Then
            names(numNames) = Trim(objSheetEmails.Cells(i, 1).Value)
            numNames = numNames + 1
        End If
    Next
    LogEvent "Names loaded for rota: " & Join(names, ", ")

    Randomize
    For i = 0 To UBound(names)
        rndIndex = Int(Rnd * (UBound(names) + 1))
        tmp = names(i)
        names(i) = names(rndIndex)
        names(rndIndex) = tmp
    Next
    LogEvent "Names shuffled for tentative rota."

    totalWeeks = 26
    LogEvent "Generating tentative rota for weeks: " & totalWeeks
    ReDim rotaDates(totalWeeks - 1)
    ReDim rotaAssignments(totalWeeks - 1, 1)
    dt = Date + ((5 - Weekday(Date, vbMonday) + 7) Mod 7)

    For i = 0 To totalWeeks - 1
        rotaDates(i) = DateAdd("ww", i, dt)
        rotaAssignments(i, 0) = names((i * 2) Mod (UBound(names) + 1))
        rotaAssignments(i, 1) = names((i * 2 + 1) Mod (UBound(names) + 1))
    Next

    rotaBody = "Tentative 6-month rota (Fridays, 2 presenters/day):" & vbCrLf & vbCrLf
    rotaBody = rotaBody & "Date         | Presenter 1          | Presenter 2" & vbCrLf
    rotaBody = rotaBody & String(50, "-") & vbCrLf
    For i = 0 To totalWeeks - 1
        rotaBody = rotaBody & FormatDateTime(rotaDates(i), 2) & " | " & _
                   PadRight(rotaAssignments(i, 0), 20) & " | " & _
                   PadRight(rotaAssignments(i, 1), 20) & vbCrLf
    Next

    csvFile = fso.GetSpecialFolder(2) & "\TentativeRota.csv"
    LogEvent "Writing tentative rota CSV to: " & csvFile
    Set txtStream = fso.CreateTextFile(csvFile, True)
    txtStream.WriteLine "Date,Presenter1,Presenter2"
    For i = 0 To totalWeeks - 1
        txtStream.WriteLine FormatDateTime(rotaDates(i), 2) & "," & _
                            rotaAssignments(i, 0) & "," & rotaAssignments(i, 1)
    Next
    txtStream.Close

    GenerateTentativeRota = rotaBody & vbCrLf & "[CSV: " & csvFile & "]"
    LogEvent "Tentative rota generation completed."
End Function

Function PadRight(str, totalLength)
    Call LogVerbose("PadRight invoked")
    Dim padding
    padding = totalLength - Len(str)
    If padding > 0 Then
        PadRight = str & String(padding, " ")
    Else
        PadRight = str
    End If
End Function

Sub WriteRotaCSV(rotaText, filePath)
    Call LogVerbose("WriteRotaCSV invoked for filePath='" & filePath & "'")
    Dim ts, lines, line, i
    LogEvent "Writing provided rota text to CSV at: " & filePath

    lines = Split(rotaText, vbCrLf)

    Set ts = fso.CreateTextFile(filePath, True)
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line <> "" Then
            line = Replace(line, "|", ",")
            ts.WriteLine line
        End If
    Next

    ts.Close
    Set ts = Nothing
    LogEvent "CSV writing completed for: " & filePath
End Sub

'==============================================================
' Workbook Helpers
'==============================================================
Function LoadMeetingWorkbook(ByRef objExcel, ByRef objWorkbook, ByRef objSheetRota, ByRef objSheetEmails)
    Call LogVerbose("LoadMeetingWorkbook invoked")
    LoadMeetingWorkbook = False

    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        LogError "Failed to create Excel.Application"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    LogEvent "Excel application created successfully."

    On Error Resume Next
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_PATH)
    If Err.Number <> 0 Then
        LogError "Failed to open workbook: " & EXCEL_PATH
        objExcel.Quit
        Set objExcel = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    LogEvent "Workbook opened: " & EXCEL_PATH

    Set objSheetRota = objWorkbook.Sheets(SHEET_ROTA)
    Set objSheetEmails = objWorkbook.Sheets(SHEET_EMAILS)
    LogEvent "Sheets obtained: Rota and Emails"

    LoadMeetingWorkbook = True
End Function

Sub CloseMeetingWorkbook(ByRef objExcel, ByRef objWorkbook, ByRef objSheetRota, ByRef objSheetEmails)
    Call LogVerbose("CloseMeetingWorkbook invoked")
    objWorkbook.Close False
    objExcel.Quit
    Set objExcel = Nothing
    Set objWorkbook = Nothing
    Set objSheetRota = Nothing
    Set objSheetEmails = Nothing
    LogEvent "Coordinator routine completed and Excel resources released."
End Sub

Sub PopulateFridayDates(ByRef fridayDates)
    Call LogVerbose("PopulateFridayDates invoked")
    fridayDates(0) = Date + ((5 - Weekday(Date, vbMonday) + 7) Mod 7)
    fridayDates(1) = fridayDates(0) + 7
    fridayDates(2) = fridayDates(0) + 14
    LogEvent "Computed Fridays | This: " & FormatDateTime(fridayDates(0), 2) & _
              " | Next: " & FormatDateTime(fridayDates(1), 2) & _
              " | Next+1: " & FormatDateTime(fridayDates(2), 2)
End Sub

Sub LocateMeetingRows(ByRef fridayDates, ByVal objSheetRota, ByRef rowIndexes)
    Call LogVerbose("LocateMeetingRows invoked")
    Dim lastRow, i, meetingDate
    lastRow = objSheetRota.Cells(objSheetRota.Rows.Count, "A").End(-4162).Row
    LogEvent "Scanning rota sheet up to row: " & lastRow

    For i = 0 To 2
        rowIndexes(i) = -1
    Next

    For i = 2 To lastRow
        meetingDate = objSheetRota.Cells(i, 1).Value
        If meetingDate = fridayDates(0) Then rowIndexes(0) = i
        If meetingDate = fridayDates(1) Then rowIndexes(1) = i
        If meetingDate = fridayDates(2) Then rowIndexes(2) = i
    Next

    LogEvent "Rota rows located | This: " & rowIndexes(0) & _
              " | Next: " & rowIndexes(1) & _
              " | Next+1: " & rowIndexes(2)
End Sub

Sub SendCriticalAlert(alertType, targetDate, recipient, objSheetEmails)
    Call LogVerbose("SendCriticalAlert invoked for alertType='" & alertType & "'")
    Dim subject, tentativeRota, csvPath
    subject = "ALERT: " & alertType & " for " & FormatDateTime(targetDate, 2)

    If EmailExists(subject, 6, False) Then Exit Sub

    If alertType = "MissingDate" Or alertType = "Unassigned" Then
        tentativeRota = GenerateTentativeRota(objSheetEmails)
        csvPath = fso.GetSpecialFolder(2) & "\TentativeRota.csv"
        WriteRotaCSV tentativeRota, csvPath
    Else
        tentativeRota = ""
        csvPath = ""
    End If

    LogEvent "Dispatching alert email | Type: " & alertType & _
              " | Date: " & FormatDateTime(targetDate, 2) & _
              " | Recipient: " & recipient

    AlertEmailWithAttachment recipient, _
        GenerateEmailAlert(alertType, FormatDateTime(targetDate, 2), recipient, tentativeRota), _
        subject, csvPath
End Sub

Sub EvaluateCriticalWindows(ByRef fridayDates, ByRef rowIndexes, ByRef objSheetRota, ByRef objSheetEmails)
    Call LogVerbose("EvaluateCriticalWindows invoked")
    Dim idx, rowNum
    Dim tmpLiterature, tmpResults, tmpAlternate, tmpRoom

    For idx = 0 To 2
        rowNum = rowIndexes(idx)

        If rowNum = -1 Then
            SendCriticalAlert "MissingDate", fridayDates(idx), MANAGER, objSheetEmails
        Else
            tmpLiterature = objSheetRota.Cells(rowNum, 2).Value
            tmpResults    = objSheetRota.Cells(rowNum, 3).Value
            tmpAlternate  = objSheetRota.Cells(rowNum, 4).Value
            tmpRoom       = objSheetRota.Cells(rowNum, 5).Value
            LogEvent "Row " & rowNum & " data | Lit: " & tmpLiterature & _
                          " | Res: " & tmpResults & " | Alt: " & tmpAlternate & _
                          " | Room: " & tmpRoom

            If Trim(tmpLiterature) = "" And Trim(tmpResults) = "" And Trim(tmpAlternate) = "" Then
                SendCriticalAlert "Unassigned", fridayDates(idx), SUPERVISOR_EMAIL, objSheetEmails
            End If

            If Trim(tmpRoom) = "" Then
                SendCriticalAlert "MissingRoom", fridayDates(idx), BUILDING_ADMIN_EMAIL, objSheetEmails
            End If
        End If
    Next
End Sub

Sub HandleCommunications(ByRef fridayDates, ByRef rowIndexes, ByRef objSheetRota, ByVal shouldSendRegularEmail)
    Call LogVerbose("HandleCommunications invoked")
    Dim currentRow, literature, results, labManager, alternate, meetingLeadName, body, header, managerAlertSubject
    currentRow = rowIndexes(0)

    If currentRow = -1 Then
        LogEvent "No rota entry found for this Friday. Regular email and reminders will not be processed."
        Exit Sub
    End If

    literature = Trim(objSheetRota.Cells(currentRow, 2).Value)
    results    = Trim(objSheetRota.Cells(currentRow, 3).Value)
    labManager = Trim(objSheetRota.Cells(currentRow, 6).Value)
    alternate  = Trim(objSheetRota.Cells(currentRow, 4).Value)
    meetingLeadName = DeriveMeetingLeadName(objSheetRota, fridayDates(0))

    If literature = "" And results = "" And alternate = "" Then
        LogEvent "No presenters or alternate assigned for " & FormatDateTime(fridayDates(0), 2) & ". Regular meeting email suppressed."
        managerAlertSubject = "ALERT: MeetingNotSet for " & FormatDateTime(fridayDates(0), 2)
        If Not EmailExists(managerAlertSubject, 5, False) Then
            LogEvent "Sending manager alert that the meeting is not set for " & FormatDateTime(fridayDates(0), 2)
            Call AlertEmail(MANAGER, _
                GenerateEmailAlert("MeetingNotSet", FormatDateTime(fridayDates(0), 2), MANAGER, ""), _
                managerAlertSubject)
        Else
            LogEvent "Manager alert already exists for " & FormatDateTime(fridayDates(0), 2) & ". Skipping duplicate notification."
        End If
        Exit Sub
    End If

    TriggerReminderEmails literature, results, labManager, meetingLeadName, fridayDates(0)

    If shouldSendRegularEmail Then
        header = BuildMeetingSubject(fridayDates(0))
        body = GenerateEmailContent(fridayDates(0), currentRow, objSheetRota, rowIndexes(1))
        LogEvent "Dispatching regular meeting email for " & FormatDateTime(fridayDates(0), 2)
        DistributeEmail header, body
    Else
        header = BuildMeetingSubject(fridayDates(0))
        LogEvent "Regular meeting email already exists. Skipping dispatch for subject: " & header
    End If
End Sub

'==============================================================
' Coordinator (Main Workflow)
'==============================================================
Sub Coordinator()
    Call LogVerbose("Coordinator invoked")
    Dim fridayDates(2), rowIndexes(2)
    Dim meetingSubject
    Dim shouldSendRegularEmail
    Dim objExcel, objWorkbook, objSheetRota, objSheetEmails

    LogEvent "Coordinator routine started"

    PopulateFridayDates fridayDates
    meetingSubject = BuildMeetingSubject(fridayDates(0))

    shouldSendRegularEmail = Not EmailExists(meetingSubject, 5, True)
    If Not shouldSendRegularEmail Then
        LogEvent "Existing meeting email detected in Sent Items. Regular message will be skipped, reminders will still be evaluated."
    End If

    If Not LoadMeetingWorkbook(objExcel, objWorkbook, objSheetRota, objSheetEmails) Then
        Exit Sub
    End If

    LocateMeetingRows fridayDates, objSheetRota, rowIndexes
    EvaluateCriticalWindows fridayDates, rowIndexes, objSheetRota, objSheetEmails
    HandleCommunications fridayDates, rowIndexes, objSheetRota, shouldSendRegularEmail

    CloseMeetingWorkbook objExcel, objWorkbook, objSheetRota, objSheetEmails
End Sub

'==============================================================
' Script Entrypoint
'==============================================================
Sub Main()
    Call LogVerbose("Main invoked")
    InitializeFileSystem
    InitializeLogging

    If Not SyncMeetingWorkbook() Then
        LogEvent "Continuing execution despite sync issues."
    End If

    If Not InitializeOutlook() Then
        CleanupSyncedWorkbook
        LogEvent "=== Script execution completed ==="
        Exit Sub
    End If

    LogEvent "Invoking Coordinator routine"
    Coordinator

    CleanupSyncedWorkbook
    LogEvent "=== Script execution completed ==="
End Sub

Main
