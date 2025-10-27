Option Explicit

Dim objShell, strCmd
Set objShell = CreateObject("WScript.Shell")

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim SCRIPT_FOLDER, LOG_FILE_PATH
SCRIPT_FOLDER = fso.GetParentFolderName(WScript.ScriptFullName)
LOG_FILE_PATH = fso.BuildPath(SCRIPT_FOLDER, "group_meetings_manager.log")

Sub LogEvent(message)
    On Error Resume Next
    Dim ts
    Set ts = fso.OpenTextFile(LOG_FILE_PATH, 8, True)
    ts.WriteLine Now & " - " & message
    ts.Close
    Set ts = Nothing
    On Error GoTo 0
End Sub

Sub LogError(context)
    Dim errMsg
    errMsg = context & " | Err: " & Err.Number & " - " & Err.Description
    LogEvent errMsg
    Err.Clear
End Sub

Call LogEvent("=== Script execution started ===")
Call LogEvent("Logging to file: " & LOG_FILE_PATH)

' Build the command
strCmd = "C:\rclone\rclone.exe copy ""ABTsharepoint:/Group Meetings/meeting_rota.xlsx"" ""C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings"""
Call LogEvent("Executing rclone command: " & strCmd)

' Run the command, show window (1) and wait until it finishes (True)
Dim rcloneExitCode
rcloneExitCode = objShell.Run(strCmd, 0, True)
Call LogEvent("rclone command completed with exit code: " & rcloneExitCode)
If rcloneExitCode <> 0 Then
    Call LogEvent("WARNING: rclone returned non-zero exit code: " & rcloneExitCode)
End If

'==============================================================
' Group Meeting Email Automation (Integrated & Optimized)
'==============================================================

'-----------------------------
' Global Configuration
'-----------------------------
Const DEBUG_MODE = False   ' Set to True to dry-run without sending emails
Const WEEKDAY_EMAIL_ONLY = True  ' <<< set False to allow weekend email sending
Const MANAGER   = "gd2613@ic.ac.uk"
Const BUILDING_ADMIN_EMAIL   = "i.gomezvilchis@imperial.ac.uk"
Const SUPERVISOR_EMAIL   = "agi@imperial.ac.uk"
Const DEFAULT_PREVIOUS_LAB_MANAGER = "Agi"

Const EMAIL_TO = "brandt-talbot-group@imperial.ac.uk"
Const EXCEL_PATH = "C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings\meeting_rota.xlsx"
Const SHEET_ROTA = "Rota"
Const SHEET_EMAILS = "Emails"

Dim gEmailDispatchAllowed
Dim gWeekendNoticeLogged
gWeekendNoticeLogged = False

Function InitializeEmailDispatchSetting()
    If Not WEEKDAY_EMAIL_ONLY Then
        InitializeEmailDispatchSetting = True
    Else
        Dim dayOfWeek
        dayOfWeek = Weekday(Date, vbMonday)
        InitializeEmailDispatchSetting = (dayOfWeek <= 5)
    End If
End Function

Function ShouldSendEmail()
    If gEmailDispatchAllowed Then
        ShouldSendEmail = True
    Else
        If Not gWeekendNoticeLogged Then
            Call LogEvent("Weekend detected (" & FormatDateTime(Date, 2) & "). Email dispatch suppressed.")
            gWeekendNoticeLogged = True
        End If
        ShouldSendEmail = False
    End If
End Function

gEmailDispatchAllowed = InitializeEmailDispatchSetting()
If gEmailDispatchAllowed Then
    Call LogEvent("Email dispatch permitted for today: " & FormatDateTime(Date, 2))
Else
    Call LogEvent("Email dispatch restricted to weekdays. Today is: " & FormatDateTime(Date, 2))
End If

Dim OL, olNS

On Error Resume Next
Set OL = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
    LogError "Failed to create Outlook.Application"
    WScript.Quit Err.Number
Else
    Call LogEvent("Outlook application created successfully.")
End If
On Error GoTo 0

On Error Resume Next
Set olNS = OL.GetNamespace("MAPI")
If Err.Number <> 0 Then
    LogError "Failed to obtain Outlook namespace"
    WScript.Quit Err.Number
End If
On Error GoTo 0
Call LogEvent("Obtained Outlook namespace for MAPI.")

'==============================================================
' Module 0: Utils
'==============================================================

Function GetExactEmail(searchName)
    If Trim(searchName) = "" Then
        GetExactEmail = Nothing
        Exit Function
    End If

    Dim olApp, ns, recip, exchUser
    Set olApp = CreateObject("Outlook.Application")
    Set ns = olApp.GetNamespace("MAPI")

    ' Try Outlook’s built-in resolver for an exact match
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
        GetExactEmail = Nothing
    End If
End Function

Function FormatBoldName(value)
    Dim trimmed
    trimmed = Trim(value)
    If trimmed = "" Then
        FormatBoldName = ""
    Else
        FormatBoldName = "**" & trimmed & "**"
    End If
End Function

'==============================================================
' Module 1: Email Verification
'==============================================================
Function EmailExists(header, folderConstant, checkOnly)
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

    Call LogEvent("Checking Outlook " & folderLabel & " for email with subject: " & header)

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
                    Call LogEvent("Confirmed email with subject '" & header & "' in " & folderLabel & ".")
                Else
                    Call LogEvent("Email with subject '" & header & "' already exists in " & folderLabel & ". Skipping related actions.")
                End If
                Exit For
            End If
        End If
    Next

    If Not found Then
        If checkMode Then
            Call LogEvent("Email with subject '" & header & "' not yet located in " & folderLabel & " during verification.")
        Else
            Call LogEvent("Email with subject '" & header & "' not found in " & folderLabel & " prior to action.")
        End If
    End If

    Call LogEvent("EmailExists('" & header & "', folder=" & folderLabel & ", checkOnly=" & checkMode & ") result: " & found)
    EmailExists = found
End Function

'==============================================================
' Module 2: Email Generation (Meeting)
'==============================================================
Function GenerateEmailContent(meetingDate, row, objSheetRota, nextFridayRow)
    Dim literature, results, alternate, room, labManager, next_lit, next_res, next_alt, content, nameParts
    Dim previousLabManager, previousFriday, previousFridayRow, searchRow, lastRow, candidateDate, meetingLeadName
    literature = Trim(objSheetRota.Cells(row, 2).Value)
    results    = Trim(objSheetRota.Cells(row, 3).Value)
    alternate  = Trim(objSheetRota.Cells(row, 4).Value)
    room       = Trim(objSheetRota.Cells(row, 5).Value)
    labManager = Trim(objSheetRota.Cells(row, 6).Value)
    Call LogEvent("Generating main email content for date " & meetingDate & _
                  " | Row: " & row & " | Lit: " & literature & _
                  " | Res: " & results & " | Alt: " & alternate & _
                  " | Room: " & room)
    
    ' Next Friday content
    If nextFridayRow <> -1 Then
        next_lit = objSheetRota.Cells(nextFridayRow, 2).Value
        next_res = objSheetRota.Cells(nextFridayRow, 3).Value
        next_alt = objSheetRota.Cells(nextFridayRow, 4).Value
        Call LogEvent("Next Friday presenters found | Row: " & nextFridayRow & _
                      " | Lit: " & next_lit & " | Res: " & next_res & _
                      " | Alt: " & next_alt)
    Else
        next_lit = ""
        next_res = ""
        next_alt = ""
        Call LogEvent("No next Friday data found. Using empty placeholders.")
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

    meetingLeadName = previousLabManager
    
    content = "Hi all," & vbCrLf
    If Trim(room) <> "" Then
        content = content & "We will have our group meeting on " & meetingDate & _
                  " at 11:00 am in " & room & " in person or via MS Teams." & vbCrLf
    Else
        content = content & "We will have our group meeting on " & meetingDate & _
                  " at 11:00 am in person or via MS Teams." & vbCrLf
    End If

    content = content & vbCrLf & "Meeting agenda:" & vbCrLf
    If Trim(alternate) <> "" Then
        content = content & "    " & FormatBoldName(alternate) & vbCrLf
    Else
        If Trim(results) <> "" Then
            content = content & "    " & FormatBoldName(results) & " will present results." & vbCrLf
        End If
        If Trim(literature) <> "" Then
            content = content & "    " & FormatBoldName(literature) & " will present literature." & vbCrLf
        End If
    End If

    content = content & vbCrLf & FormatBoldName(meetingLeadName) & " will lead the meeting and do the inclusion activity." & vbCrLf

    If labManager = "" Then
        labManager = "TBD"
    End If

    content = content & vbCrLf & vbCrLf
    content = content & "The lab manager for this week is " & FormatBoldName(labManager) & "." & vbCrLf

    If next_lit <> "" Or next_res <> "" Or next_alt <> "" Then
        content = content & vbCrLf & "Next week:" & vbCrLf
        If Trim(next_alt) <> "" Then
            content = content & "    " & FormatBoldName(next_alt) & vbCrLf
        Else
            If Trim(next_res) <> "" Then
                content = content & "    " & FormatBoldName(next_res) & " will present results." & vbCrLf
            End If
            If Trim(next_lit) <> "" Then
                content = content & "    " & FormatBoldName(next_lit) & " will present literature." & vbCrLf
            End If
        End If
    Else
        content = content & vbCrLf & "No meeting has been set for the following week." & vbCrLf
    End If

    ' Add link to rota
    content = content & "Upcoming weeks schedule: " & _
              "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx?web=1" & vbCrLf

    content = content & "Best," & vbCrLf & FormatBoldName(senderFirstName)
    GenerateEmailContent = content
    Call LogEvent("Generated email body length: " & Len(content))
End Function

'==============================================================
' Module 2b: Presenter Reminder Generation
'==============================================================
Function GeneratePresenterReminderEmail(presenterName, presentationType)
    Dim content, nameParts, normalizedType, description, greetingName, tokens, missingDescription
    presenterName = Trim(presenterName)
    presentationType = Trim(presentationType)
    Call LogEvent("Generating presenter reminder email | Presenter: " & presenterName & _
                  " | Type: " & presentationType)

    normalizedType = LCase(presentationType)
    Select Case normalizedType
        Case "literature"
            description = "the literature session"
        Case "results"
            description = "the results update"
        Case Else
            If presentationType <> "" Then
                description = "the " & presentationType & " presentation"
            Else
                description = "your presentation"
            End If
    End Select

    missingDescription = description
    If description = "your presentation" Then
        missingDescription = "the presentation slot"
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

    If presenterName = "" Then
        content = "Hi," & vbCrLf & vbCrLf
        content = content & "No presenter is currently assigned for " & missingDescription & _
                  " at this week's group meeting." & vbCrLf & vbCrLf
        content = content & "Please update the rota or advise the group so we can keep the agenda on track." & vbCrLf & vbCrLf
        content = content & "Best," & vbCrLf & FormatBoldName(senderFirstName)
        GeneratePresenterReminderEmail = content
        Call LogEvent("Generated presenter reminder length: " & Len(content))
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

    content = "Hi " & FormatBoldName(greetingName) & "," & vbCrLf & vbCrLf
    content = content & "Just a reminder that you are scheduled to present " & description & _
              " at this week's group meeting." & vbCrLf & vbCrLf
    content = content & "Please ensure your materials are ready and let me know if you need any support." & vbCrLf & vbCrLf
    content = content & "Best," & vbCrLf & FormatBoldName(senderFirstName)
    GeneratePresenterReminderEmail = content
    Call LogEvent("Generated presenter reminder length: " & Len(content))
End Function

'==============================================================
' Module 2c: Lab Manager Reminder Generation
'==============================================================
Function GenerateLabManagerReminderEmail(labManagerName, meetingDate)
    Dim content, nameParts, manager
    Dim firstName, lastName, senderDisplayName, parcelAddressee
    Dim greetingName, tokens
    Dim startText, endText, weekRange
    manager = Trim(labManagerName)
    Call LogEvent("Generating lab manager reminder email | Manager: " & manager)

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
        greetingDisplay = FormatBoldName(greetingName)
    End If

    content = "Hi " & greetingDisplay & "," & vbCrLf & vbCrLf
    content = content & "This is a reminder that you are the lab manager for the week of " & weekRange & "." & vbCrLf & vbCrLf
    content = content & "Lab cleaning starts on Friday and must be completed by next Friday " & ChrW(&H2014) & _
              " please try to finish it by Friday EOD so the lab is ready for Monday." & vbCrLf & vbCrLf
    Dim parcelDisplay
    parcelDisplay = parcelAddressee
    If LCase(parcelAddressee) <> "me" Then
        parcelDisplay = FormatBoldName(parcelAddressee)
    End If

    content = content & "On friday, please pick up all parcels addressed to " & parcelDisplay & _
              "; most of them contain the consumables ordered the last week that need to be re-stocked in the lab." & vbCrLf & vbCrLf
    content = content & "Also, please send me the list of missing items I should procure from the whiteboard and procure the rest from the stores so that we can maintain supplies available at all time." & vbCrLf & vbCrLf
    content = content & "Thank you for keeping the lab organized." & vbCrLf & vbCrLf
    content = content & "Best," & vbCrLf & FormatBoldName(senderDisplayName)
    GenerateLabManagerReminderEmail = content
    Call LogEvent("Generated lab manager reminder length: " & Len(content))
End Function

'==============================================================
' Module 3: Email Distribution
'==============================================================
Sub DistributeEmail(header, body)
    Dim mail
    Set mail = OL.CreateItem(0)
    Call LogEvent("Preparing regular meeting email with subject: " & header)
    If Not ShouldSendEmail() Then
        Call LogEvent("Skipping regular meeting email due to weekday restriction. Subject: " & header)
        Exit Sub
    End If
    With mail
        .To = EMAIL_TO
        .Subject = header
        .Body = body
        If DEBUG_MODE Then
            Call LogEvent("DEBUG_MODE enabled - email not sent. Subject: " & header)
            MsgBox "[DEBUG] Would send email:" & vbCrLf & vbCrLf & header & vbCrLf & body
        Else
            Call LogEvent("Sending email to " & EMAIL_TO & " with subject: " & header)
            .Send
        End If
    End With
End Sub

'==============================================================
' Module 4: Alert Email Generation
'==============================================================
Function GenerateEmailAlert(alertType, missingDate, recipientEmail, tentativeRota)
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

'==============================================================
' Module 5: Alert Email Distribution
'==============================================================
Sub AlertEmail(toAddress, messageBody, subjectText)
    Dim mail
    Set mail = OL.CreateItem(0)
    Call LogEvent("Preparing alert email with subject: " & subjectText & " to " & toAddress)
    If Not ShouldSendEmail() Then
        Call LogEvent("Skipping alert email due to weekday restriction. Subject: " & subjectText)
        Exit Sub
    End If
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = messageBody
        If DEBUG_MODE Then
            Call LogEvent("DEBUG_MODE enabled - alert email not sent. Subject: " & subjectText)
            MsgBox "[DEBUG] Would send ALERT email to " & toAddress & ":" & vbCrLf & _
                "Subject: " & subjectText & vbCrLf & vbCrLf & messageBody
        Else
            Call LogEvent("Sending alert email to " & toAddress & " with subject: " & subjectText)
            .Send
        End If
    End With
End Sub

'==============================================================
' Module 5b: Alert Email Distribution with Attachment
'==============================================================
Sub AlertEmailWithAttachment(toAddress, messageBody, subjectText, attachmentPath)
    Dim mail
    Set mail = OL.CreateItem(0)
    Call LogEvent("Preparing alert email with attachment. Subject: " & subjectText & _
                  " | Recipient: " & toAddress & " | Attachment: " & attachmentPath)
    If Not ShouldSendEmail() Then
        Call LogEvent("Skipping alert email with attachment due to weekday restriction. Subject: " & subjectText)
        Exit Sub
    End If
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = messageBody
        If attachmentPath <> "" Then
            .Attachments.Add attachmentPath
            Call LogEvent("Attached file to email: " & attachmentPath)
        End If
        If DEBUG_MODE Then
            Call LogEvent("DEBUG_MODE enabled - alert email with attachment not sent. Subject: " & subjectText)
            MsgBox "[DEBUG] Would send ALERT email with attachment to " & toAddress & ":" & vbCrLf & _
                   "Subject: " & subjectText & vbCrLf & vbCrLf & messageBody & vbCrLf & _
                   "Attachment: " & attachmentPath
        Else
            Call LogEvent("Sending alert email with attachment to " & toAddress & " with subject: " & subjectText)
            .Send
        End If
    End With
End Sub

'==============================================================
' Module 6: Tentative Rota Generation
'==============================================================
Function GenerateTentativeRota(objSheetEmails)
    Dim lastRow, i, numNames, names(), rndIndex, tmp
    Dim rotaDates(), rotaAssignments(), rotaBody, dt, totalWeeks
    Dim csvFile, fso, txtStream
    Call LogEvent("Starting tentative rota generation")

    '--- Collect names with column D = 1 ---
    lastRow = objSheetEmails.Cells(objSheetEmails.Rows.Count, "A").End(-4162).Row
    numNames = 0
    For i = 2 To lastRow
        If objSheetEmails.Cells(i, 4).Value = 1 Then numNames = numNames + 1
    Next
    Call LogEvent("Eligible names counted: " & numNames)
    If numNames = 0 Then
        Call LogEvent("No eligible names found for tentative rota.")
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
    Call LogEvent("Names loaded for rota: " & Join(names, ", "))

    '--- Shuffle names ---
    Randomize
    For i = 0 To UBound(names)
        rndIndex = Int(Rnd * (UBound(names) + 1))
        tmp = names(i)
        names(i) = names(rndIndex)
        names(rndIndex) = tmp
    Next
    Call LogEvent("Names shuffled for tentative rota.")

    '--- Generate tentative dates and assignments ---
    totalWeeks = 26
    Call LogEvent("Generating tentative rota for weeks: " & totalWeeks)
    ReDim rotaDates(totalWeeks - 1)
    ReDim rotaAssignments(totalWeeks - 1, 1)
    dt = Date + ((5 - Weekday(Date, vbMonday) + 7) Mod 7)

    For i = 0 To totalWeeks - 1
        rotaDates(i) = DateAdd("ww", i, dt)
        rotaAssignments(i, 0) = names((i * 2) Mod UBound(names) + 1)
        rotaAssignments(i, 1) = names((i * 2 + 1) Mod UBound(names) + 1)
    Next

    '--- Build email body ---
    rotaBody = "Tentative 6-month rota (Fridays, 2 presenters/day):" & vbCrLf & vbCrLf
    rotaBody = rotaBody & "Date         | Presenter 1          | Presenter 2" & vbCrLf
    rotaBody = rotaBody & String(50, "-") & vbCrLf
    For i = 0 To totalWeeks - 1
        rotaBody = rotaBody & FormatDateTime(rotaDates(i), 2) & " | " & _
                   PadRight(rotaAssignments(i, 0), 20) & " | " & _
                   PadRight(rotaAssignments(i, 1), 20) & vbCrLf
    Next

    '--- Create CSV ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    csvFile = fso.GetSpecialFolder(2) & "\TentativeRota.csv"
    Call LogEvent("Writing tentative rota CSV to: " & csvFile)
    Set txtStream = fso.CreateTextFile(csvFile, True)
    txtStream.WriteLine "Date,Presenter1,Presenter2"
    For i = 0 To totalWeeks - 1
        txtStream.WriteLine FormatDateTime(rotaDates(i), 2) & "," & _
                            rotaAssignments(i, 0) & "," & rotaAssignments(i, 1)
    Next
    txtStream.Close

    GenerateTentativeRota = rotaBody & vbCrLf & "[CSV: " & csvFile & "]"
    Call LogEvent("Tentative rota generation completed.")
End Function

Function PadRight(str, totalLength)
    Dim padding
    padding = totalLength - Len(str)
    If padding > 0 Then
        PadRight = str & String(padding, " ")
    Else
        PadRight = str
    End If
End Function

'==============================================================
' Module 6b: Write Tentative Rota to CSV
'==============================================================
Sub WriteRotaCSV(rotaText, filePath)
    Dim fso, ts, lines, line, i
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call LogEvent("Writing provided rota text to CSV at: " & filePath)

    ' Split rota text by lines
    lines = Split(rotaText, vbCrLf)

    ' Create file (overwrite if exists)
    Set ts = fso.CreateTextFile(filePath, True)

    ' Loop through each line and write as CSV row
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line <> "" Then
            ' Replace "|" separators with commas for CSV
            line = Replace(line, "|", ",")
            ts.WriteLine line
        End If
    Next

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Call LogEvent("CSV writing completed for: " & filePath)
End Sub

'==============================================================
' Module 7: Coordinator (Optimized)
'==============================================================
Sub Coordinator()
    Dim thisFriday, nextFriday, nextNextFriday
    Dim header, objExcel, objWorkbook, objSheetRota, objSheetEmails
    Dim lastRow, i, meetingDate
    Dim foundRow, nextFridayRow, nextNextFridayRow
    Dim body, literature, results, alternate, room, managerAlertSubject
    Call LogEvent("Coordinator routine started")

    '--- Define Fridays ---
    thisFriday = Date + ((5 - Weekday(Date, vbMonday) + 7) Mod 7)
    nextFriday = thisFriday + 7
    nextNextFriday = thisFriday + 14
    header = "ABT group meeting @ 11am (" & FormatDateTime(thisFriday, 2) & ")"
    Call LogEvent("Computed Fridays | This: " & FormatDateTime(thisFriday, 2) & _
                  " | Next: " & FormatDateTime(nextFriday, 2) & _
                  " | Next+1: " & FormatDateTime(nextNextFriday, 2))

    '--- Exit if regular meeting email already exists ---
    If EmailExists(header, 6, False) Then
        Call LogEvent("Existing meeting email detected. Coordinator will exit early.")
        Exit Sub
    End If

    '--- Open Excel ---
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        LogError "Failed to create Excel.Application"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Call LogEvent("Excel application created successfully.")

    On Error Resume Next
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_PATH)
    If Err.Number <> 0 Then
        LogError "Failed to open workbook: " & EXCEL_PATH
        objExcel.Quit
        Set objExcel = Nothing
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Call LogEvent("Workbook opened: " & EXCEL_PATH)
    Set objSheetRota = objWorkbook.Sheets(SHEET_ROTA)
    Set objSheetEmails = objWorkbook.Sheets(SHEET_EMAILS)
    Call LogEvent("Sheets obtained: Rota and Emails")

    lastRow = objSheetRota.Cells(objSheetRota.Rows.Count, "A").End(-4162).Row
    foundRow = -1: nextFridayRow = -1: nextNextFridayRow = -1
    Call LogEvent("Scanning rota sheet up to row: " & lastRow)

    '--- Locate rows for Fridays ---
    For i = 2 To lastRow
        meetingDate = objSheetRota.Cells(i, 1).Value
        If meetingDate = thisFriday Then foundRow = i
        If meetingDate = nextFriday Then nextFridayRow = i
        If meetingDate = nextNextFriday Then nextNextFridayRow = i
    Next
    Call LogEvent("Rota rows located | This: " & foundRow & _
                  " | Next: " & nextFridayRow & _
                  " | Next+1: " & nextNextFridayRow)

    '--- Prepare critical issues arrays ---
    Dim criticalDates(2), criticalTypes(2), recipientEmails(2)
    criticalDates(0) = thisFriday: criticalDates(1) = nextFriday: criticalDates(2) = nextNextFriday
    criticalTypes(0) = "": criticalTypes(1) = "": criticalTypes(2) = ""
    recipientEmails(0) = "": recipientEmails(1) = "": recipientEmails(2) = ""

    Dim idx
    Dim tmpLiterature, tmpResults, tmpAlternate, tmpRoom
    Dim rowNum

    '--- Check each Friday ---
    For idx = 0 To 2
        Select Case idx
            Case 0: rowNum = foundRow
            Case 1: rowNum = nextFridayRow
            Case 2: rowNum = nextNextFridayRow
        End Select

        If rowNum = -1 Then
            criticalTypes(idx) = "MissingDate"
            recipientEmails(idx) = MANAGER
            Call LogEvent("MissingDate detected for " & FormatDateTime(criticalDates(idx), 2) & _
                          ". Alert will be sent to " & MANAGER)
        Else
            tmpLiterature = objSheetRota.Cells(rowNum, 2).Value
            tmpResults    = objSheetRota.Cells(rowNum, 3).Value
            tmpAlternate  = objSheetRota.Cells(rowNum, 4).Value
            tmpRoom       = objSheetRota.Cells(rowNum, 5).Value
            Call LogEvent("Row " & rowNum & " data | Lit: " & tmpLiterature & _
                          " | Res: " & tmpResults & " | Alt: " & tmpAlternate & _
                          " | Room: " & tmpRoom)

            If Trim(tmpLiterature) = "" And Trim(tmpResults) = "" And Trim(tmpAlternate) = "" Then
                criticalTypes(idx) = "Unassigned"
                recipientEmails(idx) = SUPERVISOR_EMAIL
                Call LogEvent("Unassigned presenters detected for " & FormatDateTime(criticalDates(idx), 2) & _
                              ". Alert will be sent to supervisor.")
            End If

            If Trim(tmpRoom) = "" Then
                ' MissingRoom is independent; we allow sending separately
                Dim missingRoomSubject
                missingRoomSubject = "ALERT: MissingRoom for " & FormatDateTime(criticalDates(idx), 2)
                If Not EmailExists(missingRoomSubject, 6, False) Then
                    Call LogEvent("Missing room detected for " & FormatDateTime(criticalDates(idx), 2) & _
                                  ". Alert will be sent to building admin.")
                    Call AlertEmail(BUILDING_ADMIN_EMAIL, _
                        GenerateEmailAlert("MissingRoom", FormatDateTime(criticalDates(idx), 2), BUILDING_ADMIN_EMAIL, ""), _
                        missingRoomSubject)
                End If
            End If
        End If
    Next

    '--- Send alerts for MissingDate or Unassigned ---
    Dim alertSubject, tentativeRota
    For idx = 0 To 2
        If criticalTypes(idx) = "MissingDate" Or criticalTypes(idx) = "Unassigned" Then
            alertSubject = "ALERT: " & criticalTypes(idx) & " for " & FormatDateTime(criticalDates(idx), 2)
            If Not EmailExists(alertSubject, 6, False) Then
                If criticalTypes(idx) = "MissingDate" Or criticalTypes(idx) = "Unassigned" Then
                    tentativeRota = GenerateTentativeRota(objSheetEmails)
                    ' Create CSV attachment
                    Dim fso, csvPath
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    csvPath = fso.GetSpecialFolder(2) & "\TentativeRota.csv"
                    WriteRotaCSV tentativeRota, csvPath
                    Call LogEvent("Prepared tentative rota attachment for alert: " & csvPath)
                Else
                    tentativeRota = ""
                    csvPath = ""
                End If
                Call LogEvent("Dispatching alert email | Type: " & criticalTypes(idx) & _
                              " | Date: " & FormatDateTime(criticalDates(idx), 2) & _
                              " | Recipient: " & recipientEmails(idx))
                Call AlertEmailWithAttachment(recipientEmails(idx), _
                    GenerateEmailAlert(criticalTypes(idx), FormatDateTime(criticalDates(idx), 2), recipientEmails(idx), tentativeRota), _
                    alertSubject, csvPath)
            End If
        End If
    Next

    '--- Send regular meeting email for this Friday ---
    If foundRow <> -1 Then
        literature = Trim(objSheetRota.Cells(foundRow, 2).Value)
        results = Trim(objSheetRota.Cells(foundRow, 3).Value)
        alternate = Trim(objSheetRota.Cells(foundRow, 4).Value)

        If literature = "" And results = "" And alternate = "" Then
            Call LogEvent("No presenters or alternate assigned for " & FormatDateTime(thisFriday, 2) & ". Regular meeting email suppressed.")
            managerAlertSubject = "ALERT: MeetingNotSet for " & FormatDateTime(thisFriday, 2)
            If Not EmailExists(managerAlertSubject, 5, False) Then
                Call LogEvent("Sending manager alert that the meeting is not set for " & FormatDateTime(thisFriday, 2))
                Call AlertEmail(MANAGER, _
                    GenerateEmailAlert("MeetingNotSet", FormatDateTime(thisFriday, 2), MANAGER, ""), _
                    managerAlertSubject)
            Else
                Call LogEvent("Manager alert already exists for " & FormatDateTime(thisFriday, 2) & ". Skipping duplicate notification.")
            End If
        Else
            body = GenerateEmailContent(thisFriday, foundRow, objSheetRota, nextFridayRow)
            Call LogEvent("Dispatching regular meeting email for " & FormatDateTime(thisFriday, 2))
            Call DistributeEmail(header, body)
        End If
    Else
        Call LogEvent("No rota entry found for this Friday. Regular email will not be sent.")
    End If

    '--- Clean up ---
    objWorkbook.Close False
    objExcel.Quit
    Set objExcel = Nothing
    Set objWorkbook = Nothing
    Set objSheetRota = Nothing
    Set objSheetEmails = Nothing
    Call LogEvent("Coordinator routine completed and Excel resources released.")
End Sub


'==============================================================
' Module 8: Scheduler (Entrypoint)
'==============================================================
Call LogEvent("Invoking Coordinator routine")
Coordinator


' Check if the file exists before deleting
If fso.FileExists(EXCEL_PATH) Then
    Call LogEvent("Deleting local Excel file after sync: " & EXCEL_PATH)
    fso.DeleteFile EXCEL_PATH, True  ' True = force deletion (even if read-only)
Else
    Call LogEvent("Local Excel file not found during cleanup: " & EXCEL_PATH)
End If

Call LogEvent("=== Script execution completed ===")
