Option Explicit

'==============================================================
' Global Objects and State
'==============================================================
Dim objShell, fso
Dim SCRIPT_FOLDER, LOG_FILE_PATH
Dim OL, olNS
Dim managedEmailLookupUncertain, managedEmailLookupReason
Dim meetingWorkbookNeedsUpload, keepSyncedWorkbookForRecovery

'==============================================================
' Configuration Constants
'==============================================================
Const DEBUG_MODE = False   ' Set to True to dry-run without sending emails
Const EXTENDED_LOGGING = True
Const LOG_VARIANT = "[Main] "
Const DISPLAY_DEBUG_POPUPS = False
Const MANAGER   = "gd2613@ic.ac.uk"
Const BUILDING_ADMIN_EMAIL   = "i.gomezvilchis@imperial.ac.uk"
Const DEFAULT_PREVIOUS_LAB_MANAGER = "Agi"
Const INVENTORY_MANAGER_NAME = "Yilyu Qi"
Const INVENTORY_MANAGER_EMAIL = "y.qi24@imperial.ac.uk"

Const EMAIL_TO = "brandt-talbot-group@imperial.ac.uk"
Const EXCEL_PATH = "C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings\meeting_rota.xlsx"
Const RCLONE_EXE = "C:\rclone\rclone.exe"
Const RCLONE_WORKBOOK_REMOTE = "ABTsharepoint:/Group Meetings/meeting_rota.xlsx"
Const RCLONE_WORKBOOK_DEST_DIR = "C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings"
Const SHEET_ROTA = "Rota"
Const SHEET_EMAILS = "Emails"

'==============================================================
' Email Delivery Schedule (local Windows date)
'==============================================================
Function IsWeekend(checkDate)
    IsWeekend = (Weekday(checkDate, vbMonday) > 5)
End Function
Const SHEET_CHEMICAL_INVENTORY = "Chemicals"
Const INVENTORY_WORKBOOK_URL = "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx?web=1&activeCell=Chemicals!A1"
Const INVENTORY_REVIEW_DATE_HEADER = "Review date"
Const INVENTORY_REVIEW_BATCH_SIZE = 5

Const MANAGED_EMAIL_PROPERTY_NAME = "ABTManagedEmailKey"
Const MANAGED_EMAIL_KEY_PREFIX = "ABT-GROUP-MEETING"
Const MANAGED_OUTBOX_GRACE_MINUTES = 15
Const ENABLE_STALE_OUTBOX_REPLACEMENT = True
Const ENABLE_MANAGED_EMAIL_KEY_FALLBACK_BY_SUBJECT = True
Const ENABLE_OUTLOOK_TRANSPORT_DIAGNOSTICS = True
Const ENABLE_OUTBOX_DIAGNOSTICS = True
Const MANAGED_SENT_ITEMS_SCAN_LIMIT = 250
Const MANAGED_OUTBOX_SCAN_LIMIT = 100

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

Sub ResetManagedEmailLookupUncertainty()
    managedEmailLookupUncertain = False
    managedEmailLookupReason = ""
End Sub

Sub MarkManagedEmailLookupUncertain(context)
    Dim detail
    detail = context
    If Err.Number <> 0 Then
        detail = detail & " | Err: " & Err.Number & " - " & Err.Description
    End If

    managedEmailLookupUncertain = True
    managedEmailLookupReason = detail
    LogEvent "ACTION REQUIRED: Outlook managed-email lookup became uncertain. " & detail
    Err.Clear
End Sub

Function IsManagedEmailLookupUncertain()
    IsManagedEmailLookupUncertain = CBool(managedEmailLookupUncertain)
End Function

Function ManagedEmailLookupUncertaintyReason()
    ManagedEmailLookupUncertaintyReason = CStr(managedEmailLookupReason)
End Function

Sub LogManagedEmailLookupBlocked(actionText, managedKey)
    LogEvent "Queue decision: ACTION_REQUIRED"
    LogEvent "Reason for decision: Outlook Sent Items/Outbox inspection is uncertain. " & _
             ManagedEmailLookupUncertaintyReason()
    LogEvent "ACTION REQUIRED: " & actionText & " blocked for managed key '" & managedKey & _
             "'. No send, delete, or replacement was performed. Restart/sync Outlook and rerun after reviewing Outlook manually if needed."
End Sub

'==============================================================
' Formatting Helpers
'==============================================================
Function FormatMeetingDateForFilename(meetingDate)
    Dim yearPart, monthPart, dayPart
    yearPart = Right("0" & (Year(meetingDate) Mod 100), 2)
    monthPart = Right("0" & Month(meetingDate), 2)
    dayPart = Right("0" & Day(meetingDate), 2)
    FormatMeetingDateForFilename = yearPart & monthPart & dayPart
End Function

Function SafeTextValue(value, fallbackText)
    On Error Resume Next

    If IsEmpty(value) Or IsNull(value) Then
        SafeTextValue = fallbackText
        On Error GoTo 0
        Exit Function
    End If

    If IsObject(value) Then
        SafeTextValue = fallbackText
        On Error GoTo 0
        Exit Function
    End If

    SafeTextValue = CStr(value)
    If Err.Number <> 0 Then
        Err.Clear
        SafeTextValue = fallbackText
    End If
    On Error GoTo 0
End Function

Function HtmlEscape(value)
    Dim textValue
    textValue = SafeTextValue(value, "")
    textValue = Replace(textValue, "&", "&amp;")
    textValue = Replace(textValue, "<", "&lt;")
    textValue = Replace(textValue, ">", "&gt;")
    textValue = Replace(textValue, """", "&quot;")
    textValue = Replace(textValue, "'", "&#39;")
    HtmlEscape = textValue
End Function

Function IsHtmlBody(bodyText)
    Dim trimmedText
    trimmedText = LCase(Trim(SafeTextValue(bodyText, "")))
    IsHtmlBody = (Left(trimmedText, 5) = "<html" Or Left(trimmedText, 15) = "<!doctype html>")
End Function

Function HtmlParagraph(textValue)
    HtmlParagraph = "<p style=""margin:0 0 14px 0; line-height:1.5;"">" & HtmlEscape(textValue) & "</p>"
End Function

Function HtmlInfoRow(labelText, valueText)
    HtmlInfoRow = "<tr>" & _
                  "<td width=""122"" style=""width:122px; padding:5px 14px; border-top:1px solid #e1e1e6; background-color:#f5f5f5; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:10.5pt; font-weight:600; line-height:15pt; mso-line-height-rule:exactly;"">" & HtmlEscape(labelText) & "</td>" & _
                  "<td style=""padding:5px 14px; border-top:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:10.5pt; line-height:15pt; mso-line-height-rule:exactly;"">" & HtmlEscape(valueText) & "</td>" & _
                  "</tr>"
End Function

Function HtmlResponsibilityRow(roleText, personText, detailText)
    Dim detailHtml
    detailHtml = ""
    If Trim(SafeTextValue(detailText, "")) <> "" Then
        detailHtml = "<div style=""color:#708090; font-size:13px; line-height:18px; margin-top:2px; mso-line-height-rule:exactly;"">" & HtmlEscape(detailText) & "</div>"
    End If

    HtmlResponsibilityRow = "<tr>" & _
                            "<td width=""190"" style=""width:190px; padding:7px 12px 5px 14px; background-color:#f5f5f5; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:14px; font-weight:500; line-height:20px; mso-line-height-rule:exactly;"">" & HtmlEscape(roleText) & "</td>" & _
                            "<td style=""padding:7px 14px 5px 14px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:16px; font-weight:400; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(personText) & detailHtml & "</td>" & _
                            "</tr>"
End Function

Function HtmlNextResponsibilityRow(roleText, personText)
    HtmlNextResponsibilityRow = "<tr>" & _
                                "<td width=""190"" style=""width:190px; padding:3px 12px 3px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:15px; font-weight:500; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(roleText) & "</td>" & _
                                "<td style=""padding:3px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:15px; font-weight:400; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(personText) & "</td>" & _
                                "</tr>"
End Function

Function HtmlSection(titleText, innerHtml)
    HtmlSection = "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; margin:16px 0 0 0; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                  "<tr><td style=""padding:0 0 5px 0; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:12pt; font-weight:600; line-height:16pt; mso-line-height-rule:exactly;"">" & HtmlEscape(titleText) & "</td></tr>" & _
                  "<tr><td style=""padding:0; border:1px solid #e1e1e6; background-color:#ffffff;"">" & innerHtml & "</td></tr>" & _
                  "</table>"
End Function

Function GetSenderDisplayName()
    Dim currentUserText
    On Error Resume Next
    currentUserText = SafeTextValue(olNS.Accounts.Item(1).CurrentUser.Name, "")
    If currentUserText = "" Then currentUserText = SafeTextValue(olNS.Accounts.Item(1).CurrentUser, "")
    If Err.Number <> 0 Then
        Err.Clear
        currentUserText = ""
    End If
    On Error GoTo 0

    If currentUserText = "" Then currentUserText = "Giacomo"
    GetSenderDisplayName = currentUserText
End Function

Function GetSenderFirstName()
    Dim nameParts, senderName, currentUserText
    senderName = ""
    currentUserText = GetSenderDisplayName()

    nameParts = Split(currentUserText, ", ")
    If UBound(nameParts) >= 1 Then
        senderName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        senderName = Trim(nameParts(0))
    End If
    If senderName = "" Then senderName = Trim(currentUserText)
    If senderName = "" Then senderName = "Giacomo"
    GetSenderFirstName = senderName
End Function

Function FirstNameForGreeting(displayName, fallbackName)
    Dim greetingName, tokens
    greetingName = Trim(SafeTextValue(displayName, ""))
    If greetingName = "" Then greetingName = Trim(SafeTextValue(fallbackName, "there"))

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
            If UBound(tokens) >= 0 Then greetingName = Trim(tokens(0))
        End If
        If Trim(greetingName) = "" Then greetingName = Trim(SafeTextValue(displayName, fallbackName))
    End If

    If Trim(greetingName) = "" Then greetingName = "there"
    FirstNameForGreeting = greetingName
End Function

Function ReminderEmailShell(innerHtml)
    ReminderEmailShell = "<!doctype html><html><body style=""margin:0; padding:0; background-color:#ffffff;"">" & _
                         "<table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:100%; border-collapse:collapse; background-color:#ffffff;"">" & _
                         "<tr><td align=""center"" style=""padding:14px 10px 18px 10px;"">" & _
                         "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                         "<tr><td style=""padding:0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & _
                         innerHtml & _
                         "</td></tr></table></td></tr></table></body></html>"
End Function

Function ReminderGreeting(greetingText)
    ReminderGreeting = "<p style=""margin:0 0 9px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(greetingText) & "</p>"
End Function

Function ReminderSignoff(senderName)
    ReminderSignoff = "<p style=""margin:13px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">Best,<br>" & HtmlEscape(Trim(senderName)) & "</p>"
End Function

Function ReminderParagraph(textValue)
    ReminderParagraph = "<p style=""margin:10px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(textValue) & "</p>"
End Function

Function ReminderSubtleParagraph(textValue)
    ReminderSubtleParagraph = "<p style=""margin:13px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#58677a; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(textValue) & "</p>"
End Function

Function ReminderHeader(kickerText, titleText, detailText)
    Dim detailHtml
    detailHtml = ""
    If Trim(SafeTextValue(detailText, "")) <> "" Then
        detailHtml = "<tr><td style=""padding:0 0 12px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(detailText) & "</td></tr>"
    End If

    ReminderHeader = "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                     "<tr><td style=""padding:0 0 3px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:13px; font-weight:600; line-height:18px; letter-spacing:0; mso-line-height-rule:exactly;"">" & HtmlEscape(kickerText) & "</td></tr>" & _
                     "<tr><td style=""padding:0 0 3px 0; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; mso-line-height-rule:exactly;"">" & HtmlEscape(titleText) & "</td></tr>" & _
                     detailHtml & _
                     "</table>"
End Function

Function ReminderBlockStart(titleText)
    ReminderBlockStart = "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; border:2px solid #0000cd; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                         "<tr><td colspan=""2"" style=""padding:9px 14px 8px 14px; border-bottom:1px solid #e1e1e6; background-color:#ffffff; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; letter-spacing:0; mso-line-height-rule:exactly;"">" & HtmlEscape(titleText) & "</td></tr>"
End Function

Function ReminderBlockEnd()
    ReminderBlockEnd = "</table>"
End Function

Function ReminderTableRow(labelText, valueText)
    ReminderTableRow = "<tr>" & _
                       "<td width=""190"" style=""width:190px; padding:7px 12px 5px 14px; background-color:#f5f5f5; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:14px; font-weight:500; line-height:20px; mso-line-height-rule:exactly;"">" & HtmlEscape(labelText) & "</td>" & _
                       "<td style=""padding:7px 14px 5px 14px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:16px; font-weight:400; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(valueText) & "</td>" & _
                       "</tr>"
End Function

Function ReminderPlainSection(titleText, innerHtml)
    ReminderPlainSection = "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; margin:12px 0 0 0; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                           "<tr><td style=""padding:0 0 6px 0; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:14px; font-weight:600; line-height:20px; letter-spacing:0; mso-line-height-rule:exactly;"">" & HtmlEscape(titleText) & "</td></tr>" & _
                           "<tr><td style=""padding:7px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & innerHtml & "</td></tr>" & _
                           "</table>"
End Function

Function HtmlMultilineText(textValue)
    Dim escapedText
    escapedText = HtmlEscape(textValue)
    escapedText = Replace(escapedText, vbCrLf, "<br>")
    escapedText = Replace(escapedText, vbCr, "<br>")
    escapedText = Replace(escapedText, vbLf, "<br>")
    HtmlMultilineText = escapedText
End Function

Function FindHeaderColumn(objSheet, headerText, defaultColumn)
    Dim lastColumn, col, currentHeader
    lastColumn = objSheet.Cells(1, objSheet.Columns.Count).End(-4159).Column
    For col = 1 To lastColumn
        currentHeader = LCase(Trim(SafeTextValue(objSheet.Cells(1, col).Value, "")))
        If currentHeader = LCase(Trim(SafeTextValue(headerText, ""))) Then
            FindHeaderColumn = col
            Exit Function
        End If
    Next
    FindHeaderColumn = defaultColumn
End Function

Function IsFlagSet(value)
    Dim textValue
    textValue = Trim(CStr(value))
    If IsNumeric(textValue) Then
        IsFlagSet = (CDbl(textValue) <> 0)
    Else
        IsFlagSet = (LCase(textValue) = "true" Or LCase(textValue) = "yes")
    End If
End Function

Function ExtractEmailAddress(textValue)
    On Error Resume Next
    Dim re, matches, cleaned
    ExtractEmailAddress = ""
    cleaned = Replace(Trim(CStr(textValue)), "mailto:", "", 1, -1, vbTextCompare)
    If cleaned = "" Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}"
    re.IgnoreCase = True
    re.Global = False
    Set matches = re.Execute(cleaned)
    If Err.Number = 0 Then
        If matches.Count > 0 Then ExtractEmailAddress = CStr(matches.Item(0).Value)
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function ExtractNameInitials(fullName)
    Dim cleaned, parts, firstName, lastName, tokens, i, initials
    cleaned = Trim(fullName)

    If cleaned = "" Then
        ExtractNameInitials = ""
        Exit Function
    End If

    firstName = ""
    lastName = ""

    If InStr(cleaned, ",") > 0 Then
        parts = Split(cleaned, ",")
        lastName = Trim(parts(0))
        If UBound(parts) >= 1 Then
            firstName = Trim(parts(1))
        End If
    Else
        tokens = Split(cleaned, " ")
        If UBound(tokens) >= 0 Then
            firstName = Trim(tokens(0))
            lastName = Trim(tokens(UBound(tokens)))
        End If
    End If

    If LCase(firstName) = LCase(lastName) Then
        lastName = ""
    End If

    initials = ""
    If firstName <> "" Then
        initials = initials & UCase(Left(firstName, 1))
    End If
    If lastName <> "" Then
        initials = initials & UCase(Left(lastName, 1))
    End If

    ExtractNameInitials = initials
End Function

'==============================================================
' Environment Bootstrap
'==============================================================
Sub InitializeFileSystem()
    Call LogVerbose("InitializeFileSystem invoked")
    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    SCRIPT_FOLDER = fso.GetParentFolderName(WScript.ScriptFullName)
    LOG_FILE_PATH = fso.BuildPath(SCRIPT_FOLDER, "group_meetings_manager.log")
    meetingWorkbookNeedsUpload = False
    keepSyncedWorkbookForRecovery = False
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

Function QuoteShellArg(value)
    QuoteShellArg = """" & Replace(CStr(value), """", """""") & """"
End Function

Function SyncMeetingWorkbook()
    Call LogVerbose("SyncMeetingWorkbook invoked")
    Dim cmd, exitCode
    cmd = QuoteShellArg(RCLONE_EXE) & " copy " & _
          QuoteShellArg(RCLONE_WORKBOOK_REMOTE) & " " & _
          QuoteShellArg(RCLONE_WORKBOOK_DEST_DIR)
    LogEvent "Executing rclone command: " & cmd
    exitCode = objShell.Run(cmd, 0, True)
    LogEvent "rclone command completed with exit code: " & exitCode
    If exitCode <> 0 Then
        LogEvent "WARNING: rclone returned non-zero exit code: " & exitCode
    End If
    SyncMeetingWorkbook = (exitCode = 0)
End Function

Function UploadMeetingWorkbook()
    Call LogVerbose("UploadMeetingWorkbook invoked")
    Dim cmd, exitCode

    UploadMeetingWorkbook = False

    If Not fso.FileExists(RCLONE_EXE) Then
        LogEvent "ERROR: rclone executable not found for workbook upload: " & RCLONE_EXE
        Exit Function
    End If

    If Not fso.FileExists(EXCEL_PATH) Then
        LogEvent "ERROR: Local workbook not found for upload: " & EXCEL_PATH
        Exit Function
    End If

    cmd = QuoteShellArg(RCLONE_EXE) & " copyto " & _
          QuoteShellArg(EXCEL_PATH) & " " & _
          QuoteShellArg(RCLONE_WORKBOOK_REMOTE)
    LogEvent "Executing rclone upload command: " & cmd
    exitCode = objShell.Run(cmd, 0, True)
    LogEvent "rclone upload command completed with exit code: " & exitCode

    If exitCode <> 0 Then
        LogEvent "ERROR: rclone upload returned non-zero exit code: " & exitCode
        Exit Function
    End If

    UploadMeetingWorkbook = True
End Function

Sub MarkMeetingWorkbookForUpload(reason)
    meetingWorkbookNeedsUpload = True
    LogEvent "Meeting rota workbook marked for upload: " & reason
End Sub

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

Function TryRefreshOutlookSession(context)
    Call LogVerbose("TryRefreshOutlookSession invoked for context='" & context & "'")
    TryRefreshOutlookSession = False

    On Error Resume Next
    Set OL = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        MarkManagedEmailLookupUncertain "Unable to refresh Outlook.Application while " & context
        On Error GoTo 0
        Exit Function
    End If

    Set olNS = OL.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        MarkManagedEmailLookupUncertain "Unable to refresh Outlook MAPI namespace while " & context
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    LogEvent "Refreshed Outlook COM session while " & context
    TryRefreshOutlookSession = True
End Function

Sub CleanupSyncedWorkbook()
    Call LogVerbose("CleanupSyncedWorkbook invoked")
    Dim firstErrNumber, firstErrDescription

    If CBool(keepSyncedWorkbookForRecovery) Then
        LogEvent "Keeping local Excel file for recovery because the edited workbook was not confirmed uploaded: " & EXCEL_PATH
        Exit Sub
    End If

    On Error Resume Next
    If Not fso.FileExists(EXCEL_PATH) Then
        If Err.Number <> 0 Then
            LogEvent "WARNING: Unable to check local Excel cleanup path. Continuing. Path: " & _
                     EXCEL_PATH & " | Err: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            LogEvent "Local Excel file not found during cleanup: " & EXCEL_PATH
        End If
        On Error GoTo 0
        Exit Sub
    End If

    LogEvent "Deleting local Excel file after sync: " & EXCEL_PATH
    fso.DeleteFile EXCEL_PATH, True
    If Err.Number = 0 Then
        LogEvent "Deleted local Excel file after sync: " & EXCEL_PATH
        On Error GoTo 0
        Exit Sub
    End If

    firstErrNumber = Err.Number
    firstErrDescription = Err.Description
    Err.Clear

    LogEvent "WARNING: Initial cleanup delete failed; retrying once after a short wait. Path: " & _
             EXCEL_PATH & " | Err: " & firstErrNumber & " - " & firstErrDescription
    WScript.Sleep 1000
    fso.DeleteFile EXCEL_PATH, True
    If Err.Number <> 0 Then
        LogEvent "WARNING: Cleanup could not delete local Excel file. It may still be locked by Excel, OneDrive, or permissions. Continuing without failing the run. Path: " & _
                 EXCEL_PATH & " | First err: " & firstErrNumber & " - " & firstErrDescription & _
                 " | Retry err: " & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        LogEvent "Deleted local Excel file after cleanup retry: " & EXCEL_PATH
    End If
    On Error GoTo 0
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

    Dim recip, addressEntry, exchUser, exchList, resolvedAddress, explicitAddress
    GetExactEmail = ""

    explicitAddress = ExtractEmailAddress(searchName)
    If explicitAddress <> "" Then
        LogEvent "Resolved explicit email address from local display text for '" & searchName & "': " & explicitAddress
        GetExactEmail = explicitAddress
        Exit Function
    End If

    If olNS Is Nothing Then
        If Not TryRefreshOutlookSession("resolving recipient '" & searchName & "'") Then
            LogEvent "WARNING: Unable to resolve recipient '" & searchName & "' because Outlook session is unavailable."
            Exit Function
        End If
    End If

    On Error Resume Next
    Set recip = olNS.CreateRecipient(searchName)
    If Err.Number <> 0 Or recip Is Nothing Then
        LogError "Unable to create Outlook recipient for '" & searchName & "'"
        On Error GoTo 0
        Exit Function
    End If
    recip.Resolve
    If Err.Number <> 0 Then
        LogError "Unable to resolve Outlook recipient for '" & searchName & "'"
        On Error GoTo 0
        Exit Function
    End If

    If recip.Resolved Then
        Set addressEntry = recip.AddressEntry
        If Err.Number <> 0 Or addressEntry Is Nothing Then
            LogError "Unable to read Outlook AddressEntry for resolved recipient '" & searchName & "'"
            On Error GoTo 0
            Exit Function
        End If

        resolvedAddress = ""
        Set exchUser = addressEntry.GetExchangeUser()
        If Err.Number = 0 Then
            If Not exchUser Is Nothing Then resolvedAddress = Trim(CStr(exchUser.PrimarySmtpAddress))
        Else
            Err.Clear
        End If

        If resolvedAddress = "" Then
            Set exchList = addressEntry.GetExchangeDistributionList()
            If Err.Number = 0 Then
                If Not exchList Is Nothing Then resolvedAddress = Trim(CStr(exchList.PrimarySmtpAddress))
            Else
                Err.Clear
            End If
        End If

        If resolvedAddress = "" Then
            resolvedAddress = Trim(CStr(addressEntry.Address))
            If Err.Number <> 0 Then
                Err.Clear
                resolvedAddress = ""
            End If
        End If

        GetExactEmail = resolvedAddress
    End If
    On Error GoTo 0
End Function

Function EmailExists(header, folderConstant, checkOnly)
    Call LogVerbose("EmailExists invoked for subject='" & header & "'")
    Dim targetFolderId, folderLabel, sortField
    Dim targetFolder, items, mail
    Dim found, checkMode
    Dim itemCount, itemIndex, itemType, itemSubject

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
        Case 4
            folderLabel = "Outbox"
            sortField = "CreationTime"
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
    If Err.Number <> 0 Or targetFolder Is Nothing Then
        LogError "Unable to access Outlook folder id " & targetFolderId & " for subject '" & header & "'"
        On Error GoTo 0
        EmailExists = Not checkMode
        Exit Function
    End If

    Set items = targetFolder.Items
    If Err.Number <> 0 Or items Is Nothing Then
        LogError "Unable to access Outlook Items collection in " & folderLabel & " for subject '" & header & "'"
        On Error GoTo 0
        EmailExists = Not checkMode
        Exit Function
    End If
    itemCount = CLng(items.Count)
    If Err.Number <> 0 Then
        LogError "Unable to count Outlook items in " & folderLabel & " for subject '" & header & "'"
        On Error GoTo 0
        EmailExists = Not checkMode
        Exit Function
    End If
    items.Sort sortField, True
    If Err.Number <> 0 Then
        LogError "Unable to sort Outlook items in " & folderLabel & " for subject '" & header & "'"
        On Error GoTo 0
        EmailExists = Not checkMode
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    found = False
    For itemIndex = 1 To itemCount
        Set mail = Nothing
        On Error Resume Next
        Set mail = items.Item(itemIndex)
        If Err.Number <> 0 Then
            LogError "Outlook " & folderLabel & " enumeration failed at item " & CStr(itemIndex) & _
                     " of " & CStr(itemCount) & " for subject '" & header & "'"
            On Error GoTo 0
            EmailExists = Not checkMode
            Exit Function
        End If
        itemType = TypeName(mail)
        If Err.Number <> 0 Then
            LogError "Unable to inspect Outlook " & folderLabel & " item type at item " & _
                     CStr(itemIndex) & " for subject '" & header & "'"
            On Error GoTo 0
            EmailExists = Not checkMode
            Exit Function
        End If
        If itemType = "MailItem" Then
            itemSubject = CStr(mail.Subject)
            If Err.Number <> 0 Then
                LogError "Unable to read Outlook " & folderLabel & " subject at item " & _
                         CStr(itemIndex) & " for subject '" & header & "'"
                On Error GoTo 0
                EmailExists = Not checkMode
                Exit Function
            End If
            On Error GoTo 0
            If itemSubject = header Then
                found = True
                If checkMode Then
                    LogEvent "Confirmed email with subject '" & header & "' in " & folderLabel & "."
                Else
                    LogEvent "Email with subject '" & header & "' already exists in " & folderLabel & ". Skipping related actions."
                End If
                Exit For
            End If
        Else
            On Error GoTo 0
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

Function EmailExistsInSentOrOutbox(header, checkOnly)
    Call LogVerbose("EmailExistsInSentOrOutbox invoked for subject='" & header & "'")
    Dim found

    found = EmailExists(header, 5, checkOnly)
    If found Then
        EmailExistsInSentOrOutbox = True
        Exit Function
    End If

    found = EmailExists(header, 4, checkOnly)
    EmailExistsInSentOrOutbox = found
End Function

'==============================================================
' Managed Email Identity and Lookup Helpers
'==============================================================
Function BoolText(value)
    If CBool(value) Then
        BoolText = "True"
    Else
        BoolText = "False"
    End If
End Function

Function FormatManagedEmailDate(meetingDate)
    FormatManagedEmailDate = Year(meetingDate) & "-" & _
                             Right("0" & Month(meetingDate), 2) & "-" & _
                             Right("0" & Day(meetingDate), 2)
End Function

Function NormalizeManagedEmailType(emailType)
    Dim cleaned
    cleaned = UCase(Trim(CStr(emailType)))
    cleaned = Replace(cleaned, " ", "-")
    cleaned = Replace(cleaned, "_", "-")
    Do While InStr(cleaned, "--") > 0
        cleaned = Replace(cleaned, "--", "-")
    Loop

    Select Case cleaned
        Case "LITERATURE", "LITERATURE-PRESENTATION", "LITERATURE-REMINDER"
            NormalizeManagedEmailType = "LITERATURE-REMINDER"
        Case "RESULTS", "RESULTS-PRESENTATION", "RESULTS-REMINDER"
            NormalizeManagedEmailType = "RESULTS-REMINDER"
        Case "MEETING-LEAD", "MEETING-LEAD-REMINDER"
            NormalizeManagedEmailType = "MEETING-LEAD-REMINDER"
        Case "LAB-MANAGER", "LAB-MANAGER-REMINDER"
            NormalizeManagedEmailType = "LAB-MANAGER-REMINDER"
        Case "SAME-DAY", "SAME-DAY-REMINDER"
            NormalizeManagedEmailType = "SAME-DAY-REMINDER"
        Case "REGULAR", "GROUP", "GROUP-MEETING"
            NormalizeManagedEmailType = "REGULAR"
        Case Else
            If cleaned = "" Then
                NormalizeManagedEmailType = "UNKNOWN"
            Else
                NormalizeManagedEmailType = cleaned
            End If
    End Select
End Function

Function BuildManagedEmailKey(meetingDate, emailType)
    BuildManagedEmailKey = MANAGED_EMAIL_KEY_PREFIX & "|" & _
                           FormatManagedEmailDate(meetingDate) & "|" & _
                           NormalizeManagedEmailType(emailType)
    LogEvent "Managed email key built: " & BuildManagedEmailKey
End Function

Function ManagedEmailMarker(managedKey)
    ManagedEmailMarker = "<!-- " & MANAGED_EMAIL_PROPERTY_NAME & ": " & managedKey & " -->"
End Function

Sub EnsureManagedEmailHtmlMarker(ByRef mail, managedKey)
    On Error Resume Next
    Dim html, marker, bodyClosePos
    marker = ManagedEmailMarker(managedKey)
    html = mail.HTMLBody
    If Err.Number <> 0 Then
        LogError "Unable to read HTMLBody while stamping managed email key '" & managedKey & "'"
        On Error GoTo 0
        Exit Sub
    End If

    If InStr(1, html, MANAGED_EMAIL_PROPERTY_NAME & ":", vbTextCompare) > 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    bodyClosePos = InStr(1, html, "</body>", vbTextCompare)
    If bodyClosePos > 0 Then
        html = Left(html, bodyClosePos - 1) & vbCrLf & marker & vbCrLf & Mid(html, bodyClosePos)
    Else
        html = html & vbCrLf & marker
    End If
    mail.HTMLBody = html
    If Err.Number <> 0 Then
        LogError "Unable to write hidden HTML managed email marker for key '" & managedKey & "'"
    End If
    On Error GoTo 0
End Sub

Sub StampManagedEmail(ByRef mail, managedKey)
    Call LogVerbose("StampManagedEmail invoked for key='" & managedKey & "'")
    On Error Resume Next
    Dim prop
    Set prop = mail.UserProperties.Find(MANAGED_EMAIL_PROPERTY_NAME)
    If Err.Number <> 0 Then
        LogError "Unable to read Outlook UserProperties while stamping managed email key '" & managedKey & "'"
        Set prop = Nothing
    End If
    Err.Clear
    If prop Is Nothing Then
        Set prop = mail.UserProperties.Add(MANAGED_EMAIL_PROPERTY_NAME, 1, True)
    End If
    If Err.Number <> 0 Then
        LogError "Unable to create Outlook UserProperty for managed email key '" & managedKey & "'"
    Else
        prop.Value = managedKey
        If Err.Number <> 0 Then
            LogError "Unable to set Outlook UserProperty for managed email key '" & managedKey & "'"
        End If
    End If
    On Error GoTo 0

    EnsureManagedEmailHtmlMarker mail, managedKey
    LogEvent "Stamped managed email with key: " & managedKey
End Sub

Function ExtractManagedEmailKeyFromText(textValue)
    Dim sourceText, markerPos, startPos, remainder, endPos, crPos, ltPos
    ExtractManagedEmailKeyFromText = ""
    sourceText = SafeTextValue(textValue, "")
    markerPos = InStr(1, sourceText, MANAGED_EMAIL_PROPERTY_NAME & ":", vbTextCompare)
    If markerPos = 0 Then Exit Function

    startPos = markerPos + Len(MANAGED_EMAIL_PROPERTY_NAME & ":")
    remainder = Trim(Mid(sourceText, startPos))
    endPos = InStr(1, remainder, "-->", vbTextCompare)
    crPos = InStr(1, remainder, vbCrLf, vbTextCompare)
    ltPos = InStr(1, remainder, "<", vbTextCompare)

    If endPos = 0 Or (crPos > 0 And crPos < endPos) Then endPos = crPos
    If endPos = 0 Or (ltPos > 0 And ltPos < endPos) Then endPos = ltPos

    If endPos > 0 Then
        ExtractManagedEmailKeyFromText = Trim(Left(remainder, endPos - 1))
    Else
        ExtractManagedEmailKeyFromText = Trim(remainder)
    End If
End Function

Function TryGetManagedEmailKey(ByRef mail, ByRef keyValue)
    On Error Resume Next
    Dim prop, valueText, htmlText, bodyText, readable
    keyValue = ""
    readable = False

    Set prop = mail.UserProperties.Find(MANAGED_EMAIL_PROPERTY_NAME)
    If Err.Number = 0 Then
        readable = True
        If Not prop Is Nothing Then
            valueText = Trim(SafeTextValue(prop.Value, ""))
            If valueText <> "" Then
                keyValue = valueText
                On Error GoTo 0
                TryGetManagedEmailKey = True
                Exit Function
            End If
        End If
    Else
        Err.Clear
    End If

    htmlText = mail.HTMLBody
    If Err.Number = 0 Then
        readable = True
        valueText = ExtractManagedEmailKeyFromText(htmlText)
        If valueText <> "" Then
            keyValue = valueText
            On Error GoTo 0
            TryGetManagedEmailKey = True
            Exit Function
        End If
    Else
        Err.Clear
    End If

    bodyText = mail.Body
    If Err.Number = 0 Then
        readable = True
        valueText = ExtractManagedEmailKeyFromText(bodyText)
        If valueText <> "" Then keyValue = valueText
    Else
        Err.Clear
    End If
    On Error GoTo 0

    TryGetManagedEmailKey = readable
End Function

Function GetManagedEmailKey(ByRef mail)
    Dim keyValue
    GetManagedEmailKey = ""
    If TryGetManagedEmailKey(mail, keyValue) Then
        GetManagedEmailKey = keyValue
    End If
End Function

Function TryGetMailSubject(ByRef mail, ByRef subjectValue)
    On Error Resume Next
    subjectValue = CStr(mail.Subject)
    If Err.Number <> 0 Then
        Err.Clear
        subjectValue = ""
        TryGetMailSubject = False
    Else
        TryGetMailSubject = True
    End If
    On Error GoTo 0
End Function

Function SafeMailSubject(ByRef mail)
    Dim subjectValue
    If TryGetMailSubject(mail, subjectValue) Then
        SafeMailSubject = subjectValue
    Else
        SafeMailSubject = ""
    End If
End Function

Function FolderLabelFromId(folderId)
    Select Case CLng(folderId)
        Case 4
            FolderLabelFromId = "Outbox"
        Case 5
            FolderLabelFromId = "Sent Items"
        Case 6
            FolderLabelFromId = "Inbox"
        Case Else
            FolderLabelFromId = "folder id " & folderId
    End Select
End Function

Function FolderSortFieldFromId(folderId)
    Select Case CLng(folderId)
        Case 4
            FolderSortFieldFromId = "CreationTime"
        Case 5
            FolderSortFieldFromId = "SentOn"
        Case Else
            FolderSortFieldFromId = "ReceivedTime"
    End Select
End Function

Function FolderScanLimitFromId(folderId, itemCount)
    FolderScanLimitFromId = CLng(itemCount)
    Select Case CLng(folderId)
        Case 4
            If FolderScanLimitFromId > MANAGED_OUTBOX_SCAN_LIMIT Then
                FolderScanLimitFromId = MANAGED_OUTBOX_SCAN_LIMIT
            End If
        Case 5
            If FolderScanLimitFromId > MANAGED_SENT_ITEMS_SCAN_LIMIT Then
                FolderScanLimitFromId = MANAGED_SENT_ITEMS_SCAN_LIMIT
            End If
    End Select
End Function

Function FindManagedEmailInFolder(folderId, managedKey, subjectText)
    Call LogVerbose("FindManagedEmailInFolder invoked for key='" & managedKey & "'")
    Dim matches, targetFolder, items, mail, itemKey, matched, folderLabel, sortField
    Dim itemCount, itemIndex, itemType, mailSubject, scanLimit
    Set matches = CreateObject("Scripting.Dictionary")
    folderLabel = FolderLabelFromId(folderId)
    sortField = FolderSortFieldFromId(folderId)

    LogEvent "Searching Outlook " & folderLabel & " for managed email key: " & managedKey

    On Error Resume Next
    Set targetFolder = olNS.GetDefaultFolder(CLng(folderId))
    If Err.Number <> 0 Or targetFolder Is Nothing Then
        Err.Clear
        LogEvent "WARNING: Initial Outlook " & folderLabel & " lookup failed for managed key '" & _
                 managedKey & "'. Refreshing Outlook session and retrying once."
        If TryRefreshOutlookSession("searching Outlook " & folderLabel & _
                                    " for managed key '" & managedKey & "'") Then
            Set targetFolder = olNS.GetDefaultFolder(CLng(folderId))
        End If
    End If
    If Err.Number <> 0 Or targetFolder Is Nothing Then
        MarkManagedEmailLookupUncertain "Unable to access Outlook " & folderLabel & _
                                        " while searching for managed email key '" & managedKey & "'"
        On Error GoTo 0
        Set FindManagedEmailInFolder = matches
        Exit Function
    End If
    Set items = targetFolder.Items
    If Err.Number <> 0 Or items Is Nothing Then
        MarkManagedEmailLookupUncertain "Unable to access Items collection in Outlook " & folderLabel & _
                                        " for managed email key '" & managedKey & "'"
        On Error GoTo 0
        Set FindManagedEmailInFolder = matches
        Exit Function
    End If
    itemCount = CLng(items.Count)
    If Err.Number <> 0 Then
        MarkManagedEmailLookupUncertain "Unable to count Outlook " & folderLabel & _
                                        " items for managed email key '" & managedKey & "'"
        On Error GoTo 0
        Set FindManagedEmailInFolder = matches
        Exit Function
    End If
    items.Sort sortField, True
    If Err.Number <> 0 Then
        MarkManagedEmailLookupUncertain "Unable to sort Outlook " & folderLabel & _
                                        " items for managed email key '" & managedKey & "'"
        On Error GoTo 0
        Set FindManagedEmailInFolder = matches
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    scanLimit = FolderScanLimitFromId(folderId, itemCount)
    If scanLimit < itemCount Then
        LogEvent "Bounded Outlook " & folderLabel & " scan to newest " & CStr(scanLimit) & _
                 " of " & CStr(itemCount) & " item(s) for managed key '" & managedKey & "'."
    End If

    For itemIndex = 1 To scanLimit
        Set mail = Nothing
        itemType = ""

        On Error Resume Next
        Set mail = items.Item(itemIndex)
        If Err.Number <> 0 Then
            MarkManagedEmailLookupUncertain "Outlook " & folderLabel & _
                                            " item enumeration failed at item " & CStr(itemIndex) & _
                                            " of " & CStr(scanLimit) & " for managed email key '" & managedKey & "'"
            On Error GoTo 0
            Exit For
        End If
        itemType = TypeName(mail)
        If Err.Number <> 0 Then
            MarkManagedEmailLookupUncertain "Unable to inspect Outlook " & folderLabel & _
                                            " item type at item " & CStr(itemIndex) & _
                                            " for managed email key '" & managedKey & "'"
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0

        If itemType = "MailItem" Then
            matched = False
            If Not TryGetManagedEmailKey(mail, itemKey) Then
                MarkManagedEmailLookupUncertain "Unable to read managed-email marker from Outlook " & _
                                                folderLabel & " item " & CStr(itemIndex) & _
                                                " for managed email key '" & managedKey & "'"
                Exit For
            End If
            If itemKey <> "" And itemKey = managedKey Then
                matched = True
            ElseIf ENABLE_MANAGED_EMAIL_KEY_FALLBACK_BY_SUBJECT Then
                If subjectText <> "" Then
                    If Not TryGetMailSubject(mail, mailSubject) Then
                        MarkManagedEmailLookupUncertain "Unable to read Outlook " & folderLabel & _
                                                        " item subject at item " & CStr(itemIndex) & _
                                                        " for managed email key '" & managedKey & "'"
                        Exit For
                    End If
                    If mailSubject = subjectText Then
                        matched = True
                        LogEvent "Legacy subject-only fallback matched in " & folderLabel & _
                                 " for managed key '" & managedKey & "' and subject '" & subjectText & "'."
                    End If
                End If
            End If

            If matched Then
                matches.Add CStr(matches.Count), mail
                If CLng(folderId) = 5 Then
                    LogEvent "Sent Items already proves managed key '" & managedKey & _
                             "' exists; stopping scan after first match."
                    Exit For
                End If
            End If
        End If
    Next

    If IsManagedEmailLookupUncertain() Then
        LogEvent "WARNING: " & folderLabel & " search for managed key '" & managedKey & _
                 "' ended uncertain after " & CStr(matches.Count) & " match(es). Treating lookup as fail-safe."
    Else
        LogEvent folderLabel & " search result for managed key '" & managedKey & "': " & matches.Count & " match(es)."
    End If
    Set FindManagedEmailInFolder = matches
End Function

Function FindManagedEmailInSentItems(managedKey, subjectText)
    Set FindManagedEmailInSentItems = FindManagedEmailInFolder(5, managedKey, subjectText)
End Function

Function FindManagedEmailInOutbox(managedKey, subjectText)
    Set FindManagedEmailInOutbox = FindManagedEmailInFolder(4, managedKey, subjectText)
End Function

Function FindManagedEmailState(managedKey, subjectText)
    Dim state, sentMatches, outboxMatches
    Set state = CreateObject("Scripting.Dictionary")
    ResetManagedEmailLookupUncertainty
    Set sentMatches = FindManagedEmailInSentItems(managedKey, subjectText)
    Set outboxMatches = FindManagedEmailInOutbox(managedKey, subjectText)
    state.Add "SentItems", sentMatches
    state.Add "Outbox", outboxMatches
    state.Add "LookupUncertain", IsManagedEmailLookupUncertain()
    state.Add "LookupUncertaintyReason", ManagedEmailLookupUncertaintyReason()
    Set FindManagedEmailState = state
End Function

'==============================================================
' Outlook Transport and Outbox Diagnostics
'==============================================================
Function ExchangeConnectionModeName(modeValue)
    If Not IsNumeric(modeValue) Then
        ExchangeConnectionModeName = "Unavailable"
        Exit Function
    End If

    Select Case CLng(modeValue)
        Case 0
            ExchangeConnectionModeName = "olNoExchange"
        Case 100
            ExchangeConnectionModeName = "olOffline"
        Case 200
            ExchangeConnectionModeName = "olCachedOffline"
        Case 300
            ExchangeConnectionModeName = "olDisconnected"
        Case 400
            ExchangeConnectionModeName = "olCachedDisconnected"
        Case 500
            ExchangeConnectionModeName = "olCachedConnectedHeaders"
        Case 600
            ExchangeConnectionModeName = "olCachedConnectedDrizzle"
        Case 700
            ExchangeConnectionModeName = "olCachedConnectedFull"
        Case 800
            ExchangeConnectionModeName = "olOnline"
        Case Else
            ExchangeConnectionModeName = "Unknown"
    End Select
End Function

Function IsBlockedExchangeMode(modeValue)
    If Not IsNumeric(modeValue) Then
        IsBlockedExchangeMode = False
        Exit Function
    End If

    Select Case CLng(modeValue)
        Case 100, 200, 300, 400
            IsBlockedExchangeMode = True
        Case Else
            IsBlockedExchangeMode = False
    End Select
End Function

Function GetOutlookTransportState(ByRef outlookApp)
    Dim state, ns, offlineValue, modeValue, offlineAvailable, modeAvailable, blocked
    Set state = CreateObject("Scripting.Dictionary")

    offlineValue = False
    modeValue = ""
    offlineAvailable = False
    modeAvailable = False

    On Error Resume Next
    Set ns = outlookApp.GetNamespace("MAPI")
    offlineValue = CBool(ns.Offline)
    If Err.Number = 0 Then
        offlineAvailable = True
    Else
        Err.Clear
    End If

    modeValue = ns.ExchangeConnectionMode
    If Err.Number = 0 Then
        modeAvailable = True
    Else
        modeValue = ""
        Err.Clear
    End If
    On Error GoTo 0

    blocked = False
    If offlineAvailable Then
        If offlineValue Then blocked = True
    End If
    If modeAvailable Then
        If IsBlockedExchangeMode(modeValue) Then blocked = True
    End If

    state.Add "OfflineAvailable", offlineAvailable
    state.Add "Offline", offlineValue
    state.Add "ModeAvailable", modeAvailable
    state.Add "Mode", modeValue
    state.Add "ModeName", ExchangeConnectionModeName(modeValue)
    state.Add "Blocked", blocked

    Set GetOutlookTransportState = state
End Function

Function LogOutlookTransportState(contextText, ByRef outlookApp)
    Dim state, offlineText, modeText
    Set state = GetOutlookTransportState(outlookApp)

    If ENABLE_OUTLOOK_TRANSPORT_DIAGNOSTICS Then
        If state("OfflineAvailable") Then
            offlineText = BoolText(state("Offline"))
        Else
            offlineText = "Unavailable"
        End If

        If state("ModeAvailable") Then
            modeText = CStr(state("Mode")) & " (" & state("ModeName") & ")"
        Else
            modeText = "Unavailable"
        End If

        LogEvent "Outlook transport state [" & contextText & "]: Offline=" & offlineText & _
                 " | ExchangeConnectionMode=" & modeText & _
                 " | offline/disconnected/auth-blocked suspected=" & BoolText(state("Blocked"))
    End If

    Set LogOutlookTransportState = state
End Function

Sub LogOutlookTransportStateSnapshot(contextText, ByRef outlookApp)
    Dim ignoredState
    Set ignoredState = LogOutlookTransportState(contextText, outlookApp)
End Sub

Function TryGetMapiProperty(ByRef mail, schemaName, ByRef propertyValue)
    On Error Resume Next
    propertyValue = mail.PropertyAccessor.GetProperty(schemaName)
    If Err.Number <> 0 Then
        LogEvent "WARNING: MAPI property lookup failed for " & schemaName & _
                 " | Err: " & Err.Number & " - " & Err.Description
        Err.Clear
        TryGetMapiProperty = False
    ElseIf IsEmpty(propertyValue) Or IsNull(propertyValue) Then
        TryGetMapiProperty = False
    Else
        TryGetMapiProperty = True
    End If
    On Error GoTo 0
End Function

Function TryGetEffectiveDeferredDeliveryTime(ByRef mail, ByRef deferredValue)
    On Error Resume Next
    deferredValue = mail.DeferredDeliveryTime
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        TryGetEffectiveDeferredDeliveryTime = False
        Exit Function
    End If
    On Error GoTo 0

    If Not IsDate(deferredValue) Then
        TryGetEffectiveDeferredDeliveryTime = False
        Exit Function
    End If

    deferredValue = CDate(deferredValue)
    If Year(deferredValue) < 1900 Or Year(deferredValue) >= 4500 Then
        TryGetEffectiveDeferredDeliveryTime = False
    Else
        TryGetEffectiveDeferredDeliveryTime = True
    End If
End Function

Function GetManagedMailReferenceTime(ByRef mail)
    On Error Resume Next
    GetManagedMailReferenceTime = mail.CreationTime
    If Err.Number = 0 And IsDate(GetManagedMailReferenceTime) Then
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear

    GetManagedMailReferenceTime = mail.LastModificationTime
    If Err.Number <> 0 Or Not IsDate(GetManagedMailReferenceTime) Then
        Err.Clear
        GetManagedMailReferenceTime = Empty
    End If
    On Error GoTo 0
End Function

Function GetManagedMailAgeMinutes(ByRef mail, currentTime)
    Dim referenceTime
    referenceTime = GetManagedMailReferenceTime(mail)
    If IsDate(referenceTime) Then
        GetManagedMailAgeMinutes = DateDiff("n", CDate(referenceTime), CDate(currentTime))
    Else
        GetManagedMailAgeMinutes = -1
    End If
End Function

Function SafeOutlookText(ByRef mail, propertyName)
    On Error Resume Next
    Dim value
    Select Case propertyName
        Case "Subject"
            value = mail.Subject
        Case "To"
            value = mail.To
        Case "EntryID"
            value = mail.EntryID
        Case "CreationTime"
            value = mail.CreationTime
        Case "LastModificationTime"
            value = mail.LastModificationTime
        Case Else
            value = ""
    End Select

    If Err.Number <> 0 Then
        Err.Clear
        SafeOutlookText = "[unavailable]"
    Else
        SafeOutlookText = SafeTextValue(value, "[unavailable]")
    End If
    On Error GoTo 0
End Function

Function SafeAttachmentCount(ByRef mail)
    On Error Resume Next
    SafeAttachmentCount = CStr(mail.Attachments.Count)
    If Err.Number <> 0 Then
        Err.Clear
        SafeAttachmentCount = "[unavailable]"
    End If
    On Error GoTo 0
End Function

Sub LogOutboxDiagnosticSnapshot(ByRef mail, managedKey)
    If Not ENABLE_OUTBOX_DIAGNOSTICS Then Exit Sub

    Dim deferredValue, deferredText, ageMinutes, flagsValue, submitTimeValue
    Dim flagsText, submitText, submitFlagText, unsentFlagText

    ageMinutes = GetManagedMailAgeMinutes(mail, Now)
    deferredText = "Absent"
    If TryGetEffectiveDeferredDeliveryTime(mail, deferredValue) Then
        If CDate(deferredValue) > Now Then
            deferredText = CStr(deferredValue) & " (future)"
        Else
            deferredText = CStr(deferredValue) & " (passed)"
        End If
    End If

    flagsText = "[unavailable]"
    submitFlagText = "[unavailable]"
    unsentFlagText = "[unavailable]"
    If TryGetMapiProperty(mail, "http://schemas.microsoft.com/mapi/proptag/0x0E070003", flagsValue) Then
        flagsText = SafeTextValue(flagsValue, "[unavailable]")
        If IsNumeric(flagsValue) Then
            submitFlagText = BoolText((CLng(flagsValue) And 4) <> 0)
            unsentFlagText = BoolText((CLng(flagsValue) And 8) <> 0)
        End If
    End If

    submitText = "[unavailable]"
    If TryGetMapiProperty(mail, "http://schemas.microsoft.com/mapi/proptag/0x00390040", submitTimeValue) Then
        submitText = SafeTextValue(submitTimeValue, "[unavailable]")
    End If

    LogEvent "Outbox diagnostic snapshot | ManagedKey=" & managedKey & _
             " | Subject=" & SafeOutlookText(mail, "Subject") & _
             " | To=" & SafeOutlookText(mail, "To") & _
             " | EntryID=" & SafeOutlookText(mail, "EntryID") & _
             " | CreationTime=" & SafeOutlookText(mail, "CreationTime") & _
             " | LastModificationTime=" & SafeOutlookText(mail, "LastModificationTime") & _
             " | AgeMinutes=" & CStr(ageMinutes) & _
             " | DeferredDeliveryTime=" & deferredText & _
             " | AttachmentCount=" & SafeAttachmentCount(mail) & _
             " | PR_MESSAGE_FLAGS=" & flagsText & _
             " | MSGFLAG_SUBMIT=" & submitFlagText & _
             " | MSGFLAG_UNSENT=" & unsentFlagText & _
             " | PR_CLIENT_SUBMIT_TIME=" & submitText
End Sub

Function IsItemInOutbox(ByRef mail)
    On Error Resume Next
    Dim outbox
    Set outbox = olNS.GetDefaultFolder(4)
    IsItemInOutbox = (CStr(mail.Parent.EntryID) = CStr(outbox.EntryID))
    If Err.Number <> 0 Then
        Err.Clear
        IsItemInOutbox = False
    End If
    On Error GoTo 0
End Function

Function NewQueueDecision(decisionText, reasonText)
    Dim decision
    Set decision = CreateObject("Scripting.Dictionary")
    decision.Add "Decision", decisionText
    decision.Add "Reason", reasonText
    Set NewQueueDecision = decision
End Function

Function GetQueuedManagedEmailDecision(ByRef outlookApp, ByRef queuedMail, managedKey, currentTime, graceMinutes, expectedDeferredDeliveryTime)
    Dim transportState, deferredValue, ageMinutes, sentMatches
    Set transportState = LogOutlookTransportState("queued managed email " & managedKey, outlookApp)

    If transportState("Blocked") Then
        Set GetQueuedManagedEmailDecision = NewQueueDecision("KEEP_AND_ALERT", _
            "Outlook appears offline/disconnected/auth-blocked; user action required.")
        Exit Function
    End If

    If TryGetEffectiveDeferredDeliveryTime(queuedMail, deferredValue) Then
        If CDate(deferredValue) > CDate(currentTime) Then
            Set GetQueuedManagedEmailDecision = NewQueueDecision("KEEP_WAITING", _
                "Deferred delivery time has not been reached.")
            Exit Function
        End If
    End If

    ageMinutes = GetManagedMailAgeMinutes(queuedMail, currentTime)
    If ageMinutes < 0 Then
        Set GetQueuedManagedEmailDecision = NewQueueDecision("KEEP_AND_ALERT", _
            "Unable to determine queued item age safely; manual review required.")
        Exit Function
    End If

    If ageMinutes < CLng(graceMinutes) Then
        Set GetQueuedManagedEmailDecision = NewQueueDecision("KEEP_WAITING", _
            "Recently queued; still within grace period.")
        Exit Function
    End If

    Set sentMatches = FindManagedEmailInSentItems(managedKey, SafeMailSubject(queuedMail))
    If sentMatches.Count > 0 Then
        Set GetQueuedManagedEmailDecision = NewQueueDecision("ALREADY_SENT", _
            "Email reached Sent Items during recheck.")
        Exit Function
    End If
    If IsManagedEmailLookupUncertain() Then
        Set GetQueuedManagedEmailDecision = NewQueueDecision("KEEP_AND_ALERT", _
            "Sent Items recheck is uncertain; stale Outbox replacement is not safe.")
        Exit Function
    End If

    If Not ENABLE_STALE_OUTBOX_REPLACEMENT Then
        Set GetQueuedManagedEmailDecision = NewQueueDecision("KEEP_AND_ALERT", _
            "Managed email appears stale, but stale Outbox replacement is disabled.")
        Exit Function
    End If

    Set GetQueuedManagedEmailDecision = NewQueueDecision("REPLACE_OUTBOX_ITEM", _
        "Managed email is stale in Outbox while Outlook appears connected and delivery time has passed.")
End Function

'==============================================================
' Managed Email Dispatch State Machine
'==============================================================
Function GetFirstManagedEmailMatch(ByRef matches)
    Dim matchKey
    Set GetFirstManagedEmailMatch = Nothing
    For Each matchKey In matches.Keys
        Set GetFirstManagedEmailMatch = matches.Item(matchKey)
        Exit Function
    Next
End Function

Function FindNewestManagedOutboxItem(ByRef matches)
    Dim matchKey, mail, newest, mailTime, newestTime
    Set newest = Nothing

    For Each matchKey In matches.Keys
        Set mail = matches.Item(matchKey)
        If newest Is Nothing Then
            Set newest = mail
        Else
            mailTime = GetManagedMailReferenceTime(mail)
            newestTime = GetManagedMailReferenceTime(newest)
            If IsDate(mailTime) And IsDate(newestTime) Then
                If CDate(mailTime) > CDate(newestTime) Then Set newest = mail
            End If
        End If
    Next

    Set FindNewestManagedOutboxItem = newest
End Function

Function AllMatchesHaveExpectedManagedKey(ByRef matches, managedKey)
    Dim matchKey, mail
    AllMatchesHaveExpectedManagedKey = True
    For Each matchKey In matches.Keys
        Set mail = matches.Item(matchKey)
        If GetManagedEmailKey(mail) <> managedKey Then
            AllMatchesHaveExpectedManagedKey = False
            Exit Function
        End If
    Next
End Function

Function ConsolidateManagedOutboxMatches(ByRef matches, managedKey, subjectText)
    Dim result, sentMatches, keptMail, matchKey, mail
    Set ConsolidateManagedOutboxMatches = Nothing

    If matches.Count <= 1 Then
        Set ConsolidateManagedOutboxMatches = matches
        Exit Function
    End If

    LogEvent "Multiple matching managed Outbox items detected for key '" & managedKey & "': " & matches.Count

    If Not AllMatchesHaveExpectedManagedKey(matches, managedKey) Then
        LogEvent "ERROR: Multiple Outbox matches include legacy subject-only or unkeyed items. No cleanup or new send will occur; manual Outlook review required for subject '" & subjectText & "'."
        Exit Function
    End If

    Set sentMatches = FindManagedEmailInSentItems(managedKey, subjectText)
    If sentMatches.Count > 0 Then
        LogEvent "Queue decision: ALREADY_SENT"
        LogEvent "Reason for decision: Email reached Sent Items during duplicate-Outbox safety recheck."
        Set result = CreateObject("Scripting.Dictionary")
        Set ConsolidateManagedOutboxMatches = result
        Exit Function
    End If
    If IsManagedEmailLookupUncertain() Then
        LogEvent "ERROR: Duplicate-Outbox consolidation aborted because Sent Items recheck is uncertain."
        Exit Function
    End If

    Set keptMail = FindNewestManagedOutboxItem(matches)
    If keptMail Is Nothing Then
        LogEvent "ERROR: Unable to choose a newest managed Outbox item for key '" & managedKey & "'. Manual review required."
        Exit Function
    End If

    For Each matchKey In matches.Keys
        Set mail = matches.Item(matchKey)
        If Not mail Is keptMail Then
            If GetManagedEmailKey(mail) = managedKey And IsItemInOutbox(mail) Then
                On Error Resume Next
                LogEvent "Deleting older duplicate managed Outbox item for key '" & managedKey & "' | EntryID=" & SafeOutlookText(mail, "EntryID")
                mail.Delete
                If Err.Number <> 0 Then
                    LogError "Failed to delete older duplicate managed Outbox item for key '" & managedKey & "'"
                    On Error GoTo 0
                    Exit Function
                End If
                On Error GoTo 0
            Else
                LogEvent "ERROR: Older duplicate failed ownership/folder guard for key '" & managedKey & "'. Manual review required."
                Exit Function
            End If
        End If
    Next

    Set result = CreateObject("Scripting.Dictionary")
    result.Add "0", keptMail
    LogEvent "Consolidated duplicate managed Outbox items for key '" & managedKey & "'; kept newest active copy."
    Set ConsolidateManagedOutboxMatches = result
End Function

Function CreateManagedEmailFromSpec(ByRef outlookApp, managedKey, subjectText, toAddress, bodyText, deferredDeliveryTime)
    Set CreateManagedEmailFromSpec = CreateManagedEmailFromSpecWithCc(outlookApp, managedKey, subjectText, toAddress, "", bodyText, deferredDeliveryTime)
End Function

Function CreateManagedEmailFromSpecWithCc(ByRef outlookApp, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime)
    Dim mail
    Set mail = outlookApp.CreateItem(0)
    With mail
        .To = toAddress
        If Trim(CStr(ccAddress)) <> "" Then .CC = ccAddress
        .Subject = subjectText
        If IsHtmlBody(bodyText) Then
            .BodyFormat = 2
            .HTMLBody = bodyText
        Else
            .Body = bodyText
        End If
        If IsDate(deferredDeliveryTime) Then
            .DeferredDeliveryTime = CDate(deferredDeliveryTime)
        End If
    End With
    StampManagedEmail mail, managedKey
    Set CreateManagedEmailFromSpecWithCc = mail
End Function

Sub VerifyManagedEmailAfterSend(managedKey, subjectText)
    Dim sentMatches, outboxMatches
    ResetManagedEmailLookupUncertainty
    Set sentMatches = FindManagedEmailInSentItems(managedKey, subjectText)
    Set outboxMatches = FindManagedEmailInOutbox(managedKey, subjectText)

    LogEvent "Post-send verification result for managed key '" & managedKey & "': SentItems=" & _
             sentMatches.Count & " | Outbox=" & outboxMatches.Count

    If sentMatches.Count > 0 Or outboxMatches.Count > 0 Then
        LogEvent "Managed email observed in Sent Items or Outbox after send attempt for key '" & managedKey & "'."
    ElseIf IsManagedEmailLookupUncertain() Then
        LogEvent "ACTION REQUIRED: Post-send verification is uncertain for managed key '" & managedKey & _
                 "'. " & ManagedEmailLookupUncertaintyReason()
    Else
        LogEvent "WARNING: Managed email not observed in Sent Items or Outbox after send attempt for key '" & managedKey & "'."
    End If
End Sub

Sub VerifyManagedEmailAfterSendIfUrgent(isUrgentSend, managedKey, subjectText)
    If CBool(isUrgentSend) Then
        VerifyManagedEmailAfterSend managedKey, subjectText
    Else
        LogEvent "Immediate post-send verification skipped for non-urgent managed email key '" & managedKey & "'. Next script run will check Sent Items or Outbox by managed key."
    End If
End Sub

Function SafeToReplaceManagedOutboxItem(ByRef outlookApp, ByRef queuedMail, managedKey, subjectText, currentTime, graceMinutes)
    Dim sentMatches, transportState, deferredValue, ageMinutes
    SafeToReplaceManagedOutboxItem = False

    If GetManagedEmailKey(queuedMail) <> managedKey Then
        LogEvent "Replacement guard failed: queued item does not have expected managed key '" & managedKey & "'."
        Exit Function
    End If

    If Not IsItemInOutbox(queuedMail) Then
        LogEvent "Replacement guard failed: managed item is no longer in Outlook Outbox for key '" & managedKey & "'."
        Exit Function
    End If

    Set sentMatches = FindManagedEmailInSentItems(managedKey, subjectText)
    If sentMatches.Count > 0 Then
        LogEvent "Replacement guard failed: Sent Items now contains managed key '" & managedKey & "'."
        Exit Function
    End If
    If IsManagedEmailLookupUncertain() Then
        LogEvent "Replacement guard failed: Sent Items recheck is uncertain for key '" & managedKey & _
                 "'. " & ManagedEmailLookupUncertaintyReason()
        Exit Function
    End If

    Set transportState = LogOutlookTransportState("pre-replacement guard " & managedKey, outlookApp)
    If transportState("Blocked") Then
        LogEvent "Replacement guard failed: Outlook appears offline/disconnected/auth-blocked for key '" & managedKey & "'."
        Exit Function
    End If

    ageMinutes = GetManagedMailAgeMinutes(queuedMail, currentTime)
    If ageMinutes < CLng(graceMinutes) Then
        LogEvent "Replacement guard failed: queued managed email age " & CStr(ageMinutes) & _
                 " minute(s) is below grace period " & CStr(graceMinutes) & " for key '" & managedKey & "'."
        Exit Function
    End If

    If TryGetEffectiveDeferredDeliveryTime(queuedMail, deferredValue) Then
        If CDate(deferredValue) > CDate(currentTime) Then
            LogEvent "Replacement guard failed: deferred delivery time is still in the future for key '" & managedKey & "'."
            Exit Function
        End If
    End If

    If Not ENABLE_STALE_OUTBOX_REPLACEMENT Then
        LogEvent "Replacement guard failed: stale Outbox replacement is disabled."
        Exit Function
    End If

    SafeToReplaceManagedOutboxItem = True
End Function

Function ReplaceStaleManagedOutboxItem(ByRef outlookApp, ByRef queuedMail, managedKey, subjectText, toAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
    ReplaceStaleManagedOutboxItem = ReplaceStaleManagedOutboxItemWithCc(outlookApp, queuedMail, managedKey, subjectText, toAddress, "", bodyText, deferredDeliveryTime, emailType, isUrgentSend)
End Function

Function ReplaceStaleManagedOutboxItemWithCc(ByRef outlookApp, ByRef queuedMail, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
    Dim replacement
    ReplaceStaleManagedOutboxItemWithCc = "REPLACE_ABORTED"

    If Not SafeToReplaceManagedOutboxItem(outlookApp, queuedMail, managedKey, subjectText, Now, MANAGED_OUTBOX_GRACE_MINUTES) Then
        LogEvent "ERROR: Replacement aborted by guardrails for managed key '" & managedKey & "'."
        Exit Function
    End If

    LogEvent "Managed email appears stale in Outbox while Outlook is connected and delivery time has passed. Replacing stale managed Outbox item."

    On Error Resume Next
    LogEvent "Deleting stale managed Outbox item for key '" & managedKey & "' | EntryID=" & SafeOutlookText(queuedMail, "EntryID")
    queuedMail.Delete
    If Err.Number <> 0 Then
        LogError "Failed to delete stale managed Outbox item for key '" & managedKey & "'"
        On Error GoTo 0
        ReplaceStaleManagedOutboxItemWithCc = "DELETE_ERROR"
        Exit Function
    End If
    On Error GoTo 0

    Set replacement = CreateManagedEmailFromSpecWithCc(outlookApp, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime)
    LogEvent "Created clean replacement managed email for key '" & managedKey & "' and type '" & emailType & "'."
    LogOutlookTransportStateSnapshot "before replacement send " & managedKey, outlookApp
    On Error Resume Next
    replacement.Send
    If Err.Number <> 0 Then
        LogError "Failed to send replacement managed email for key '" & managedKey & "'"
        On Error GoTo 0
        ReplaceStaleManagedOutboxItemWithCc = "SEND_ERROR"
        Exit Function
    End If
    On Error GoTo 0
    LogOutlookTransportStateSnapshot "after replacement send " & managedKey, outlookApp
    LogEvent "Replaced stale email and submitted replacement for managed key '" & managedKey & "'."
    VerifyManagedEmailAfterSendIfUrgent isUrgentSend, managedKey, subjectText
    ReplaceStaleManagedOutboxItemWithCc = "REPLACED"
End Function

Function DispatchManagedEmail(ByRef outlookApp, managedKey, subjectText, toAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
    DispatchManagedEmail = DispatchManagedEmailWithCc(outlookApp, managedKey, subjectText, toAddress, "", bodyText, deferredDeliveryTime, emailType, isUrgentSend)
End Function

Function DispatchManagedEmailWithCc(ByRef outlookApp, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
    DispatchManagedEmailWithCc = DispatchManagedEmailInternal(outlookApp, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
End Function

Function DispatchManagedEmailInternal(ByRef outlookApp, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
    Call LogVerbose("DispatchManagedEmail invoked for key='" & managedKey & "'")
    Dim sentMatches, outboxMatches, consolidatedMatches, queuedMail, mail, decision, queuedKey
    Dim ccLogText

    DispatchManagedEmailInternal = "BLOCKED"
    ccLogText = ""
    If Trim(CStr(ccAddress)) <> "" Then ccLogText = " | Cc: " & ccAddress

    LogEvent "Preparing managed email | Key: " & managedKey & " | Type: " & emailType & _
             " | Subject: " & subjectText & " | To: " & toAddress & ccLogText
    LogOutlookTransportStateSnapshot "before managed dispatch " & managedKey, outlookApp

    ResetManagedEmailLookupUncertainty
    Set sentMatches = FindManagedEmailInSentItems(managedKey, subjectText)
    If sentMatches.Count > 0 Then
        LogEvent "Queue decision: ALREADY_SENT"
        LogEvent "Reason for decision: Matching managed email exists in Sent Items. Do not recreate or resend."
        DispatchManagedEmailInternal = "ALREADY_SENT"
        Exit Function
    End If
    If IsManagedEmailLookupUncertain() Then
        LogManagedEmailLookupBlocked "Managed email dispatch", managedKey
        DispatchManagedEmailInternal = "LOOKUP_UNCERTAIN"
        Exit Function
    End If

    Set outboxMatches = FindManagedEmailInOutbox(managedKey, subjectText)
    If IsManagedEmailLookupUncertain() Then
        LogManagedEmailLookupBlocked "Managed email dispatch", managedKey
        DispatchManagedEmailInternal = "LOOKUP_UNCERTAIN"
        Exit Function
    End If
    ' Recheck in case this run continued past midnight into the weekend.
    If IsWeekend(Date) Then
        LogEvent "Weekend: email delivery skipped. Subject: " & .Subject
        Exit Sub
    End If
    If outboxMatches.Count = 0 Then
        If DEBUG_MODE Then
            LogEvent "DEBUG_MODE enabled - managed email not sent. Key: " & managedKey & " | Subject: " & subjectText
            If DISPLAY_DEBUG_POPUPS Then
                MsgBox "[DEBUG] Would send managed email to " & toAddress & ":" & vbCrLf & _
                       "Key: " & managedKey & vbCrLf & "Subject: " & subjectText & vbCrLf & vbCrLf & bodyText
            Else
                LogVerbose "Managed email preview suppressed for key='" & managedKey & "'"
            End If
            DispatchManagedEmailInternal = "DEBUG_SKIPPED"
            Exit Function
        End If

        Set mail = CreateManagedEmailFromSpecWithCc(outlookApp, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime)
        LogEvent "Created new managed email for key '" & managedKey & "'."
        LogOutlookTransportStateSnapshot "before send " & managedKey, outlookApp
        On Error Resume Next
        mail.Send
        If Err.Number <> 0 Then
            LogError "Failed to send managed email for key '" & managedKey & "'"
            On Error GoTo 0
            DispatchManagedEmailInternal = "SEND_ERROR"
            Exit Function
        End If
        On Error GoTo 0
        LogOutlookTransportStateSnapshot "after send " & managedKey, outlookApp
        LogEvent "Submitted managed email for key '" & managedKey & "'."
        VerifyManagedEmailAfterSendIfUrgent isUrgentSend, managedKey, subjectText
        DispatchManagedEmailInternal = "SUBMITTED"
        Exit Function
    End If

    LogEvent "Outbox search result for managed key '" & managedKey & "': active queued attempt detected."
    Set consolidatedMatches = ConsolidateManagedOutboxMatches(outboxMatches, managedKey, subjectText)
    If consolidatedMatches Is Nothing Then
        LogEvent "Queue decision: KEEP_AND_ALERT"
        LogEvent "Reason for decision: Multiple Outbox matches could not be safely consolidated; manual review required."
        LogEvent "ACTION REQUIRED: Managed email has ambiguous Outbox copies. Keeping existing Outbox item(s); review Outlook before rerunning automation."
        DispatchManagedEmailInternal = "KEEP_AND_ALERT"
        Exit Function
    End If

    If consolidatedMatches.Count = 0 Then
        DispatchManagedEmailInternal = "ALREADY_SENT"
        Exit Function
    End If

    Set queuedMail = GetFirstManagedEmailMatch(consolidatedMatches)
    queuedKey = GetManagedEmailKey(queuedMail)
    LogOutboxDiagnosticSnapshot queuedMail, managedKey

    If queuedKey <> managedKey Then
        LogEvent "Queue decision: KEEP_WAITING"
        LogEvent "Reason for decision: Legacy subject-only Outbox match has no managed key; treating it as an active send attempt but not deleting or replacing it."
        DispatchManagedEmailInternal = "KEEP_WAITING"
        Exit Function
    End If

    Set decision = GetQueuedManagedEmailDecision(outlookApp, queuedMail, managedKey, Now, MANAGED_OUTBOX_GRACE_MINUTES, deferredDeliveryTime)
    LogEvent "Queue decision: " & decision("Decision")
    LogEvent "Reason for decision: " & decision("Reason")
    DispatchManagedEmailInternal = decision("Decision")

    Select Case decision("Decision")
        Case "ALREADY_SENT"
            LogEvent "Managed email has already reached Sent Items. No recreate or resend for key '" & managedKey & "'."
        Case "KEEP_AND_ALERT"
            If InStr(1, decision("Reason"), "offline/disconnected/auth-blocked", vbTextCompare) > 0 Then
                LogEvent "ACTION REQUIRED: Managed email is queued because Outlook appears offline/disconnected/auth-blocked. Keeping existing Outbox item. User should reconnect Outlook, disable Work Offline, or complete O365 authentication."
            Else
                LogEvent "ACTION REQUIRED: Managed email remains queued and requires manual review. Keeping existing Outbox item."
            End If
        Case "KEEP_WAITING"
            If InStr(1, decision("Reason"), "grace period", vbTextCompare) > 0 Then
                LogEvent "Queued managed email is still within grace period. Keeping existing Outbox item."
            Else
                LogEvent "Queued managed email is waiting for Outlook delivery conditions. Keeping existing Outbox item."
            End If
        Case "REPLACE_OUTBOX_ITEM"
            DispatchManagedEmailInternal = ReplaceStaleManagedOutboxItemWithCc(outlookApp, queuedMail, managedKey, subjectText, toAddress, ccAddress, bodyText, deferredDeliveryTime, emailType, isUrgentSend)
        Case Else
            LogEvent "ERROR: Unknown managed email queue decision '" & decision("Decision") & "' for key '" & managedKey & "'."
            DispatchManagedEmailInternal = "UNKNOWN_DECISION"
    End Select
End Function

Function HasAlertBeenSentToday(subjectText)
    Call LogVerbose("HasAlertBeenSentToday invoked for subject='" & subjectText & "'")
    Dim sentItems, mail, itemCount, itemIndex, itemType, itemSubject, itemSentOn

    On Error Resume Next
    Set sentItems = olNS.GetDefaultFolder(5).Items
    If Err.Number <> 0 Or sentItems Is Nothing Then
        LogError "Unable to access Sent Items for alert subject '" & subjectText & "'"
        On Error GoTo 0
        LogEvent "ACTION REQUIRED: Alert duplicate check is uncertain; suppressing alert to avoid duplicate subject '" & subjectText & "'."
        HasAlertBeenSentToday = True
        Exit Function
    End If

    sentItems.Sort "SentOn", True
    If Err.Number <> 0 Then
        LogError "Unable to sort Sent Items for alert subject '" & subjectText & "'"
        On Error GoTo 0
        LogEvent "ACTION REQUIRED: Alert duplicate check is uncertain; suppressing alert to avoid duplicate subject '" & subjectText & "'."
        HasAlertBeenSentToday = True
        Exit Function
    End If
    itemCount = CLng(sentItems.Count)
    If Err.Number <> 0 Then
        LogError "Unable to count Sent Items for alert subject '" & subjectText & "'"
        On Error GoTo 0
        LogEvent "ACTION REQUIRED: Alert duplicate check is uncertain; suppressing alert to avoid duplicate subject '" & subjectText & "'."
        HasAlertBeenSentToday = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    HasAlertBeenSentToday = False
    For itemIndex = 1 To itemCount
        Set mail = Nothing
        On Error Resume Next
        Set mail = sentItems.Item(itemIndex)
        If Err.Number <> 0 Then
            LogError "Sent Items enumeration failed at item " & CStr(itemIndex) & _
                     " of " & CStr(itemCount) & " for alert subject '" & subjectText & "'"
            On Error GoTo 0
            LogEvent "ACTION REQUIRED: Alert duplicate check is uncertain; suppressing alert to avoid duplicate subject '" & subjectText & "'."
            HasAlertBeenSentToday = True
            Exit Function
        End If
        itemType = TypeName(mail)
        If Err.Number <> 0 Then
            LogError "Unable to inspect Sent Items item type at item " & CStr(itemIndex) & _
                     " for alert subject '" & subjectText & "'"
            On Error GoTo 0
            LogEvent "ACTION REQUIRED: Alert duplicate check is uncertain; suppressing alert to avoid duplicate subject '" & subjectText & "'."
            HasAlertBeenSentToday = True
            Exit Function
        End If
        If itemType = "MailItem" Then
            itemSentOn = mail.SentOn
            itemSubject = CStr(mail.Subject)
            If Err.Number <> 0 Then
                LogError "Unable to read Sent Items subject/date at item " & CStr(itemIndex) & _
                         " for alert subject '" & subjectText & "'"
                On Error GoTo 0
                LogEvent "ACTION REQUIRED: Alert duplicate check is uncertain; suppressing alert to avoid duplicate subject '" & subjectText & "'."
                HasAlertBeenSentToday = True
                Exit Function
            End If
            On Error GoTo 0
            If DateValue(itemSentOn) < Date Then Exit For
            If itemSubject = subjectText And DateValue(itemSentOn) = Date Then
                LogEvent "Alert with subject '" & subjectText & "' already sent today."
                HasAlertBeenSentToday = True
                Exit For
            End If
        Else
            On Error GoTo 0
        End If
    Next

    LogEvent "HasAlertBeenSentToday('" & subjectText & "') result: " & HasAlertBeenSentToday
End Function

'==============================================================
' Meeting Content Builders
'==============================================================
Function DeriveMeetingLeadName(objSheetRota, meetingDate)
    Call LogVerbose("DeriveMeetingLeadName invoked")
    Dim currentLabManager, currentFridayRow, searchRow, lastRow, candidateDate

    currentFridayRow = -1
    lastRow = objSheetRota.Cells(objSheetRota.Rows.Count, "A").End(-4162).Row

    For searchRow = 2 To lastRow
        candidateDate = objSheetRota.Cells(searchRow, 1).Value
        If candidateDate = meetingDate Then
            currentFridayRow = searchRow
            Exit For
        End If
    Next

    If currentFridayRow <> -1 Then
        currentLabManager = Trim(objSheetRota.Cells(currentFridayRow, 6).Value)
    Else
        currentLabManager = ""
    End If

    DeriveMeetingLeadName = currentLabManager
    LogEvent "Derived meeting lead as current lab responsible '" & currentLabManager & "' for meeting date " & FormatDateTime(meetingDate, 2)
End Function

Function BuildMeetingSubject(meetingDate, room)
    Call LogVerbose("BuildMeetingSubject invoked")
    Dim roomText, subjectSeparator, meetingDateText
    roomText = Trim(SafeTextValue(room, ""))
    If roomText = "" Then roomText = "TBC"
    subjectSeparator = " " & ChrW(&HB7) & " "
    meetingDateText = WeekdayName(Weekday(meetingDate), False) & " " & FormatDateTime(meetingDate, 2)
    BuildMeetingSubject = "Group Meeting" & subjectSeparator & meetingDateText & subjectSeparator & "11:00" & subjectSeparator & roomText
End Function

Function BuildGroupEmailSubject(meetingDate, alternateText, room)
    Call LogVerbose("BuildGroupEmailSubject invoked")
    If IsNoMeetingEvent(alternateText) Then
        BuildGroupEmailSubject = "No Group Meeting " & ChrW(&HB7) & " " & WeekdayName(Weekday(meetingDate), False) & " " & FormatDateTime(meetingDate, 2)
    Else
        BuildGroupEmailSubject = BuildMeetingSubject(meetingDate, room)
    End If
End Function

Function IsCollegeClosureEvent(alternateText)
    Dim normalizedAlternate
    normalizedAlternate = LCase(Trim(CStr(alternateText)))
    IsCollegeClosureEvent = (InStr(normalizedAlternate, "college closure") > 0)
End Function

Function IsNoMeetingEvent(alternateText)
    Dim normalizedAlternate
    normalizedAlternate = LCase(Trim(CStr(alternateText)))
    IsNoMeetingEvent = (InStr(normalizedAlternate, "college closure") > 0 Or _
                        InStr(normalizedAlternate, "no meeting") > 0 Or _
                        InStr(normalizedAlternate, "no-meeting") > 0)
End Function

Function GetNoMeetingReasonText(alternateText)
    If IsCollegeClosureEvent(alternateText) Then
        GetNoMeetingReasonText = "because of college closure"
    Else
        GetNoMeetingReasonText = ""
    End If
End Function

Function IsSameDayBeforeMeeting(meetingDate)
    IsSameDayBeforeMeeting = (DateValue(Date) = DateValue(meetingDate) And TimeValue(Now) < TimeSerial(11, 0, 0))
End Function

Function FormatDayMonth(targetDate)
    FormatDayMonth = Right("0" & Day(targetDate), 2) & "/" & Right("0" & Month(targetDate), 2)
End Function

Function FormatShortDate(targetDate)
    FormatShortDate = Right("0" & Day(targetDate), 2) & "/" & _
                      Right("0" & Month(targetDate), 2) & "/" & _
                      Right("0" & (Year(targetDate) Mod 100), 2)
End Function

Function GetLabManagerWeekStart(meetingDate)
    GetLabManagerWeekStart = DateAdd("d", -4, meetingDate)
End Function

Function GetLabManagerWeekEnd(meetingDate)
    GetLabManagerWeekEnd = meetingDate
End Function

Function CellText(ByRef objSheet, row, col)
    Dim value
    CellText = ""
    If CLng(col) <= 0 Then Exit Function

    On Error Resume Next
    value = objSheet.Cells(row, col).Value
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If IsNull(value) Or IsEmpty(value) Then Exit Function
    CellText = Trim(CStr(value))
End Function

Function TryGetWorksheet(ByRef objWorkbook, sheetName)
    On Error Resume Next
    Set TryGetWorksheet = objWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        Err.Clear
        Set TryGetWorksheet = Nothing
    End If
    On Error GoTo 0
End Function

Function GetInventoryWorksheet(ByRef objWorkbook)
    Dim objSheet
    Set objSheet = TryGetWorksheet(objWorkbook, SHEET_CHEMICAL_INVENTORY)
    If objSheet Is Nothing Then Set objSheet = TryGetWorksheet(objWorkbook, "Chemical Inventory")
    If objSheet Is Nothing Then Set objSheet = TryGetWorksheet(objWorkbook, "Inventory")
    Set GetInventoryWorksheet = objSheet
End Function

Function FindHeaderColumnByAliases(ByRef objSheet, aliases, defaultColumn)
    Dim lastCol, col, headerText, aliasText
    FindHeaderColumnByAliases = CLng(defaultColumn)

    On Error Resume Next
    lastCol = objSheet.Cells(1, objSheet.Columns.Count).End(-4159).Column
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    For col = 1 To lastCol
        headerText = LCase(CellText(objSheet, 1, col))
        For Each aliasText In aliases
            If headerText = LCase(Trim(CStr(aliasText))) Then
                FindHeaderColumnByAliases = col
                Exit Function
            End If
        Next
    Next
End Function

Function FindInventoryReviewDateColumn(ByRef objSheet)
    FindInventoryReviewDateColumn = FindHeaderColumnByAliases(objSheet, _
        Array(INVENTORY_REVIEW_DATE_HEADER, "Reviewed date", "Inventory review date", "Inventory reviewed date"), 0)
End Function

Function EnsureInventoryReviewDateColumn(ByRef objSheet)
    Dim reviewCol, lastCol
    reviewCol = FindInventoryReviewDateColumn(objSheet)
    If CLng(reviewCol) > 0 Then
        EnsureInventoryReviewDateColumn = reviewCol
        Exit Function
    End If

    lastCol = objSheet.Cells(1, objSheet.Columns.Count).End(-4159).Column
    reviewCol = CLng(lastCol) + 1
    objSheet.Cells(1, reviewCol).Value = INVENTORY_REVIEW_DATE_HEADER
    objSheet.Cells(1, reviewCol).Font.Bold = True
    EnsureInventoryReviewDateColumn = reviewCol
    MarkMeetingWorkbookForUpload "created inventory review-date column"
End Function

Function IsBlankCellValue(value)
    Dim textValue
    On Error Resume Next
    IsBlankCellValue = False
    If IsNull(value) Or IsEmpty(value) Then
        IsBlankCellValue = True
    Else
        textValue = CStr(value)
        If Err.Number <> 0 Then
            Err.Clear
            IsBlankCellValue = False
        Else
            IsBlankCellValue = (Trim(textValue) = "")
        End If
    End If
    If Err.Number <> 0 Then
        Err.Clear
        IsBlankCellValue = False
    End If
    On Error GoTo 0
End Function

Function IsSameReviewDate(value, meetingDate)
    Dim reviewDate
    IsSameReviewDate = False
    If TryGetReviewDateValue(value, reviewDate) Then
        On Error Resume Next
        IsSameReviewDate = (DateValue(reviewDate) = DateValue(CDate(meetingDate)))
        If Err.Number <> 0 Then
            Err.Clear
            IsSameReviewDate = False
        End If
        On Error GoTo 0
    End If
End Function

Function IsInventoryReviewDue(value, meetingDate)
    Dim reviewDate
    IsInventoryReviewDue = IsBlankCellValue(value)
    If IsInventoryReviewDue Then Exit Function
    If TryGetReviewDateValue(value, reviewDate) Then
        IsInventoryReviewDue = (reviewDate < DateAdd("yyyy", -1, DateValue(CDate(meetingDate))))
    End If
End Function

Function IsMalformedReviewDateValue(value)
    Dim reviewDate
    IsMalformedReviewDateValue = False
    If IsBlankCellValue(value) Then Exit Function
    If TryGetReviewDateValue(value, reviewDate) Then Exit Function
    IsMalformedReviewDateValue = True
End Function

Function TryGetReviewDateValue(value, ByRef reviewDate)
    Dim numericValue, candidateDate
    TryGetReviewDateValue = False
    If IsBlankCellValue(value) Then Exit Function

    On Error Resume Next
    If IsDate(value) Then
        candidateDate = DateValue(CDate(value))
        If Err.Number = 0 Then
            If IsPlausibleReviewDate(candidateDate) Then
                reviewDate = candidateDate
                TryGetReviewDateValue = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If IsNumeric(value) Then
        numericValue = CDbl(value)
        candidateDate = DateValue(CDate(numericValue))
        If Err.Number = 0 Then
            If IsPlausibleReviewDate(candidateDate) Then
                reviewDate = candidateDate
                TryGetReviewDateValue = True
            End If
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function IsPlausibleReviewDate(value)
    On Error Resume Next
    IsPlausibleReviewDate = (DateValue(CDate(value)) >= DateSerial(2000, 1, 1) And _
                             DateValue(CDate(value)) <= DateSerial(2100, 12, 31))
    If Err.Number <> 0 Then
        Err.Clear
        IsPlausibleReviewDate = False
    End If
    On Error GoTo 0
End Function

Function LastInventoryDataRow(ByRef objSheet, barcodeCol, productCol)
    Dim lastRow, candidateRow
    lastRow = 1

    If CLng(barcodeCol) > 0 Then
        candidateRow = objSheet.Cells(objSheet.Rows.Count, CLng(barcodeCol)).End(-4162).Row
        If CLng(candidateRow) > CLng(lastRow) Then lastRow = CLng(candidateRow)
    End If

    If CLng(productCol) > 0 Then
        candidateRow = objSheet.Cells(objSheet.Rows.Count, CLng(productCol)).End(-4162).Row
        If CLng(candidateRow) > CLng(lastRow) Then lastRow = CLng(candidateRow)
    End If

    LastInventoryDataRow = lastRow
End Function

Function InventoryRowHasData(ByRef objSheet, row, ByRef cols)
    InventoryRowHasData = (CellText(objSheet, row, cols("Barcode")) <> "" Or _
                           CellText(objSheet, row, cols("Product")) <> "")
End Function

Function FirstNonBlankText(values)
    Dim value
    FirstNonBlankText = ""
    For Each value In values
        If Trim(CStr(value)) <> "" Then
            FirstNonBlankText = Trim(CStr(value))
            Exit Function
        End If
    Next
End Function

Function BuildInventoryQuantityText(ByRef objSheet, row, ByRef cols)
    Dim currentQty, normalizedCurrent, normalizedQuantity, normalizedUnit, originalQty
    currentQty = CellText(objSheet, row, cols("CurrentQuantity"))
    If currentQty <> "" Then
        BuildInventoryQuantityText = currentQty
        Exit Function
    End If

    normalizedCurrent = CellText(objSheet, row, cols("NormalizedCurrentQuantity"))
    normalizedQuantity = CellText(objSheet, row, cols("NormalizedQuantity"))
    normalizedUnit = CellText(objSheet, row, cols("NormalizedUnit"))
    If normalizedCurrent <> "" Then
        If normalizedUnit <> "" Then
            BuildInventoryQuantityText = normalizedCurrent & " " & normalizedUnit
        Else
            BuildInventoryQuantityText = normalizedCurrent
        End If
        Exit Function
    End If

    If normalizedQuantity <> "" Then
        If normalizedUnit <> "" Then
            BuildInventoryQuantityText = normalizedQuantity & " " & normalizedUnit
        Else
            BuildInventoryQuantityText = normalizedQuantity
        End If
        Exit Function
    End If

    originalQty = CellText(objSheet, row, cols("Quantity"))
    BuildInventoryQuantityText = originalQty
End Function

Function BuildInventoryReviewItem(ByRef objSheet, row, ByRef cols, needsReviewDateStamp)
    Dim item
    Set item = CreateObject("Scripting.Dictionary")
    item.Add "Row", CLng(row)
    item.Add "NeedsReviewDateStamp", CBool(needsReviewDateStamp)
    item.Add "Barcode", CellText(objSheet, row, cols("Barcode"))
    item.Add "Product", CellText(objSheet, row, cols("Product"))
    item.Add "CAS", CellText(objSheet, row, cols("CAS"))
    item.Add "Brand", CellText(objSheet, row, cols("Brand"))
    item.Add "ProductNumber", CellText(objSheet, row, cols("ProductNumber"))
    item.Add "Location", FirstNonBlankText(Array( _
        CellText(objSheet, row, cols("LocationPath")), _
        CellText(objSheet, row, cols("Location")), _
        CellText(objSheet, row, cols("Sublocation")), _
        CellText(objSheet, row, cols("Laboratory")), _
        CellText(objSheet, row, cols("Building")), _
        CellText(objSheet, row, cols("Site"))))
    item.Add "Quantity", BuildInventoryQuantityText(objSheet, row, cols)
    item.Add "Hazards", CellText(objSheet, row, cols("Hazards"))
    item.Add "StorageTemperature", CellText(objSheet, row, cols("StorageTemperature"))
    Set BuildInventoryReviewItem = item
End Function

Sub AddInventoryItemIfEligible(ByRef items, ByRef objSheet, row, ByRef cols, needsReviewDateStamp)
    Dim item
    If items.Count >= INVENTORY_REVIEW_BATCH_SIZE Then Exit Sub
    If Not InventoryRowHasData(objSheet, row, cols) Then Exit Sub
    Set item = BuildInventoryReviewItem(objSheet, row, cols, needsReviewDateStamp)
    items.Add CStr(items.Count + 1), item
End Sub

Function SelectInventoryReviewItems(ByRef objWorkbook, meetingDate)
    Dim items, objSheet, cols, reviewCol, lastRow, row, reviewValue
    Set items = CreateObject("Scripting.Dictionary")
    Set objSheet = GetInventoryWorksheet(objWorkbook)

    If objSheet Is Nothing Then
        LogEvent "Inventory review skipped: no '" & SHEET_CHEMICAL_INVENTORY & "', 'Chemical Inventory', or 'Inventory' sheet found in meeting rota workbook."
        Set SelectInventoryReviewItems = items
        Exit Function
    End If

    Set cols = CreateObject("Scripting.Dictionary")
    cols.Add "Barcode", FindHeaderColumnByAliases(objSheet, Array("Barcode"), 1)
    cols.Add "CAS", FindHeaderColumnByAliases(objSheet, Array("CAS", "CAS number", "CAS No."), 2)
    cols.Add "Product", FindHeaderColumnByAliases(objSheet, Array("Product name", "Product", "Chemical name", "Name"), 3)
    cols.Add "Hazards", FindHeaderColumnByAliases(objSheet, Array("Hazard Label(s)", "Hazard labels", "Hazards"), 4)
    cols.Add "StorageTemperature", FindHeaderColumnByAliases(objSheet, Array("Storage temperature", "Storage temp"), 5)
    cols.Add "Site", FindHeaderColumnByAliases(objSheet, Array("Site", "Campus"), 6)
    cols.Add "Building", FindHeaderColumnByAliases(objSheet, Array("Building"), 7)
    cols.Add "Laboratory", FindHeaderColumnByAliases(objSheet, Array("Laboratory", "Lab"), 8)
    cols.Add "Location", FindHeaderColumnByAliases(objSheet, Array("Location"), 9)
    cols.Add "LocationPath", FindHeaderColumnByAliases(objSheet, Array("Location path", "Location Path"), 10)
    cols.Add "Sublocation", FindHeaderColumnByAliases(objSheet, Array("Sublocation", "Sub-location", "Sub location"), 11)
    cols.Add "Brand", FindHeaderColumnByAliases(objSheet, Array("Brand", "Manufacturer"), 12)
    cols.Add "ProductNumber", FindHeaderColumnByAliases(objSheet, Array("Product number", "Product no.", "Catalogue number", "Catalog number"), 13)
    cols.Add "Quantity", FindHeaderColumnByAliases(objSheet, Array("Quantity"), 14)
    cols.Add "CurrentQuantity", FindHeaderColumnByAliases(objSheet, Array("Current quantity (Net)", "Current quantity", "Current Quantity"), 17)
    cols.Add "NormalizedQuantity", FindHeaderColumnByAliases(objSheet, Array("Normalized quantity", "Normalised quantity"), 18)
    cols.Add "NormalizedCurrentQuantity", FindHeaderColumnByAliases(objSheet, Array("Normalized current quantity", "Normalised current quantity"), 19)
    cols.Add "NormalizedUnit", FindHeaderColumnByAliases(objSheet, Array("Normalized unit", "Normalised unit"), 20)

    reviewCol = FindInventoryReviewDateColumn(objSheet)
    lastRow = LastInventoryDataRow(objSheet, cols("Barcode"), cols("Product"))

    If CLng(reviewCol) > 0 Then
        For row = 2 To lastRow
            reviewValue = objSheet.Cells(row, reviewCol).Value
            If IsSameReviewDate(reviewValue, meetingDate) Then
                AddInventoryItemIfEligible items, objSheet, row, cols, False
            End If
            If items.Count >= INVENTORY_REVIEW_BATCH_SIZE Then Exit For
        Next
    End If

    If items.Count < INVENTORY_REVIEW_BATCH_SIZE Then
        For row = 2 To lastRow
            If CLng(reviewCol) = 0 Then
                reviewValue = ""
            Else
                reviewValue = objSheet.Cells(row, reviewCol).Value
                If IsMalformedReviewDateValue(reviewValue) Then
                    LogEvent "WARNING: Inventory row " & row & " has a non-empty malformed review date value '" & _
                             SafeTextValue(reviewValue, "[unreadable]") & "'. Row will not be selected for review unless the value is cleared or corrected."
                End If
            End If

            If IsInventoryReviewDue(reviewValue, meetingDate) Then
                AddInventoryItemIfEligible items, objSheet, row, cols, True
            End If
            If items.Count >= INVENTORY_REVIEW_BATCH_SIZE Then Exit For
        Next
    End If

    LogEvent "Selected " & items.Count & " inventory item(s) for review on " & FormatDateTime(meetingDate, 2) & "."
    Set SelectInventoryReviewItems = items
End Function

Sub StampInventoryReviewItems(ByRef objWorkbook, ByRef inventoryItems, meetingDate)
    Dim objSheet, reviewCol, item, itemKey, row, currentValue, changedCount
    If inventoryItems Is Nothing Then Exit Sub
    If inventoryItems.Count = 0 Then Exit Sub

    Set objSheet = GetInventoryWorksheet(objWorkbook)
    If objSheet Is Nothing Then
        LogEvent "WARNING: Cannot stamp inventory review dates because inventory sheet is unavailable."
        Exit Sub
    End If

    On Error Resume Next
    reviewCol = EnsureInventoryReviewDateColumn(objSheet)
    If Err.Number <> 0 Then
        LogEvent "ERROR: Failed to prepare inventory review-date column. Keeping local workbook for recovery. Err: " & Err.Number & " - " & Err.Description
        keepSyncedWorkbookForRecovery = True
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    changedCount = 0

    For Each itemKey In inventoryItems.Keys
        Set item = inventoryItems.Item(itemKey)
        row = CLng(item("Row"))
        currentValue = objSheet.Cells(row, reviewCol).Value
        If Not IsSameReviewDate(currentValue, meetingDate) Then
            objSheet.Cells(row, reviewCol).Value = DateValue(CDate(meetingDate))
            objSheet.Cells(row, reviewCol).NumberFormat = "dd/mm/yyyy"
            If Err.Number <> 0 Then
                LogEvent "ERROR: Failed to stamp inventory review date on row " & row & ". Keeping local workbook for recovery. Err: " & Err.Number & " - " & Err.Description
                keepSyncedWorkbookForRecovery = True
                Err.Clear
            Else
                changedCount = changedCount + 1
            End If
        End If
    Next
    On Error GoTo 0

    If changedCount > 0 Then
        MarkMeetingWorkbookForUpload "stamped " & changedCount & " inventory review date(s)"
    Else
        LogEvent "Inventory review dates already stamped for " & FormatDateTime(meetingDate, 2) & "."
    End If
End Sub

Function InventoryTableCell(value)
    Dim textValue
    textValue = Trim(SafeTextValue(value, ""))
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, vbTab, " ")
    If textValue = "" Then textValue = "-"
    InventoryTableCell = textValue
End Function

Function HexByte(value)
    HexByte = Right("0" & Hex(CLng(value) And &HFF), 2)
End Function

Function PercentEncodeByte(value)
    PercentEncodeByte = "%" & HexByte(value)
End Function

Function MailtoUrlEncode(value)
    Dim textValue, output, i, ch, code
    textValue = CStr(value)
    output = ""

    For i = 1 To Len(textValue)
        ch = Mid(textValue, i, 1)
        code = AscW(ch)
        If CLng(code) < 0 Then code = CLng(code) + 65536

        If (CLng(code) >= 48 And CLng(code) <= 57) Or _
           (CLng(code) >= 65 And CLng(code) <= 90) Or _
           (CLng(code) >= 97 And CLng(code) <= 122) Or _
           CLng(code) = 45 Or CLng(code) = 46 Or CLng(code) = 95 Or CLng(code) = 126 Then
            output = output & ch
        ElseIf CLng(code) = 13 Then
            If i < Len(textValue) Then
                If AscW(Mid(textValue, i + 1, 1)) = 10 Then i = i + 1
            End If
            output = output & "%0D%0A"
        ElseIf CLng(code) = 10 Then
            output = output & "%0D%0A"
        Else
            If CLng(code) < 128 Then
                output = output & PercentEncodeByte(code)
            ElseIf CLng(code) < 2048 Then
                output = output & PercentEncodeByte(&HC0 Or (CLng(code) \ 64)) & _
                                  PercentEncodeByte(&H80 Or (CLng(code) And &H3F))
            Else
                output = output & PercentEncodeByte(&HE0 Or (CLng(code) \ 4096)) & _
                                  PercentEncodeByte(&H80 Or ((CLng(code) \ 64) And &H3F)) & _
                                  PercentEncodeByte(&H80 Or (CLng(code) And &H3F))
            End If
        End If
    Next

    MailtoUrlEncode = output
End Function

Function BuildInventoryFeedbackMailto(ByRef inventoryItems, meetingDate)
    Dim subjectText, bodyText, recipientName

    If Trim(CStr(INVENTORY_MANAGER_EMAIL)) = "" Then
        BuildInventoryFeedbackMailto = ""
        Exit Function
    End If

    If inventoryItems Is Nothing Then
        BuildInventoryFeedbackMailto = ""
        Exit Function
    End If
    If inventoryItems.Count = 0 Then
        BuildInventoryFeedbackMailto = ""
        Exit Function
    End If

    recipientName = Trim(CStr(INVENTORY_MANAGER_NAME))
    If recipientName = "" Then recipientName = "Inventory team"

    subjectText = "Labcup inventory feedback: week ending " & FormatDateTime(meetingDate, 2)

    bodyText = "Hi " & recipientName & "," & vbCrLf & vbCrLf & _
               "I reviewed the following Labcup inventory records for the week ending " & _
               FormatDateTime(meetingDate, 2) & ". The batch details are in the weekly responsibility email." & _
               vbCrLf & vbCrLf & _
               "Feedback / corrections:" & vbCrLf & _
               "- Items confirmed with no changes:" & vbCrLf & _
               "- Chemical identity/name corrections:" & vbCrLf & _
               "- Location corrections:" & vbCrLf & _
               "- Quantity corrections:" & vbCrLf & _
               "- Hazards/storage corrections:" & vbCrLf & _
               "- Other notes:" & vbCrLf & vbCrLf & _
               "Please update Labcup with any corrections and update the reviewed items' review dates in the Chemicals sheet of meeting_rota.xlsx, which the script uses to schedule future reviews." & vbCrLf & _
               "Meeting rota - Chemicals: " & INVENTORY_WORKBOOK_URL & vbCrLf & vbCrLf & _
               "Best,"

    BuildInventoryFeedbackMailto = "mailto:" & Trim(CStr(INVENTORY_MANAGER_EMAIL)) & _
                                   "?cc=" & MailtoUrlEncode(MANAGER) & _
                                   "&subject=" & MailtoUrlEncode(subjectText) & _
                                   "&body=" & MailtoUrlEncode(bodyText)
End Function

Function BuildInventoryReviewTable(ByRef inventoryItems, meetingDate)
    Dim tableText, item, itemKey, index, feedbackLink
    BuildInventoryReviewTable = ""
    If inventoryItems Is Nothing Then Exit Function
    If inventoryItems.Count = 0 Then Exit Function

    tableText = "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                "<tr>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">No.</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">Barcode</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">Product</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">CAS</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">Location</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">Quantity</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">Hazards</td>" & _
                "<td style=""padding:5px 8px; background-color:#f5f5f5; color:#708090; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; font-weight:600; mso-line-height-rule:exactly;"">Storage</td>" & _
                "</tr>"
    index = 1
    For Each itemKey In inventoryItems.Keys
        Set item = inventoryItems.Item(itemKey)
        tableText = tableText & "<tr>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(index) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("Barcode"))) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("Product"))) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("CAS"))) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("Location"))) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("Quantity"))) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("Hazards"))) & "</td>" & _
                    "<td style=""padding:5px 8px; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; font-size:13px; line-height:18px; mso-line-height-rule:exactly;"">" & HtmlEscape(InventoryTableCell(item("StorageTemperature"))) & "</td>" & _
                    "</tr>"

        index = index + 1
    Next
    tableText = tableText & "</table>"

    feedbackLink = BuildInventoryFeedbackMailto(inventoryItems, meetingDate)
    If Trim(feedbackLink) <> "" Then
        tableText = tableText & "<p style=""margin:12px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;""><a href=""" & HtmlEscape(feedbackLink) & """ style=""color:#000000; text-decoration:underline; font-weight:600;"">Open feedback email for this batch</a> <span style=""color:#708090;"">to " & HtmlEscape(INVENTORY_MANAGER_NAME) & ", CC Giacomo. Use the table above as the batch reference.</span></p>"
        tableText = tableText & "<p style=""margin:8px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;""><a href=""" & HtmlEscape(INVENTORY_WORKBOOK_URL) & """ style=""color:#000000; text-decoration:underline;"">Meeting rota - Chemicals</a></p>"
    End If

    BuildInventoryReviewTable = tableText
End Function

Function InventoryManagerCcFor(toAddress)
    Dim target, inventoryManager
    target = LCase(Trim(CStr(toAddress)))
    inventoryManager = LCase(Trim(CStr(INVENTORY_MANAGER_EMAIL)))

    If inventoryManager = "" Or target = inventoryManager Then
        InventoryManagerCcFor = ""
    Else
        InventoryManagerCcFor = INVENTORY_MANAGER_EMAIL
    End If
End Function

Function DispatchStatusSubmitted(statusText)
    Select Case UCase(Trim(CStr(statusText)))
        Case "SUBMITTED", "REPLACED"
            DispatchStatusSubmitted = True
        Case Else
            DispatchStatusSubmitted = False
    End Select
End Function

Function GenerateEmailContent(meetingDate, row, objSheetRota, nextFridayRow)
    Call LogVerbose("GenerateEmailContent invoked for row " & row)
    Dim literature, results, alternate, room, labManager, next_lit, next_res, next_alt, content, nameParts
    Dim meetingLeadName, nextMeetingLeadName
    Dim meetingDateText, meetingShortDate, nextFridayDate, nextFridayText, nextFridayShortDate
    Dim labManagerRangeStart, labManagerRangeEnd, nextLabManagerRangeStart, nextLabManagerRangeEnd
    Dim nextLabManager
    Dim currentResults, currentLiterature, currentLabManager, currentChair
    Dim nextResults, nextLiterature, nextChair
    Dim currentAlternate, nextAlternate
    Dim rotaUrl, meetingModeText, meetingLocationText, detailsHtml, thisWeekHtml, nextWeekHtml, supportHtml
    Dim setupNoteText, occupiedLocationText, noMeetingMessage, nextMeetingNotice
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
        nextLabManager = Trim(objSheetRota.Cells(nextFridayRow, 6).Value)
        nextFridayDate = objSheetRota.Cells(nextFridayRow, 1).Value
        LogEvent "Next Friday presenters found | Row: " & nextFridayRow & _
                  " | Lit: " & next_lit & " | Res: " & next_res & _
                  " | Alt: " & next_alt
    Else
        next_lit = ""
        next_res = ""
        next_alt = ""
        nextLabManager = ""
        nextFridayDate = DateAdd("d", 7, meetingDate)
        LogEvent "No next Friday data found. Using empty placeholders."
    End If

    Dim senderFirstName
    senderFirstName = GetSenderFirstName()

    meetingLeadName = DeriveMeetingLeadName(objSheetRota, meetingDate)
    nextMeetingLeadName = DeriveMeetingLeadName(objSheetRota, nextFridayDate)

    meetingDateText = WeekdayName(Weekday(meetingDate), False) & " " & FormatDateTime(meetingDate, 2)
    meetingShortDate = FormatDayMonth(meetingDate)
    nextFridayText = WeekdayName(Weekday(nextFridayDate), False) & " " & FormatDateTime(nextFridayDate, 2)
    nextFridayShortDate = FormatDayMonth(nextFridayDate)

    labManagerRangeStart = FormatDayMonth(DateAdd("d", -4, meetingDate))
    labManagerRangeEnd = FormatDayMonth(meetingDate)
    nextLabManagerRangeStart = FormatDayMonth(DateAdd("d", -4, nextFridayDate))
    nextLabManagerRangeEnd = FormatDayMonth(nextFridayDate)

    currentResults = Trim(results)
    If currentResults = "" Then
        currentResults = "TBC (X)"
    End If

    currentLiterature = Trim(literature)
    If currentLiterature = "" Then
        currentLiterature = "TBC (X)"
    End If

    currentLabManager = Trim(labManager)
    If currentLabManager = "" Then
        currentLabManager = "TBC (X)"
    End If

    currentChair = Trim(meetingLeadName)
    If currentChair = "" Then
        currentChair = "TBC (X)"
    End If

    currentAlternate = Trim(alternate)

    nextResults = Trim(next_res)
    If nextResults = "" Then
        nextResults = "TBC (X)"
    End If

    nextLiterature = Trim(next_lit)
    If nextLiterature = "" Then
        nextLiterature = "TBC (X)"
    End If

    If Trim(nextLabManager) = "" Then
        nextLabManager = "TBC (X)"
    End If

    nextAlternate = Trim(next_alt)

    nextChair = Trim(nextMeetingLeadName)
    If nextChair = "" Then
        nextChair = "TBC (X)"
    End If

    rotaUrl = "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx?web=1"

    If Trim(room) <> "" Then
        meetingLocationText = room
    Else
        meetingLocationText = "TBC"
    End If
    meetingModeText = "In person or via MS Teams (see calendar invite for the link)."

    detailsHtml = "<table role=""presentation"" data-od-id=""action-header"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                  "<tr><td data-od-id=""email-kicker"" style=""padding:0 0 3px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:13px; font-weight:600; line-height:18px; letter-spacing:0; mso-line-height-rule:exactly;"">GROUP MEETING &middot; THIS WEEK</td></tr>" & _
                  "<tr><td data-od-id=""meeting-summary"" style=""padding:0 0 3px 0; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; mso-line-height-rule:exactly;"">" & HtmlEscape(meetingDateText) & " &middot; 11:00 &middot; " & HtmlEscape(meetingLocationText) & "</td></tr>" & _
                  "<tr><td data-od-id=""meeting-location-format"" style=""padding:0 0 12px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(meetingModeText) & "</td></tr>" & _
                  "</table>"

    If IsNoMeetingEvent(currentAlternate) Then
        detailsHtml = "<table role=""presentation"" data-od-id=""action-header"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                      "<tr><td data-od-id=""email-kicker"" style=""padding:0 0 3px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:13px; font-weight:600; line-height:18px; letter-spacing:0; mso-line-height-rule:exactly;"">GROUP MEETING &middot; THIS WEEK</td></tr>" & _
                      "<tr><td data-od-id=""meeting-summary"" style=""padding:0 0 3px 0; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; mso-line-height-rule:exactly;"">" & HtmlEscape(meetingDateText) & "</td></tr>" & _
                      "<tr><td data-od-id=""meeting-location-format"" style=""padding:0 0 12px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">No regular meeting is scheduled for this date.</td></tr>" & _
                      "</table>"

        noMeetingMessage = "Please note that no group meeting will take place on " & meetingDateText
        If Trim(GetNoMeetingReasonText(currentAlternate)) <> "" Then
            noMeetingMessage = noMeetingMessage & " " & GetNoMeetingReasonText(currentAlternate)
        End If
        noMeetingMessage = noMeetingMessage & "."
        nextMeetingNotice = ""
        If nextFridayRow <> -1 Then
            nextMeetingNotice = "<p style=""margin:8px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">The next scheduled group meeting is on " & HtmlEscape(nextFridayText) & " at 11:00.</p>"
        End If

        content = "<!doctype html><html><body style=""margin:0; padding:0; background-color:#ffffff;"">" & _
                  "<table role=""presentation"" data-od-id=""email-shell"" width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:100%; border-collapse:collapse; background-color:#ffffff;"">" & _
                  "<tr><td align=""center"" style=""padding:14px 10px 18px 10px;"">" & _
                  "<table role=""presentation"" data-od-id=""email-content"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                  "<tr><td style=""padding:0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & _
                  "<p style=""margin:0 0 9px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">Hi all,</p>" & _
                  detailsHtml & _
                  "<table role=""presentation"" data-od-id=""no-meeting-section"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; border:2px solid #0000cd; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                  "<tr><td style=""padding:9px 14px 8px 14px; border-bottom:1px solid #e1e1e6; background-color:#ffffff; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; letter-spacing:0; mso-line-height-rule:exactly;"">NO REGULAR MEETING</td></tr>" & _
                  "<tr><td style=""padding:8px 14px 9px 14px; background-color:#ffffff; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(noMeetingMessage) & nextMeetingNotice & "</td></tr>" & _
                  "</table>" & _
                  "<table role=""presentation"" data-od-id=""schedule-link-section"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                  "<tr><td height=""18"" style=""height:18px; font-size:1px; line-height:1px; mso-line-height-rule:exactly;"">&nbsp;</td></tr>" & _
                  "<tr><td style=""padding:8px 0 0 0; border-top:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:15px; line-height:22px; mso-line-height-rule:exactly;""><span style=""font-weight:600;"">Later weeks</span> &rarr; <a href=""" & HtmlEscape(rotaUrl) & """ style=""color:#000000; text-decoration:underline;"">Open the meeting rota</a></td></tr>" & _
                  "</table>" & _
                  "<p data-od-id=""signoff"" style=""margin:13px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">Best,<br>" & HtmlEscape(Trim(senderFirstName)) & "</p>" & _
                  "</td></tr></table></td></tr></table></body></html>"

        GenerateEmailContent = content
        LogEvent "Generated no-meeting email body length: " & Len(content)
        Exit Function
    End If

    If meetingLocationText <> "TBC" Then
        occupiedLocationText = meetingLocationText
    Else
        occupiedLocationText = "the room"
    End If
    setupNoteText = "Please arrive/join 5-10 minutes early to set up. We start at 11:00, or 11:05 if " & occupiedLocationText & " is occupied until 11:00."

    thisWeekHtml = "<table role=""presentation"" data-od-id=""this-week-section"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; border:2px solid #0000cd; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                   "<tr><td colspan=""2"" data-od-id=""this-week-heading"" style=""padding:9px 14px 8px 14px; border-bottom:1px solid #e1e1e6; background-color:#ffffff; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; letter-spacing:0; mso-line-height-rule:exactly;"">THIS WEEK <span style=""font-weight:400; letter-spacing:0;"">(" & HtmlEscape(labManagerRangeStart) & " - " & HtmlEscape(labManagerRangeEnd) & ")</span></td></tr>"
    If currentAlternate <> "" Then
        thisWeekHtml = thisWeekHtml & HtmlResponsibilityRow("Programme", currentAlternate, "")
    Else
        thisWeekHtml = thisWeekHtml & HtmlResponsibilityRow("Results", currentResults, "")
        thisWeekHtml = thisWeekHtml & HtmlResponsibilityRow("Literature", currentLiterature, "")
    End If
    thisWeekHtml = thisWeekHtml & HtmlResponsibilityRow("Lab manager + chair/timekeeper", currentLabManager, "")
    thisWeekHtml = thisWeekHtml & "<tr><td colspan=""2"" data-od-id=""setup-note"" style=""padding:8px 14px 9px 14px; background-color:#ffffff; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & HtmlEscape(setupNoteText) & "</td></tr></table>"

    supportHtml = "<table role=""presentation"" data-od-id=""lab-practice-support"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                  "<tr><td height=""12"" style=""height:12px; font-size:1px; line-height:1px; mso-line-height-rule:exactly;"">&nbsp;</td></tr>" & _
                  "<tr><td style=""padding:0 14px 0 14px; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & _
                  "<span style=""font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; font-weight:600; line-height:22px; mso-line-height-rule:exactly;"">Lab-practice support this week.</span> This week's lab manager is the first point of contact for minor lab-practice questions and reminders. Persistent issues can be raised constructively in Friday's lab-status update, with the focus on behaviours and improvements rather than blame." & _
                  "</td></tr></table>"

    nextWeekHtml = "<table role=""presentation"" data-od-id=""next-week-section"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                   "<tr><td height=""18"" style=""height:18px; font-size:1px; line-height:1px; mso-line-height-rule:exactly;"">&nbsp;</td></tr>" & _
                   "<tr><td colspan=""2"" data-od-id=""next-week-heading"" style=""padding:0 0 6px 0; border-bottom:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:14px; font-weight:600; line-height:20px; letter-spacing:0; mso-line-height-rule:exactly;"">NEXT WEEK <span style=""font-weight:400; letter-spacing:0;"">(" & HtmlEscape(nextLabManagerRangeStart) & " - " & HtmlEscape(nextLabManagerRangeEnd) & ")</span></td></tr>"
    If nextAlternate <> "" Then
        nextWeekHtml = nextWeekHtml & HtmlNextResponsibilityRow("Programme", nextAlternate)
    Else
        nextWeekHtml = nextWeekHtml & HtmlNextResponsibilityRow("Results", nextResults)
        nextWeekHtml = nextWeekHtml & HtmlNextResponsibilityRow("Literature", nextLiterature)
    End If
    If Not IsNoMeetingEvent(nextAlternate) Then
        nextWeekHtml = nextWeekHtml & HtmlNextResponsibilityRow("Lab manager + chair/timekeeper", nextLabManager)
    End If
    nextWeekHtml = nextWeekHtml & "</table>"

    content = "<!doctype html><html><body style=""margin:0; padding:0; background-color:#ffffff;"">" & _
              "<table role=""presentation"" data-od-id=""email-shell"" width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:100%; border-collapse:collapse; background-color:#ffffff;"">" & _
              "<tr><td align=""center"" style=""padding:14px 10px 18px 10px;"">" & _
              "<table role=""presentation"" data-od-id=""email-content"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
              "<tr><td style=""padding:0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & _
              "<p style=""margin:0 0 9px 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">Hi all,</p>" & _
              detailsHtml & _
              thisWeekHtml & _
              supportHtml & _
              nextWeekHtml & _
              "<table role=""presentation"" data-od-id=""schedule-link-section"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
              "<tr><td height=""18"" style=""height:18px; font-size:1px; line-height:1px; mso-line-height-rule:exactly;"">&nbsp;</td></tr>" & _
              "<tr><td style=""padding:8px 0 0 0; border-top:1px solid #e1e1e6; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & _
              "<span style=""font-weight:600;"">Later weeks</span> &rarr; " & _
              "<a href=""" & HtmlEscape(rotaUrl) & """ style=""color:#000000; text-decoration:underline;"">Open the meeting rota</a>" & _
              "</td></tr></table>" & _
              "<p data-od-id=""signoff"" style=""margin:13px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">Best,<br>" & HtmlEscape(Trim(senderFirstName)) & "</p>" & _
              "</td></tr></table></td></tr></table></body></html>"

    GenerateEmailContent = content
    LogEvent "Generated email body length: " & Len(content)
End Function

Function GeneratePresenterReminderEmail(presenterName, presentationType, meetingDate)
    Call LogVerbose("GeneratePresenterReminderEmail invoked for presenter='" & presenterName & "'")
    Dim content, trimmedType, normalizedType, description, reminderSubject, meetingDateText, typeLabel
    Dim greetingName, meetingDateCode, presenterInitials, senderFirstName
    Dim enDash, apostrophe, dotText, filePattern, fileExample, headerTitle, blockHtml
    Dim clarityText, literatureHtml, pointersHtml

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
    meetingDateCode = FormatMeetingDateForFilename(meetingDate)
    presenterInitials = ExtractNameInitials(presenterName)
    If presenterInitials = "" Then presenterInitials = "XX"
    senderFirstName = GetSenderFirstName()
    dotText = " " & ChrW(&HB7) & " "

    If presenterName = "" Then
        content = ReminderGreeting("Hi,") & _
                  ReminderHeader("PRESENTATION REMINDER" & dotText & "GROUP MEETING", "Presenter not assigned" & dotText & meetingDateText, "") & _
                  ReminderParagraph("No presenter is currently assigned for the " & reminderSubject & " scheduled on " & meetingDateText & ".") & _
                  ReminderParagraph("Please update the rota or advise the group so we can keep the agenda on track.") & _
                  ReminderSignoff(senderFirstName)
        content = ReminderEmailShell(content)
        GeneratePresenterReminderEmail = content
        LogEvent "Generated presenter reminder length: " & Len(content)
        Exit Function
    End If

    greetingName = FirstNameForGreeting(presenterName, "there")
    enDash = ChrW(&H2013)
    apostrophe = Chr(39)
    filePattern = "YYMMDD_[NameInitials]_" & normalizedType & "_presentation.pptx"
    fileExample = meetingDateCode & "_" & presenterInitials & "_" & normalizedType & "_presentation.pptx"
    headerTitle = typeLabel & dotText & meetingDateText

    blockHtml = ReminderBlockStart("FOR THE GROUP DISCUSSION") & _
                ReminderTableRow("Timing", "The slot is 15 minutes for your presentation plus 5 minutes for Q&A. As a rough guide, aim for about 1 minute per slide (so around 10" & enDash & "12 slides in total).") & _
                ReminderTableRow("Rehearsal", "The group meeting leader will keep time to ensure we don" & apostrophe & "t overrun, so please rehearse your talk beforehand to check that you are within the time limit.") & _
                ReminderTableRow("File name", fileExample) & _
                "<tr><td colspan=""2"" style=""padding:8px 14px 9px 14px; background-color:#ffffff; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;"">" & _
                "Please save your presentation to the group SharePoint folder (<a href=""https://imperiallondon.sharepoint.com/:f:/r/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/Group%20Meeting%20Presentations?d=w9f1e10972eb04b19a63ff26d79b9d6df&amp;csf=1&amp;web=1&amp;e=4hsiRl"" style=""color:#000000; text-decoration:underline;"">Group Meeting Presentations</a>) using the naming pattern " & HtmlEscape(filePattern) & ".</td></tr>" & _
                ReminderBlockEnd()

    clarityText = "For clarity: use the meeting date as YYMMDD, add your initials, and include the talk type (e.g., " & description & "), ending with _presentation.pptx (e.g., " & fileExample & ")."

    literatureHtml = ""
    If normalizedType = "literature" Then
        literatureHtml = ReminderPlainSection("LITERATURE-SPECIFIC NOTES", _
            ReminderParagraph("Share the article you" & apostrophe & "ll present on Friday by Wednesday, so anyone interested can skim it in advance and come prepared with questions" & enDash & "making the discussion more enjoyable and productive.") & _
            ReminderParagraph("For literature presentations, please choose high-quality papers that may be of general interest, and focus on the higher-level implications (e.g., key scientific principles and/or relevance to work within the group)."))
    End If

    pointersHtml = "<p style=""margin:0 0 7px 0;"">A few pointers for preparing your slides:</p>" & _
                   "<ul style=""margin:0; padding:0 0 0 18px;"">" & _
                   "<li style=""margin:0 0 7px 0;"">" & HtmlEscape("Aim for a balanced, well-planned presentation, rather than trying to include everything.") & "</li>" & _
                   "<li style=""margin:0 0 7px 0;"">" & HtmlEscape("Use no more than ~3 concise bullet points per slide to support (not duplicate) what you say, so the audience focuses on you rather than reading the screen.") & "</li>" & _
                   "<li style=""margin:0 0 7px 0;"">" & HtmlEscape("Let the slides support your performance, not be the performance" & enDash & "they shouldn" & apostrophe & "t steal your stage.") & "</li>" & _
                   "<li style=""margin:0 0 7px 0;"">" & HtmlEscape("One clear, well-chosen image is often more effective than a lot of text, especially as many scientists are visual learners.") & "</li>" & _
                   "<li style=""margin:0;"">" & HtmlEscape("Assume a left-to-right, top-to-bottom reading pattern: place the most important information at the top left and less critical details toward the bottom right.") & "</li>" & _
                   "</ul>"

    content = ReminderGreeting("Hi " & Trim(greetingName) & ",") & _
              ReminderHeader("PRESENTATION REMINDER" & dotText & "GROUP MEETING", headerTitle, "15 minutes for the presentation + 5 minutes for Q&A") & _
              ReminderParagraph("Just a reminder that you are scheduled to present " & description & " at this week's group meeting on " & meetingDateText & ".") & _
              blockHtml & _
              ReminderParagraph(clarityText) & _
              literatureHtml & _
              ReminderPlainSection("SLIDE POINTERS", pointersHtml) & _
              ReminderSubtleParagraph("Please ensure your materials are ready and let me know if you need any support.") & _
              ReminderSignoff(senderFirstName)
    content = ReminderEmailShell(content)
    GeneratePresenterReminderEmail = content
    LogEvent "Generated presenter reminder length: " & Len(content)
End Function

Function GenerateMeetingLeadReminderEmail(meetingLeadName, meetingDate)
    Call LogVerbose("GenerateMeetingLeadReminderEmail invoked for lead='" & meetingLeadName & "'")
    Dim content, greetingName, meetingDateText, senderFirstName, dotText, blockHtml

    meetingLeadName = Trim(meetingLeadName)
    meetingDateText = FormatDateTime(meetingDate, 2)
    senderFirstName = GetSenderFirstName()
    greetingName = FirstNameForGreeting(meetingLeadName, "there")
    dotText = " " & ChrW(&HB7) & " "

    blockHtml = ReminderBlockStart("GROUP MEETING RESPONSIBILITIES") & _
                ReminderTableRow("Lead", "Just a reminder that you will lead this week's group meeting on " & meetingDateText & ".") & _
                ReminderTableRow("Timing", "Each presenter has 15 minutes for their talk plus 5 minutes for Q&A.") & _
                ReminderTableRow("Setup", "Please arrive about 5 minutes early to set up the room so the computer is ready right at meeting time.") & _
                ReminderTableRow("Meeting flow", "Bring the red/yellow timing cards (5 minutes/1 minute left) to signal time to each presenter, and keep presenters on schedule while running the inclusion activity.") & _
                ReminderBlockEnd()

    content = ReminderGreeting("Hi " & greetingName & ",") & _
              ReminderHeader("MEETING LEAD REMINDER" & dotText & "GROUP MEETING", "Meeting lead" & dotText & meetingDateText, "") & _
              blockHtml & _
              ReminderSignoff(senderFirstName)
    content = ReminderEmailShell(content)

    GenerateMeetingLeadReminderEmail = content
    LogEvent "Generated meeting lead reminder length: " & Len(content)
End Function

Function GenerateLabManagerReminderEmail(labManagerName, meetingDate)
    Call LogVerbose("GenerateLabManagerReminderEmail invoked for manager='" & labManagerName & "'")
    Dim content, nameParts, manager
    Dim firstName, lastName, senderDisplayName, parcelAddressee
    Dim greetingName, blockHtml, dotText, parcelDisplay
    Dim weekStartDate, weekEndDate, startText, endText, weekRange
    manager = Trim(labManagerName)
    LogEvent "Generating lab manager reminder email | Manager: " & manager

    If manager = "" Then
        manager = "there"
    End If

    greetingName = FirstNameForGreeting(manager, "there")

    nameParts = Split(GetSenderDisplayName(), ", ")
    firstName = ""
    lastName = ""

    If UBound(nameParts) >= 1 Then
        lastName = Trim(nameParts(0))
        firstName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        firstName = Trim(nameParts(0))
    End If

    senderDisplayName = firstName
    If senderDisplayName = "" Then senderDisplayName = GetSenderFirstName()

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

    weekStartDate = GetLabManagerWeekStart(meetingDate)
    weekEndDate = GetLabManagerWeekEnd(meetingDate)
    startText = FormatDateTime(weekStartDate, 2)
    endText = FormatDateTime(weekEndDate, 2)
    weekRange = startText & " " & ChrW(&H2013) & " " & endText
    dotText = " " & ChrW(&HB7) & " "

    parcelDisplay = parcelAddressee
    If LCase(parcelAddressee) <> "me" Then
        parcelDisplay = Trim(parcelAddressee)
    End If

    blockHtml = ReminderBlockStart("LAB MANAGEMENT RESPONSIBILITIES") & _
                ReminderTableRow("Cleaning", "Lab cleaning starts on Monday and must be completed by this Friday; please aim to finish by Friday EOD so the lab is ready for Monday.") & _
                ReminderTableRow("Parcels", "On Friday, please pick up all parcels addressed to " & parcelDisplay & "; most of them contain the consumables ordered the last week that need to be re-stocked in the lab.") & _
                ReminderTableRow("Stores", "Also, please procure the list of missing items on the whiteboard from the stores so that we can maintain supplies at all times.") & _
                ReminderBlockEnd()

    content = ReminderGreeting("Hi " & greetingName & ",") & _
              ReminderHeader("LAB MANAGEMENT REMINDER", "Lab manager" & dotText & weekRange, "") & _
              ReminderParagraph("This is a reminder that you are the lab manager for the week of " & weekRange & ".") & _
              blockHtml & _
              ReminderSubtleParagraph("Keeping the lab organized is part of the weekly group rhythm.") & _
              ReminderSignoff(senderDisplayName)
    content = ReminderEmailShell(content)
    GenerateLabManagerReminderEmail = content
    LogEvent "Generated lab manager reminder length: " & Len(content)
End Function

Function GenerateWeeklyResponsibilityReminderEmail(responsibleName, meetingDate, ByRef inventoryReviewItems)
    Call LogVerbose("GenerateWeeklyResponsibilityReminderEmail invoked for responsible='" & responsibleName & "'")
    Dim content, nameParts, manager
    Dim firstName, lastName, senderDisplayName, parcelAddressee
    Dim greetingName, greetingDisplay
    Dim weekStartDate, weekEndDate, startText, endText, weekRange
    Dim meetingDateText, inventoryTable
    Dim groupBlockHtml, labBlockHtml, inventorySectionHtml

    manager = Trim(responsibleName)
    If manager = "" Then manager = "there"

    greetingName = FirstNameForGreeting(manager, "there")

    nameParts = Split(GetSenderDisplayName(), ", ")
    firstName = ""
    lastName = ""
    If UBound(nameParts) >= 1 Then
        lastName = Trim(nameParts(0))
        firstName = Trim(nameParts(1))
    ElseIf UBound(nameParts) = 0 Then
        firstName = Trim(nameParts(0))
    End If

    senderDisplayName = firstName
    If senderDisplayName = "" Then senderDisplayName = GetSenderFirstName()

    parcelAddressee = senderDisplayName
    If lastName <> "" Then
        If senderDisplayName <> "" Then
            parcelAddressee = senderDisplayName & " " & lastName
        Else
            parcelAddressee = lastName
        End If
    End If
    If parcelAddressee = "" Then parcelAddressee = "me"
    If senderDisplayName = "" Then senderDisplayName = parcelAddressee

    weekStartDate = GetLabManagerWeekStart(meetingDate)
    weekEndDate = GetLabManagerWeekEnd(meetingDate)
    startText = FormatDateTime(weekStartDate, 2)
    endText = FormatDateTime(weekEndDate, 2)
    weekRange = startText & " " & ChrW(&H2013) & " " & endText
    meetingDateText = FormatDateTime(meetingDate, 2)
    If LCase(greetingName) = "there" Then
        greetingDisplay = greetingName
    Else
        greetingDisplay = Trim(greetingName)
    End If

    groupBlockHtml = ReminderBlockStart("GROUP MEETING RESPONSIBILITIES") & _
                     ReminderTableRow("Lead", "Lead this week's group meeting on " & meetingDateText & " at 11:00.") & _
                     ReminderTableRow("Timing", "Keep each presenter to 15 minutes for the talk plus 5 minutes for Q&A.") & _
                     ReminderTableRow("Setup", "Arrive about 5 minutes early so the room computer is ready at the start.") & _
                     ReminderTableRow("Meeting flow", "Bring the red/yellow timing cards and keep the inclusion activity moving.") & _
                     ReminderTableRow("Lab-status update", "At the end of the meeting, give a brief lab-status update: what is working well, any minor reminders, and any persistent issues that need group attention.") & _
                     ReminderBlockEnd()

    labBlockHtml = "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; margin:12px 0 0 0; border:2px solid #0000cd; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
                   "<tr><td colspan=""2"" style=""padding:9px 14px 8px 14px; border-bottom:1px solid #e1e1e6; background-color:#ffffff; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; letter-spacing:0; mso-line-height-rule:exactly;"">LAB MANAGEMENT RESPONSIBILITIES</td></tr>" & _
                   ReminderTableRow("Lab-practice support", "Be the supportive first point of contact for minor lab-practice questions and reminders during the week.") & _
                   ReminderTableRow("Agreed practices", "Help people adopt the agreed lab practices by validating what is working and gently redirecting small issues before they grow. This is not an enforcement role; persistent concerns can be raised constructively during the Friday lab-status update, focusing on behaviours rather than blame.") & _
                   ReminderTableRow("Cleaning", "Lab cleaning starts on Monday and must be completed by this Friday; please aim to finish by Friday EOD so the lab is ready for Monday.") & _
                   ReminderTableRow("Parcels", "On Friday, please pick up parcels addressed to " & parcelAddressee & "; most contain consumables ordered the previous week that need re-stocking.") & _
                   ReminderTableRow("Stores", "Please procure the missing items listed on the whiteboard from stores so we can maintain supplies.") & _
                   ReminderBlockEnd()

    inventoryTable = BuildInventoryReviewTable(inventoryReviewItems, meetingDate)
    inventorySectionHtml = ""
    If Trim(inventoryTable) <> "" Then
        inventorySectionHtml = ReminderPlainSection("INVENTORY REVIEW", _
            ReminderParagraph("Please check the following Labcup records physically in the lab. Confirm the chemical identity, location, quantity, hazards, and storage temperature. Complete the review and send feedback to " & INVENTORY_MANAGER_NAME & " by Friday using the feedback link below. I will be copied on the feedback email.") & _
            inventoryTable)
    End If

    content = ReminderGreeting("Hi " & greetingDisplay & ",") & _
              ReminderHeader("WEEKLY RESPONSIBILITIES", "Lab manager, group meeting chair/timekeeper, and weekly inventory reviewer", weekRange) & _
              ReminderParagraph("This is a reminder that you are the lab manager, group meeting chair/timekeeper, and weekly inventory reviewer for the week of " & weekRange & ".") & _
              groupBlockHtml & _
              labBlockHtml & _
              inventorySectionHtml & _
              ReminderSubtleParagraph("Keeping the meeting and lab work tidy and easy to follow is part of the weekly group rhythm.") & _
              ReminderSignoff(senderDisplayName)
    content = ReminderEmailShell(content)

    GenerateWeeklyResponsibilityReminderEmail = content
    LogEvent "Generated weekly responsibility reminder length: " & Len(content)
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

Function DispatchReminderEmail(toAddress, subjectText, bodyText, managedKey, emailType)
    Call LogVerbose("DispatchReminderEmail invoked for subject='" & subjectText & "'")
    DispatchReminderEmail = DispatchManagedEmail(OL, managedKey, subjectText, toAddress, bodyText, Empty, emailType, False)
End Function

Function DispatchReminderEmailWithCc(toAddress, ccAddress, subjectText, bodyText, managedKey, emailType)
    Call LogVerbose("DispatchReminderEmailWithCc invoked for subject='" & subjectText & "'")
    DispatchReminderEmailWithCc = DispatchManagedEmailWithCc(OL, managedKey, subjectText, toAddress, ccAddress, bodyText, Empty, emailType, False)
End Function

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

Function BuildSameDayMeetingReminderSubject(meetingDate, room)
    Call LogVerbose("BuildSameDayMeetingReminderSubject invoked")
    Dim meetingDateText, subjectSeparator
    subjectSeparator = " " & ChrW(&HB7) & " "
    meetingDateText = WeekdayName(Weekday(meetingDate), False) & " " & FormatDateTime(meetingDate, 2)
    If Trim(room) <> "" Then
        BuildSameDayMeetingReminderSubject = "Reminder: Group Meeting" & subjectSeparator & meetingDateText & subjectSeparator & "11:00" & subjectSeparator & Trim(room)
    Else
        BuildSameDayMeetingReminderSubject = "Reminder: Group Meeting" & subjectSeparator & meetingDateText & subjectSeparator & "11:00"
    End If
End Function

Function BuildLabManagerReminderSubject(labManagerName, meetingDate)
    Call LogVerbose("BuildLabManagerReminderSubject invoked")
    Dim weekStartDate, weekEndDate, startText, endText
    weekStartDate = GetLabManagerWeekStart(meetingDate)
    weekEndDate = GetLabManagerWeekEnd(meetingDate)
    startText = FormatDateTime(weekStartDate, 2)
    endText = FormatDateTime(weekEndDate, 2)
    BuildLabManagerReminderSubject = "Lab management reminder (" & startText & " " & ChrW(&H2013) & " " & endText & ")"
End Function

Function BuildWeeklyResponsibilityReminderSubject(meetingDate)
    Call LogVerbose("BuildWeeklyResponsibilityReminderSubject invoked")
    Dim weekStartDate, weekEndDate, startText, endText
    weekStartDate = GetLabManagerWeekStart(meetingDate)
    weekEndDate = GetLabManagerWeekEnd(meetingDate)
    startText = FormatDateTime(weekStartDate, 2)
    endText = FormatDateTime(weekEndDate, 2)
    BuildWeeklyResponsibilityReminderSubject = "Weekly lab, meeting, and inventory responsibilities (" & startText & " " & ChrW(&H2013) & " " & endText & ")"
End Function

Sub TriggerReminderEmails(literaturePresenter, resultsPresenter, labManagerName, meetingDate, ByRef inventoryReviewItems, ByRef objWorkbook)
    Call LogVerbose("TriggerReminderEmails invoked")
    LogEvent "Triggering reminder workflow for meeting date " & FormatDateTime(meetingDate, 2)
    MaybeSendPresenterReminder literaturePresenter, "Literature", meetingDate
    MaybeSendPresenterReminder resultsPresenter, "Results", meetingDate
    MaybeSendWeeklyResponsibilityReminder labManagerName, meetingDate, inventoryReviewItems, objWorkbook
End Sub

Sub MaybeSendPresenterReminder(presenterName, presentationType, meetingDate)
    Call LogVerbose("MaybeSendPresenterReminder invoked for presenter='" & presenterName & "'")
    Dim trimmedName, subjectLine, targetEmail, body, emailType, managedKey
    trimmedName = Trim(presenterName)

    If trimmedName = "" Then
        LogEvent "Skipping " & presentationType & " presenter reminder because no presenter is assigned."
        Exit Sub
    End If

    targetEmail = ResolveEmailOrManager(trimmedName)
    subjectLine = BuildPresenterReminderSubject(presentationType, meetingDate)
    emailType = NormalizeManagedEmailType(presentationType & "-REMINDER")
    managedKey = BuildManagedEmailKey(meetingDate, emailType)

    body = GeneratePresenterReminderEmail(trimmedName, presentationType, meetingDate)
    DispatchReminderEmail targetEmail, subjectLine, body, managedKey, emailType
End Sub

Sub MaybeSendMeetingLeadReminder(meetingLeadName, meetingDate)
    Call LogVerbose("MaybeSendMeetingLeadReminder invoked for meeting lead='" & meetingLeadName & "'")
    Dim trimmedName, subjectLine, targetEmail, body, emailType, managedKey
    trimmedName = Trim(meetingLeadName)

    targetEmail = ResolveEmailOrManager(trimmedName)
    subjectLine = BuildMeetingLeadReminderSubject(meetingDate)
    emailType = "MEETING-LEAD-REMINDER"
    managedKey = BuildManagedEmailKey(meetingDate, emailType)

    body = GenerateMeetingLeadReminderEmail(trimmedName, meetingDate)
    DispatchReminderEmail targetEmail, subjectLine, body, managedKey, emailType
End Sub

Function GenerateSameDayMeetingReminderEmail(meetingDate, room)
    Call LogVerbose("GenerateSameDayMeetingReminderEmail invoked")
    Dim content, senderFirstName, meetingDateText, locationText, dotText, summaryText

    meetingDateText = WeekdayName(Weekday(meetingDate), False) & " " & FormatDateTime(meetingDate, 2)
    senderFirstName = GetSenderFirstName()
    dotText = " " & ChrW(&HB7) & " "
    If Trim(room) <> "" Then
        locationText = Trim(room)
    Else
        locationText = "TBC"
    End If
    summaryText = meetingDateText & dotText & "11:00" & dotText & locationText

    content = ReminderGreeting("Hi all,") & _
              "<table role=""presentation"" width=""660"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:660px; border-collapse:collapse; border:2px solid #0000cd; mso-table-lspace:0pt; mso-table-rspace:0pt;"">" & _
              "<tr><td style=""padding:9px 14px 3px 14px; font-family:'Imperial Sans Text', Arial, sans-serif; color:#708090; font-size:13px; font-weight:600; line-height:18px; letter-spacing:0; mso-line-height-rule:exactly;"">TODAY &middot; GROUP MEETING</td></tr>" & _
              "<tr><td style=""padding:0 14px 9px 14px; font-family:'Imperial Sans Display', 'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:18px; font-weight:600; line-height:24px; mso-line-height-rule:exactly;"">" & HtmlEscape(summaryText) & "</td></tr>" & _
              "</table>" & _
              ReminderSubtleParagraph("Just a quick reminder that our group meeting is today. See you soon.") & _
              ReminderSignoff(senderFirstName)
    content = ReminderEmailShell(content)

    GenerateSameDayMeetingReminderEmail = content
    LogEvent "Generated same-day meeting reminder length: " & Len(content)
End Function

Sub MaybeSendSameDayMeetingReminder(meetingDate, room, alternateText)
    Call LogVerbose("MaybeSendSameDayMeetingReminder invoked")
    Dim subjectLine, body, emailType, managedKey

    If IsNoMeetingEvent(alternateText) Then
        LogEvent "Skipping same-day meeting reminder " & GetNoMeetingReasonText(alternateText) & "."
        Exit Sub
    End If

    If Not IsSameDayBeforeMeeting(meetingDate) Then
        LogEvent "Skipping same-day meeting reminder because it is not the meeting day before 11:00."
        Exit Sub
    End If

    subjectLine = BuildSameDayMeetingReminderSubject(meetingDate, room)
    emailType = "SAME-DAY-REMINDER"
    managedKey = BuildManagedEmailKey(meetingDate, emailType)

    body = GenerateSameDayMeetingReminderEmail(meetingDate, room)
    DispatchReminderEmail EMAIL_TO, subjectLine, body, managedKey, emailType
End Sub

Sub MaybeSendWeeklyResponsibilityReminder(labManagerName, meetingDate, ByRef inventoryReviewItems, ByRef objWorkbook)
    Call LogVerbose("MaybeSendWeeklyResponsibilityReminder invoked for manager='" & labManagerName & "'")
    Dim trimmedName, subjectLine, targetEmail, ccAddress, body, emailType, managedKey, dispatchStatus
    trimmedName = Trim(labManagerName)

    If trimmedName = "" Then
        LogEvent "WARNING: No lab manager assigned for " & FormatDateTime(meetingDate, 2) & ". Weekly responsibility reminder will fall back to the manager address."
    End If

    targetEmail = ResolveEmailOrManager(trimmedName)
    ccAddress = InventoryManagerCcFor(targetEmail)
    subjectLine = BuildWeeklyResponsibilityReminderSubject(meetingDate)
    emailType = "LAB-MANAGER-REMINDER"
    managedKey = BuildManagedEmailKey(meetingDate, emailType)

    body = GenerateWeeklyResponsibilityReminderEmail(trimmedName, meetingDate, inventoryReviewItems)
    dispatchStatus = DispatchReminderEmailWithCc(targetEmail, ccAddress, subjectLine, body, managedKey, emailType)

    If DispatchStatusSubmitted(dispatchStatus) Then
        StampInventoryReviewItems objWorkbook, inventoryReviewItems, meetingDate
    ElseIf Not (inventoryReviewItems Is Nothing) Then
        If inventoryReviewItems.Count > 0 Then
            LogEvent "Inventory review dates not stamped because weekly responsibility reminder dispatch status was '" & dispatchStatus & "'."
        End If
    End If
End Sub

Sub MaybeSendLabManagerReminder(labManagerName, meetingDate)
    Call LogVerbose("MaybeSendLabManagerReminder invoked for manager='" & labManagerName & "'")
    Dim trimmedName, subjectLine, targetEmail, body, emailType, managedKey
    trimmedName = Trim(labManagerName)

    targetEmail = ResolveEmailOrManager(trimmedName)
    subjectLine = BuildLabManagerReminderSubject(trimmedName, meetingDate)
    emailType = "LAB-MANAGER-REMINDER"
    managedKey = BuildManagedEmailKey(meetingDate, emailType)

    body = GenerateLabManagerReminderEmail(trimmedName, meetingDate)
    DispatchReminderEmail targetEmail, subjectLine, body, managedKey, emailType
End Sub

'==============================================================
' Standard Email Dispatch
'==============================================================
Sub DistributeEmail(header, body, meetingDate)
    Call LogVerbose("DistributeEmail invoked with subject='" & header & "'")
    Dim managedKey, emailType
    LogEvent "Preparing regular meeting email with subject: " & header

    emailType = "REGULAR"
    managedKey = BuildManagedEmailKey(meetingDate, emailType)
    DispatchManagedEmail OL, managedKey, header, EMAIL_TO, body, Empty, emailType, False
End Sub

'==============================================================
' Alert Email Helpers
'==============================================================
Function GenerateEmailAlert(alertType, missingDate, recipientEmail, tentativeRota)
    Call LogVerbose("GenerateEmailAlert invoked for alertType='" & alertType & "'")
    Dim content, blockHtml, extraHtml, titleText, blockTitle, contactText
    Dim roomRequestMailto, dotText
    dotText = " " & ChrW(&HB7) & " "
    extraHtml = ""
    blockTitle = "ROTA CHECK"
    titleText = "Rota check" & dotText & missingDate

    Select Case alertType
        Case "MissingDate"
            titleText = "Missing rota entry" & dotText & missingDate
            blockTitle = "ROTA ENTRY"
            blockHtml = ReminderBlockStart(blockTitle) & _
                        ReminderTableRow("Rota date", missingDate) & _
                        ReminderTableRow("Issue", "There is no entry in the rota for " & missingDate & ".") & _
                        ReminderTableRow("Rota update", "Please update the rota accordingly.") & _
                        ReminderBlockEnd()
            If Trim(SafeTextValue(tentativeRota, "")) <> "" Then
                extraHtml = ReminderPlainSection("TENTATIVE ROTA", "<div style=""font-family:'Courier New', Consolas, monospace; font-size:13px; line-height:18px; color:#000000;"">" & HtmlMultilineText(tentativeRota) & "</div>")
            End If
        Case "Unassigned"
            titleText = "Unassigned presenters" & dotText & missingDate
            blockTitle = "PRESENTER ASSIGNMENT"
            blockHtml = ReminderBlockStart(blockTitle) & _
                        ReminderTableRow("Rota date", missingDate) & _
                        ReminderTableRow("Issue", "The rota entry for " & missingDate & " has no assigned presenters.") & _
                        ReminderTableRow("Rota update", "Please review the tentative 6-month rota below and approve or adjust:") & _
                        ReminderBlockEnd()
            If Trim(SafeTextValue(tentativeRota, "")) <> "" Then
                extraHtml = ReminderPlainSection("TENTATIVE ROTA", "<div style=""font-family:'Courier New', Consolas, monospace; font-size:13px; line-height:18px; color:#000000;"">" & HtmlMultilineText(tentativeRota) & "</div>")
            End If
        Case "MissingRoom"
            titleText = "Missing room reservation" & dotText & missingDate
            blockTitle = "ROOM RESERVATION"
            blockHtml = ReminderBlockStart(blockTitle) & _
                        ReminderTableRow("Rota date", missingDate) & _
                        ReminderTableRow("Issue", "The rota entry for " & missingDate & " has no room reservation.") & _
                        ReminderTableRow("Rota update", "Please reserve a room for the meeting.") & _
                        ReminderBlockEnd()
            roomRequestMailto = "mailto:" & BUILDING_ADMIN_EMAIL & "?subject=Room%20Reservation%20Request%20for%20ABT%20Group%20Meetings%20%E2%80%93%20Upcoming%20Term&body=Hi%20building%20manager,%0D%0A%0D%0AI%20hope%20you%20are%20well.%0D%0A%0D%0AI%20would%20like%20to%20request%20a%20recurring%20booking%20for%20our%20ABT%20group%20meetings%20for%20the%20upcoming%20academic%20term.%20Could%20you%20please%20reserve%20G25%20(MSRH)%20for%20Fridays%20from%2011:00%20to%2012:00,%20starting%20from%20the%20next%20available%20week?%0D%0A%0D%0ACould%20you%20also%20let%20me%20know%20the%20exact%20date%20the%20reservation%20ends?%0D%0A%0D%0AThank%20you%20very%20much%20for%20your%20help.%0D%0A%0D%0ABest%20regards,%0D%0AGiacomo"
            extraHtml = "<p style=""margin:12px 0 0 0; font-family:'Imperial Sans Text', Arial, sans-serif; color:#000000; font-size:15px; line-height:22px; mso-line-height-rule:exactly;""><a href=""" & HtmlEscape(roomRequestMailto) & """ style=""color:#000000; text-decoration:underline; font-weight:600;"">Send email to the building manager to request room for the upcoming term</a></p>"
        Case "MeetingNotSet"
            titleText = "Meeting not set" & dotText & missingDate
            blockTitle = "MEETING SETUP"
            blockHtml = ReminderBlockStart(blockTitle) & _
                        ReminderTableRow("Rota date", missingDate) & _
                        ReminderTableRow("Issue", "The rota entry for " & missingDate & " has no presenters or alternate assigned, so the regular meeting email was not sent to the distribution list.") & _
                        ReminderTableRow("Rota update", "Please update the rota with presenters so the meeting can be scheduled.") & _
                        ReminderBlockEnd()
        Case Else
            blockHtml = ReminderBlockStart(blockTitle) & _
                        ReminderTableRow("Rota date", missingDate) & _
                        ReminderTableRow("Issue", "Please review the meeting rota.") & _
                        ReminderBlockEnd()
    End Select

    contactText = "Contact directly via email: " & recipientEmail
    content = ReminderGreeting("Hi,") & _
              ReminderHeader("ROTA CHECK", titleText, "") & _
              blockHtml & _
              extraHtml & _
              ReminderSubtleParagraph(contactText) & _
              ReminderSignoff("Automated Meeting Scheduler")
    content = ReminderEmailShell(content)
    GenerateEmailAlert = content
End Function

Sub AlertEmail(toAddress, messageBody, subjectText)
    Call LogVerbose("AlertEmail invoked to='" & toAddress & "'")
    Dim mail
    If HasAlertBeenSentToday(subjectText) Then
        LogEvent "Alert email for subject '" & subjectText & "' was already sent today. Skipping duplicate dispatch."
        Exit Sub
    End If

    LogEvent "Preparing alert email with subject: " & subjectText & " to " & toAddress
    Set mail = OL.CreateItem(0)
    With mail
        .To = toAddress
        .Subject = subjectText
        If IsHtmlBody(messageBody) Then
            .BodyFormat = 2
            .HTMLBody = messageBody
        Else
            .Body = messageBody
        End If
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
            ' Recheck in case this run continued past midnight into the weekend.
            If IsWeekend(Date) Then
                LogEvent "Weekend: email delivery skipped. Subject: " & .Subject
                Exit Sub
            End If
            .Send
            If EmailExists(subjectText, 5, True) Then
                LogEvent "Negative feedback loop confirmed for alert subject: " & subjectText
            Else
                LogEvent "WARNING: Alert subject not observed in Sent Items after sending: " & subjectText
            End If
        End If
    End With
End Sub

Sub AlertEmailWithAttachment(toAddress, messageBody, subjectText, attachmentPath)
    Call LogVerbose("AlertEmailWithAttachment invoked to='" & toAddress & "' with attachment='" & attachmentPath & "'")
    Dim mail
    If HasAlertBeenSentToday(subjectText) Then
        LogEvent "Alert email with attachment for subject '" & subjectText & "' was already sent today. Skipping duplicate dispatch."
        Exit Sub
    End If

    LogEvent "Preparing alert email with attachment. Subject: " & subjectText & _
                  " | Recipient: " & toAddress & " | Attachment: " & attachmentPath
    Set mail = OL.CreateItem(0)
    With mail
        .To = toAddress
        .Subject = subjectText
        If IsHtmlBody(messageBody) Then
            .BodyFormat = 2
            .HTMLBody = messageBody
        Else
            .Body = messageBody
        End If
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
            ' Recheck in case this run continued past midnight into the weekend.
            If IsWeekend(Date) Then
                LogEvent "Weekend: email delivery skipped. Subject: " & .Subject
                Exit Sub
            End If
            .Send
            If EmailExists(subjectText, 5, True) Then
                LogEvent "Negative feedback loop confirmed for alert subject: " & subjectText
            Else
                LogEvent "WARNING: Alert subject not observed in Sent Items after sending: " & subjectText
            End If
        End If
    End With
End Sub

'==============================================================
' Tentative Rota Helpers
'==============================================================
Function GenerateTentativeRota(objSheetEmails)
    Call LogVerbose("GenerateTentativeRota invoked")
    Dim lastRow, presenterColumn, i, numNames, names(), rndIndex, tmp
    Dim rotaDates(), rotaAssignments(), rotaBody, dt, totalWeeks
    Dim csvFile, txtStream
    LogEvent "Starting tentative rota generation"

    lastRow = objSheetEmails.Cells(objSheetEmails.Rows.Count, "A").End(-4162).Row
    presenterColumn = FindHeaderColumn(objSheetEmails, "Presenter", 5)
    LogEvent "Using Emails sheet Presenter column: " & presenterColumn
    numNames = 0
    For i = 2 To lastRow
        If IsFlagSet(objSheetEmails.Cells(i, presenterColumn).Value) Then numNames = numNames + 1
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
        If IsFlagSet(objSheetEmails.Cells(i, presenterColumn).Value) Then
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
    objExcel.DisplayAlerts = False
    objExcel.AskToUpdateLinks = False
    On Error GoTo 0
    LogEvent "Excel application created successfully."

    On Error Resume Next
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_PATH, 0, False)
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
    Dim savedForUpload
    savedForUpload = False

    If CBool(meetingWorkbookNeedsUpload) Then
        On Error Resume Next
        LogEvent "Saving edited meeting rota workbook before upload."
        objWorkbook.Save
        If Err.Number <> 0 Then
            LogEvent "ERROR: Failed to save edited meeting rota workbook. Keeping local workbook for recovery. Err: " & Err.Number & " - " & Err.Description
            keepSyncedWorkbookForRecovery = True
            Err.Clear
        Else
            savedForUpload = True
            LogEvent "Edited meeting rota workbook saved locally."
        End If
        On Error GoTo 0
    End If

    objWorkbook.Close False
    objExcel.Quit
    Set objExcel = Nothing
    Set objWorkbook = Nothing
    Set objSheetRota = Nothing
    Set objSheetEmails = Nothing
    LogEvent "Coordinator routine completed and Excel resources released."

    If CBool(savedForUpload) Then
        If UploadMeetingWorkbook() Then
            meetingWorkbookNeedsUpload = False
            LogEvent "Edited meeting rota workbook uploaded successfully."
        Else
            keepSyncedWorkbookForRecovery = True
            LogEvent "ACTION REQUIRED: Edited meeting rota workbook upload failed. Local workbook kept for recovery: " & EXCEL_PATH
        End If
    End If
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

    If HasAlertBeenSentToday(subject) Then
        LogEvent "Alert '" & subject & "' already dispatched today. Skipping duplicate send."
        Exit Sub
    End If

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
                SendCriticalAlert "Unassigned", fridayDates(idx), MANAGER, objSheetEmails
            End If

            If Trim(tmpRoom) = "" And Not IsNoMeetingEvent(tmpAlternate) Then
                SendCriticalAlert "MissingRoom", fridayDates(idx), MANAGER, objSheetEmails
            ElseIf Trim(tmpRoom) = "" Then
                LogEvent "Skipping missing-room alert for " & FormatDateTime(fridayDates(idx), 2) & " " & GetNoMeetingReasonText(tmpAlternate) & "."
            End If
        End If
    Next
End Sub

Sub HandleCommunications(ByRef fridayDates, ByRef rowIndexes, ByRef objSheetRota, ByRef objWorkbook)
    Call LogVerbose("HandleCommunications invoked")
    Dim currentRow, literature, results, labManager, alternate, room, body, header, managerAlertSubject
    Dim inventoryReviewItems
    currentRow = rowIndexes(0)

    If currentRow = -1 Then
        LogEvent "No rota entry found for this Friday. Regular email and reminders will not be processed."
        Exit Sub
    End If

    literature = Trim(objSheetRota.Cells(currentRow, 2).Value)
    results    = Trim(objSheetRota.Cells(currentRow, 3).Value)
    labManager = Trim(objSheetRota.Cells(currentRow, 6).Value)
    alternate  = Trim(objSheetRota.Cells(currentRow, 4).Value)
    room       = Trim(objSheetRota.Cells(currentRow, 5).Value)

    If IsNoMeetingEvent(alternate) Then
        LogEvent "No-meeting event detected for " & FormatDateTime(fridayDates(0), 2) & " " & GetNoMeetingReasonText(alternate) & ". Sending no-meeting notice and suppressing meeting reminders."
        header = BuildGroupEmailSubject(fridayDates(0), alternate, room)
        body = GenerateEmailContent(fridayDates(0), currentRow, objSheetRota, rowIndexes(1))
        DistributeEmail header, body, fridayDates(0)
        Exit Sub
    End If

    If literature = "" And results = "" And alternate = "" Then
        LogEvent "No presenters or alternate assigned for " & FormatDateTime(fridayDates(0), 2) & ". Regular meeting email suppressed."
        managerAlertSubject = "ALERT: MeetingNotSet for " & FormatDateTime(fridayDates(0), 2)
        If Not HasAlertBeenSentToday(managerAlertSubject) Then
            LogEvent "Sending manager alert that the meeting is not set for " & FormatDateTime(fridayDates(0), 2)
            Call AlertEmail(MANAGER, _
                GenerateEmailAlert("MeetingNotSet", FormatDateTime(fridayDates(0), 2), MANAGER, ""), _
                managerAlertSubject)
        Else
            LogEvent "Manager alert already exists for " & FormatDateTime(fridayDates(0), 2) & ". Skipping duplicate notification."
        End If
        Exit Sub
    End If

    Set inventoryReviewItems = SelectInventoryReviewItems(objWorkbook, fridayDates(0))
    TriggerReminderEmails literature, results, labManager, fridayDates(0), inventoryReviewItems, objWorkbook
    MaybeSendSameDayMeetingReminder fridayDates(0), room, alternate

    header = BuildGroupEmailSubject(fridayDates(0), alternate, room)
    body = GenerateEmailContent(fridayDates(0), currentRow, objSheetRota, rowIndexes(1))
    LogEvent "Dispatching regular meeting email for " & FormatDateTime(fridayDates(0), 2)
    DistributeEmail header, body, fridayDates(0)
End Sub

'==============================================================
' Coordinator (Main Workflow)
'==============================================================
Sub Coordinator()
    Call LogVerbose("Coordinator invoked")
    Dim fridayDates(2), rowIndexes(2)
    Dim objExcel, objWorkbook, objSheetRota, objSheetEmails

    LogEvent "Coordinator routine started"

    PopulateFridayDates fridayDates

    If Not LoadMeetingWorkbook(objExcel, objWorkbook, objSheetRota, objSheetEmails) Then
        Exit Sub
    End If

    LocateMeetingRows fridayDates, objSheetRota, rowIndexes
    EvaluateCriticalWindows fridayDates, rowIndexes, objSheetRota, objSheetEmails
    HandleCommunications fridayDates, rowIndexes, objSheetRota, objWorkbook

    CloseMeetingWorkbook objExcel, objWorkbook, objSheetRota, objSheetEmails
End Sub

'==============================================================
' Script Entrypoint
'==============================================================
Sub Main()
    Call LogVerbose("Main invoked")
    InitializeFileSystem
    InitializeLogging

    If IsWeekend(Date) Then
        LogEvent "Weekend: skipping all group meeting emails, reminders, and alerts."
        LogEvent "=== Script execution completed ==="
        Exit Sub
    End If

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
