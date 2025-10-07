Option Explicit

Dim objShell, strCmd
Set objShell = CreateObject("WScript.Shell")

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Build the command
strCmd = "C:\rclone\rclone.exe copy ""ABTsharepoint:/Group Meetings/meeting_rota.xlsx"" ""C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings"""

' Run the command, show window (1) and wait until it finishes (True)
objShell.Run strCmd, 0, True

'==============================================================
' Group Meeting Email Automation (Integrated & Optimized)
'==============================================================

'-----------------------------
' Global Configuration
'-----------------------------
Const DEBUG_MODE = False   ' <<< set False to actually send emails
Const MANAGER   = "gd2613@ic.ac.uk"
Const BUILDING_ADMIN_EMAIL   = "i.gomezvilchis@imperial.ac.uk"
Const SUPERVISOR_EMAIL   = "agi@imperial.ac.uk"

Const EMAIL_TO = "brandt-talbot-group@imperial.ac.uk"
Const EXCEL_PATH = "C:\Users\giaco\OneDrive - Imperial College London\Work\group_meetings\meeting_rota.xlsx"
Const SHEET_ROTA = "Rota"
Const SHEET_EMAILS = "Emails"

Dim OL, olNS
Set OL = CreateObject("Outlook.Application")
Set olNS = OL.GetNamespace("MAPI")

'==============================================================
' Module 1: Email Verification
'==============================================================
Function EmailExists(header)
    Dim inbox, items, mail, found
    found = False
    Set inbox = olNS.GetDefaultFolder(6) ' Inbox
    Set items = inbox.Items
    items.Sort "ReceivedTime", True
    For Each mail In items
        If mail.Subject = header Then
            found = True
            Exit For
        End If
    Next
    EmailExists = found
End Function

'==============================================================
' Module 2: Email Generation (Meeting)
'==============================================================
Function GenerateEmailContent(meetingDate, row, objSheetRota, nextFridayRow)
    Dim literature, results, alternate, room, next_lit, next_res, next_alt, content, nameParts
    literature = objSheetRota.Cells(row, 2).Value
    results    = objSheetRota.Cells(row, 3).Value
    alternate  = objSheetRota.Cells(row, 4).Value
    room       = objSheetRota.Cells(row, 5).Value
    
    ' Next Friday content
    If nextFridayRow <> -1 Then
        next_lit = objSheetRota.Cells(nextFridayRow, 2).Value
        next_res = objSheetRota.Cells(nextFridayRow, 3).Value
        next_alt = objSheetRota.Cells(nextFridayRow, 4).Value
    Else
        next_lit = ""
        next_res = ""
        next_alt = ""
    End If
    
    nameParts = Split(olNS.Accounts.Item(1).CurrentUser, ", ")
    
    content = "Hi all," & vbCrLf & vbCrLf
    If Trim(room) <> "" Then
        content = content & "We will have our group meeting on " & meetingDate & _
                  " at 11:00 am in " & room & " in person or via MS Teams." & vbCrLf & vbCrLf
    Else
        content = content & "We will have our group meeting on " & meetingDate & _
                  " at 11:00 am in person or via MS Teams." & vbCrLf & vbCrLf
    End If
    
    content = content & "Meeting agenda:" & vbCrLf
    If alternate <> "" Then
        content = content & "     " & alternate & vbCrLf
    Else
        content = content & "     " & results & " will present results." & vbCrLf
        content = content & "     " & literature & " will present literature." & vbCrLf & vbCrLf
    End If
    
    If next_lit <> "" And next_res <> "" Then
        content = content & "The week after:" & vbCrLf
        If next_alt <> "" Then
            content = content & "     " & next_alt & vbCrLf
        Else
            content = content & "     " & next_res & " will present results." & vbCrLf
            content = content & "     " & next_lit & " will present literature." & vbCrLf & vbCrLf
        End If
    End If

    ' Add link to rota
    content = content & vbCrLf & "Upcoming weeks schedule: " & _
              "https://imperiallondon.sharepoint.com/sites/Brandt-TalbotGroup/Shared%20Documents/Group%20Meetings/meeting_rota.xlsx?web=1" & vbCrLf & vbCrLf

    content = content & vbCrLf & "Best," & vbCrLf & Trim(nameParts(1))
    GenerateEmailContent = content
End Function

'==============================================================
' Module 3: Email Distribution
'==============================================================
Sub DistributeEmail(header, body)
    Dim mail
    Set mail = OL.CreateItem(0)
    With mail
        .To = EMAIL_TO
        .Subject = header
        .Body = body
        If DEBUG_MODE Then
            MsgBox "[DEBUG] Would send email:" & vbCrLf & vbCrLf & header & vbCrLf & body
        Else
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
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = messageBody
        If DEBUG_MODE Then
            MsgBox "[DEBUG] Would send ALERT email to " & toAddress & ":" & vbCrLf & _
                "Subject: " & subjectText & vbCrLf & vbCrLf & messageBody
        Else
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
    With mail
        .To = toAddress
        .Subject = subjectText
        .Body = messageBody
        If attachmentPath <> "" Then
            .Attachments.Add attachmentPath
        End If
        If DEBUG_MODE Then
            MsgBox "[DEBUG] Would send ALERT email with attachment to " & toAddress & ":" & vbCrLf & _
                   "Subject: " & subjectText & vbCrLf & vbCrLf & messageBody & vbCrLf & _
                   "Attachment: " & attachmentPath
        Else
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

    '--- Collect names with column D = 1 ---
    lastRow = objSheetEmails.Cells(objSheetEmails.Rows.Count, "A").End(-4162).Row
    numNames = 0
    For i = 2 To lastRow
        If objSheetEmails.Cells(i, 4).Value = 1 Then numNames = numNames + 1
    Next
    If numNames = 0 Then
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

    '--- Shuffle names ---
    Randomize
    For i = 0 To UBound(names)
        rndIndex = Int(Rnd * (UBound(names) + 1))
        tmp = names(i)
        names(i) = names(rndIndex)
        names(rndIndex) = tmp
    Next

    '--- Generate tentative dates and assignments ---
    totalWeeks = 26
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
    Set txtStream = fso.CreateTextFile(csvFile, True)
    txtStream.WriteLine "Date,Presenter1,Presenter2"
    For i = 0 To totalWeeks - 1
        txtStream.WriteLine FormatDateTime(rotaDates(i), 2) & "," & _
                            rotaAssignments(i, 0) & "," & rotaAssignments(i, 1)
    Next
    txtStream.Close

    GenerateTentativeRota = rotaBody & vbCrLf & "[CSV: " & csvFile & "]"
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
End Sub

'==============================================================
' Module 7: Coordinator (Optimized)
'==============================================================
Sub Coordinator()
    Dim thisFriday, nextFriday, nextNextFriday
    Dim header, objExcel, objWorkbook, objSheetRota, objSheetEmails 
    Dim lastRow, i, meetingDate
    Dim foundRow, nextFridayRow, nextNextFridayRow
    Dim body, literature, results, alternate, room

    '--- Define Fridays ---
    thisFriday = Date + ((5 - Weekday(Date, vbMonday) + 7) Mod 7)
    nextFriday = thisFriday + 7
    nextNextFriday = thisFriday + 14
    header = "ABT group meeting @ 11am (" & FormatDateTime(thisFriday, 2) & ")"

    '--- Exit if regular meeting email already exists ---
    If EmailExists(header) Then Exit Sub

    '--- Open Excel ---
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_PATH)
    Set objSheetRota = objWorkbook.Sheets(SHEET_ROTA)
    Set objSheetEmails = objWorkbook.Sheets(SHEET_EMAILS)

    lastRow = objSheetRota.Cells(objSheetRota.Rows.Count, "A").End(-4162).Row
    foundRow = -1: nextFridayRow = -1: nextNextFridayRow = -1

    '--- Locate rows for Fridays ---
    For i = 2 To lastRow
        meetingDate = objSheetRota.Cells(i, 1).Value
        If meetingDate = thisFriday Then foundRow = i
        If meetingDate = nextFriday Then nextFridayRow = i
        If meetingDate = nextNextFriday Then nextNextFridayRow = i
    Next

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
        Else
            tmpLiterature = objSheetRota.Cells(rowNum, 2).Value
            tmpResults    = objSheetRota.Cells(rowNum, 3).Value
            tmpAlternate  = objSheetRota.Cells(rowNum, 4).Value
            tmpRoom       = objSheetRota.Cells(rowNum, 5).Value

            If Trim(tmpLiterature) = "" And Trim(tmpResults) = "" And Trim(tmpAlternate) = "" Then
                criticalTypes(idx) = "Unassigned"
                recipientEmails(idx) = SUPERVISOR_EMAIL
            End If

            If Trim(tmpRoom) = "" Then
                ' MissingRoom is independent; we allow sending separately
                Dim missingRoomSubject
                missingRoomSubject = "ALERT: MissingRoom for " & FormatDateTime(criticalDates(idx), 2)
                If Not EmailExists(missingRoomSubject) Then
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
            If Not EmailExists(alertSubject) Then
                If criticalTypes(idx) = "MissingDate" Or criticalTypes(idx) = "Unassigned" Then
                    tentativeRota = GenerateTentativeRota(objSheetEmails)
                    ' Create CSV attachment
                    Dim fso, csvPath
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    csvPath = fso.GetSpecialFolder(2) & "\TentativeRota.csv"
                    WriteRotaCSV tentativeRota, csvPath
                Else
                    tentativeRota = ""
                    csvPath = ""
                End If
                Call AlertEmailWithAttachment(recipientEmails(idx), _
                    GenerateEmailAlert(criticalTypes(idx), FormatDateTime(criticalDates(idx), 2), recipientEmails(idx), tentativeRota), _
                    alertSubject, csvPath)
            End If
        End If
    Next

    '--- Send regular meeting email for this Friday ---
    If foundRow <> -1 Then
        body = GenerateEmailContent(thisFriday, foundRow, objSheetRota, nextFridayRow)
        Call DistributeEmail(header, body)
    End If

    '--- Clean up ---
    objWorkbook.Close False
    objExcel.Quit
    Set objExcel = Nothing
    Set objWorkbook = Nothing
    Set objSheetRota = Nothing
    Set objSheetEmails = Nothing
End Sub


'==============================================================
' Module 8: Scheduler (Entrypoint)
'==============================================================
Coordinator


' Check if the file exists before deleting
If fso.FileExists(EXCEL_PATH) Then
    fso.DeleteFile EXCEL_PATH, True  ' True = force deletion (even if read-only)
End If