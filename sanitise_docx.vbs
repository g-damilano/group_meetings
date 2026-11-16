Option Explicit

' -----------------------------
' Standalone Word sanitizer UI
' -----------------------------

Const wdAlertsNone = 0
Const wdFindStop = 0
Const wdReplaceAll = 2
Const wdFormatDocument = 0          ' .doc
Const wdFormatXMLDocument = 12      ' .docx
Const wdRDIAll = 99                 ' late-bound doc info removal (works in most Word versions)
Const msoDocInspectorStatusIssueFound = 2

Dim selectionResult, sourcePath, wantsSaveDialog
selectionResult = ShowOpenDialogWithCheckbox()
If Len(selectionResult) = 0 Then WScript.Quit 0

sourcePath = Split(selectionResult, "|")(0)
wantsSaveDialog = (LCase(Split(selectionResult, "|")(1)) = "true")

If Len(sourcePath) = 0 Then WScript.Quit 0
If Not FileExists(sourcePath) Then
    MsgBox "File not found:" & vbCrLf & sourcePath, vbExclamation, "Sanitize Word document"
    WScript.Quit 1
End If

Dim defaultOutputPath, outputPath
defaultOutputPath = BuildSanitizedPath(sourcePath, "_sanitized")

If wantsSaveDialog Then
    outputPath = ShowSaveAsDialog(defaultOutputPath)
    If Len(outputPath) = 0 Then WScript.Quit 0
Else
    outputPath = FindAvailablePath(defaultOutputPath)
End If

Dim errorText
errorText = SanitizeWordDocument(sourcePath, outputPath)

If Len(errorText) > 0 Then
    MsgBox "Sanitization failed:" & vbCrLf & errorText, vbCritical, "Sanitize Word document"
    WScript.Quit 2
Else
    MsgBox "Done." & vbCrLf & "Saved sanitized copy to:" & vbCrLf & outputPath, vbInformation, "Sanitize Word document"
End If


' -----------------------------
' Core sanitization
' -----------------------------
Function SanitizeWordDocument(ByVal inputPath, ByVal outputPath)
    On Error Resume Next

    Dim wordApplication, documentObject
    Set wordApplication = CreateObject("Word.Application")
    wordApplication.Visible = False
    wordApplication.DisplayAlerts = wdAlertsNone

    Set documentObject = wordApplication.Documents.Open(inputPath, False, False, False)

    If Err.Number <> 0 Then
        SanitizeWordDocument = "Could not open document. " & Err.Description
        CleanupWord wordApplication, documentObject
        Exit Function
    End If

    On Error GoTo 0
    On Error Resume Next

    ' 1) Accept all tracked changes, stop tracking
    If documentObject.Revisions.Count > 0 Then
        documentObject.AcceptAllRevisions
    End If
    documentObject.TrackRevisions = False

    ' 2) Delete all comments
    DeleteAllComments documentObject

    ' 3) Delete hidden text (Font.Hidden = True)
    DeleteHiddenText documentObject

    ' 4) Remove personal info / doc info (best-effort)
    Err.Clear
    documentObject.RemovePersonalInformation = True
    documentObject.RemoveDocumentInformation wdRDIAll

    ' 5) Run Document Inspectors (best-effort)
    RunDocumentInspectors documentObject

    ' 6) Save output (choose format based on extension)
    Dim fileFormat
    fileFormat = FileFormatForPath(outputPath)

    Err.Clear
    documentObject.SaveAs2 outputPath, fileFormat
    If Err.Number <> 0 Then
        SanitizeWordDocument = "Could not save output. " & Err.Description
        CleanupWord wordApplication, documentObject
        Exit Function
    End If

    documentObject.Close False
    wordApplication.Quit
    Set documentObject = Nothing
    Set wordApplication = Nothing

    SanitizeWordDocument = ""
End Function

Sub CleanupWord(wordApplication, documentObject)
    On Error Resume Next
    If Not documentObject Is Nothing Then documentObject.Close False
    If Not wordApplication Is Nothing Then wordApplication.Quit
End Sub

Sub DeleteAllComments(documentObject)
    On Error Resume Next
    Dim indexValue
    For indexValue = documentObject.Comments.Count To 1 Step -1
        documentObject.Comments(indexValue).Delete
    Next
End Sub

Sub DeleteHiddenText(documentObject)
    On Error Resume Next
    Dim searchRange
    Set searchRange = documentObject.Range

    With searchRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting

        .Text = ""
        .Replacement.Text = ""

        .Format = True
        .Font.Hidden = True

        .Forward = True
        .Wrap = wdFindStop

        ' VBScript does NOT support named args. Use positional Execute args.
        ' Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        '         MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
        .Execute "", False, False, False, False, False, True, wdFindStop, True, "", wdReplaceAll
    End With
End Sub

Sub RunDocumentInspectors(documentObject)
    On Error Resume Next

    Dim inspectors, inspector, results, statusValue, indexValue
    Set inspectors = documentObject.DocumentInspectors
    If inspectors Is Nothing Then Exit Sub

    For indexValue = 1 To inspectors.Count
        Set inspector = inspectors(indexValue)
        statusValue = inspector.Inspect(results)
        If statusValue = msoDocInspectorStatusIssueFound Then
            inspector.Fix results
        End If
    Next
End Sub

Function FileFormatForPath(ByVal filePath)
    Dim lowerPath
    lowerPath = LCase(filePath)

    If Right(lowerPath, 5) = ".docx" Then
        FileFormatForPath = wdFormatXMLDocument
    ElseIf Right(lowerPath, 4) = ".doc" Then
        FileFormatForPath = wdFormatDocument
    Else
        FileFormatForPath = wdFormatXMLDocument
    End If
End Function


' -----------------------------
' UI: Open + checkbox (single popup)
' -----------------------------
Function ShowOpenDialogWithCheckbox()
    Dim ps, outputText
    ps = ""
    ps = ps & "Add-Type -AssemblyName System.Windows.Forms; "
    ps = ps & "$form = New-Object System.Windows.Forms.Form; "
    ps = ps & "$form.Text = 'Select Word document to sanitize'; "
    ps = ps & "$form.Width = 680; $form.Height = 190; $form.StartPosition = 'CenterScreen'; "

    ps = ps & "$label = New-Object System.Windows.Forms.Label; "
    ps = ps & "$label.Text = 'Word file:'; $label.Left = 10; $label.Top = 20; $label.AutoSize = $true; "
    ps = ps & "$form.Controls.Add($label); "

    ps = ps & "$textbox = New-Object System.Windows.Forms.TextBox; "
    ps = ps & "$textbox.Left = 80; $textbox.Top = 16; $textbox.Width = 480; "
    ps = ps & "$form.Controls.Add($textbox); "

    ps = ps & "$browse = New-Object System.Windows.Forms.Button; "
    ps = ps & "$browse.Text = 'Browse...'; $browse.Left = 570; $browse.Top = 14; $browse.Width = 80; "
    ps = ps & "$form.Controls.Add($browse); "

    ps = ps & "$check = New-Object System.Windows.Forms.CheckBox; "
    ps = ps & "$check.Text = 'Choose save location (Save As...)'; $check.Left = 80; $check.Top = 55; $check.Width = 350; "
    ps = ps & "$form.Controls.Add($check); "

    ps = ps & "$ok = New-Object System.Windows.Forms.Button; "
    ps = ps & "$ok.Text = 'OK'; $ok.Left = 490; $ok.Top = 95; $ok.Width = 75; "
    ps = ps & "$cancel = New-Object System.Windows.Forms.Button; "
    ps = ps & "$cancel.Text = 'Cancel'; $cancel.Left = 575; $cancel.Top = 95; $cancel.Width = 75; "
    ps = ps & "$form.Controls.Add($ok); $form.Controls.Add($cancel); "

    ps = ps & "$browse.Add_Click({ "
    ps = ps & "  $dlg = New-Object System.Windows.Forms.OpenFileDialog; "
    ps = ps & "  $dlg.Filter = 'Word Documents (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*'; "
    ps = ps & "  $dlg.Multiselect = $false; "
    ps = ps & "  if ($dlg.ShowDialog() -eq 'OK') { $textbox.Text = $dlg.FileName } "
    ps = ps & "}); "

    ps = ps & "$ok.Add_Click({ "
    ps = ps & "  if ([string]::IsNullOrWhiteSpace($textbox.Text)) { "
    ps = ps & "    [System.Windows.Forms.MessageBox]::Show('Please choose a file.'); return "
    ps = ps & "  } "
    ps = ps & "  if (-not (Test-Path $textbox.Text)) { "
    ps = ps & "    [System.Windows.Forms.MessageBox]::Show('File not found.'); return "
    ps = ps & "  } "
    ps = ps & "  $form.Tag = ($textbox.Text + '|' + $check.Checked); $form.Close(); "
    ps = ps & "}); "

    ps = ps & "$cancel.Add_Click({ $form.Tag = ''; $form.Close(); }); "
    ps = ps & "$form.Topmost = $true; "
    ps = ps & "$form.ShowDialog() | Out-Null; "
    ps = ps & "Write-Output $form.Tag;"

    outputText = RunPowerShell(ps)
    ShowOpenDialogWithCheckbox = Trim(outputText)
End Function

' -----------------------------
' UI: Save As dialog (only if checkbox ticked)
' -----------------------------
Function ShowSaveAsDialog(ByVal defaultPath)
    Dim escaped, ps, outputText
    escaped = Replace(defaultPath, "'", "''")

    ps = ""
    ps = ps & "Add-Type -AssemblyName System.Windows.Forms; "
    ps = ps & "$p = '" & escaped & "'; "
    ps = ps & "$dlg = New-Object System.Windows.Forms.SaveFileDialog; "
    ps = ps & "$dlg.Filter = 'Word Document (*.docx)|*.docx|Word 97-2003 (*.doc)|*.doc|All files (*.*)|*.*'; "
    ps = ps & "$dlg.InitialDirectory = Split-Path $p; "
    ps = ps & "$dlg.FileName = Split-Path $p -Leaf; "
    ps = ps & "if ($dlg.ShowDialog() -eq 'OK') { Write-Output $dlg.FileName }"

    outputText = RunPowerShell(ps)
    ShowSaveAsDialog = Trim(outputText)
End Function

Function RunPowerShell(ByVal scriptText)
    Dim sh, ex, cmd
    Set sh = CreateObject("WScript.Shell")
    cmd = "powershell -NoProfile -STA -ExecutionPolicy Bypass -Command " & QuoteForCmd(scriptText)
    Set ex = sh.Exec(cmd)
    RunPowerShell = ex.StdOut.ReadAll
End Function

Function QuoteForCmd(ByVal textValue)
    QuoteForCmd = Chr(34) & Replace(textValue, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function


' -----------------------------
' Path helpers
' -----------------------------
Function BuildSanitizedPath(ByVal originalPath, ByVal suffixText)
    Dim folderPath, fileName, baseName, extensionName
    folderPath = GetFolderPath(originalPath)
    fileName = GetFileName(originalPath)
    baseName = RemoveExtension(fileName)
    extensionName = GetExtension(fileName)

    If Len(extensionName) = 0 Then extensionName = "docx"
    BuildSanitizedPath = folderPath & "\" & baseName & suffixText & "." & extensionName
End Function

Function FindAvailablePath(ByVal desiredPath)
    Dim candidatePath, counterValue
    candidatePath = desiredPath
    counterValue = 1

    Do While FileExists(candidatePath)
        candidatePath = InsertCounterBeforeExtension(desiredPath, counterValue)
        counterValue = counterValue + 1
    Loop

    FindAvailablePath = candidatePath
End Function

Function InsertCounterBeforeExtension(ByVal filePath, ByVal counterValue)
    Dim folderPath, fileName, baseName, extensionName
    folderPath = GetFolderPath(filePath)
    fileName = GetFileName(filePath)
    baseName = RemoveExtension(fileName)
    extensionName = GetExtension(fileName)

    If Len(extensionName) > 0 Then
        InsertCounterBeforeExtension = folderPath & "\" & baseName & "_" & CStr(counterValue) & "." & extensionName
    Else
        InsertCounterBeforeExtension = folderPath & "\" & baseName & "_" & CStr(counterValue)
    End If
End Function

Function GetFolderPath(ByVal filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderPath = fso.GetParentFolderName(filePath)
End Function

Function GetFileName(ByVal filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(filePath)
End Function

Function RemoveExtension(ByVal fileName)
    Dim pos
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        RemoveExtension = Left(fileName, pos - 1)
    Else
        RemoveExtension = fileName
    End If
End Function

Function GetExtension(ByVal fileName)
    Dim pos
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        GetExtension = Mid(fileName, pos + 1)
    Else
        GetExtension = ""
    End If
End Function

Function FileExists(ByVal filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function