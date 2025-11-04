'======================================================================
' test_dynamic_links.vbs — Load external HTML, inject dynamic links, display mail
'======================================================================
Option Explicit
SetLocale 1033

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName)
If Len(folder)=0 Then folder = "."

' Load mailer code
ExecuteGlobal fso.OpenTextFile(folder & "\mail_utils_mailer_dynamiclinks.vbs", 1).ReadAll

' Load HTML template (UTF-8)
Dim htmlFile, htmlText
htmlFile = folder & "\template.html"

If fso.FileExists(htmlFile) Then
  Dim s: Set s = CreateObject("ADODB.Stream")
  s.Type = 2
  s.Charset = "utf-8"
  s.Open
  s.LoadFromFile htmlFile
  htmlText = s.ReadText
  s.Close
Else
  MsgBox "Missing template.html in the same folder.", vbExclamation
  WScript.Quit
End If

Dim subj, ok
subj = "Dynamic Link Injection (External HTML)"
ok = CreateMail("", "", "", subj, htmlText, Array(), "display", "", "", "", 0, "")

If Not ok Then
  MsgBox "CreateMail failed. Check mail_util.log.", vbCritical
Else
  MsgBox "Mail created successfully using external template.", vbInformation
End If
