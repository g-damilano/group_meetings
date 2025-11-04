' test_send_external_html.vbs — uses an external HTML file as the body
Option Explicit
SetLocale 1033

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim folder: folder = fso.GetParentFolderName(WScript.ScriptFullName) : If Len(folder)=0 Then folder = "."

' Load mailer (which loads the shield core)
ExecuteGlobal fso.OpenTextFile(folder & "\mail_utils_mailer.vbs", 1).ReadAll

Dim toAddr, subjectText, templatePath, inlineImgFolder, action, logPath, ok
toAddr          = "gd2613@ic.ac.uk"
subjectText     = "UTF‑8 Plus – External HTML Template Test"
templatePath    = folder & "\sample_template_utf8.html"
inlineImgFolder = folder & "\assets"    ' will embed images as cid:filename
action          = "display"             ' change to "send" after review
logPath         = ""

If Not fso.FileExists(templatePath) Then
  WScript.Echo "Template not found: " & templatePath
  WScript.Quit 1
End If

If Not fso.FolderExists(inlineImgFolder) Then
  fso.CreateFolder inlineImgFolder
End If

' Create a basic logo if missing (optional)
If Not fso.FileExists(inlineImgFolder & "\logo.png") Then
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  Dim dom: Set dom = CreateObject("MSXML2.DOMDocument")
  Dim node: Set node = dom.createElement("b64")
  node.dataType = "bin.base64"
  node.Text = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII="
  Dim bytes: bytes = node.nodeTypedValue
  stm.Type = 1
  stm.Open
  stm.Write bytes
  stm.SaveToFile inlineImgFolder & "\logo.png", 2
  stm.Close
End If

ok = CreateMail( _
      toAddr, "", "", _
      subjectText, _
      "", _
      Array(), _
      action, _
      "", _
      templatePath, _
      inlineImgFolder, _
      0, _
      logPath _
     )

If Not ok Then WScript.Echo "CreateMail failed. See mail_util.log for details."
