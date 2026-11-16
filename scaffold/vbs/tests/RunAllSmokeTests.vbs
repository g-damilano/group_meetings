Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell
Set shell = CreateObject("WScript.Shell")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)

Function QuoteArg(value)
  QuoteArg = Chr(34) & Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function IsRunnableTest(fileName)
  IsRunnableTest = (LCase(fso.GetExtensionName(fileName)) = "vbs" And LCase(fileName) <> "runallsmoketests.vbs")
End Function

Sub SortNames(ByRef names, count)
  Dim i, j, temp
  For i = 0 To count - 2
    For j = i + 1 To count - 1
      If LCase(names(j)) < LCase(names(i)) Then
        temp = names(i)
        names(i) = names(j)
        names(j) = temp
      End If
    Next
  Next
End Sub

Function CollectTestNames()
  Dim folder
  Set folder = fso.GetFolder(scriptFolder)

  Dim names()
  ReDim names(0)
  Dim count
  count = 0

  Dim file
  For Each file In folder.Files
    If IsRunnableTest(file.Name) Then
      ReDim Preserve names(count)
      names(count) = file.Name
      count = count + 1
    End If
  Next

  If count > 1 Then SortNames names, count
  CollectTestNames = names
End Function

Function RunTest(fileName)
  Dim result
  Set result = CreateObject("Scripting.Dictionary")

  Dim command
  command = "cscript.exe //NoLogo " & QuoteArg(fso.BuildPath(scriptFolder, fileName))

  Dim exec
  Set exec = shell.Exec(command)

  Dim output
  output = ""
  Do While exec.Status = 0
    Do While Not exec.StdOut.AtEndOfStream
      output = output & exec.StdOut.ReadLine & vbCrLf
    Loop
    Do While Not exec.StdErr.AtEndOfStream
      output = output & exec.StdErr.ReadLine & vbCrLf
    Loop
    WScript.Sleep 50
  Loop

  Do While Not exec.StdOut.AtEndOfStream
    output = output & exec.StdOut.ReadLine & vbCrLf
  Loop
  Do While Not exec.StdErr.AtEndOfStream
    output = output & exec.StdErr.ReadLine & vbCrLf
  Loop

  result("exit_code") = exec.ExitCode
  result("output") = output
  Set RunTest = result
End Function

Dim testNames
testNames = CollectTestNames()

Dim total
total = UBound(testNames) + 1
If total = 0 Then
  WScript.Echo "FAILED: no VBS smoke tests found."
  WScript.Quit 1
End If

Dim failures
failures = 0

Dim i
For i = 0 To UBound(testNames)
  WScript.Echo "RUN: " & testNames(i)
  Dim runResult
  Set runResult = RunTest(testNames(i))
  If Len(runResult("output")) > 0 Then WScript.Echo runResult("output")
  If runResult("exit_code") <> 0 Then
    WScript.Echo "FAILED: " & testNames(i) & " exited " & runResult("exit_code")
    failures = failures + 1
  Else
    WScript.Echo "OK: " & testNames(i)
  End If
Next

If failures > 0 Then
  WScript.Echo "VBS smoke suite failed: " & failures & " of " & total & " test(s)."
  WScript.Quit 1
End If

WScript.Echo "VBS smoke suite passed: " & total & " test(s)."
