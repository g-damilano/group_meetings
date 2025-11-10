Function IsValidEmail(email)
  Dim candidate: candidate = ""

  If IsObject(email) Then
    IsValidEmail = False
    Exit Function
  End If

  If Not IsNull(email) Then
    If Not IsEmpty(email) Then candidate = CStr(email)
  End If

  candidate = Trim(candidate)
  If Len(candidate) = 0 Then
    IsValidEmail = False
    Exit Function
  End If

  Dim regex: Set regex = Nothing
  On Error Resume Next
  Set regex = CreateObject("VBScript.RegExp")
  If Err.Number <> 0 Then
    Err.Clear
    Set regex = Nothing
  End If
  On Error GoTo 0

  If regex Is Nothing Then
    If InStr(candidate, "@") > 0 Then
      Dim atPos: atPos = InStr(candidate, "@")
      Dim dotPos: dotPos = InStrRev(candidate, ".")
      IsValidEmail = (dotPos > atPos And dotPos < Len(candidate))
    Else
      IsValidEmail = False
    End If
    Exit Function
  End If

  regex.Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,63}$"
  regex.IgnoreCase = True
  regex.Global = False

  IsValidEmail = regex.Test(candidate)
End Function
