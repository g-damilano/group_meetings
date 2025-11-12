' Contact-centric helpers for name resolution diagnostics

Function BuildMatchTokenList(tokens)
  Set BuildMatchTokenList = Nothing
  If tokens Is Nothing Then Exit Function

  Dim result: Set result = CreateArrayList()
  Dim i
  For i = 0 To tokens.Count - 1
    Dim lowered: lowered = SafeLower(tokens(i))
    If Len(lowered) > 0 Then result.Add lowered
  Next

  If result.Count = 0 Then Exit Function

  Set BuildMatchTokenList = result
End Function

Sub OutputComponentSummary(components, tokens)
  WScript.Echo ""
  WScript.Echo "Component analysis of target value:"

  If components Is Nothing Then
    WScript.Echo "  (No components parsed.)"
  Else
    WScript.Echo "  Display: " & SafeText(components("Display"))
    WScript.Echo "  Normalized: " & SafeText(components("Normalized"))
    WScript.Echo "  Reverse display: " & SafeText(components("ReverseDisplay"))
    WScript.Echo "  First name: " & SafeText(components("First"))
    WScript.Echo "  Surname: " & SafeText(components("Last"))
    WScript.Echo "  Alternate first name: " & SafeText(components("AlternateFirst"))
    WScript.Echo "  Alternate surname: " & SafeText(components("AlternateLast"))
  End If

  Dim tokenSummary
  tokenSummary = ""
  If Not tokens Is Nothing Then
    Dim i
    For i = 0 To tokens.Count - 1
      If Len(tokenSummary) = 0 Then
        tokenSummary = tokens(i)
      Else
        tokenSummary = tokenSummary & ", " & tokens(i)
      End If
    Next
  End If
  If Len(tokenSummary) = 0 Then
    WScript.Echo "  Extracted tokens: (none)"
  Else
    WScript.Echo "  Extracted tokens: " & tokenSummary
  End If
End Sub

Sub OutputContactSummary(contact, indent)
  If Len(indent) = 0 Then indent = ""
  If contact Is Nothing Then
    WScript.Echo indent & "(Contact reference is Nothing)"
    Exit Sub
  End If

  If Not IsContactLikeItem(contact) Then
    WScript.Echo indent & "(Item is not a contact: " & DescribeItemClass(contact) & ")"
    Exit Sub
  End If

  WScript.Echo indent & "Contact: " & SafeText(ContactFieldValue(contact, "FullName"))
  WScript.Echo indent & "  FileAs: " & SafeText(ContactFieldValue(contact, "FileAs"))
  WScript.Echo indent & "  First name: " & SafeText(ContactFieldValue(contact, "FirstName"))
  WScript.Echo indent & "  Surname: " & SafeText(ContactFieldValue(contact, "LastName"))
  WScript.Echo indent & "  Primary email: " & SafeText(ContactFieldValue(contact, "Email1Address"))
  WScript.Echo indent & "  Secondary email: " & SafeText(ContactFieldValue(contact, "Email2Address"))
  WScript.Echo indent & "  Tertiary email: " & SafeText(ContactFieldValue(contact, "Email3Address"))
  WScript.Echo indent & "  Company: " & SafeText(ContactFieldValue(contact, "CompanyName"))
  WScript.Echo indent & "  Department: " & SafeText(ContactFieldValue(contact, "Department"))
End Sub

Function ContactItemMatchesTokens(contact, tokens)
  ContactItemMatchesTokens = False
  If contact Is Nothing Then Exit Function
  If tokens Is Nothing Then Exit Function
  If tokens.Count = 0 Then Exit Function
  If Not IsContactLikeItem(contact) Then Exit Function

  Dim fields: fields = Array( _
    SafeLower(ContactFieldValue(contact, "FullName")), _
    SafeLower(ContactFieldValue(contact, "FileAs")), _
    SafeLower(ContactFieldValue(contact, "FirstName")), _
    SafeLower(ContactFieldValue(contact, "LastName")), _
    SafeLower(ContactFieldValue(contact, "Email1Address")), _
    SafeLower(ContactFieldValue(contact, "Email2Address")), _
    SafeLower(ContactFieldValue(contact, "Email3Address")), _
    SafeLower(ContactFieldValue(contact, "CompanyName")), _
    SafeLower(ContactFieldValue(contact, "Department")) _
  )

  Dim i
  For i = LBound(fields) To UBound(fields)
    If ContainsAllTokens(fields(i), tokens) Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
  Next

  Dim firstLower: firstLower = SafeLower(ContactFieldValue(contact, "FirstName"))
  Dim lastLower: lastLower = SafeLower(ContactFieldValue(contact, "LastName"))

  If tokens.Count = 2 Then
    Dim t0: t0 = SafeLower(tokens(0))
    Dim t1: t1 = SafeLower(tokens(1))
    If InStr(firstLower, t0) > 0 And InStr(lastLower, t1) > 0 Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
    If InStr(firstLower, t1) > 0 And InStr(lastLower, t0) > 0 Then
      ContactItemMatchesTokens = True
      Exit Function
    End If
  End If
End Function

Function ContactFieldValue(contact, fieldName)
  Dim value: value = ""
  If contact Is Nothing Then
    ContactFieldValue = ""
    Exit Function
  End If

  On Error Resume Next
  Select Case LCase(fieldName)
    Case "fullname"
      value = contact.FullName
    Case "fileas"
      value = contact.FileAs
    Case "firstname"
      value = contact.FirstName
    Case "lastname"
      value = contact.LastName
    Case "email1address"
      value = contact.Email1Address
    Case "email2address"
      value = contact.Email2Address
    Case "email3address"
      value = contact.Email3Address
    Case "companyname"
      value = contact.CompanyName
    Case "department"
      value = contact.Department
    Case Else
      value = ""
  End Select
  If Err.Number <> 0 Then
    Err.Clear
    value = ""
  End If
  On Error GoTo 0

  ContactFieldValue = SafeText(value)
End Function

Function IsContactLikeItem(item)
  IsContactLikeItem = False
  If item Is Nothing Then Exit Function

  Dim classValue: classValue = -1
  On Error Resume Next
  classValue = item.Class
  If Err.Number <> 0 Then
    Err.Clear
    classValue = -1
  End If

  Dim messageClass: messageClass = ""
  messageClass = item.MessageClass
  If Err.Number <> 0 Then
    Err.Clear
    messageClass = ""
  End If
  On Error GoTo 0

  If classValue = 40 Or classValue = 69 Then
    IsContactLikeItem = True
    Exit Function
  End If

  Dim lowered: lowered = LCase(messageClass)
  If Len(lowered) = 0 Then Exit Function
  If Left(lowered, 11) = "ipm.contact" Then
    IsContactLikeItem = True
  ElseIf lowered = "ipm.distlist" Then
    IsContactLikeItem = True
  End If
End Function

Function DescribeItemClass(item)
  DescribeItemClass = "unknown"
  If item Is Nothing Then Exit Function

  Dim classValue: classValue = "?"
  Dim messageClass: messageClass = ""

  On Error Resume Next
  classValue = item.Class
  If Err.Number <> 0 Then
    Err.Clear
    classValue = "?"
  End If
  messageClass = item.MessageClass
  If Err.Number <> 0 Then
    Err.Clear
    messageClass = ""
  End If
  On Error GoTo 0

  If Len(messageClass) > 0 Then
    DescribeItemClass = "Class=" & classValue & ", MessageClass='" & messageClass & "'"
  Else
    DescribeItemClass = "Class=" & classValue
  End If
End Function

Function ContainsAllTokens(text, tokens)
  ContainsAllTokens = False
  If tokens Is Nothing Then Exit Function
  If tokens.Count = 0 Then Exit Function

  Dim haystack: haystack = SafeLower(text)
  If Len(haystack) = 0 Then Exit Function

  Dim i
  For i = 0 To tokens.Count - 1
    Dim needle: needle = SafeLower(tokens(i))
    If Len(needle) = 0 Then
      ContainsAllTokens = False
      Exit Function
    End If
    If InStr(haystack, needle) = 0 Then
      ContainsAllTokens = False
      Exit Function
    End If
  Next

  ContainsAllTokens = True
End Function

