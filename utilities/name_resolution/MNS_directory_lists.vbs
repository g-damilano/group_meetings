' Directory list discovery helpers for name resolution diagnostics

Function BuildDirectorySearchList(ns)
  Dim list: Set list = CreateObject("System.Collections.ArrayList")

  Call AddListNameIfMissing(list, "Global Address List")
  Call AddListNameIfMissing(list, "Offline Global Address List")

  Dim contactNames: contactNames = CollectLocalContactsListNames(ns)
  If IsArray(contactNames) Then
    Dim i
    For i = LBound(contactNames) To UBound(contactNames)
      Call AddListNameIfMissing(list, contactNames(i))
    Next
  End If

  If list.Count = 0 Then
    BuildDirectorySearchList = Array()
  Else
    BuildDirectorySearchList = list.ToArray()
  End If
End Function

Sub AddListNameIfMissing(list, listName)
  If list Is Nothing Then Exit Sub

  Dim trimmed: trimmed = ""
  If Not IsNull(listName) Then
    If Not IsEmpty(listName) Then trimmed = Trim(CStr(listName))
  End If
  If Len(trimmed) = 0 Then Exit Sub

  On Error Resume Next
  If list.IndexOf(trimmed) = -1 Then list.Add trimmed
  On Error GoTo 0
End Sub

Function CollectLocalContactsListNames(ns)
  Dim results: Set results = CreateObject("System.Collections.ArrayList")
  If ns Is Nothing Then
    CollectLocalContactsListNames = results.ToArray()
    Exit Function
  End If

  Dim defaultListName: defaultListName = GetDefaultContactsAddressListName(ns)
  Call AppendAvailableContactList(results, ns, defaultListName)
  Call AppendAvailableContactList(results, ns, "Contacts")
  Call AppendAvailableContactList(results, ns, "Contacts (This computer only)")

  If results.Count = 0 Then
    CollectLocalContactsListNames = Array()
  Else
    CollectLocalContactsListNames = results.ToArray()
  End If
End Function

Sub AppendAvailableContactList(target, ns, listName)
  If target Is Nothing Then Exit Sub
  If ns Is Nothing Then Exit Sub

  Dim trimmed: trimmed = ""
  If Not IsNull(listName) Then
    If Not IsEmpty(listName) Then trimmed = Trim(CStr(listName))
  End If
  If Len(trimmed) = 0 Then Exit Sub

  On Error Resume Next
  If target.IndexOf(trimmed) <> -1 Then
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0

  Dim addressList
  Set addressList = Nothing

  On Error Resume Next
  Set addressList = ns.AddressLists(trimmed)
  If Err.Number <> 0 Then
    Err.Clear
    Set addressList = Nothing
  End If
  On Error GoTo 0

  If addressList Is Nothing Then Exit Sub
  If Not ShouldTreatAsLocalContactsList(addressList) Then Exit Sub

  target.Add trimmed
End Sub

Function ShouldTreatAsLocalContactsList(addressList)
  ShouldTreatAsLocalContactsList = False
  If addressList Is Nothing Then Exit Function

  Dim listType
  listType = -1
  On Error Resume Next
  listType = CLng(addressList.AddressListType)
  If Err.Number <> 0 Then
    Err.Clear
    listType = -1
  End If
  On Error GoTo 0

  Const olOutlookAddressList = 4
  If listType = olOutlookAddressList Then
    ShouldTreatAsLocalContactsList = True
    Exit Function
  End If

  Dim name: name = ""
  On Error Resume Next
  name = addressList.Name
  If Err.Number <> 0 Then
    Err.Clear
    name = ""
  End If
  On Error GoTo 0
  If Len(name) = 0 Then Exit Function

  Dim lowered: lowered = SafeLower(name)
  If lowered = "contacts" Or lowered = "contacts (this computer only)" Then
    ShouldTreatAsLocalContactsList = True
  End If
End Function

Function GetDefaultContactsAddressListName(ns)
  GetDefaultContactsAddressListName = ""
  If ns Is Nothing Then Exit Function

  Dim contactsFolder
  Set contactsFolder = Nothing

  On Error Resume Next
  Set contactsFolder = ns.GetDefaultFolder(10)
  If Err.Number <> 0 Then
    Err.Clear
    Set contactsFolder = Nothing
  End If
  On Error GoTo 0

  If contactsFolder Is Nothing Then Exit Function

  Dim addressBookName: addressBookName = ""

  On Error Resume Next
  addressBookName = contactsFolder.AddressBookName
  If Err.Number <> 0 Then
    Err.Clear
    addressBookName = contactsFolder.Name
  End If
  On Error GoTo 0

  GetDefaultContactsAddressListName = Trim(CStr(addressBookName))
End Function
