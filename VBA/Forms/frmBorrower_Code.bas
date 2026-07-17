' frmBorrower - Borrower Management Form
' Purpose: Add, Edit, Search, and Delete borrower records
' Date: 2026-07-17

Option Explicit

Private currentBorrowerID As String

Private Sub UserForm_Initialize()
    Me.StartupPosition = 1 ' Center Owner
    
    ' Clear fields
    ClearForm
    
    ' Load borrowers into listbox
    LoadBorrowers
    
    ' Set focus
    txtLastName.SetFocus
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo ErrorHandler
    
    ' Validate fields
    If Len(txtLastName.Value) = 0 Or Len(txtFirstName.Value) = 0 Then
        MsgBox "Last Name and First Name are required.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Add borrower
    If AddBorrower(txtLastName.Value, txtFirstName.Value, txtMiddleName.Value, _
                   txtAddress.Value, txtContactNo.Value, txtOccupation.Value, _
                   txtPhotoPath.Value, cbxStatus.Value) Then
        MsgBox "Borrower added successfully!", vbInformation, "Success"
        ClearForm
        LoadBorrowers
        txtLastName.SetFocus
    Else
        MsgBox "Error adding borrower.", vbCritical, "Error"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo ErrorHandler
    
    If Len(currentBorrowerID) = 0 Then
        MsgBox "Please select a borrower to update.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    If UpdateBorrower(currentBorrowerID, txtLastName.Value, txtFirstName.Value, _
                      txtMiddleName.Value, txtAddress.Value, txtContactNo.Value, _
                      txtOccupation.Value, txtPhotoPath.Value, cbxStatus.Value) Then
        MsgBox "Borrower updated successfully!", vbInformation, "Success"
        ClearForm
        LoadBorrowers
    Else
        MsgBox "Error updating borrower.", vbCritical, "Error"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrorHandler
    
    If Len(currentBorrowerID) = 0 Then
        MsgBox "Please select a borrower to delete.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this borrower?", vbYesNo, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    If DeleteBorrower(currentBorrowerID) Then
        MsgBox "Borrower deleted successfully!", vbInformation, "Success"
        ClearForm
        LoadBorrowers
    Else
        MsgBox "Error deleting borrower.", vbCritical, "Error"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim searchTerm As String
    searchTerm = txtSearchName.Value
    
    If Len(searchTerm) = 0 Then
        LoadBorrowers
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    lstBorrowers.Clear
    
    For Each lr In tbl.ListRows
        Dim lastName As String
        Dim firstName As String
        lastName = lr.Range.Cells(1, 2).Value
        firstName = lr.Range.Cells(1, 3).Value
        
        If InStr(1, lastName & " " & firstName, searchTerm, vbTextCompare) > 0 Then
            lstBorrowers.AddItem lr.Range.Cells(1, 1).Value & " - " & lastName & ", " & firstName
        End If
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub lstBorrowers_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim borrowerID As String
    
    selectedText = lstBorrowers.Value
    borrowerID = Left(selectedText, InStr(selectedText, " ") - 1)
    
    currentBorrowerID = borrowerID
    
    ' Get and display borrower data
    Dim borrowerData As Variant
    borrowerData = GetBorrower(borrowerID)
    
    If Not IsNothing(borrowerData) Then
        txtBorrowerID.Value = borrowerData(1)
        txtLastName.Value = borrowerData(2)
        txtFirstName.Value = borrowerData(3)
        txtMiddleName.Value = borrowerData(4)
        txtAddress.Value = borrowerData(5)
        txtContactNo.Value = borrowerData(6)
        txtOccupation.Value = borrowerData(7)
        txtPhotoPath.Value = borrowerData(8)
        cbxStatus.Value = borrowerData(9)
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ClearForm()
    txtBorrowerID.Value = ""
    txtLastName.Value = ""
    txtFirstName.Value = ""
    txtMiddleName.Value = ""
    txtAddress.Value = ""
    txtContactNo.Value = ""
    txtOccupation.Value = ""
    txtPhotoPath.Value = ""
    cbxStatus.Value = "Active"
    txtSearchName.Value = ""
    currentBorrowerID = ""
End Sub

Private Sub LoadBorrowers()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    lstBorrowers.Clear
    
    For Each lr In tbl.ListRows
        lstBorrowers.AddItem lr.Range.Cells(1, 1).Value & " - " & _
                             lr.Range.Cells(1, 2).Value & ", " & _
                             lr.Range.Cells(1, 3).Value
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading borrowers: " & Err.Description, vbCritical, "Error"
End Sub

Private Function IsNothing(obj As Variant) As Boolean
    IsNothing = (TypeName(obj) = "Nothing")
End Function