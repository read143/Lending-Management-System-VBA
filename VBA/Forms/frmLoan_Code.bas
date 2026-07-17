' frmLoan - Loan Management Form
' Purpose: Create, View, and Manage loan records
' Date: 2026-07-17

Option Explicit

Private currentLoanID As String

Private Sub UserForm_Initialize()
    Me.StartupPosition = 1 ' Center Owner
    
    ' Populate dropdowns
    LoadBorrowerCombo
    LoadCollectorCombo
    
    cbxLoanType.AddItem "Personal Loan"
    cbxLoanType.AddItem "Business Loan"
    cbxLoanType.AddItem "Emergency Loan"
    cbxLoanType.AddItem "Education Loan"
    cbxLoanType.Value = "Personal Loan"
    
    cbxStatus.AddItem "Active"
    cbxStatus.AddItem "Completed"
    cbxStatus.AddItem "Defaulted"
    cbxStatus.Value = "Active"
    
    ' Set default dates
    txtReleaseDate.Value = Format(Now(), "mm/dd/yyyy")
    txtDueDate.Value = Format(DateAdd("m", 12, Now()), "mm/dd/yyyy")
    
    ' Clear form
    ClearForm
    
    ' Load loans
    LoadLoans
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo ErrorHandler
    
    ' Validate
    If cbxBorrowerID.ListIndex = -1 Then
        MsgBox "Please select a borrower.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If Val(txtPrincipal.Value) <= 0 Then
        MsgBox "Principal must be greater than 0.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If Val(txtInterestRate.Value) < 0 Then
        MsgBox "Interest rate cannot be negative.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Add loan
    Dim loanID As String
    loanID = AddLoan(cbxBorrowerID.Value, cbxLoanType.Value, CDbl(txtPrincipal.Value), _
                     CDbl(txtInterestRate.Value), CDate(txtReleaseDate.Value), _
                     CDate(txtDueDate.Value), CInt(txtTerms.Value), _
                     CDbl(txtDailyPayment.Value), cbxCollectorID.Value, cbxStatus.Value)
    
    If Len(loanID) > 0 Then
        MsgBox "Loan created successfully!" & vbCrLf & "Loan ID: " & loanID, vbInformation, "Success"
        ClearForm
        LoadLoans
    Else
        MsgBox "Error creating loan.", vbCritical, "Error"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim searchTerm As String
    searchTerm = txtSearchLoanID.Value
    
    If Len(searchTerm) = 0 Then
        LoadLoans
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    lstLoans.Clear
    
    For Each lr In tbl.ListRows
        If InStr(1, lr.Range.Cells(1, 1).Value, searchTerm, vbTextCompare) > 0 Then
            lstLoans.AddItem lr.Range.Cells(1, 1).Value & " - " & lr.Range.Cells(1, 2).Value
        End If
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub lstLoans_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim loanID As String
    
    selectedText = lstLoans.Value
    loanID = Left(selectedText, InStr(selectedText, " ") - 1)
    
    currentLoanID = loanID
    
    ' Get and display loan data
    Dim loanData As Variant
    loanData = GetLoan(loanID)
    
    If Not IsNothing(loanData) Then
        txtLoanID.Value = loanData(1)
        cbxBorrowerID.Value = loanData(2)
        cbxLoanType.Value = loanData(3)
        txtPrincipal.Value = loanData(4)
        txtInterestRate.Value = loanData(5)
        lblInterestAmount.Caption = "Interest: PHP " & Format(loanData(6), "#,##0.00")
        lblTotalCollectible.Caption = "Total: PHP " & Format(loanData(7), "#,##0.00")
        txtReleaseDate.Value = Format(loanData(8), "mm/dd/yyyy")
        txtDueDate.Value = Format(loanData(9), "mm/dd/yyyy")
        txtTerms.Value = loanData(10)
        txtDailyPayment.Value = loanData(11)
        lblAmountPaid.Caption = "Paid: PHP " & Format(loanData(12), "#,##0.00")
        lblBalance.Caption = "Balance: PHP " & Format(loanData(13), "#,##0.00")
        cbxCollectorID.Value = loanData(14)
        cbxStatus.Value = loanData(15)
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ClearForm()
    txtLoanID.Value = ""
    cbxBorrowerID.Value = ""
    cbxLoanType.Value = "Personal Loan"
    txtPrincipal.Value = ""
    txtInterestRate.Value = ""
    lblInterestAmount.Caption = "Interest: PHP 0.00"
    lblTotalCollectible.Caption = "Total: PHP 0.00"
    txtReleaseDate.Value = Format(Now(), "mm/dd/yyyy")
    txtDueDate.Value = Format(DateAdd("m", 12, Now()), "mm/dd/yyyy")
    txtTerms.Value = 12
    txtDailyPayment.Value = ""
    lblAmountPaid.Caption = "Paid: PHP 0.00"
    lblBalance.Caption = "Balance: PHP 0.00"
    cbxCollectorID.Value = ""
    cbxStatus.Value = "Active"
    txtSearchLoanID.Value = ""
    currentLoanID = ""
End Sub

Private Sub LoadLoans()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    lstLoans.Clear
    
    For Each lr In tbl.ListRows
        lstLoans.AddItem lr.Range.Cells(1, 1).Value & " - " & lr.Range.Cells(1, 2).Value
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading loans: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub LoadBorrowerCombo()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    cbxBorrowerID.Clear
    
    For Each lr In tbl.ListRows
        cbxBorrowerID.AddItem lr.Range.Cells(1, 1).Value
    Next lr
End Sub

Private Sub LoadCollectorCombo()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Collectors")
    Set tbl = ws.ListObjects("tblCollectors")
    
    cbxCollectorID.Clear
    
    For Each lr In tbl.ListRows
        cbxCollectorID.AddItem lr.Range.Cells(1, 1).Value
    Next lr
End Sub

Private Function IsNothing(obj As Variant) As Boolean
    IsNothing = (TypeName(obj) = "Nothing")
End Function