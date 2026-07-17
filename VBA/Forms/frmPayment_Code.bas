' frmPayment - Payment Recording Form
' Purpose: Record loan payments and track payment history
' Date: 2026-07-17

Option Explicit

Private Sub UserForm_Initialize()
    Me.StartupPosition = 1 ' Center Owner
    
    ' Load loans into combo
    LoadLoanCombo
    
    ' Load collectors into combo
    LoadCollectorCombo
    
    ' Set payment date to today
    txtPaymentDate.Value = Format(Now(), "mm/dd/yyyy")
    
    ' Load payment history
    LoadPaymentHistory
    
    ' Clear form
    ClearForm
End Sub

Private Sub cmdSearchLoan_Click()
    On Error GoTo ErrorHandler
    
    Dim loanID As String
    loanID = txtLoanID.Value
    
    If Len(loanID) = 0 Then
        MsgBox "Please enter a Loan ID.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Get loan data
    Dim loanData As Variant
    loanData = GetLoan(loanID)
    
    If Not IsNothing(loanData) Then
        cbxLoanID.Value = loanID
        lblBorrowerID.Caption = "Borrower: " & loanData(2)
        lblBalance.Caption = "Current Balance: PHP " & Format(loanData(13), "#,##0.00")
        lblDailyPayment.Caption = "Daily Payment: PHP " & Format(loanData(11), "#,##0.00")
    Else
        MsgBox "Loan not found.", vbExclamation, "Not Found"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cbxLoanID_Change()
    On Error GoTo ErrorHandler
    
    If cbxLoanID.ListIndex = -1 Then Exit Sub
    
    Dim loanID As String
    loanID = cbxLoanID.Value
    
    ' Get loan data
    Dim loanData As Variant
    loanData = GetLoan(loanID)
    
    If Not IsNothing(loanData) Then
        lblBorrowerID.Caption = "Borrower: " & loanData(2)
        lblBalance.Caption = "Current Balance: PHP " & Format(loanData(13), "#,##0.00")
        lblDailyPayment.Caption = "Daily Payment: PHP " & Format(loanData(11), "#,##0.00")
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdRecordPayment_Click()
    On Error GoTo ErrorHandler
    
    ' Validate
    If cbxLoanID.ListIndex = -1 Then
        MsgBox "Please select a loan.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If Val(txtPaymentAmount.Value) <= 0 Then
        MsgBox "Payment amount must be greater than 0.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If Len(txtORNumber.Value) = 0 Then
        MsgBox "OR Number is required.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If cbxCollectorID.ListIndex = -1 Then
        MsgBox "Please select a collector.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Record payment
    Dim paymentID As String
    paymentID = RecordPayment(cbxLoanID.Value, CDbl(txtPaymentAmount.Value), _
                             CDate(txtPaymentDate.Value), txtORNumber.Value, _
                             cbxCollectorID.Value)
    
    If Len(paymentID) > 0 Then
        MsgBox "Payment recorded successfully!" & vbCrLf & "Payment ID: " & paymentID, vbInformation, "Success"
        ClearForm
        LoadPaymentHistory
        ' Refresh loan info
        cbxLoanID_Change
    Else
        MsgBox "Error recording payment.", vbCritical, "Error"
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
    cbxLoanID.Value = ""
    lblBorrowerID.Caption = "Borrower: --"
    lblBalance.Caption = "Current Balance: PHP 0.00"
    lblDailyPayment.Caption = "Daily Payment: PHP 0.00"
    txtPaymentAmount.Value = ""
    txtPaymentDate.Value = Format(Now(), "mm/dd/yyyy")
    txtORNumber.Value = ""
    cbxCollectorID.Value = ""
End Sub

Private Sub LoadPaymentHistory()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Payments")
    Set tbl = ws.ListObjects("tblPayments")
    
    lstPaymentHistory.Clear
    
    For Each lr In tbl.ListRows
        lstPaymentHistory.AddItem lr.Range.Cells(1, 1).Value & " - " & _
                                   "Loan: " & lr.Range.Cells(1, 2).Value & " - " & _
                                   "PHP " & Format(lr.Range.Cells(1, 4).Value, "#,##0.00")
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading payment history: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub LoadLoanCombo()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    cbxLoanID.Clear
    
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 15).Value = "Active" Then
            cbxLoanID.AddItem lr.Range.Cells(1, 1).Value
        End If
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