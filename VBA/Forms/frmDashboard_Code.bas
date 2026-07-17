' frmDashboard - Main Navigation Form
' Purpose: Main menu and dashboard for the system
' Date: 2026-07-17

Option Explicit

Private Sub UserForm_Initialize()
    ' Center form on screen
    Me.StartupPosition = 1 ' Center Owner
    
    ' Display logged-in user info
    lblUserInfo.Caption = "Welcome, " & ThisWorkbook.Sheets(1).Range("LoggedInUser").Value & _
                          " (" & ThisWorkbook.Sheets(1).Range("UserRole").Value & ")"
    
    ' Load statistics
    LoadStatistics
End Sub

Private Sub cmdBorrowers_Click()
    frmBorrower.Show
End Sub

Private Sub cmdLoans_Click()
    frmLoan.Show
End Sub

Private Sub cmdPayments_Click()
    frmPayment.Show
End Sub

Private Sub cmdCollectors_Click()
    frmCollector.Show
End Sub

Private Sub cmdLogout_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub LoadStatistics()
    On Error Resume Next
    
    Dim wsLoans As Worksheet
    Dim tblLoans As ListObject
    Dim activeLoans As Long
    Dim totalCollectible As Currency
    Dim totalPaid As Currency
    
    Set wsLoans = ThisWorkbook.Sheets("Loans")
    Set tblLoans = wsLoans.ListObjects("tblLoans")
    
    activeLoans = 0
    totalCollectible = 0
    totalPaid = 0
    
    ' Calculate statistics
    Dim lr As ListRow
    For Each lr In tblLoans.ListRows
        If lr.Range.Cells(1, 15).Value = "Active" Then
            activeLoans = activeLoans + 1
        End If
        totalCollectible = totalCollectible + lr.Range.Cells(1, 7).Value
        totalPaid = totalPaid + lr.Range.Cells(1, 12).Value
    Next lr
    
    ' Display statistics
    lblActiveLoans.Caption = activeLoans & " Active Loans"
    lblTotalCollectible.Caption = "PHP " & Format(totalCollectible, "#,##0.00")
    lblTotalPaid.Caption = "PHP " & Format(totalPaid, "#,##0.00")
    
End Sub
