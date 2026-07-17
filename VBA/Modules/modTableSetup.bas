Option Explicit

' modTableSetup - Automated Excel Tables Creation
' Purpose: Create all database tables with proper structure and formatting
' Author: Lending Management System
' Date: 2026-07-17

Public Sub CreateAllTables()
    ' Main procedure to create all tables
    On Error GoTo ErrorHandler
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' Create each table
    CreateBorrowersTable
    CreateLoansTable
    CreatePaymentsTable
    CreateCollectorsTable
    CreateUsersTable
    
    Application.ScreenUpdating = True
    MsgBox "All tables created successfully!", vbInformation, "Success"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error creating tables: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' Create tblBorrowers Table
' ============================================================
Private Sub CreateBorrowersTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Create worksheet
    Set ws = CreateWorksheet("Borrowers")
    
    ' Add headers
    With ws.Range("A1")
        .Value = "BorrowerID"
        .Offset(0, 1).Value = "LastName"
        .Offset(0, 2).Value = "FirstName"
        .Offset(0, 3).Value = "MiddleName"
        .Offset(0, 4).Value = "Address"
        .Offset(0, 5).Value = "ContactNo"
        .Offset(0, 6).Value = "Occupation"
        .Offset(0, 7).Value = "PhotoPath"
        .Offset(0, 8).Value = "Status"
    End With
    
    ' Format as table
    Set rng = ws.Range("A1:I1")
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblBorrowers"
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Set column widths
    ws.Columns("A:I").AutoFit
    
    ' Add data validation for Status column
    AddDataValidation ws.Range("I2:I1000"), "Active,Inactive,Deceased,Blacklisted"
    
End Sub

' ============================================================
' Create tblLoans Table
' ============================================================
Private Sub CreateLoansTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Create worksheet
    Set ws = CreateWorksheet("Loans")
    
    ' Add headers
    With ws.Range("A1")
        .Value = "LoanID"
        .Offset(0, 1).Value = "BorrowerID"
        .Offset(0, 2).Value = "LoanType"
        .Offset(0, 3).Value = "Principal"
        .Offset(0, 4).Value = "InterestRate"
        .Offset(0, 5).Value = "InterestAmount"
        .Offset(0, 6).Value = "TotalCollectible"
        .Offset(0, 7).Value = "ReleaseDate"
        .Offset(0, 8).Value = "DueDate"
        .Offset(0, 9).Value = "Terms"
        .Offset(0, 10).Value = "DailyPayment"
        .Offset(0, 11).Value = "AmountPaid"
        .Offset(0, 12).Value = "Balance"
        .Offset(0, 13).Value = "CollectorID"
        .Offset(0, 14).Value = "Status"
    End With
    
    ' Format as table
    Set rng = ws.Range("A1:O1")
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblLoans"
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Set column formats
    ws.Columns("D:F").NumberFormat = "$#,##0.00"   ' Currency columns
    ws.Columns("G").NumberFormat = "$#,##0.00"      ' Total Collectible
    ws.Columns("H:I").NumberFormat = "mm/dd/yyyy"   ' Date columns
    ws.Columns("K:M").NumberFormat = "$#,##0.00"   ' Currency columns
    
    ' Add data validation for Status column
    AddDataValidation ws.Range("O2:O1000"), "Active,Completed,Defaulted,Cancelled,Closed"
    
    ' Set column widths
    ws.Columns("A:O").AutoFit
    
End Sub

' ============================================================
' Create tblPayments Table
' ============================================================
Private Sub CreatePaymentsTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Create worksheet
    Set ws = CreateWorksheet("Payments")
    
    ' Add headers
    With ws.Range("A1")
        .Value = "PaymentID"
        .Offset(0, 1).Value = "LoanID"
        .Offset(0, 2).Value = "PaymentDate"
        .Offset(0, 3).Value = "AmountPaid"
        .Offset(0, 4).Value = "RemainingBalance"
        .Offset(0, 5).Value = "ORNumber"
        .Offset(0, 6).Value = "CollectorID"
    End With
    
    ' Format as table
    Set rng = ws.Range("A1:G1")
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblPayments"
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Set column formats
    ws.Columns("C").NumberFormat = "mm/dd/yyyy"    ' PaymentDate
    ws.Columns("D:E").NumberFormat = "$#,##0.00"   ' Currency columns
    
    ' Set column widths
    ws.Columns("A:G").AutoFit
    
End Sub

' ============================================================
' Create tblCollectors Table
' ============================================================
Private Sub CreateCollectorsTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Create worksheet
    Set ws = CreateWorksheet("Collectors")
    
    ' Add headers
    With ws.Range("A1")
        .Value = "CollectorID"
        .Offset(0, 1).Value = "Name"
        .Offset(0, 2).Value = "ContactNo"
        .Offset(0, 3).Value = "Commission"
    End With
    
    ' Format as table
    Set rng = ws.Range("A1:D1")
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblCollectors"
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Set column formats
    ws.Columns("D").NumberFormat = "$#,##0.00"     ' Commission
    
    ' Set column widths
    ws.Columns("A:D").AutoFit
    
End Sub

' ============================================================
' Create tblUsers Table
' ============================================================
Private Sub CreateUsersTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Create worksheet
    Set ws = CreateWorksheet("Users")
    
    ' Add headers
    With ws.Range("A1")
        .Value = "Username"
        .Offset(0, 1).Value = "Password"
        .Offset(0, 2).Value = "Role"
    End With
    
    ' Format as table
    Set rng = ws.Range("A1:C1")
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblUsers"
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Add data validation for Role column
    AddDataValidation ws.Range("C2:C1000"), "Admin,Collector,Manager,Viewer"
    
    ' Set column widths
    ws.Columns("A:C").AutoFit
    
End Sub

' ============================================================
' Helper: Create or Return Worksheet
' ============================================================
Private Function CreateWorksheet(wsName As String) As Worksheet
    Dim ws As Worksheet
    
    ' Check if worksheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new worksheet
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = wsName
    Else
        ' Clear existing worksheet
        ws.Cells.Delete
    End If
    
    Set CreateWorksheet = ws
End Function

' ============================================================
' Helper: Add Data Validation
' ============================================================
Private Sub AddDataValidation(rng As Range, validValues As String)
    ' validValues: comma-separated string like "Active,Inactive,Deceased,Blacklisted"
    
    With rng.Validation
        .Delete
        .Add Type:=3, Formula1:=validValues
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid Entry"
        .ErrorMessage = "Please select a valid option from the list."
        .ShowError = True
    End With
End Sub

' ============================================================
' Delete All Tables (Cleanup)
' ============================================================
Public Sub DeleteAllTables()
    Dim ws As Worksheet
    Dim response As Integer
    
    response = MsgBox("Are you sure you want to delete all tables? This cannot be undone.", vbYesNo, "Confirm Delete")
    
    If response = vbNo Then Exit Sub
    
    On Error Resume Next
    
    ' Delete worksheets
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "Borrowers" Or ws.Name Like "Loans" Or ws.Name Like "Payments" Or _
           ws.Name Like "Collectors" Or ws.Name Like "Users" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    On Error GoTo 0
    
    MsgBox "All tables deleted successfully!", vbInformation, "Success"
End Sub
