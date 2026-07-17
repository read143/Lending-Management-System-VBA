Option Explicit

' modDatabaseFunctions - CRUD Operations for Lending Management System
' Purpose: Provide reusable functions to Create, Read, Update, Delete records
' Author: Lending Management System
' Date: 2026-07-17

' ============================================================
' BORROWER FUNCTIONS
' ============================================================

' Add a new borrower record
Public Function AddBorrower(lastName As String, firstName As String, middleName As String, _
                           address As String, contactNo As String, occupation As String, _
                           photoPath As String, Optional status As String = "Active") As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim borrowerID As String
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    ' Validate required fields
    If Len(lastName) = 0 Or Len(firstName) = 0 Then
        MsgBox "Last Name and First Name are required.", vbExclamation, "Validation Error"
        Exit Function
    End If
    
    ' Generate BorrowerID
    borrowerID = GenerateID("tblBorrowers", "BorrowerID", "BOR")
    
    ' Add new row to table
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = borrowerID
        .Cells(1, 2).Value = lastName
        .Cells(1, 3).Value = firstName
        .Cells(1, 4).Value = middleName
        .Cells(1, 5).Value = address
        .Cells(1, 6).Value = contactNo
        .Cells(1, 7).Value = occupation
        .Cells(1, 8).Value = photoPath
        .Cells(1, 9).Value = status
    End With
    
    AddBorrower = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding borrower: " & Err.Description, vbCritical, "Error"
    AddBorrower = False
End Function

' Get borrower information by BorrowerID
Public Function GetBorrower(borrowerID As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim borrowerData As Variant
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    ReDim borrowerData(1 To 9)
    found = False
    
    ' Search for borrower
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = borrowerID Then
            borrowerData(1) = lr.Range.Cells(1, 1).Value ' BorrowerID
            borrowerData(2) = lr.Range.Cells(1, 2).Value ' LastName
            borrowerData(3) = lr.Range.Cells(1, 3).Value ' FirstName
            borrowerData(4) = lr.Range.Cells(1, 4).Value ' MiddleName
            borrowerData(5) = lr.Range.Cells(1, 5).Value ' Address
            borrowerData(6) = lr.Range.Cells(1, 6).Value ' ContactNo
            borrowerData(7) = lr.Range.Cells(1, 7).Value ' Occupation
            borrowerData(8) = lr.Range.Cells(1, 8).Value ' PhotoPath
            borrowerData(9) = lr.Range.Cells(1, 9).Value ' Status
            found = True
            Exit For
        End If
    Next lr
    
    If found Then
        GetBorrower = borrowerData
    Else
        GetBorrower = Nothing
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error retrieving borrower: " & Err.Description, vbCritical, "Error"
    GetBorrower = Nothing
End Function

' Update borrower information
Public Function UpdateBorrower(borrowerID As String, lastName As String, firstName As String, _
                              middleName As String, address As String, contactNo As String, _
                              occupation As String, photoPath As String, status As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    ' Validate required fields
    If Len(lastName) = 0 Or Len(firstName) = 0 Then
        MsgBox "Last Name and First Name are required.", vbExclamation, "Validation Error"
        Exit Function
    End If
    
    found = False
    
    ' Find and update borrower
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = borrowerID Then
            With lr.Range
                .Cells(1, 2).Value = lastName
                .Cells(1, 3).Value = firstName
                .Cells(1, 4).Value = middleName
                .Cells(1, 5).Value = address
                .Cells(1, 6).Value = contactNo
                .Cells(1, 7).Value = occupation
                .Cells(1, 8).Value = photoPath
                .Cells(1, 9).Value = status
            End With
            found = True
            Exit For
        End If
    Next lr
    
    If Not found Then
        MsgBox "Borrower not found.", vbExclamation, "Not Found"
        UpdateBorrower = False
    Else
        UpdateBorrower = True
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error updating borrower: " & Err.Description, vbCritical, "Error"
    UpdateBorrower = False
End Function

' Delete borrower record
Public Function DeleteBorrower(borrowerID As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    found = False
    
    ' Find and delete borrower
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = borrowerID Then
            lr.Delete
            found = True
            Exit For
        End If
    Next lr
    
    If Not found Then
        MsgBox "Borrower not found.", vbExclamation, "Not Found"
        DeleteBorrower = False
    Else
        DeleteBorrower = True
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error deleting borrower: " & Err.Description, vbCritical, "Error"
    DeleteBorrower = False
End Function

' ============================================================
' LOAN FUNCTIONS
' ============================================================

' Add a new loan record
Public Function AddLoan(borrowerID As String, loanType As String, principal As Currency, _
                       interestRate As Double, releaseDate As Date, dueDate As Date, _
                       terms As Integer, dailyPayment As Currency, collectorID As String, _
                       Optional status As String = "Active") As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim loanID As String
    Dim interestAmount As Currency
    Dim totalCollectible As Currency
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    ' Validate required fields
    If Len(borrowerID) = 0 Or principal <= 0 Then
        MsgBox "BorrowerID and Principal are required and must be valid.", vbExclamation, "Validation Error"
        Exit Function
    End If
    
    ' Generate LoanID
    loanID = GenerateID("tblLoans", "LoanID", "LN")
    
    ' Calculate interest and total
    interestAmount = principal * interestRate
    totalCollectible = principal + interestAmount
    
    ' Add new row to table
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = loanID
        .Cells(1, 2).Value = borrowerID
        .Cells(1, 3).Value = loanType
        .Cells(1, 4).Value = principal
        .Cells(1, 5).Value = interestRate
        .Cells(1, 6).Value = interestAmount
        .Cells(1, 7).Value = totalCollectible
        .Cells(1, 8).Value = releaseDate
        .Cells(1, 9).Value = dueDate
        .Cells(1, 10).Value = terms
        .Cells(1, 11).Value = dailyPayment
        .Cells(1, 12).Value = 0 ' AmountPaid - starts at 0
        .Cells(1, 13).Value = totalCollectible ' Balance - starts at total
        .Cells(1, 14).Value = collectorID
        .Cells(1, 15).Value = status
    End With
    
    AddLoan = loanID
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding loan: " & Err.Description, vbCritical, "Error"
    AddLoan = ""
End Function

' Get loan information by LoanID
Public Function GetLoan(loanID As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim loanData As Variant
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    ReDim loanData(1 To 15)
    found = False
    
    ' Search for loan
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = loanID Then
            Dim i As Integer
            For i = 1 To 15
                loanData(i) = lr.Range.Cells(1, i).Value
            Next i
            found = True
            Exit For
        End If
    Next lr
    
    If found Then
        GetLoan = loanData
    Else
        GetLoan = Nothing
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error retrieving loan: " & Err.Description, vbCritical, "Error"
    GetLoan = Nothing
End Function

' ============================================================
' PAYMENT FUNCTIONS
' ============================================================

' Record a payment
Public Function RecordPayment(loanID As String, paymentAmount As Currency, paymentDate As Date, _
                             orNumber As String, collectorID As String) As String
    On Error GoTo ErrorHandler
    
    Dim wsPayments As Worksheet
    Dim wsLoans As Worksheet
    Dim tblPayments As ListObject
    Dim tblLoans As ListObject
    Dim newRow As ListRow
    Dim paymentID As String
    Dim loanRow As ListRow
    Dim remainingBalance As Currency
    Dim amountPaid As Currency
    Dim found As Boolean
    
    Set wsPayments = ThisWorkbook.Sheets("Payments")
    Set wsLoans = ThisWorkbook.Sheets("Loans")
    Set tblPayments = wsPayments.ListObjects("tblPayments")
    Set tblLoans = wsLoans.ListObjects("tblLoans")
    
    ' Validate
    If Len(loanID) = 0 Or paymentAmount <= 0 Then
        MsgBox "LoanID and Payment Amount are required and must be valid.", vbExclamation, "Validation Error"
        Exit Function
    End If
    
    found = False
    
    ' Find loan and update balance
    For Each loanRow In tblLoans.ListRows
        If loanRow.Range.Cells(1, 1).Value = loanID Then
            ' Get current balance and amount paid
            remainingBalance = loanRow.Range.Cells(1, 13).Value - paymentAmount
            amountPaid = loanRow.Range.Cells(1, 12).Value + paymentAmount
            
            ' Validate payment doesn't exceed balance
            If paymentAmount > loanRow.Range.Cells(1, 13).Value Then
                MsgBox "Payment amount exceeds remaining balance.", vbExclamation, "Validation Error"
                Exit Function
            End If
            
            ' Update loan record
            loanRow.Range.Cells(1, 12).Value = amountPaid ' AmountPaid
            loanRow.Range.Cells(1, 13).Value = remainingBalance ' Balance
            
            ' Update status if fully paid
            If remainingBalance <= 0 Then
                loanRow.Range.Cells(1, 15).Value = "Completed"
            End If
            
            found = True
            Exit For
        End If
    Next loanRow
    
    If Not found Then
        MsgBox "Loan not found.", vbExclamation, "Not Found"
        Exit Function
    End If
    
    ' Generate PaymentID
    paymentID = GenerateID("tblPayments", "PaymentID", "PM")
    
    ' Add payment record
    Set newRow = tblPayments.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = paymentID
        .Cells(1, 2).Value = loanID
        .Cells(1, 3).Value = paymentDate
        .Cells(1, 4).Value = paymentAmount
        .Cells(1, 5).Value = remainingBalance
        .Cells(1, 6).Value = orNumber
        .Cells(1, 7).Value = collectorID
    End With
    
    RecordPayment = paymentID
    Exit Function
    
ErrorHandler:
    MsgBox "Error recording payment: " & Err.Description, vbCritical, "Error"
    RecordPayment = ""
End Function

' ============================================================
' COLLECTOR FUNCTIONS
' ============================================================

' Add a new collector
Public Function AddCollector(name As String, contactNo As String, commission As Currency) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim collectorID As String
    
    Set ws = ThisWorkbook.Sheets("Collectors")
    Set tbl = ws.ListObjects("tblCollectors")
    
    ' Validate
    If Len(name) = 0 Then
        MsgBox "Collector name is required.", vbExclamation, "Validation Error"
        Exit Function
    End If
    
    ' Generate CollectorID
    collectorID = GenerateID("tblCollectors", "CollectorID", "COL")
    
    ' Add new row
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = collectorID
        .Cells(1, 2).Value = name
        .Cells(1, 3).Value = contactNo
        .Cells(1, 4).Value = commission
    End With
    
    AddCollector = collectorID
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding collector: " & Err.Description, vbCritical, "Error"
    AddCollector = ""
End Function

' Get collector information
Public Function GetCollector(collectorID As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim collectorData As Variant
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Collectors")
    Set tbl = ws.ListObjects("tblCollectors")
    
    ReDim collectorData(1 To 4)
    found = False
    
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = collectorID Then
            collectorData(1) = lr.Range.Cells(1, 1).Value
            collectorData(2) = lr.Range.Cells(1, 2).Value
            collectorData(3) = lr.Range.Cells(1, 3).Value
            collectorData(4) = lr.Range.Cells(1, 4).Value
            found = True
            Exit For
        End If
    Next lr
    
    If found Then
        GetCollector = collectorData
    Else
        GetCollector = Nothing
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error retrieving collector: " & Err.Description, vbCritical, "Error"
    GetCollector = Nothing
End Function

' ============================================================
' USER FUNCTIONS
' ============================================================

' Add a new user
Public Function AddUser(username As String, password As String, role As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Users")
    Set tbl = ws.ListObjects("tblUsers")
    
    ' Validate
    If Len(username) = 0 Or Len(password) = 0 Or Len(role) = 0 Then
        MsgBox "Username, Password, and Role are required.", vbExclamation, "Validation Error"
        Exit Function
    End If
    
    ' Check if username already exists
    Dim lr As ListRow
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = username Then
            MsgBox "Username already exists.", vbExclamation, "Duplicate Error"
            Exit Function
        End If
    Next lr
    
    ' Add new user
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = username
        .Cells(1, 2).Value = EncryptPassword(password)
        .Cells(1, 3).Value = role
    End With
    
    AddUser = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding user: " & Err.Description, vbCritical, "Error"
    AddUser = False
End Function

' Verify user credentials
Public Function VerifyUser(username As String, password As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Users")
    Set tbl = ws.ListObjects("tblUsers")
    
    found = False
    
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = username Then
            If lr.Range.Cells(1, 2).Value = EncryptPassword(password) Then
                VerifyUser = True
                found = True
                Exit For
            End If
        End If
    Next lr
    
    If Not found Then
        VerifyUser = False
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error verifying user: " & Err.Description, vbCritical, "Error"
    VerifyUser = False
End Function

' Get user role
Public Function GetUserRole(username As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Users")
    Set tbl = ws.ListObjects("tblUsers")
    
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 1).Value = username Then
            GetUserRole = lr.Range.Cells(1, 3).Value
            Exit Function
        End If
    Next lr
    
    GetUserRole = ""
    Exit Function
ErrorHandler:
    MsgBox "Error retrieving user role: " & Err.Description, vbCritical, "Error"
    GetUserRole = ""
End Function

' ============================================================
' HELPER FUNCTIONS
' ============================================================

' Generate unique ID with format: PREFIX + sequential number
Private Function GenerateID(tableName As String, fieldName As String, prefix As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim maxNum As Long
    Dim newID As String
    
    Set ws = ThisWorkbook.Sheets(GetTableSheet(tableName))
    Set tbl = ws.ListObjects(tableName)
    
    maxNum = 0
    
    ' Find the highest number in the ID field
    For Each lr In tbl.ListRows
        Dim currentValue As String
        Dim numPart As String
        currentValue = lr.Range.Cells(1, GetFieldColumn(tableName, fieldName)).Value
        
        If Len(currentValue) > 0 Then
            numPart = Replace(currentValue, prefix, "")
            If IsNumeric(numPart) Then
                If CLng(numPart) > maxNum Then
                    maxNum = CLng(numPart)
                End If
            End If
        End If
    Next lr
    
    ' Generate new ID
    newID = prefix & Format(maxNum + 1, "000000")
    GenerateID = newID
    
    Exit Function
ErrorHandler:
    GenerateID = prefix & Format(1, "000000")
End Function

' Get worksheet name for a table
Private Function GetTableSheet(tableName As String) As String
    Select Case tableName
        Case "tblBorrowers"
            GetTableSheet = "Borrowers"
        Case "tblLoans"
            GetTableSheet = "Loans"
        Case "tblPayments"
            GetTableSheet = "Payments"
        Case "tblCollectors"
            GetTableSheet = "Collectors"
        Case "tblUsers"
            GetTableSheet = "Users"
        Case Else
            GetTableSheet = ""
    End Select
End Function

' Get column number for a field in a table
Private Function GetFieldColumn(tableName As String, fieldName As String) As Integer
    Dim columnMap As Object
    Set columnMap = CreateObject("Scripting.Dictionary")
    
    Select Case tableName
        Case "tblBorrowers"
            columnMap.Add "BorrowerID", 1
            columnMap.Add "LastName", 2
            columnMap.Add "FirstName", 3
            columnMap.Add "MiddleName", 4
            columnMap.Add "Address", 5
            columnMap.Add "ContactNo", 6
            columnMap.Add "Occupation", 7
            columnMap.Add "PhotoPath", 8
            columnMap.Add "Status", 9
            
        Case "tblLoans"
            columnMap.Add "LoanID", 1
            columnMap.Add "BorrowerID", 2
            columnMap.Add "LoanType", 3
            columnMap.Add "Principal", 4
            columnMap.Add "InterestRate", 5
            columnMap.Add "InterestAmount", 6
            columnMap.Add "TotalCollectible", 7
            columnMap.Add "ReleaseDate", 8
            columnMap.Add "DueDate", 9
            columnMap.Add "Terms", 10
            columnMap.Add "DailyPayment", 11
            columnMap.Add "AmountPaid", 12
            columnMap.Add "Balance", 13
            columnMap.Add "CollectorID", 14
            columnMap.Add "Status", 15
            
        Case "tblPayments"
            columnMap.Add "PaymentID", 1
            columnMap.Add "LoanID", 2
            columnMap.Add "PaymentDate", 3
            columnMap.Add "AmountPaid", 4
            columnMap.Add "RemainingBalance", 5
            columnMap.Add "ORNumber", 6
            columnMap.Add "CollectorID", 7
            
        Case "tblCollectors"
            columnMap.Add "CollectorID", 1
            columnMap.Add "Name", 2
            columnMap.Add "ContactNo", 3
            columnMap.Add "Commission", 4
            
        Case "tblUsers"
            columnMap.Add "Username", 1
            columnMap.Add "Password", 2
            columnMap.Add "Role", 3
    End Select
    
    If columnMap.Exists(fieldName) Then
        GetFieldColumn = columnMap(fieldName)
    Else
        GetFieldColumn = 0
    End If
End Function

' Check if string is numeric
Private Function IsNumeric(str As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsNull(CLng(str))
    On Error GoTo 0
End Function

' Simple password encryption (basic obfuscation - enhance for production)
Private Function EncryptPassword(password As String) As String
    ' Note: This is basic encryption. For production, use proper encryption library
    ' This just reverses and adds a suffix for demo purposes
    EncryptPassword = StrReverse(password) & "###"
End Function

' Decrypt password (for verification)
Public Function DecryptPassword(encryptedPassword As String) As String
    ' Remove suffix
    Dim decrypted As String
    decrypted = Left(encryptedPassword, Len(encryptedPassword) - 3)
    DecryptPassword = StrReverse(decrypted)
End Function
