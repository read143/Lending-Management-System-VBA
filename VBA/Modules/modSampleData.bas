' modSampleData - Populate System with Test Data
' Purpose: Create sample data for development and testing
' Date: 2026-07-17

Option Explicit

' Main procedure to populate all sample data
Public Sub PopulateAllSampleData()
    On Error GoTo ErrorHandler
    
    MsgBox "This will populate the system with test data." & vbCrLf & _
           "Continue?", vbYesNo, "Populate Sample Data"
    
    If MsgBox("Are you sure? This may overwrite existing data.", vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    
    ' Clear existing data
    ClearAllData
    
    ' Populate each table
    PopulateUsers
    PopulateBorrowers
    PopulateCollectors
    PopulateLoans
    PopulatePayments
    
    MsgBox "Sample data populated successfully!" & vbCrLf & _
           "You can now test the system.", vbInformation, "Success"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' CLEAR DATA PROCEDURES
' ============================================================

Private Sub ClearAllData()
    On Error Resume Next
    
    ' Clear all tables
    ClearTableData "tblUsers"
    ClearTableData "tblBorrowers"
    ClearTableData "tblCollectors"
    ClearTableData "tblLoans"
    ClearTableData "tblPayments"
    
End Sub

Private Sub ClearTableData(tableName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowCount As Long
    Dim i As Long
    
    ' Get worksheet name from table name
    Dim wsName As String
    wsName = GetTableSheet(tableName)
    
    Set ws = ThisWorkbook.Sheets(wsName)
    Set tbl = ws.ListObjects(tableName)
    
    ' Delete all rows except header
    rowCount = tbl.ListRows.Count
    For i = rowCount To 1 Step -1
        tbl.ListRows(i).Delete
    Next i
    
    Exit Sub
ErrorHandler:
    ' Table might already be empty
End Sub

' ============================================================
' POPULATE USERS TABLE
' ============================================================

Private Sub PopulateUsers()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Users")
    Set tbl = ws.ListObjects("tblUsers")
    
    ' User 1: Admin Account
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = "admin"
        .Cells(1, 2).Value = EncryptPassword("admin123")
        .Cells(1, 3).Value = "Admin"
    End With
    
    ' User 2: Collector Account
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = "collector1"
        .Cells(1, 2).Value = EncryptPassword("collector123")
        .Cells(1, 3).Value = "Collector"
    End With
    
    ' User 3: Manager Account
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = "manager"
        .Cells(1, 2).Value = EncryptPassword("manager123")
        .Cells(1, 3).Value = "Manager"
    End With
    
    ' User 4: Assistant Account
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = "assistant"
        .Cells(1, 2).Value = EncryptPassword("assistant123")
        .Cells(1, 3).Value = "Assistant"
    End With
    
    Exit Sub
ErrorHandler:
    MsgBox "Error populating users: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' POPULATE BORROWERS TABLE
' ============================================================

Private Sub PopulateBorrowers()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Borrowers")
    Set tbl = ws.ListObjects("tblBorrowers")
    
    ' Sample borrower data
    Dim borrowerData As Variant
    borrowerData = Array( _
        Array("Smith", "John", "Michael", "123 Main St, Metro Manila", "09171234567", "Engineer", "", "Active"), _
        Array("Johnson", "Maria", "Santos", "456 Oak Ave, Quezon City", "09187654321", "Teacher", "", "Active"), _
        Array("Garcia", "Carlos", "Rodriguez", "789 Pine Rd, Makati", "09161112222", "Business Owner", "", "Active"), _
        Array("Martinez", "Rosa", "Cruz", "321 Elm St, Pasig", "09159876543", "Accountant", "", "Active"), _
        Array("Lopez", "Pedro", "Reyes", "654 Maple Dr, Las Piñas", "09165555555", "Mechanic", "", "Active"), _
        Array("Hernandez", "Angeles", "Gonzales", "987 Cedar Lane, Cebu", "09177777777", "Nurse", "", "Active"), _
        Array("Torres", "Miguel", "Fernandez", "147 Birch St, Davao", "09188888888", "Construction Manager", "", "Active"), _
        Array("Ramirez", "Lucia", "Morales", "258 Spruce Ave, Cagayan de Oro", "09169999999", "Hairdresser", "", "Active"), _
        Array("Castillo", "Antonio", "Navarro", "369 Walnut Rd, Iloilo", "09174444444", "Farmer", "", "Active"), _
        Array("Rivera", "Isabel", "Delgado", "741 Chestnut Dr, Bacolod", "09183333333", "Seamstress", "", "Active") _
    )
    
    ' Add borrowers
    For i = LBound(borrowerData) To UBound(borrowerData)
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = "BOR" & Format(i + 1, "000000")
            .Cells(1, 2).Value = borrowerData(i)(0)
            .Cells(1, 3).Value = borrowerData(i)(1)
            .Cells(1, 4).Value = borrowerData(i)(2)
            .Cells(1, 5).Value = borrowerData(i)(3)
            .Cells(1, 6).Value = borrowerData(i)(4)
            .Cells(1, 7).Value = borrowerData(i)(5)
            .Cells(1, 8).Value = borrowerData(i)(6)
            .Cells(1, 9).Value = borrowerData(i)(7)
        End With
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "Error populating borrowers: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' POPULATE COLLECTORS TABLE
' ============================================================

Private Sub PopulateCollectors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Collectors")
    Set tbl = ws.ListObjects("tblCollectors")
    
    ' Sample collector data
    Dim collectorData As Variant
    collectorData = Array( _
        Array("Juan Dela Cruz", "09201234567", 0.05), _
        Array("Maria Santos", "09209876543", 0.05), _
        Array("Carlos Reyes", "09205555555", 0.05), _
        Array("Rosa Garcia", "09207777777", 0.05), _
        Array("Pedro Gonzales", "09203333333", 0.05) _
    )
    
    ' Add collectors
    For i = LBound(collectorData) To UBound(collectorData)
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = "COL" & Format(i + 1, "000000")
            .Cells(1, 2).Value = collectorData(i)(0)
            .Cells(1, 3).Value = collectorData(i)(1)
            .Cells(1, 4).Value = collectorData(i)(2)
        End With
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "Error populating collectors: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' POPULATE LOANS TABLE
' ============================================================

Private Sub PopulateLoans()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim i As Integer
    Dim loanID As String
    Dim principal As Currency
    Dim interestRate As Double
    Dim interestAmount As Currency
    Dim totalCollectible As Currency
    Dim releaseDate As Date
    Dim dueDate As Date
    Dim terms As Integer
    Dim dailyPayment As Currency
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    ' Create 10 sample loans
    For i = 1 To 10
        ' Generate loan data
        loanID = "LN" & Format(i, "000000")
        principal = (5000 + (i * 1000)) ' 6000 to 15000
        interestRate = 0.15 ' 15% interest
        interestAmount = principal * interestRate
        totalCollectible = principal + interestAmount
        releaseDate = DateAdd("d", -30, Now()) ' 30 days ago
        dueDate = DateAdd("m", 12, releaseDate) ' 12 months from release
        terms = 12
        dailyPayment = totalCollectible / 360 ' Divided by 360 days
        
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = loanID
            .Cells(1, 2).Value = "BOR" & Format(i, "000000")
            .Cells(1, 3).Value = IIf(i Mod 3 = 1, "Personal Loan", IIf(i Mod 3 = 2, "Business Loan", "Emergency Loan"))
            .Cells(1, 4).Value = principal
            .Cells(1, 5).Value = interestRate
            .Cells(1, 6).Value = interestAmount
            .Cells(1, 7).Value = totalCollectible
            .Cells(1, 8).Value = releaseDate
            .Cells(1, 9).Value = dueDate
            .Cells(1, 10).Value = terms
            .Cells(1, 11).Value = dailyPayment
            .Cells(1, 12).Value = 0 ' AmountPaid starts at 0
            .Cells(1, 13).Value = totalCollectible ' Balance starts at total
            .Cells(1, 14).Value = "COL" & Format((i Mod 5) + 1, "000000")
            .Cells(1, 15).Value = "Active"
        End With
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "Error populating loans: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' POPULATE PAYMENTS TABLE
' ============================================================

Private Sub PopulatePayments()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim loanWs As Worksheet
    Dim loanTbl As ListObject
    Dim loanRow As ListRow
    Dim paymentCount As Integer
    Dim paymentAmount As Currency
    Dim remainingBalance As Currency
    Dim paymentDate As Date
    Dim j As Integer
    
    Set ws = ThisWorkbook.Sheets("Payments")
    Set tbl = ws.ListObjects("tblPayments")
    
    Set loanWs = ThisWorkbook.Sheets("Loans")
    Set loanTbl = loanWs.ListObjects("tblLoans")
    
    paymentCount = 0
    
    ' Add payments for first 3 loans (3 payments each)
    For Each loanRow In loanTbl.ListRows
        If paymentCount < 3 Then
            ' Add 3 payments per loan
            For j = 1 To 3
                paymentAmount = CDbl(loanRow.Range.Cells(1, 11).Value) * 30 ' 30 days worth
                remainingBalance = CDbl(loanRow.Range.Cells(1, 13).Value) - paymentAmount
                paymentDate = DateAdd("d", j * 30, CDDate(loanRow.Range.Cells(1, 8).Value))
                
                Set newRow = tbl.ListRows.Add
                With newRow.Range
                    .Cells(1, 1).Value = "PM" & Format(paymentCount * 3 + j, "000000")
                    .Cells(1, 2).Value = loanRow.Range.Cells(1, 1).Value
                    .Cells(1, 3).Value = paymentDate
                    .Cells(1, 4).Value = paymentAmount
                    .Cells(1, 5).Value = remainingBalance
                    .Cells(1, 6).Value = "OR" & Format(paymentCount * 3 + j, "000000")
                    .Cells(1, 7).Value = loanRow.Range.Cells(1, 14).Value
                End With
            Next j
            paymentCount = paymentCount + 1
        End If
    Next loanRow
    
    Exit Sub
ErrorHandler:
    MsgBox "Error populating payments: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' HELPER FUNCTIONS
' ============================================================

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

Private Function EncryptPassword(password As String) As String
    ' Simple encryption - reverse + suffix
    EncryptPassword = StrReverse(password) & "###"
End Function
