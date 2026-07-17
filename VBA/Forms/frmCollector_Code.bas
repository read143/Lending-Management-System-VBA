' frmCollector - Collector Management Form
' Purpose: Add and manage loan collectors/agents
' Date: 2026-07-17

Option Explicit

Private currentCollectorID As String

Private Sub UserForm_Initialize()
    Me.StartupPosition = 1 ' Center Owner
    
    ' Clear form
    ClearForm
    
    ' Load collectors
    LoadCollectors
    
    ' Set focus
    txtName.SetFocus
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo ErrorHandler
    
    ' Validate
    If Len(txtName.Value) = 0 Then
        MsgBox "Collector name is required.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Add collector
    Dim collectorID As String
    collectorID = AddCollector(txtName.Value, txtContactNo.Value, CDbl(txtCommission.Value))
    
    If Len(collectorID) > 0 Then
        MsgBox "Collector added successfully!" & vbCrLf & "Collector ID: " & collectorID, vbInformation, "Success"
        ClearForm
        LoadCollectors
        txtName.SetFocus
    Else
        MsgBox "Error adding collector.", vbCritical, "Error"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdViewAssignedLoans_Click()
    On Error GoTo ErrorHandler
    
    If Len(currentCollectorID) = 0 Then
        MsgBox "Please select a collector.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Load assigned loans into list
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Loans")
    Set tbl = ws.ListObjects("tblLoans")
    
    lstAssignedLoans.Clear
    
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, 14).Value = currentCollectorID Then
            lstAssignedLoans.AddItem lr.Range.Cells(1, 1).Value & " - " & _
                                      "Borrower: " & lr.Range.Cells(1, 2).Value & " - " & _
                                      "Balance: PHP " & Format(lr.Range.Cells(1, 13).Value, "#,##0.00")
        End If
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub lstCollectors_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim collectorID As String
    
    selectedText = lstCollectors.Value
    collectorID = Left(selectedText, InStr(selectedText, " ") - 1)
    
    currentCollectorID = collectorID
    
    ' Get and display collector data
    Dim collectorData As Variant
    collectorData = GetCollector(collectorID)
    
    If Not IsNothing(collectorData) Then
        txtCollectorID.Value = collectorData(1)
        txtName.Value = collectorData(2)
        txtContactNo.Value = collectorData(3)
        txtCommission.Value = collectorData(4)
    End If
    
    ' Load assigned loans
    cmdViewAssignedLoans_Click
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ClearForm()
    txtCollectorID.Value = ""
    txtName.Value = ""
    txtContactNo.Value = ""
    txtCommission.Value = 0
    lstAssignedLoans.Clear
    currentCollectorID = ""
End Sub

Private Sub LoadCollectors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set ws = ThisWorkbook.Sheets("Collectors")
    Set tbl = ws.ListObjects("tblCollectors")
    
    lstCollectors.Clear
    
    For Each lr In tbl.ListRows
        lstCollectors.AddItem lr.Range.Cells(1, 1).Value & " - " & lr.Range.Cells(1, 2).Value
    Next lr
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading collectors: " & Err.Description, vbCritical, "Error"
End Sub

Private Function IsNothing(obj As Variant) As Boolean
    IsNothing = (TypeName(obj) = "Nothing")
End Function