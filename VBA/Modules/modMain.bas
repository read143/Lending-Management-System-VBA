' ================================================================
' MODULE: modMain
' PURPOSE: Application entry point, navigation, and orchestration
' INTEGRATES WITH: modDatabaseFunctions for all data operations
' ================================================================

Option Explicit

' ================================================================
' GLOBAL DECLARATIONS
' ================================================================

' Application constants
Public Const APP_NAME As String = "Lending Management System"
Public Const APP_VERSION As String = "1.0.0"
Public Const COMPANY_NAME As String = "Your Company Name"
Public Const MAX_LOGIN_ATTEMPTS As Integer = 3

' Session variables
Public g_CurrentUser As String
Public g_UserRole As String
Public g_LoginTime As Date
Public g_IsLoggedIn As Boolean
Public g_LoginAttempts As Integer

' Global objects for current context
Public g_CurrentBorrowerID As String
Public g_CurrentLoanID As String
Public g_CurrentPaymentID As String
Public g_CurrentCollectorID As String

' ================================================================
' APPLICATION STARTUP
' ================================================================

' Auto-run when workbook opens
Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Set application title
    Application.Caption = APP_NAME & " v" & APP_VERSION
    
    ' Initialize session
    Call InitializeSession
    
    ' Show login form
    frmLogin.Show
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Startup Error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub

' Initialize session variables
Private Sub InitializeSession()
    g_CurrentUser = ""
    g_UserRole = ""
    g_LoginTime = #1/1/1900#
    g_IsLoggedIn = False
    g_LoginAttempts = 0
    
    g_CurrentBorrowerID = ""
    g_CurrentLoanID = ""
    g_CurrentPaymentID = ""
    g_CurrentCollectorID = ""
End Sub

' ================================================================
' LOGIN MANAGEMENT
' ================================================================

' Handle login from frmLogin
Public Function ProcessLogin(username As String, password As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if user exists and password matches
    If modDatabaseFunctions.VerifyUser(username, password) Then
        ' Set session variables
        g_CurrentUser = username
        g_UserRole = modDatabaseFunctions.GetUserRole(username)
        g_LoginTime = Now
        g_IsLoggedIn = True
        g_LoginAttempts = 0
        
        ' Log successful login
        Call LogEvent("Login", "User " & username & " logged in successfully")
        
        ProcessLogin = True
    Else
        g_LoginAttempts = g_LoginAttempts + 1
        
        ' Log failed attempt
        Call LogEvent("Login Failed", "Failed login attempt for user: " & username)
        
        ' Check if max attempts reached
        If g_LoginAttempts >= MAX_LOGIN_ATTEMPTS Then
            MsgBox "Maximum login attempts exceeded. Application will close.", vbCritical
            Call ApplicationShutdown
        End If
        
        ProcessLogin = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Login error: " & Err.Description, vbCritical
    ProcessLogin = False
End Function

' Logout current user
Public Sub LogoutUser()
    On Error GoTo ErrorHandler
    
    If g_IsLoggedIn Then
        Call LogEvent("Logout", "User " & g_CurrentUser & " logged out")
    End If
    
    ' Unload all forms
    Call UnloadAllForms
    
    ' Reset session
    Call InitializeSession
    
    ' Show login form
    frmLogin.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Logout error: " & Err.Description, vbCritical
End Sub

' ================================================================
' NAVIGATION FUNCTIONS
' ================================================================

' Check if user is logged in
Private Function CheckLogin() As Boolean
    If Not g_IsLoggedIn Then
        MsgBox "Please login first.", vbExclamation
        frmLogin.Show
        CheckLogin = False
    Else
        CheckLogin = True
    End If
End Function

' Open Dashboard
Public Sub OpenDashboard()
    If Not CheckLogin Then Exit Sub
    
    On Error Resume Next
    Unload frmDashboard
    On Error GoTo 0
    
    frmDashboard.Show
End Sub

' Open Borrower Form (Add mode or Edit mode)
Public Sub OpenBorrowerForm(Optional ByVal BorrowerID As String = "")
    If Not CheckLogin Then Exit Sub
    
    On Error Resume Next
    Unload frmBorrower
    On Error GoTo 0
    
    ' Set current borrower ID if editing
    If Len(BorrowerID) > 0 Then
        g_CurrentBorrowerID = BorrowerID
    Else
        g_CurrentBorrowerID = ""
    End If
    
    frmBorrower.Show
End Sub

' Open Loan Form (Add mode or Edit mode)
Public Sub OpenLoanForm(Optional ByVal LoanID As String = "")
    If Not CheckLogin Then Exit Sub
    
    On Error Resume Next
    Unload frmLoan
    On Error GoTo 0
    
    If Len(LoanID) > 0 Then
        g_CurrentLoanID = LoanID
    Else
        g_CurrentLoanID = ""
    End If
    
    frmLoan.Show
End Sub

' Open Payment Form
Public Sub OpenPaymentForm(Optional ByVal LoanID As String = "")
    If Not CheckLogin Then Exit Sub
    
    On Error Resume Next
    Unload frmPayment
    On Error GoTo 0
    
    If Len(LoanID) > 0 Then
        g_CurrentLoanID = LoanID
    Else
        g_CurrentLoanID = ""
    End If
    
    frmPayment.Show
End Sub

' Open Collector Form
Public Sub OpenCollectorForm(Optional ByVal CollectorID As String = "")
    If Not CheckLogin Then Exit Sub
    
    ' Check permission
    If Not HasPermission("ManageCollectors") Then
        MsgBox "You don't have permission to manage collectors.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Unload frmCollector
    On Error GoTo 0
    
    If Len(CollectorID) > 0 Then
        g_CurrentCollectorID = CollectorID
    Else
        g_CurrentCollectorID = ""
    End If
    
    frmCollector.Show
End Sub

' ================================================================
' REPORT FUNCTIONS
' ================================================================

' Generate reports based on type
Public Sub GenerateReport(reportType As String, Optional params As Variant)
    If Not CheckLogin Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Select Case reportType
        Case "Borrowers"
            Call modReports.GenerateBorrowerReport
        Case "Loans"
            Call modReports.GenerateLoanReport
        Case "Payments"
            Call modReports.GeneratePaymentReport
        Case "Aging"
            Call modReports.GenerateAgingReport
        Case "Collection"
            Call modReports.GenerateCollectionReport
        Case "Dashboard"
            Call modReports.GenerateDashboardReport
        Case Else
            MsgBox "Unknown report type: " & reportType, vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating report: " & Err.Description, vbCritical
End Sub

' ================================================================
' PERMISSION FUNCTIONS
' ================================================================

' Check if current user has permission
Public Function HasPermission(permission As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Admin has all permissions
    If g_UserRole = "Administrator" Then
        HasPermission = True
        Exit Function
    End If
    
    ' Define role-based permissions
    Select Case g_UserRole
        Case "Collector"
            ' Collectors can only view and record payments
            Select Case permission
                Case "ViewBorrowers", "ViewLoans", "RecordPayments", "ViewReports"
                    HasPermission = True
                Case Else
                    HasPermission = False
            End Select
            
        Case "Manager"
            ' Managers can do almost everything except user management
            Select Case permission
                Case "ViewBorrowers", "ViewLoans", "RecordPayments", "ViewReports", _
                     "AddBorrower", "EditBorrower", "AddLoan", "EditLoan", _
                     "ManageCollectors"
                    HasPermission = True
                Case "ManageUsers"
                    HasPermission = False
                Case Else
                    HasPermission = False
            End Select
            
        Case Else
            HasPermission = False
    End Select
    
    Exit Function
    
ErrorHandler:
    HasPermission = False
End Function

' ================================================================
' UTILITY FUNCTIONS
' ================================================================

' Unload all forms
Private Sub UnloadAllForms()
    On Error Resume Next
    Unload frmLogin
    Unload frmDashboard
    Unload frmBorrower
    Unload frmLoan
    Unload frmPayment
    Unload frmCollector
    On Error GoTo 0
End Sub

' Log events to error log sheet
Public Sub LogEvent(eventType As String, description As String)
    On Error Resume Next
    
    Dim wsLog As Worksheet
    Dim nextRow As Long
    
    ' Check if error log sheet exists
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("ErrorLog")
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "ErrorLog"
        wsLog.Range("A1:E1") = Array("Date", "User", "Event Type", "Description", "IP Address")
        wsLog.Range("A1:E1").Font.Bold = True
        wsLog.Columns("A:E").AutoFit
    End If
    
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    wsLog.Cells(nextRow, 1) = Now
    wsLog.Cells(nextRow, 2) = IIf(g_CurrentUser <> "", g_CurrentUser, "SYSTEM")
    wsLog.Cells(nextRow, 3) = eventType
    wsLog.Cells(nextRow, 4) = description
    
    ' Auto-save if possible
    On Error Resume Next
    wsLog.Parent.Save
    On Error GoTo 0
End Sub

' Application shutdown
Public Sub ApplicationShutdown()
    On Error Resume Next
    
    ' Log shutdown
    If g_IsLoggedIn Then
        Call LogEvent("Shutdown", "User " & g_CurrentUser & " closed the application")
    End If
    
    ' Unload all forms
    Call UnloadAllForms
    
    ' Reset session
    Call InitializeSession
    
    ' Reset Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.Caption = "Microsoft Excel"
    
    ' Save workbook
    If Not ThisWorkbook.Saved Then
        ThisWorkbook.Save
    End If
    
    ' Quit Excel (optional - comment out if you want to keep Excel open)
    ' Application.Quit
End Sub

' ================================================================
' MENU FUNCTIONS (for Ribbon or Quick Access Toolbar)
' ================================================================

Public Sub Menu_Dashboard()
    Call OpenDashboard
End Sub

Public Sub Menu_AddBorrower()
    Call OpenBorrowerForm
End Sub

Public Sub Menu_ViewBorrowers()
    Call OpenBorrowerForm
End Sub

Public Sub Menu_AddLoan()
    Call OpenLoanForm
End Sub

Public Sub Menu_ViewLoans()
    Call OpenLoanForm
End Sub

Public Sub Menu_RecordPayment()
    Call OpenPaymentForm
End Sub

Public Sub Menu_ViewPayments()
    Call OpenPaymentForm
End Sub

Public Sub Menu_ManageCollectors()
    Call OpenCollectorForm
End Sub

Public Sub Menu_Reports()
    ' Show report selection form (you may need to create this)
    frmReports.Show
End Sub

Public Sub Menu_Logout()
    Call LogoutUser
End Sub

Public Sub Menu_Exit()
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo) = vbYes Then
        Call ApplicationShutdown
    End If
End Sub

' ================================================================
' HOTKEY FUNCTIONS
' ================================================================

Public Sub Hotkey_Dashboard()
    Call OpenDashboard
End Sub

Public Sub Hotkey_AddBorrower()
    Call OpenBorrowerForm
End Sub

Public Sub Hotkey_AddLoan()
    Call OpenLoanForm
End Sub

Public Sub Hotkey_RecordPayment()
    Call OpenPaymentForm
End Sub

' ================================================================
' WORKBOOK EVENT HANDLERS (Call from ThisWorkbook)
' ================================================================

Public Sub Workbook_Open()
    Call Auto_Open
End Sub

Public Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    If g_IsLoggedIn Then
        If MsgBox("Are you sure you want to close " & APP_NAME & "?", _
                  vbQuestion + vbYesNo) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    Call ApplicationShutdown
End Sub

' ================================================================
' END OF MODULE
' ================================================================
