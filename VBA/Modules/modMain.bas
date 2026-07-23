' ================================================================
' MODULE: modMain
' PURPOSE: Application entry point and navigation
' ================================================================

Option Explicit

' Global Session Variables
Public g_CurrentUser As String
Public g_UserRole As String
Public g_LoginTime As Date
Public g_IsLoggedIn As Boolean

' ================================================================
' AUTO-START - Runs when workbook opens
' ================================================================

Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Hide ribbon
    ' Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    ' Show login form
    frmLogin.Show
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Startup Error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub

' ================================================================
' NAVIGATION FUNCTIONS
' ================================================================

Public Sub OpenDashboard()
    If Not g_IsLoggedIn Then
        MsgBox "Please login first.", vbExclamation
        frmLogin.Show
        Exit Sub
    End If
    
    frmDashboard.Show
End Sub

Public Sub OpenBorrowerForm(Optional ByVal BorrowerID As Long = 0)
    If Not g_IsLoggedIn Then
        MsgBox "Please login first.", vbExclamation
        frmLogin.Show
        Exit Sub
    End If
    
    frmBorrower.Show
End Sub

Public Sub OpenLoanForm(Optional ByVal LoanID As Long = 0)
    If Not g_IsLoggedIn Then
        MsgBox "Please login first.", vbExclamation
        frmLogin.Show
        Exit Sub
    End If
    
    frmLoan.Show
End Sub

Public Sub OpenPaymentForm(Optional ByVal PaymentID As Long = 0)
    If Not g_IsLoggedIn Then
        MsgBox "Please login first.", vbExclamation
        frmLogin.Show
        Exit Sub
    End If
    
    frmPayment.Show
End Sub

Public Sub OpenCollectorForm()
    If Not g_IsLoggedIn Then
        MsgBox "Please login first.", vbExclamation
        frmLogin.Show
        Exit Sub
    End If
    
    frmCollector.Show
End Sub

Public Sub LogoutUser()
    g_CurrentUser = ""
    g_UserRole = ""
    g_IsLoggedIn = False
    g_LoginTime = #1/1/1900#
    
    Unload frmDashboard
    Unload frmBorrower
    Unload frmLoan
    Unload frmPayment
    
    frmLogin.Show
End Sub

' ================================================================
' MENU FUNCTIONS (For Ribbon Buttons)
' ================================================================

Public Sub Menu_Dashboard()
    Call OpenDashboard
End Sub

Public Sub Menu_AddBorrower()
    Call OpenBorrowerForm
End Sub

Public Sub Menu_AddLoan()
    Call OpenLoanForm
End Sub

Public Sub Menu_RecordPayment()
    Call OpenPaymentForm
End Sub

Public Sub Menu_ManageCollectors()
    Call OpenCollectorForm
End Sub

Public Sub Menu_Logout()
    Call LogoutUser
End Sub

' ================================================================
' END OF MODULE
' ================================================================
