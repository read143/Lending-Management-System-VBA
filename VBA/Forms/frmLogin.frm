' frmLogin - User Authentication Form
' Purpose: Authenticate users before accessing the system
' Date: 2026-07-17

Option Explicit

Private Sub UserForm_Initialize()
    ' Center form on screen
    Me.StartupPosition = 1 ' Center Owner
    
    ' Set focus to username field
    txtUsername.SetFocus
    
    ' Clear fields
    txtUsername.Value = ""
    txtPassword.Value = ""
    lblErrorMessage.Caption = ""
    lblErrorMessage.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub cmdLogin_Click()
    On Error GoTo ErrorHandler
    
    ' Validate fields
    If Len(txtUsername.Value) = 0 Then
        ShowError "Username is required."
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If Len(txtPassword.Value) = 0 Then
        ShowError "Password is required."
        txtPassword.SetFocus
        Exit Sub
    End If
    
    ' Verify user credentials
    If VerifyUser(txtUsername.Value, txtPassword.Value) Then
        ' Store logged-in user info in a global variable or worksheet
        ThisWorkbook.Sheets(1).Range("LoggedInUser").Value = txtUsername.Value
        ThisWorkbook.Sheets(1).Range("UserRole").Value = GetUserRole(txtUsername.Value)
        
        lblErrorMessage.Caption = "Login successful!"
        lblErrorMessage.ForeColor = RGB(0, 176, 0)
        
        ' Open Dashboard
        Unload Me
        frmDashboard.Show
    Else
        ShowError "Invalid username or password."
        txtPassword.Value = ""
        txtUsername.SetFocus
    End If
    
    Exit Sub
ErrorHandler:
    ShowError "Error during login: " & Err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    ThisWorkbook.Close
End Sub

Private Sub txtPassword_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow Enter key to login
    If KeyAscii = 13 Then
        cmdLogin_Click
        KeyAscii = 0
    End If
End Sub

Private Sub ShowError(message As String)
    lblErrorMessage.Caption = message
    lblErrorMessage.ForeColor = RGB(255, 0, 0)
End Sub
