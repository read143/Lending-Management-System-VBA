# Lending Management System - UserForm Design Guide
## Complete Instructions for Creating All Forms

---

## TABLE OF CONTENTS
1. frmLogin - User Authentication Form
2. frmDashboard - Main Navigation Form
3. frmBorrower - Borrower Management Form
4. frmLoan - Loan Management Form
5. frmPayment - Payment Recording Form
6. frmCollector - Collector Management Form

---

## FORM 1: frmLogin - User Authentication Form

### Purpose
Authenticate users before accessing the system with username and password validation.

### Form Properties
| Property | Value |
|----------|-------|
| Name | frmLogin |
| Caption | Lending Management System - Login |
| Width | 400 |
| Height | 300 |
| StartupPosition | 1 (Center Owner) |
| BackColor | White |

### Controls to Add

#### 1. Title Label
- **Name:** lblTitle
- **Caption:** Lending Management System
- **Left:** 20, **Top:** 20
- **Width:** 350, **Height:** 25
- **Font:** Size 16, Bold

#### 2. Logo/Image (Optional)
- **Name:** imgLogo
- **Left:** 150, **Top:** 50
- **Width:** 100, **Height:** 50
- **Picture:** Your company logo

#### 3. Username Label
- **Name:** lblUsername
- **Caption:** Username:
- **Left:** 20, **Top:** 120
- **Width:** 80, **Height:** 20
- **Font:** Size 10

#### 4. Username TextBox
- **Name:** txtUsername
- **Left:** 110, **Top:** 120
- **Width:** 250, **Height:** 20
- **BackColor:** White
- **BorderStyle:** 1 (Single)

#### 5. Password Label
- **Name:** lblPassword
- **Caption:** Password:
- **Left:** 20, **Top:** 150
- **Width:** 80, **Height:** 20
- **Font:** Size 10

#### 6. Password TextBox
- **Name:** txtPassword
- **Left:** 110, **Top:** 150
- **Width:** 250, **Height:** 20
- **PasswordChar:** *
- **BackColor:** White
- **BorderStyle:** 1 (Single)

#### 7. Error Message Label
- **Name:** lblErrorMessage
- **Caption:** (empty)
- **Left:** 20, **Top:** 180
- **Width:** 340, **Height:** 20
- **ForeColor:** Red
- **Font:** Size 9

#### 8. Login Button
- **Name:** cmdLogin
- **Caption:** Login
- **Left:** 110, **Top:** 220
- **Width:** 100, **Height:** 30
- **BackColor:** Green
- **ForeColor:** White
- **Font:** Size 11, Bold

#### 9. Cancel Button
- **Name:** cmdCancel
- **Caption:** Cancel
- **Left:** 260, **Top:** 220
- **Width:** 100, **Height:** 30
- **BackColor:** Red
- **ForeColor:** White
- **Font:** Size 11, Bold

### Form Layout Diagram
```
┌─────────────────────────────────────┐
│  Lending Management System - Login  │
│                                     │
│     [Logo Image - Optional]         │
│                                     │
│  Username: [________________]       │
│                                     │
│  Password: [________________]       │
│                                     │
│  [Error Message Display Area]       │
│                                     │
│     [Login]         [Cancel]        │
└─────────────────────────────────────┘
```

---

## FORM 2: frmDashboard - Main Navigation Form

### Purpose
Main menu and dashboard showing system statistics and navigation to all modules.

### Form Properties
| Property | Value |
|----------|-------|
| Name | frmDashboard |
| Caption | Lending Management System - Dashboard |
| Width | 600 |
| Height | 500 |
| StartupPosition | 1 (Center Owner) |
| BackColor | Light Gray (RGB 240, 240, 240) |

### Controls to Add

#### 1. Welcome Label
- **Name:** lblUserInfo
- **Caption:** Welcome, Admin (Admin)
- **Left:** 20, **Top:** 20
- **Width:** 550, **Height:** 25
- **Font:** Size 12, Bold
- **ForeColor:** Blue

#### 2. Statistics Frame
- **Name:** fraStatistics
- **Caption:** Quick Statistics
- **Left:** 20, **Top:** 60
- **Width:** 550, **Height:** 100
- **Font:** Size 10, Bold

#### 3. Active Loans Label (inside Frame)
- **Name:** lblActiveLoans
- **Caption:** 0 Active Loans
- **Left:** 40, **Top:** 90
- **Width:** 160, **Height:** 30
- **Font:** Size 11, Bold
- **ForeColor:** Green

#### 4. Total Collectible Label (inside Frame)
- **Name:** lblTotalCollectible
- **Caption:** PHP 0.00
- **Left:** 220, **Top:** 90
- **Width:** 160, **Height:** 30
- **Font:** Size 11, Bold
- **ForeColor:** Blue

#### 5. Total Paid Label (inside Frame)
- **Name:** lblTotalPaid
- **Caption:** PHP 0.00
- **Left:** 400, **Top:** 90
- **Width:** 160, **Height:** 30
- **Font:** Size 11, Bold
- **ForeColor:** Green

#### 6. Borrowers Button
- **Name:** cmdBorrowers
- **Caption:** 👥 Borrowers
- **Left:** 20, **Top:** 180
- **Width:** 130, **Height:** 60
- **BackColor:** Light Blue
- **Font:** Size 11, Bold

#### 7. Loans Button
- **Name:** cmdLoans
- **Caption:** 💰 Loans
- **Left:** 165, **Top:** 180
- **Width:** 130, **Height:** 60
- **BackColor:** Light Green
- **Font:** Size 11, Bold

#### 8. Payments Button
- **Name:** cmdPayments
- **Caption:** 💳 Payments
- **Left:** 310, **Top:** 180
- **Width:** 130, **Height:** 60
- **BackColor:** Light Yellow
- **Font:** Size 11, Bold

#### 9. Collectors Button
- **Name:** cmdCollectors
- **Caption:** 👤 Collectors
- **Left:** 455, **Top:** 180
- **Width:** 115, **Height:** 60
- **BackColor:** Light Orange
- **Font:** Size 11, Bold

#### 10. Logout Button
- **Name:** cmdLogout
- **Caption:** 🚪 Logout
- **Left:** 230, **Top:** 280
- **Width:** 140, **Height:** 40
- **BackColor:** Red
- **ForeColor:** White
- **Font:** Size 11, Bold

### Form Layout Diagram
```
┌──────────────────────────────────────────────┐
│  Welcome, Admin (Admin)                      │
│                                              │
│  ┌────────────────────────────────────────┐ │
│  │ Quick Statistics                       │ │
│  │ 0 Active Loans | PHP 0.00 | PHP 0.00  │ │
│  └────────────────────────────────────────┘ │
│                                              │
│  [👥 Borrowers] [💰 Loans]  [💳 Payments]  [👤 Collectors] │
│                                              │
│              [🚪 Logout]                     │
└──────────────────────────────────────────────┘
```

---

## FORM 3: frmBorrower - Borrower Management Form

### Purpose
Add, search, view, edit, and delete borrower records.

### Form Properties
| Property | Value |
|----------|-------|
| Name | frmBorrower |
| Caption | Borrower Management |
| Width | 700 |
| Height | 600 |
| StartupPosition | 1 (Center Owner) |

### Controls to Add

#### Left Section - Form Fields

#### 1. Title Label
- **Name:** lblTitle
- **Caption:** Borrower Information
- **Left:** 20, **Top:** 20
- **Width:** 300, **Height:** 25
- **Font:** Size 12, Bold

#### 2. BorrowerID Label
- **Name:** lblBorrowerID
- **Caption:** Borrower ID:
- **Left:** 20, **Top:** 60
- **Width:** 80, **Height:** 20

#### 3. BorrowerID TextBox (Read-only)
- **Name:** txtBorrowerID
- **Left:** 110, **Top:** 60
- **Width:** 200, **Height:** 20
- **Locked:** True
- **BackColor:** Light Gray

#### 4. Last Name Label
- **Name:** lblLastName
- **Caption:** Last Name:
- **Left:** 20, **Top:** 90
- **Width:** 80, **Height:** 20

#### 5. Last Name TextBox
- **Name:** txtLastName
- **Left:** 110, **Top:** 90
- **Width:** 200, **Height:** 20

#### 6. First Name Label
- **Name:** lblFirstName
- **Caption:** First Name:
- **Left:** 20, **Top:** 120
- **Width:** 80, **Height:** 20

#### 7. First Name TextBox
- **Name:** txtFirstName
- **Left:** 110, **Top:** 120
- **Width:** 200, **Height:** 20

#### 8. Middle Name Label
- **Name:** lblMiddleName
- **Caption:** Middle Name:
- **Left:** 20, **Top:** 150
- **Width:** 80, **Height:** 20

#### 9. Middle Name TextBox
- **Name:** txtMiddleName
- **Left:** 110, **Top:** 150
- **Width:** 200, **Height:** 20

#### 10. Address Label
- **Name:** lblAddress
- **Caption:** Address:
- **Left:** 20, **Top:** 180
- **Width:** 80, **Height:** 20

#### 11. Address TextBox
- **Name:** txtAddress
- **Left:** 110, **Top:** 180
- **Width:** 200, **Height:** 40
- **MultiLine:** True

#### 12. Contact No Label
- **Name:** lblContactNo
- **Caption:** Contact No:
- **Left:** 20, **Top:** 230
- **Width:** 80, **Height:** 20

#### 13. Contact No TextBox
- **Name:** txtContactNo
- **Left:** 110, **Top:** 230
- **Width:** 200, **Height:** 20

#### 14. Occupation Label
- **Name:** lblOccupation
- **Caption:** Occupation:
- **Left:** 20, **Top:** 260
- **Width:** 80, **Height:** 20

#### 15. Occupation TextBox
- **Name:** txtOccupation
- **Left:** 110, **Top:** 260
- **Width:** 200, **Height:** 20

#### 16. Photo Path Label
- **Name:** lblPhotoPath
- **Caption:** Photo Path:
- **Left:** 20, **Top:** 290
- **Width:** 80, **Height:** 20

#### 17. Photo Path TextBox
- **Name:** txtPhotoPath
- **Left:** 110, **Top:** 290
- **Width:** 200, **Height:** 20

#### 18. Status Label
- **Name:** lblStatus
- **Caption:** Status:
- **Left:** 20, **Top:** 320
- **Width:** 80, **Height:** 20

#### 19. Status ComboBox
- **Name:** cbxStatus
- **Left:** 110, **Top:** 320
- **Width:** 200, **Height:** 20
- **Style:** 2 (Dropdown List)
- **List Items:** Active, Inactive, Deceased, Blacklisted

#### Right Section - Action Buttons & List

#### 20. Add Button
- **Name:** cmdAdd
- **Caption:** ➕ Add
- **Left:** 330, **Top:** 60
- **Width:** 100, **Height:** 30
- **BackColor:** Green
- **ForeColor:** White

#### 21. Update Button
- **Name:** cmdUpdate
- **Caption:** ✏️ Update
- **Left:** 440, **Top:** 60
- **Width:** 100, **Height:** 30
- **BackColor:** Blue
- **ForeColor:** White

#### 22. Delete Button
- **Name:** cmdDelete
- **Caption:** 🗑️ Delete
- **Left:** 550, **Top:** 60
- **Width:** 100, **Height:** 30
- **BackColor:** Red
- **ForeColor:** White

#### 23. Search Label
- **Name:** lblSearch
- **Caption:** Search by Name:
- **Left:** 330, **Top:** 110
- **Width:** 100, **Height:** 20

#### 24. Search TextBox
- **Name:** txtSearchName
- **Left:** 440, **Top:** 110
- **Width:** 210, **Height:** 20

#### 25. Search Button
- **Name:** cmdSearch
- **Caption:** 🔍 Search
- **Left:** 330, **Top:** 140
- **Width:** 320, **Height:** 25

#### 26. Borrowers ListBox
- **Name:** lstBorrowers
- **Left:** 330, **Top:** 180
- **Width:** 320, **Height:** 300
- **ColumnCount:** 1
- **BorderStyle:** 1

#### 27. Close Button
- **Name:** cmdClose
- **Caption:** Close
- **Left:** 540, **Top:** 530
- **Width:** 110, **Height:** 30
- **BackColor:** Gray

### Form Layout Diagram
```
┌─────────────────────────────────────────────────┐
│ Borrower Information                            │
│                          [➕ Add] [✏️ Update] [🗑️ Delete] │
│ Borrower ID: [_________]                        │
│ Last Name: [_________]   Search: [________]     │
│ First Name: [_________]  [🔍 Search]            │
│ Middle Name: [_________] │                      │
│ Address: [_________]     │  ┌────────────────┐ │
│           [_________]    │  │ BOR001 Smith.. │ │
│ Contact No: [_________]  │  │ BOR002 Jones.. │ │
│ Occupation: [_________]  │  │ BOR003 Garcia. │ │
│ Photo Path: [_________]  │  │                │ │
│ Status: [▼ Active]       │  │                │ │
│                          │  └────────────────┘ │
│                      [Close]                    │
└─────────────────────────────────────────────────┘
```

---

## FORM 4: frmLoan - Loan Management Form

### Purpose
Create, search, and manage loan records with automatic interest calculation.

### Form Properties
| Property | Value |
|----------|-------|
| Name | frmLoan |
| Caption | Loan Management |
| Width | 800 |
| Height | 650 |
| StartupPosition | 1 (Center Owner) |

### Controls to Add

#### Left Section - Loan Form

#### 1. Title Label
- **Name:** lblTitle
- **Caption:** Loan Information
- **Left:** 20, **Top:** 20
- **Width:** 300, **Height:** 25
- **Font:** Size 12, Bold

#### 2. Loan ID Label & TextBox
- **Label Name:** lblLoanID, **Caption:** Loan ID:, **Left:** 20, **Top:** 60
- **TextBox Name:** txtLoanID, **Left:** 110, **Top:** 60, **Width:** 200, **Height:** 20, **Locked:** True

#### 3. Borrower ID Label & ComboBox
- **Label Name:** lblBorrowerID, **Caption:** Borrower ID:, **Left:** 20, **Top:** 90
- **ComboBox Name:** cbxBorrowerID, **Left:** 110, **Top:** 90, **Width:** 200, **Height:** 20, **Style:** 2

#### 4. Loan Type Label & ComboBox
- **Label Name:** lblLoanType, **Caption:** Loan Type:, **Left:** 20, **Top:** 120
- **ComboBox Name:** cbxLoanType, **Left:** 110, **Top:** 120, **Width:** 200, **Height:** 20, **Style:** 2

#### 5. Principal Label & TextBox
- **Label Name:** lblPrincipal, **Caption:** Principal Amount:, **Left:** 20, **Top:** 150
- **TextBox Name:** txtPrincipal, **Left:** 110, **Top:** 150, **Width:** 200, **Height:** 20

#### 6. Interest Rate Label & TextBox
- **Label Name:** lblInterestRate, **Caption:** Interest Rate (%):, **Left:** 20, **Top:** 180
- **TextBox Name:** txtInterestRate, **Left:** 110, **Top:** 180, **Width:** 200, **Height:** 20

#### 7. Interest Amount Label (Display)
- **Name:** lblInterestAmount
- **Caption:** Interest: PHP 0.00
- **Left:** 20, **Top:** 210
- **Width:** 290, **Height:** 20
- **ForeColor:** Blue
- **Font:** Bold

#### 8. Total Collectible Label (Display)
- **Name:** lblTotalCollectible
- **Caption:** Total: PHP 0.00
- **Left:** 20, **Top:** 240
- **Width:** 290, **Height:** 20
- **ForeColor:** Green
- **Font:** Bold

#### 9. Release Date Label & TextBox
- **Label Name:** lblReleaseDate, **Caption:** Release Date:, **Left:** 20, **Top:** 270
- **TextBox Name:** txtReleaseDate, **Left:** 110, **Top:** 270, **Width:** 200, **Height:** 20

#### 10. Due Date Label & TextBox
- **Label Name:** lblDueDate, **Caption:** Due Date:, **Left:** 20, **Top:** 300
- **TextBox Name:** txtDueDate, **Left:** 110, **Top:** 300, **Width:** 200, **Height:** 20

#### 11. Terms Label & TextBox
- **Label Name:** lblTerms, **Caption:** Terms (months):, **Left:** 20, **Top:** 330
- **TextBox Name:** txtTerms, **Left:** 110, **Top:** 330, **Width:** 200, **Height:** 20

#### 12. Daily Payment Label & TextBox
- **Label Name:** lblDailyPayment, **Caption:** Daily Payment:, **Left:** 20, **Top:** 360
- **TextBox Name:** txtDailyPayment, **Left:** 110, **Top:** 360, **Width:** 200, **Height:** 20

#### 13. Amount Paid Label (Display)
- **Name:** lblAmountPaid
- **Caption:** Paid: PHP 0.00
- **Left:** 20, **Top:** 390
- **Width:** 290, **Height:** 20
- **ForeColor:** Blue

#### 14. Balance Label (Display)
- **Name:** lblBalance
- **Caption:** Balance: PHP 0.00
- **Left:** 20, **Top:** 420
- **Width:** 290, **Height:** 20
- **ForeColor:** Red
- **Font:** Bold

#### 15. Collector ID Label & ComboBox
- **Label Name:** lblCollectorID, **Caption:** Collector ID:, **Left:** 20, **Top:** 450
- **ComboBox Name:** cbxCollectorID, **Left:** 110, **Top:** 450, **Width:** 200, **Height:** 20, **Style:** 2

#### 16. Status Label & ComboBox
- **Label Name:** lblStatus, **Caption:** Status:, **Left:** 20, **Top:** 480
- **ComboBox Name:** cbxStatus, **Left:** 110, **Top:** 480, **Width:** 200, **Height:** 20, **Style:** 2

#### Right Section - Buttons & List

#### 17. Add Button
- **Name:** cmdAdd
- **Caption:** ➕ Add Loan
- **Left:** 330, **Top:** 60
- **Width:** 120, **Height:** 35
- **BackColor:** Green
- **ForeColor:** White

#### 18. Search Label & TextBox
- **Label Name:** lblSearchLoanID, **Caption:** Search Loan ID:, **Left:** 330, **Top:** 110
- **TextBox Name:** txtSearchLoanID, **Left:** 470, **Top:** 110, **Width:** 200, **Height:** 20

#### 19. Search Button
- **Name:** cmdSearch
- **Caption:** 🔍 Search
- **Left:** 330, **Top:** 140
- **Width:** 340, **Height:** 25

#### 20. Loans ListBox
- **Name:** lstLoans
- **Left:** 330, **Top:** 180
- **Width:** 340, **Height:** 350
- **BorderStyle:** 1

#### 21. Close Button
- **Name:** cmdClose
- **Caption:** Close
- **Left:** 550, **Top:** 600
- **Width:** 120, **Height:** 30

### Form Layout Diagram
```
┌──────────────────────────────────────────────────────┐
│ Loan Information                    [➕ Add Loan]    │
│ Loan ID: [_________]                                 │
│ Borrower ID: [▼_______]  Search: [________] [🔍]   │
│ Loan Type: [▼_________]  │                          │
│ Principal: [_________]   │  ┌────────────────────┐ │
│ Interest Rate: [____%]   │  │ LN001 - BOR001     │ │
│ Interest: PHP 0.00       │  │ LN002 - BOR002     │ │
│ Total: PHP 0.00          │  │ LN003 - BOR003     │ │
│ Release Date: [___/___]  │  │                    │ │
│ Due Date: [___/___]      │  │                    │ │
│ Terms: [__] months       │  │                    │ │
│ Daily Payment: [____]    │  │                    │ │
│ Paid: PHP 0.00           │  │                    │ │
│ Balance: PHP 0.00        │  │                    │ │
│ Collector: [▼_________]  │  │                    │ │
│ Status: [▼ Active]       │  │                    │ │
│                          │  └────────────────────┘ │
│                              [Close]               │
└──────────────────────────────────────────────────────┘
```

---

## FORM 5: frmPayment - Payment Recording Form

### Purpose
Record loan payments and track payment history with automatic balance updates.

### Form Properties
| Property | Value |
|----------|-------|
| Name | frmPayment |
| Caption | Payment Recording |
| Width | 800 |
| Height | 650 |
| StartupPosition | 1 (Center Owner) |

### Controls to Add

#### Left Section - Payment Form

#### 1. Title Label
- **Name:** lblTitle
- **Caption:** Record Payment
- **Left:** 20, **Top:** 20
- **Width:** 300, **Height:** 25
- **Font:** Size 12, Bold

#### 2. Loan ID Label & TextBox
- **Label Name:** lblLoanID, **Caption:** Loan ID:, **Left:** 20, **Top:** 60
- **TextBox Name:** txtLoanID, **Left:** 110, **Top:** 60, **Width:** 150, **Height:** 20

#### 3. Search Button
- **Name:** cmdSearchLoan
- **Caption:** 🔍 Search
- **Left:** 270, **Top:** 60
- **Width:** 70, **Height:** 20

#### 4. Loan ID ComboBox
- **Name:** cbxLoanID
- **Left:** 20, **Top:** 90
- **Width:** 320, **Height:** 20
- **Style:** 2 (Dropdown List)

#### 5. Borrower ID Label (Display)
- **Name:** lblBorrowerID
- **Caption:** Borrower: --
- **Left:** 20, **Top:** 120
- **Width:** 320, **Height:** 20
- **Font:** Bold
- **ForeColor:** Blue

#### 6. Current Balance Label (Display)
- **Name:** lblBalance
- **Caption:** Current Balance: PHP 0.00
- **Left:** 20, **Top:** 150
- **Width:** 320, **Height:** 20
- **Font:** Bold
- **ForeColor:** Red

#### 7. Daily Payment Label (Display)
- **Name:** lblDailyPayment
- **Caption:** Daily Payment: PHP 0.00
- **Left:** 20, **Top:** 180
- **Width:** 320, **Height:** 20
- **ForeColor:** Green

#### 8. Payment Amount Label & TextBox
- **Label Name:** lblPaymentAmount, **Caption:** Payment Amount:, **Left:** 20, **Top:** 220
- **TextBox Name:** txtPaymentAmount, **Left:** 150, **Top:** 220, **Width:** 190, **Height:** 20

#### 9. Payment Date Label & TextBox
- **Label Name:** lblPaymentDate, **Caption:** Payment Date:, **Left:** 20, **Top:** 250
- **TextBox Name:** txtPaymentDate, **Left:** 150, **Top:** 250, **Width:** 190, **Height:** 20

#### 10. OR Number Label & TextBox
- **Label Name:** lblORNumber, **Caption:** OR Number:, **Left:** 20, **Top:** 280
- **TextBox Name:** txtORNumber, **Left:** 150, **Top:** 280, **Width:** 190, **Height:** 20

#### 11. Collector ID Label & ComboBox
- **Label Name:** lblCollectorID, **Caption:** Collector ID:, **Left:** 20, **Top:** 310
- **ComboBox Name:** cbxCollectorID, **Left:** 150, **Top:** 310, **Width:** 190, **Height:** 20, **Style:** 2

#### 12. Record Payment Button
- **Name:** cmdRecordPayment
- **Caption:** 💳 Record Payment
- **Left:** 110, **Top:** 360
- **Width:** 140, **Height:** 35
- **BackColor:** Green
- **ForeColor:** White

#### Right Section - Payment History

#### 13. Payment History Title Label
- **Name:** lblHistoryTitle
- **Caption:** Payment History
- **Left:** 380, **Top:** 20
- **Width:** 200, **Height:** 25
- **Font:** Size 11, Bold

#### 14. Payment History ListBox
- **Name:** lstPaymentHistory
- **Left:** 380, **Top:** 60
- **Width:** 380, **Height:** 480
- **BorderStyle:** 1

#### 15. Close Button
- **Name:** cmdClose
- **Caption:** Close
- **Left:** 640, **Top:** 600
- **Width:** 120, **Height:** 30

### Form Layout Diagram
```
┌────────────────────────────────────────────────────────────┐
│ Record Payment                    Payment History          │
│ Loan ID: [_______] [🔍]          ┌──────────────────────┐ │
│ [▼ Select Loan________]           │ PM001 - Loan: LN001  │ │
│ Borrower: --                      │ PHP 1,000.00         │ │
│ Current Balance: PHP 0.00         │ PM002 - Loan: LN001  │ │
│ Daily Payment: PHP 0.00           │ PHP 1,000.00         │ │
│                                   │ PM003 - Loan: LN002  │ │
│ Payment Amount: [_________]       │ PHP 2,000.00         │ │
│ Payment Date: [___/___/____]      │                      │ │
│ OR Number: [_________]            │                      │ │
│ Collector ID: [▼_________]        │                      │ │
│                                   │                      │ │
│    [💳 Record Payment]            │                      │ │
│                                   └──────────────────────┘ │
│                                           [Close]         │
└────────────────────────────────────────────────────────────┘
```

---

## FORM 6: frmCollector - Collector Management Form

### Purpose
Add and manage loan collectors/agents and view their assigned loans.

### Form Properties
| Property | Value |
|----------|-------|
| Name | frmCollector |
| Caption | Collector Management |
| Width | 750 |
| Height | 600 |
| StartupPosition | 1 (Center Owner) |

### Controls to Add

#### Left Section - Collector Form

#### 1. Title Label
- **Name:** lblTitle
- **Caption:** Collector Information
- **Left:** 20, **Top:** 20
- **Width:** 300, **Height:** 25
- **Font:** Size 12, Bold

#### 2. Collector ID Label & TextBox
- **Label Name:** lblCollectorID, **Caption:** Collector ID:, **Left:** 20, **Top:** 60
- **TextBox Name:** txtCollectorID, **Left:** 120, **Top:** 60, **Width:** 200, **Height:** 20, **Locked:** True

#### 3. Name Label & TextBox
- **Label Name:** lblName, **Caption:** Name:, **Left:** 20, **Top:** 90
- **TextBox Name:** txtName, **Left:** 120, **Top:** 90, **Width:** 200, **Height:** 20

#### 4. Contact No Label & TextBox
- **Label Name:** lblContactNo, **Caption:** Contact No:, **Left:** 20, **Top:** 120
- **TextBox Name:** txtContactNo, **Left:** 120, **Top:** 120, **Width:** 200, **Height:** 20

#### 5. Commission Label & TextBox
- **Label Name:** lblCommission, **Caption:** Commission (%):, **Left:** 20, **Top:** 150
- **TextBox Name:** txtCommission, **Left:** 120, **Top:** 150, **Width:** 200, **Height:** 20

#### 6. Add Button
- **Name:** cmdAdd
- **Caption:** ➕ Add
- **Left:** 120, **Top:** 190
- **Width:** 200, **Height:** 30
- **BackColor:** Green
- **ForeColor:** White

#### 7. View Assigned Loans Button
- **Name:** cmdViewAssignedLoans
- **Caption:** 📋 View Assigned Loans
- **Left:** 20, **Top:** 240
- **Width:** 300, **Height:** 30
- **BackColor:** Blue
- **ForeColor:** White

#### Right Section - Collectors List & Assigned Loans

#### 8. Collectors Title Label
- **Name:** lblCollectorsTitle
- **Caption:** Collectors List
- **Left:** 340, **Top:** 20
- **Width:** 200, **Height:** 25
- **Font:** Bold

#### 9. Collectors ListBox
- **Name:** lstCollectors
- **Left:** 340, **Top:** 60
- **Width:** 370, **Height:** 150
- **BorderStyle:** 1

#### 10. Assigned Loans Title Label
- **Name:** lblAssignedLoansTitle
- **Caption:** Assigned Loans
- **Left:** 340, **Top:** 220
- **Width:** 200, **Height:** 25
- **Font:** Bold

#### 11. Assigned Loans ListBox
- **Name:** lstAssignedLoans
- **Left:** 340, **Top:** 260
- **Width:** 370, **Height:** 280
- **BorderStyle:** 1

#### 12. Close Button
- **Name:** cmdClose
- **Caption:** Close
- **Left:** 590, **Top:** 555
- **Width:** 120, **Height:** 30

### Form Layout Diagram
```
┌───────────────────────────────────────────────────────┐
│ Collector Information               Collectors List   │
│ Collector ID: [_________]          ┌────────────────┐│
│ Name: [_________]                  │ COL001 - Smith ││
│ Contact No: [_________]            │ COL002 - Jones ││
│ Commission (%): [_________]        │ COL003 - Garcia││
│                                    │ COL004 - Brown ││
│ [➕ Add]                           └────────────────┘│
│ [📋 View Assigned Loans]                             │
│                                                      │
│                          Assigned Loans             │
│                         ┌────────────────┐          │
│                         │ LN001 - BOR001 │          │
│                         │ Balance: PHP... │          │
│                         │ LN002 - BOR002 │          │
│                         │ Balance: PHP... │          │
│                         │                │          │
│                         │                │          │
│                         │                │          │
│                         │                │          │
│                         └────────────────┘          │
│                                [Close]              │
└───────────────────────────────────────────────────────┘
```

---

## STEP-BY-STEP INSTRUCTIONS TO CREATE FORMS IN EXCEL VBA

### Creating a UserForm

1. **Open VBA Editor**
   - Press `Alt + F11`

2. **Create New UserForm**
   - Right-click on your project in Project Explorer
   - Select `Insert` → `UserForm`
   - A new UserForm will appear

3. **Rename the UserForm**
   - In Properties panel (bottom-left), change Name to your form name (e.g., `frmLogin`)

4. **Add Controls**
   - Use the Toolbox on the left to drag controls onto the form
   - Adjust size and position using the Properties panel
   - Set control Name, Caption, and other properties

5. **Add Event Code**
   - Double-click a button to open the code editor
   - Paste the corresponding code from `frmXXX_Code.bas` files

6. **Test the Form**
   - Press `F5` or click Run to test the form

### Example: Adding a TextBox

1. In Toolbox, click the TextBox icon
2. Click and drag on the form to create the textbox
3. In Properties panel:
   - Set `Name` to `txtUsername`
   - Set `Width` to 250 and `Height` to 20
   - Adjust `Left` and `Top` for positioning

---

## IMPORTANT NOTES

✅ **Control Naming Convention:**
- TextBox: `txt` prefix (txtUsername, txtPassword, etc.)
- ComboBox: `cbx` prefix (cbxStatus, cbxBorrowerID, etc.)
- ListBox: `lst` prefix (lstBorrowers, lstLoans, etc.)
- Label: `lbl` prefix (lblTitle, lblError, etc.)
- Button: `cmd` prefix (cmdAdd, cmdDelete, etc.)
- Frame: `fra` prefix (fraStatistics)

✅ **Before Testing:**
- Ensure all table names match exactly: tblBorrowers, tblLoans, etc.
- Check that modDatabaseFunctions and modTableSetup are in your project
- Add named ranges: `LoggedInUser` and `UserRole` in Sheet1 for login tracking

✅ **Color Codes:**
- Green buttons: Actions (Add, Record)
- Blue buttons: View/Search
- Red buttons: Delete/Logout
- Gray buttons: Close

✅ **Font Recommendations:**
- Use Calibri or Arial size 10-11 for normal text
- Bold for titles and headers
- Blue for information, Red for warnings, Green for success

---

## Next Steps

After creating all forms:
1. Copy each form code from the VBA/Forms folder
2. Paste into the corresponding UserForm code module
3. Test each form with sample data
4. Customize colors and fonts as needed
5. Proceed to Step 3: Sample Data Population

Good luck! 🎉
