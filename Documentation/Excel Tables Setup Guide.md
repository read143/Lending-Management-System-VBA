# Excel Tables Setup Guide

## Overview
This guide will help you create all database tables in your Excel workbook using structured table references.

---

## Method 1: Manual Setup (Step-by-Step)

### Prerequisites
- Microsoft Excel 2016 or later
- Your workbook: `Lending Management System.xlsm`

### Steps for Each Table

#### 1. Create tblBorrowers
1. Create a new worksheet named "Borrowers"
2. Add headers in row 1:
   - A1: BorrowerID
   - B1: LastName
   - C1: FirstName
   - D1: MiddleName
   - E1: Address
   - F1: ContactNo
   - G1: Occupation
   - H1: PhotoPath
   - I1: Status

3. Select the header row (A1:I1)
4. Go to **Data** → **From Text/CSV** → **Format as Table**
5. Name the table: `tblBorrowers`
6. Click **OK**

#### 2. Create tblLoans
1. Create a new worksheet named "Loans"
2. Add headers in row 1:
   - A1: LoanID
   - B1: BorrowerID
   - C1: LoanType
   - D1: Principal
   - E1: InterestRate
   - F1: InterestAmount
   - G1: TotalCollectible
   - H1: ReleaseDate
   - I1: DueDate
   - J1: Terms
   - K1: DailyPayment
   - L1: AmountPaid
   - M1: Balance
   - N1: CollectorID
   - O1: Status

3. Format as table and name: `tblLoans`

#### 3. Create tblPayments
1. Create a new worksheet named "Payments"
2. Add headers in row 1:
   - A1: PaymentID
   - B1: LoanID
   - C1: PaymentDate
   - D1: AmountPaid
   - E1: RemainingBalance
   - F1: ORNumber
   - G1: CollectorID

3. Format as table and name: `tblPayments`

#### 4. Create tblCollectors
1. Create a new worksheet named "Collectors"
2. Add headers in row 1:
   - A1: CollectorID
   - B1: Name
   - C1: ContactNo
   - D1: Commission

3. Format as table and name: `tblCollectors`

#### 5. Create tblUsers
1. Create a new worksheet named "Users"
2. Add headers in row 1:
   - A1: Username
   - B1: Password
   - C1: Role

3. Format as table and name: `tblUsers`

---

## Method 2: Automated Setup (VBA Script)

### How to Use:
1. Open your workbook in Excel
2. Press **Alt + F11** to open the VBA Editor
3. Insert a new module (Insert → Module)
4. Copy the code from `modTableSetup.bas` into the module
5. Run the macro: `CreateAllTables()`
6. All tables will be created automatically

### Benefits:
- ✅ Creates all tables at once
- ✅ Sets proper formatting
- ✅ Configures data types
- ✅ Adds validation rules
- ✅ No manual steps needed

---

## Setting Column Data Types

After creating tables, set these formats for better data validation:

### tblBorrowers
- ContactNo: Text
- Status: Text (validation list: Active, Inactive, Deceased, Blacklisted)

### tblLoans
- Principal: Currency
- InterestRate: Percentage
- InterestAmount: Currency (formula: =Principal*InterestRate)
- TotalCollectible: Currency (formula: =Principal+InterestAmount)
- ReleaseDate: Date
- DueDate: Date
- Terms: Number
- DailyPayment: Currency
- AmountPaid: Currency
- Balance: Currency (formula: =TotalCollectible-AmountPaid)
- Status: Text (validation list: Active, Completed, Defaulted, Cancelled, Closed)

### tblPayments
- PaymentDate: Date
- AmountPaid: Currency
- RemainingBalance: Currency

### tblCollectors
- Commission: Currency

---

## Data Validation Rules

Add these validation rules to ensure data integrity:

### Status Fields (Dropdown Lists)
1. Select the Status column in tblBorrowers
2. Go to **Data** → **Data Validation**
3. Set:
   - Allow: List
   - Source: Active, Inactive, Deceased, Blacklisted
   - Show Error Alert: ✓ Checked

### Similar process for tblLoans Status column with:
- Active, Completed, Defaulted, Cancelled, Closed

---

## Next Steps

After creating tables:
1. ✅ Create VBA functions to interact with tables
2. ✅ Set up form bindings
3. ✅ Add sample data
4. ✅ Create validation procedures
