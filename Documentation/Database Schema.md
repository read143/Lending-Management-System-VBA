# Lending Management System - Database Schema

## Overview
All data is stored in Excel Tables within the workbook for better structure, validation, and VBA integration using structured references.

---

## Table: tblBorrowers
**Purpose**: Store borrower information

| Field | Type | Description | Key |
|-------|------|-------------|-----|
| BorrowerID | Text | Unique borrower identifier | Primary Key |
| LastName | Text | Borrower's last name | |
| FirstName | Text | Borrower's first name | |
| MiddleName | Text | Borrower's middle name (optional) | |
| Address | Text | Complete address | |
| ContactNo | Text | Phone/mobile number | |
| Occupation | Text | Borrower's occupation | |
| PhotoPath | Text | Path to borrower's photo file (optional) | |
| Status | Text | Active/Inactive | |

---

## Table: tblLoans
**Purpose**: Store loan transaction details

| Field | Type | Description | Key |
|-------|------|-------------|-----|
| LoanID | Text | Unique loan identifier | Primary Key |
| BorrowerID | Text | Reference to tblBorrowers | Foreign Key |
| LoanType | Text | Type of loan (Personal, Business, etc.) | |
| Principal | Currency | Initial loan amount | |
| InterestRate | Percentage | Annual interest rate | |
| InterestAmount | Currency | Calculated interest | |
| TotalCollectible | Currency | Principal + Interest | |
| ReleaseDate | Date | Date loan was released | |
| DueDate | Date | Final payment due date | |
| Terms | Number | Number of payment terms/months | |
| DailyPayment | Currency | Daily payment amount (if applicable) | |
| AmountPaid | Currency | Total amount paid to date | |
| Balance | Currency | Remaining balance | |
| CollectorID | Text | Reference to tblCollectors | Foreign Key |
| Status | Text | Active/Completed/Defaulted/Closed | |

---

## Table: tblPayments
**Purpose**: Record individual payment transactions

| Field | Type | Description | Key |
|-------|------|-------------|-----|
| PaymentID | Text | Unique payment identifier | Primary Key |
| LoanID | Text | Reference to tblLoans | Foreign Key |
| PaymentDate | Date | Date of payment | |
| AmountPaid | Currency | Payment amount | |
| RemainingBalance | Currency | Balance after this payment | |
| ORNumber | Text | Official Receipt number | |
| CollectorID | Text | Reference to tblCollectors | Foreign Key |

---

## Table: tblCollectors
**Purpose**: Store collector/agent information

| Field | Type | Description | Key |
|-------|------|-------------|-----|
| CollectorID | Text | Unique collector identifier | Primary Key |
| Name | Text | Full name of collector | |
| ContactNo | Text | Phone/mobile number | |
| Commission | Currency | Commission rate or amount | |

---

## Table: tblUsers
**Purpose**: Store user authentication and role information

| Field | Type | Description | Key |
|-------|------|-------------|-----|
| Username | Text | Unique username | Primary Key |
| Password | Text | Encrypted password (stored securely) | |
| Role | Text | User role (Admin, Collector, Viewer, etc.) | |

---

## Relationships & Constraints

### Foreign Key Relationships
- **tblLoans.BorrowerID** → **tblBorrowers.BorrowerID**
- **tblLoans.CollectorID** → **tblCollectors.CollectorID**
- **tblPayments.LoanID** → **tblLoans.LoanID**
- **tblPayments.CollectorID** → **tblCollectors.CollectorID**

### Calculated Fields
- **tblLoans.InterestAmount** = Principal × InterestRate
- **tblLoans.TotalCollectible** = Principal + InterestAmount
- **tblLoans.Balance** = TotalCollectible - AmountPaid
- **tblPayments.RemainingBalance** = Previous Balance - AmountPaid

### Status Values
- **tblBorrowers.Status**: Active, Inactive, Deceased, Blacklisted
- **tblLoans.Status**: Active, Completed, Defaulted, Cancelled, Closed
- **tblUsers.Role**: Admin, Collector, Manager, Viewer

---

## Notes
- All ID fields should be auto-generated with sequential numbering or GUIDs
- Dates should use Excel date format (mm/dd/yyyy or dd/mm/yyyy as per locale)
- Currency fields should use proper number formatting
- Password fields should be encrypted before storage
- Consider adding audit fields: DateCreated, DateModified, CreatedBy, ModifiedBy
