# Security Patterns

[Back to README](../README.md) | [Code Metrics](code-metrics.md) | [Complexity Analysis](complexity-analysis.md) | [Dependency Analysis](dependency-analysis.md) | [Tech Debt](tech-debt.md)

## Security Overview

The SKS application has **critical security deficiencies** across authentication, authorization, data access, and data protection. The security posture is characteristic of a VB6 desktop application from an era when desktop apps were considered inherently trusted.

## Authentication — Severity: High

### Current Implementation
- **Login form** (`frmLogin.frm`): Queries `Users` table with username/password
- **Status**: Login is **bypassed** — `Sub Main()` hardcodes admin credentials:
  ```vb
  ' In modMain.bas Sub Main():
  CurrentUserAdmin = True
  UserFullname = "Admin"
  UserLevel = "Administrator"
  UserId = "admin"
  ```
- The login form exists but is never shown to the user

### Vulnerabilities
| Issue | Location | Severity |
|-------|----------|----------|
| Login bypassed entirely | modMain.bas, Sub Main | High |
| SQL injection in login query | frmLogin.frm | High |
| Plaintext password storage | Users table | High |
| Plaintext password comparison | frmLogin.frm | High |
| No brute-force protection | frmLogin.frm | Medium |
| No session timeout | Application-wide | Medium |
| No password complexity rules | frmUsersManage.frm | Medium |

### Login SQL (Vulnerable)
```vb
' frmLogin.frm - SQL injection vulnerability:
ExecuteSql "SELECT * FROM Users WHERE username = '" & txtUserName.Text & "' and password = '" & txtPassword.Text & "'"
```
An attacker could enter `' OR '1'='1` as username to bypass authentication.

## Authorization — Severity: Medium

### Current Implementation
- **Two-tier**: Admin vs. Non-admin (based on `CurrentUserAdmin` global)
- **Admin check**: Only in `frmUsersManage` (prevents non-admins from deleting admin-level users)
- **No form-level access control**: All menu items are visible to all users
- **No field-level access control**: All data fields are editable by all users

### Gaps
- No role-based access control (RBAC) beyond admin/non-admin
- No audit trail for who made changes (except stock operations)
- No data-level access restrictions (e.g., a user can see all customers)

## Data Access Security — Severity: High

### SQL Injection
**Every SQL query** in the application is constructed via string concatenation. Examples:

```vb
' frmUsersManage.frm:
ExecuteSql "SELECT * FROM Users WHERE Username='" & txtUsername & "'"

' frmOrderRequest.frm:
ExecuteSql "SELECT * FROM Customers WHERE CompanyName like '%" & txtCompanyName.Text & "%'"

' frmSearch.frm:
ExecuteSql2 "SELECT * FROM " & Table & " WHERE " & fieldToSearch & " like '%" & Replace(itemToSearch, "'", "") & "%'"
```

Note: `frmSearch.frm` has a minimal mitigation — it strips trailing single quotes from search input. This is insufficient protection.

**Affected forms**: All 13+ forms that execute SQL queries.

## Data Protection — Severity: High

| Asset | Protection | Risk |
|-------|-----------|------|
| User passwords | None (plaintext in SQLite) | Full credential exposure |
| Customer data | None (plaintext in SQLite) | Data breach via file access |
| Order data | None (plaintext in SQLite) | Financial data exposure |
| Database file | None (no encryption) | File copy = full data access |
| Connection string | Hardcoded in source | Reveals database path |

## Positive Security Patterns

| Pattern | Location | Notes |
|---------|----------|-------|
| "Last user" deletion protection | frmUsersManage.frm | Prevents deleting the last user account |
| Admin-level user protection | frmUsersManage.frm | Non-admins cannot modify admin accounts |
| Status-based action prevention | frmActionOrderReception.frm | Cannot re-approve or re-cancel orders |
| Wingdings password display | frmUsersManage.frm | Passwords displayed in Wingdings font (visual obscuring only) |
