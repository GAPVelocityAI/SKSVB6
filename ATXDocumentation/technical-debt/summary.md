# Technical Debt Summary

[Back to README](../README.md) | [Outdated Components](outdated-components.md) | [Maintenance Burden](maintenance-burden.md) | [Remediation Plan](remediation-plan.md)

## Overview

The SKS application carries substantial technical debt that spans every layer of the system. The VB6 runtime has been end-of-life for over 15 years, all ActiveX dependencies are deprecated, and the codebase contains critical security vulnerabilities that cannot be patched without a full rewrite.

## Findings by Severity

### High Severity (7 items)
1. **VB6 Runtime EOL** — Runtime end-of-life since 2008; no security patches available
2. **VB6 Language EOL** — Language/IDE end-of-life; no modern tooling support
3. **SQL Injection Vulnerabilities** — All database queries use string concatenation; every form is exploitable
4. **Plaintext Password Storage** — User passwords stored without hashing in the Users table
5. **Hardcoded Admin Bypass** — `Sub Main()` bypasses authentication, granting admin to all users
6. **No Input Sanitization** — User input flows directly into SQL and business logic without validation
7. **Unencrypted Database** — SQLite file has no encryption; data accessible by file copy

### Medium Severity (6 items)
1. **All ActiveX Controls Deprecated** — 13 COM/ActiveX dependencies (MSCOMCTL.OCX, etc.) are all deprecated, 32-bit only
2. **ADODB 2.5 Legacy** — ADO 2.5 is the data access layer; no modern equivalent in VB6
3. **Global Shared Recordsets** — Two shared recordsets (`rs`, `rs2`) serve the entire application
4. **No Separation of Concerns** — Business logic, data access, and UI are mixed in form code
5. **Tightly Coupled Forms** — Forms directly reference globals and each other's controls
6. **No Error Logging** — Error handling is inconsistent; no logging infrastructure

### Low Severity (5 items)
1. **Dead Code** — OfrmAccountProfile.frm references non-existent tables
2. **Inconsistent Error Handling** — Mix of `On Error Resume Next` and `On Error GoTo`; some handlers empty
3. **No Automated Tests** — Zero test coverage; no testing framework
4. **Hardcoded Connection String** — Database path fixed in source code
5. **Inconsistent Naming** — Mix of Hungarian notation and plain names; misspelled identifiers (e.g., "Aproval")
