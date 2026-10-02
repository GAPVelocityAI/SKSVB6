# Technical Debt Report — Salmon King Seafood (SKS)

[Back to README](README.md)

## 🎯 AWS Transformation Recommendation

### **RECOMMENDED TRANSFORMATIONS: None**

This is a Visual Basic 6.0 desktop application. No current AWS-managed transformation directly targets VB6-to-modern-language migration. The codebase includes Mobilize.NET WebMAP solution files for migration to C#/.NET Framework 4.7, C#/.NET Core 3.1, and VB.NET/.NET Framework 4.7. The recommended next steps are to first migrate from VB6 to .NET using the existing Mobilize.NET tooling or manual rewrite, then apply AWS-managed transformations such as `AWS/dotnet-modernization` or `AWS/dotnet-lambda-upgrade` to modernize the resulting .NET application for cloud deployment.

---

## Executive Summary

The SKS application carries **critical technical debt** across every dimension: the runtime (VB6) has been end-of-life since 2008, all external dependencies are legacy ActiveX controls with no modern support, the codebase has severe security vulnerabilities (SQL injection, plaintext passwords, hardcoded admin bypass), and the architecture uses anti-patterns (global state, no separation of concerns) that prevent incremental modernization.

**Total Technical Debt Items**: 18 findings across 3 severity levels

| Severity | Count | Description |
|----------|-------|-------------|
| **High** | 7 | EOL runtime, deprecated framework, critical security flaws |
| **Medium** | 6 | Outdated dependencies, architectural anti-patterns |
| **Low** | 5 | Code quality, maintainability, missing features |

## Critical Findings (Priority Order)

### 1. EOL Runtime and Framework — Severity: High
- **VB6 Runtime**: End-of-life since April 2008. No security patches, no modern OS guarantees
- **VB6 IDE**: Last released in 1998, unsupported since 2008
- **Impact**: Cannot receive security updates; increasingly incompatible with modern Windows versions
- See: [Outdated Components](technical-debt/outdated-components.md)

### 2. Security Vulnerabilities — Severity: High
- **SQL Injection**: Every SQL query uses string concatenation — all 20+ forms are vulnerable
- **Plaintext Passwords**: User passwords stored and compared without hashing
- **Hardcoded Admin Bypass**: `Sub Main()` bypasses login, granting admin access to all users
- **No Input Validation**: Minimal sanitization across the application
- See: [Security Patterns](analysis/security-patterns.md)

### 3. Outdated ActiveX Dependencies — Severity: Medium
- All UI controls (MSCOMCTL.OCX, MSFlexGrid, ADODC, etc.) are legacy COM components
- No vendor updates available; 32-bit only
- See: [Dependencies](architecture/dependencies.md)

### 4. Architectural Anti-Patterns — Severity: Medium
- Global shared ADODB recordsets (rs, rs2) create concurrency issues
- No separation between UI, business logic, and data access
- Tightly coupled forms with direct SQL execution
- See: [Patterns](architecture/patterns.md)

### 5. Code Quality — Severity: Low
- Unused legacy form (OfrmAccountProfile) referencing non-existent tables
- Inconsistent error handling (some `On Error Resume Next`, some `On Error GoTo`)
- No automated tests
- See: [Tech Debt](analysis/tech-debt.md)

## Navigation
- [Detailed Summary](technical-debt/summary.md)
- [Outdated Components](technical-debt/outdated-components.md)
- [Maintenance Burden](technical-debt/maintenance-burden.md)
- [Remediation Plan](technical-debt/remediation-plan.md)
