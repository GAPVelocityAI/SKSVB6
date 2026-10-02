# Remediation Plan

[Back to README](../README.md) | [Summary](summary.md) | [Outdated Components](outdated-components.md) | [Maintenance Burden](maintenance-burden.md)

## Prioritized Remediation Actions

### Priority 1: Migrate from VB6 to Modern .NET — Severity: High

**Rationale**: The VB6 runtime, language, and all ActiveX dependencies are end-of-life. No incremental fix can address the fundamental platform obsolescence. Full migration is the only viable path.

**Recommended Approach**:
1. Use existing Mobilize.NET WebMAP solution files as a starting point (C#/.NET or VB.NET targets)
2. Target **.NET 8.0+** with WinForms for desktop compatibility
3. After .NET migration, consider `AWS/dotnet-modernization` or `AWS/dotnet-lambda-upgrade` for cloud deployment

**Key Decisions**:
- **Target Language**: C# (broader ecosystem, better tooling) over VB.NET
- **Target Framework**: .NET 8.0+ (LTS, cross-platform capable)
- **UI Framework**: WinForms (closest to VB6 forms) or consider web-based UI for future flexibility
- **Database**: Keep SQLite for desktop; consider PostgreSQL/DynamoDB for cloud deployment

### Priority 2: Eliminate SQL Injection — Severity: High

**Rationale**: Every database query is vulnerable. This must be addressed during migration.

**Actions**:
- Replace all string-concatenated SQL with parameterized queries
- Use an ORM (Entity Framework Core) or micro-ORM (Dapper) for data access
- Implement input validation at all form boundaries
- **Scope**: All 13+ files containing SQL queries

### Priority 3: Implement Proper Authentication — Severity: High

**Actions**:
- Remove hardcoded admin bypass in `Sub Main()`
- Implement password hashing (bcrypt/Argon2)
- Enable the login form
- Add session management and timeout
- Implement role-based access control beyond admin/non-admin

### Priority 4: Refactor Architecture — Severity: Medium

**Actions**:
- Eliminate global state (CurrentConnection, rs, rs2, user globals)
- Implement separation of concerns (UI / Business Logic / Data Access layers)
- Extract duplicated code from order/reception form pairs into shared services
- Implement dependency injection
- Add proper error handling and logging

### Priority 5: Encrypt and Secure Database — Severity: Medium

**Actions**:
- Encrypt SQLite database (SQLCipher) or migrate to a server database
- Add proper foreign key constraints
- Implement schema migration tooling
- Add audit logging at the database level

### Priority 6: Establish Development Practices — Severity: Low

**Actions**:
- Add automated test suite (unit tests for business logic, integration tests for data access)
- Set up CI/CD pipeline
- Implement code review process
- Add application logging framework
- Remove dead code (OfrmAccountProfile.frm)
- Fix naming inconsistencies

## Migration Sequence

See [Migration Component Order](../migration/component-order.md) for the recommended order of migrating components.
