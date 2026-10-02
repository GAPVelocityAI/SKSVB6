# Outdated Components

[Back to README](../README.md) | [Summary](summary.md) | [Maintenance Burden](maintenance-burden.md) | [Remediation Plan](remediation-plan.md)

## Runtime and Language — Severity: High

| Component | Current Version | EOL Date | Replacement |
|-----------|----------------|----------|-------------|
| **Visual Basic 6.0 IDE** | VB6 SP6 (1998) | April 8, 2008 | Visual Studio 2022+ with C# or VB.NET |
| **VB6 Runtime (MSVBVM60.DLL)** | 6.0 | Supported for runtime only on Windows | .NET Runtime 8.0+ |
| **32-bit Execution Model** | x86 only | N/A | x64 / ARM64 via .NET |

**Assessment**: VB6 IDE and development support ended in 2008. The VB6 runtime DLL continues to ship with Windows for backward compatibility, but Microsoft has stated no new features or enhancements will be made. The 32-bit execution model limits memory to ~2GB and is increasingly incompatible with modern deployment targets.

## ActiveX/COM Dependencies — Severity: Medium

All OCX/DLL dependencies below are legacy 32-bit COM components with no active development:

| Component | File | Last Updated | Modern Equivalent |
|-----------|------|-------------|-------------------|
| Microsoft Common Controls 6.0 | MSCOMCTL.OCX | ~2012 (security patch) | WinForms ToolStrip, ListView, TreeView |
| Microsoft Common Controls-2 6.0 | MSCOMCT2.OCX | ~2012 | WinForms DateTimePicker, NumericUpDown |
| Microsoft Common Dialog | COMDLG32.OCX | ~2009 | .NET OpenFileDialog, SaveFileDialog |
| Microsoft Masked Edit | MSMASK32.OCX | ~2008 | WinForms MaskedTextBox |
| Microsoft ADO Data Control 6.0 | MSADODC.OCX | ~2008 | .NET BindingSource + DataAdapter |
| Microsoft DataList Controls 6.0 | MSDATLST.OCX | ~2008 | WinForms ComboBox, ListBox with data binding |
| Microsoft DataGrid Control 6.0 | MSDATGRD.OCX | ~2008 | WinForms DataGridView |
| Microsoft FlexGrid Control 6.0 | MSFLXGRD.OCX | ~2008 | WinForms DataGridView |
| Microsoft Chart Control 6.0 | MSCHRT20.OCX | ~2008 | .NET Chart control or third-party |
| Microsoft Data Binding Collection | MSBIND.DLL | ~2008 | .NET data binding |
| Microsoft Data Formatting | MSSTDFMT.DLL | ~2008 | .NET formatting |

## Data Access — Severity: Medium

| Component | Current | Replacement |
|-----------|---------|-------------|
| **ADODB 2.5** (msado25.tlb) | ADO 2.5 via COM | ADO.NET / Entity Framework / Dapper |
| **SQLite3 ODBC Driver** | ODBC connectivity | Microsoft.Data.Sqlite / System.Data.SQLite |

## Database — Severity: Medium

| Issue | Current State | Recommendation |
|-------|--------------|----------------|
| **No Encryption** | Plain SQLite file | Use SQLCipher or encrypted connection |
| **No Schema Migrations** | Manual schema | Use migration framework (EF Migrations, Flyway) |
| **No Constraints Enforcement** | Loose schema | Add foreign keys, NOT NULL, CHECK constraints |

## Build/Development Tooling — Severity: Low

| Component | Current | Status |
|-----------|---------|--------|
| VB6 IDE | Visual Basic 6.0 SP6 | EOL since 2008; requires legacy Windows for installation |
| Project Format | .vbp | No modern tooling support |
| No Build Automation | Manual compilation | No CI/CD pipeline possible with VB6 |
| No Package Manager | Manual OCX registration | No dependency management |
