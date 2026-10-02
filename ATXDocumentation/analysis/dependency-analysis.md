# Dependency Analysis

[Back to README](../README.md) | [Code Metrics](code-metrics.md) | [Complexity Analysis](complexity-analysis.md) | [Security Patterns](security-patterns.md) | [Tech Debt](tech-debt.md)

## External Dependency Analysis

### COM/ActiveX Libraries (from SKS.vbp)

| # | Library | Version | Criticality | Status | Risk |
|---|---------|---------|-------------|--------|------|
| 1 | stdole2.tlb (OLE Automation) | 2.0 | Low | Ships with Windows | Low |
| 2 | MSSTDFMT.DLL (Data Formatting) | 6.0 | Low | Deprecated | Medium |
| 3 | MSCOMCTL.OCX (Common Controls) | 6.0 | High | Deprecated, 32-bit | High |
| 4 | COMDLG32.OCX (Common Dialog) | 6.0 | Low | Deprecated, 32-bit | Medium |
| 5 | MSCOMCT2.OCX (Common Controls 2) | 6.0 | Medium | Deprecated, 32-bit | High |
| 6 | MSMASK32.OCX (Masked Edit) | 6.0 | Low | Deprecated, 32-bit | Medium |
| 7 | MSADODC.OCX (ADO Data Control) | 6.0 | High | Deprecated | High |
| 8 | msado25.tlb (ADO 2.5) | 2.5 | High | Legacy | High |
| 9 | MSDATLST.OCX (DataList) | 6.0 | Low | Deprecated | Medium |
| 10 | MSDATGRD.OCX (DataGrid) | 6.0 | Low | Deprecated | Medium |
| 11 | MSFLXGRD.OCX (FlexGrid) | 6.0 | High | Deprecated | High |
| 12 | MSCHRT20.OCX (Chart) | 6.0 | Low | Deprecated, unused | Low |
| 13 | MSBIND.DLL (Data Binding) | 6.0 | Medium | Deprecated | Medium |

### Runtime Dependencies

| Dependency | Criticality | Notes |
|-----------|-------------|-------|
| MSVBVM60.DLL (VB6 Runtime) | Critical | Must be present; ships with Windows |
| SQLite3 ODBC Driver | Critical | Must be separately installed |
| Windows OS (32-bit compat) | Critical | Required for OCX registration and execution |

## Internal Dependency Map

### Module Coupling

```
                    ┌──────────────┐
                    │   modMain    │
                    │  (globals)   │
                    └──────┬───────┘
                           │ used by 14 forms
                    ┌──────┴───────┐
                    │modConnection │
                    │ (rs, rs2)    │
                    └──────┬───────┘
                           │ used by 17 forms
                    ┌──────┴───────┐
                    │ modFunctions │
                    │ (utilities)  │
                    └──────┬───────┘
                           │ used by 9 forms
```

### Form-to-Form Dependencies

| Source Form | Target Form | Relationship |
|------------|-------------|-------------|
| frmMain | All MDI children | Creates via menu handlers |
| frmOrderRequest | frmAddProductTo | Opens for product assignment |
| frmOrderReception | frmAddProductTo | Opens for product assignment |
| frmOrderAproval | frmActionOrderReception | Opens for order action |
| frmRequestAproval | frmActionOrderRequest | Opens for order action |
| frmCustomers | frmSearch | Opens for customer search |
| frmProducts | frmSearch | Opens for product search |
| frmProviders | frmSearch | Opens for provider search |

## Dependency Criticality Assessment

| Category | Count | Risk Level |
|----------|-------|-----------|
| Critical (blocks execution) | 3 | VB6 Runtime, SQLite ODBC, Windows |
| High (core functionality) | 4 | MSCOMCTL, MSADODC, ADO 2.5, MSFLXGRD |
| Medium (feature support) | 5 | MSCOMCT2, MSBIND, MSDATGRD, MSDATLST, MSSTDFMT |
| Low (minimal usage) | 3 | COMDLG32, MSMASK32, MSCHRT20 |

## Transitive Dependencies
VB6 COM components do not have formal transitive dependencies in the npm/Maven sense. However, all OCX controls depend on:
- Windows COM infrastructure (ole32.dll, oleaut32.dll)
- VB6 runtime (MSVBVM60.DLL)
- Registry entries for CLSID registration
