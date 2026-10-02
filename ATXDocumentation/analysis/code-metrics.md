# Code Metrics

[Back to README](../README.md) | [Complexity Analysis](complexity-analysis.md) | [Dependency Analysis](dependency-analysis.md) | [Security Patterns](security-patterns.md) | [Tech Debt](tech-debt.md)

## Summary Metrics

| Metric | Value |
|--------|-------|
| **Total Source Lines** | ~9,617 |
| **Source Files** | 24 (.frm, .bas, .vbp) |
| **Binary Resource Files** | 13 (.frx) |
| **Programming Language** | Visual Basic 6.0 |
| **Database Size** | ~108 KB (SQLite) |
| **Database Tables** | 15 |

## File-Level Metrics

| File | Type | Approx. Lines | Complexity |
|------|------|--------------|------------|
| frmOrderRequest.frm | Form (Order) | ~600 | High |
| frmOrderReception.frm | Form (Order) | ~550 | High |
| frmOrderAproval.frm | Form (Approval) | ~380 | Medium |
| frmRequestAproval.frm | Form (Approval) | ~350 | Medium |
| frmProducts.frm | Form (Master) | ~350 | Medium |
| frmAdjustStockManual.frm | Form (Inventory) | ~350 | Medium |
| frmCustomers.frm | Form (Master) | ~300 | Low |
| frmAddStockManual.frm | Form (Inventory) | ~300 | Medium |
| frmActionOrderReception.frm | Form (Action) | ~300 | Medium |
| frmAddProductTo.frm | Form (Utility) | ~300 | Medium |
| frmProviders.frm | Form (Master) | ~280 | Low |
| frmUsersManage.frm | Form (Admin) | ~250 | Medium |
| modFunctions.bas | Module | ~250 | Medium |
| frmActionOrderRequest.frm | Form (Action) | ~200 | Low |
| OfrmAccountProfile.frm | Form (Legacy) | ~200 | Low |
| frmSearch.frm | Form (Utility) | ~150 | Medium |
| frmMain.frm | MDI Parent | ~120 | Low |
| modMain.bas | Module | ~80 | Low |
| frmLogin.frm | Form (Auth) | ~60 | Low |
| frmAbout.frm | Form (Dialog) | ~40 | Low |
| modConnection.bas | Module | ~40 | Low |
| frmSplash.frm | Form (Dialog) | ~30 | Low |

## Quality Indicators

| Indicator | Assessment |
|-----------|-----------|
| **Test Coverage** | 0% — No automated tests exist |
| **Documentation** | Minimal — No inline documentation beyond occasional comments |
| **Error Handling** | Inconsistent — Mix of On Error Resume Next, On Error GoTo, and none |
| **Code Duplication** | High — Order/Reception form pairs share ~80% identical logic |
| **Naming Consistency** | Low — Mix of Hungarian notation, abbreviations, and misspellings |
| **Dead Code** | Present — OfrmAccountProfile.frm references non-existent tables |
| **Magic Numbers** | Present — Grid column indices, action codes used without constants |
| **Input Validation** | Minimal — Basic "is empty" checks only |
