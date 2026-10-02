# Maintenance Burden

[Back to README](../README.md) | [Summary](summary.md) | [Outdated Components](outdated-components.md) | [Remediation Plan](remediation-plan.md)

## High Maintenance Areas

### 1. Security Vulnerability Surface — Severity: High
- **Scope**: All 20 forms with database access
- **Issue**: SQL injection in every query requires reviewing and rewriting all SQL code
- **Burden**: Any new feature or bug fix touching SQL perpetuates the vulnerability pattern
- **Files**: frmLogin.frm, frmOrderRequest.frm, frmOrderReception.frm, frmUsersManage.frm, frmAddStockManual.frm, frmAdjustStockManual.frm, frmOrderAproval.frm, frmRequestAproval.frm, frmActionOrderReception.frm, frmActionOrderRequest.frm, frmAddProductTo.frm, frmSearch.frm, modFunctions.bas

### 2. Global State Management — Severity: Medium
- **Scope**: modMain.bas, modConnection.bas, all forms
- **Issue**: 15+ global variables shared across the entire application; changes to any global affect all forms
- **Burden**: Every form modification requires understanding the global state impact; debugging state corruption is extremely difficult
- **Files**: modMain.bas (globals), modConnection.bas (shared recordsets)

### 3. Duplicated Code Patterns — Severity: Medium
- **Scope**: Order/reception forms, approval forms
- **Issue**: frmOrderRequest and frmOrderReception contain near-identical code (~80% similarity). frmOrderAproval and frmRequestAproval are similarly duplicated. frmActionOrderReception and frmActionOrderRequest share the same structure.
- **Burden**: Bug fixes must be applied in multiple places; inconsistencies arise from partial fixes
- **Files**: frmOrderRequest.frm / frmOrderReception.frm, frmOrderAproval.frm / frmRequestAproval.frm, frmActionOrderReception.frm / frmActionOrderRequest.frm

### 4. ActiveX Control Dependencies — Severity: Medium
- **Scope**: All forms using OCX controls
- **Issue**: 13 ActiveX controls require registration on every target machine; 32-bit only; no longer updated
- **Burden**: Deployment requires manual registration; Windows updates can break OCX compatibility; finding replacement developers who understand these controls is difficult

### 5. Inconsistent Error Handling — Severity: Low
- **Scope**: All forms and modules
- **Issue**: Some procedures use `On Error Resume Next` (silently swallowing errors), others use `On Error GoTo` with varying handler quality, some have no error handling at all
- **Burden**: Debugging is difficult; errors may be silently ignored, causing data corruption

### 6. No Test Infrastructure — Severity: Low
- **Scope**: Entire codebase
- **Issue**: Zero automated tests; VB6 has no built-in testing framework
- **Burden**: Every change requires manual testing of all affected workflows; regression risk is high

## Maintenance Effort Distribution

| Area | Relative Effort | Primary Concern |
|------|----------------|-----------------|
| Security fixes | High | Cannot be done incrementally in VB6; requires rewrite |
| Bug fixes | Medium | Global state makes isolation difficult |
| New features | High | Must work around deprecated controls and global state |
| Deployment | Medium | Manual OCX registration, 32-bit compatibility |
| Developer onboarding | High | VB6 expertise is scarce; no documentation or tests |
