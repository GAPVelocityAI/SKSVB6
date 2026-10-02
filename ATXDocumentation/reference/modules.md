# Modules

[Back to README](../README.md) | [Program Structure](program-structure.md) | [Interfaces](interfaces.md) | [Data Models](data-models.md) | [API Reference](api-reference.md)

## Module Organization

The SKS application uses three VB6 standard modules (.bas files) and 20 form modules (.frm files). VB6 does not have a namespace or package system; all public symbols are globally accessible.

### Standard Modules

#### modMain.bas — Application Bootstrap
- **Role**: Entry point and global application state
- **Dependencies**: modConnection, frmMain, frmSplash
- **Exports**: Global variables (user state, connection string), LogStatus/ClearLogStatus subs
- **Critical Path**: `Sub Main()` → `OpenConnection()` → show `frmMain`

#### modConnection.bas — Data Access Layer
- **Role**: Database connection management and SQL execution
- **Dependencies**: modMain (ConnectionString)
- **Exports**: CurrentConnection, rs, rs2, OpenConnection, ExecuteSql, ExecuteSql2
- **Usage**: Every form that accesses the database calls ExecuteSql/ExecuteSql2

#### modFunctions.bas — Shared Utilities
- **Role**: Reusable helper functions for validation, UI population, and ID generation
- **Dependencies**: modConnection (ExecuteSql, rs)
- **Exports**: 16+ public functions (validation, combo loading, search, ID generation)
- **Usage**: Called by most forms for common operations

### Form Modules — Functional Grouping

#### Group 1: Application Shell
| Module | Dependencies | Notes |
|--------|-------------|-------|
| frmMain.frm | modMain | MDI parent; creates all child forms via menu |
| frmSplash.frm | — | Self-contained; Timer-based auto-close |
| frmLogin.frm | modMain, modConnection | Currently bypassed in Sub Main |
| frmAbout.frm | — | Self-contained dialog |

#### Group 2: Master Data (ADODC-based)
| Module | Dependencies | Notes |
|--------|-------------|-------|
| frmCustomers.frm | modMain, modFunctions | ADODC data binding; Toolbar CRUD |
| frmProducts.frm | modMain, modFunctions | ADODC + LoadCombo for categories |
| frmProviders.frm | modMain, modFunctions | ADODC data binding; Toolbar CRUD |

#### Group 3: User Management
| Module | Dependencies | Notes |
|--------|-------------|-------|
| frmUsersManage.frm | modMain, modConnection, modFunctions | ListView-based; manual SQL |

#### Group 4: Order Processing
| Module | Dependencies | Notes |
|--------|-------------|-------|
| frmOrderRequest.frm | modMain, modConnection, modFunctions | Customer orders; MSFlexGrid |
| frmOrderReception.frm | modMain, modConnection, modFunctions | Supplier orders; MSFlexGrid |
| frmRequestAproval.frm | modMain, modConnection | Customer order approval list |
| frmOrderAproval.frm | modMain, modConnection | Supplier order approval list |
| frmActionOrderRequest.frm | modConnection | Approve/cancel customer orders |
| frmActionOrderReception.frm | modConnection | Approve/cancel supplier orders + stock |

#### Group 5: Inventory
| Module | Dependencies | Notes |
|--------|-------------|-------|
| frmAddStockManual.frm | modMain, modConnection, modFunctions | Add stock entries |
| frmAdjustStockManual.frm | modMain, modConnection, modFunctions | Modify stock entries |

#### Group 6: Utilities
| Module | Dependencies | Notes |
|--------|-------------|-------|
| frmSearch.frm | modConnection | Dynamic column search |
| frmAddProductTo.frm | modConnection | Product-entity linking |

#### Group 7: Legacy
| Module | Dependencies | Notes |
|--------|-------------|-------|
| OfrmAccountProfile.frm | modConnection | Dead code; references missing tables |

## Dependency Graph

```
modMain ◄─── frmMain, frmLogin, frmCustomers, frmProducts, frmProviders,
             frmUsersManage, frmOrderRequest, frmOrderReception,
             frmAddStockManual, frmAdjustStockManual, frmOrderAproval,
             frmRequestAproval

modConnection ◄─── ALL forms except frmSplash, frmAbout

modFunctions ◄─── frmCustomers, frmProducts, frmProviders, frmUsersManage,
                   frmOrderRequest, frmOrderReception, frmAddStockManual,
                   frmAdjustStockManual
```
