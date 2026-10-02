# Components

[Back to README](../README.md) | [System Overview](system-overview.md) | [Dependencies](dependencies.md) | [Patterns](patterns.md)

## Component Overview

The application is organized into functional areas through VB6 forms (UI+logic) and shared modules (utilities).

### Shared Modules

| Module | File | Responsibility |
|--------|------|---------------|
| **modMain** | modMain.bas | Application entry point, global variables (user state, connection string), status bar management |
| **modConnection** | modConnection.bas | ADODB connection lifecycle, SQL execution wrappers (`ExecuteSql`, `ExecuteSql2`) |
| **modFunctions** | modFunctions.bas | Reusable utilities: validation, combo loading, ID generation, search, type conversion |

### UI Components — MDI Container

| Component | File | Type | Responsibility |
|-----------|------|------|---------------|
| **Main Window** | frmMain.frm | MDI Parent | Application shell, menu system, status bar |
| **Splash Screen** | frmSplash.frm | Form | 3-second splash on startup |
| **Login** | frmLogin.frm | Form | User authentication (currently bypassed) |
| **About** | frmAbout.frm | Dialog | Application version information |

### UI Components — Master Data Management

| Component | File | Type | Responsibility |
|-----------|------|------|---------------|
| **Customers** | frmCustomers.frm | MDI Child | Customer CRUD with ADODC data binding, 15 fields |
| **Products** | frmProducts.frm | MDI Child | Product CRUD with category combo, discontinued flag |
| **Providers** | frmProviders.frm | MDI Child | Supplier CRUD with ADODC data binding, 15 fields |
| **Users** | frmUsersManage.frm | MDI Child | User account CRUD with ListView, access levels |
| **Search** | frmSearch.frm | Dialog | Reusable entity search across Customers/Products/Providers |
| **Product Assignment** | frmAddProductTo.frm | Dialog | Many-to-many product-to-customer/provider linking |

### UI Components — Order Processing

| Component | File | Type | Responsibility |
|-----------|------|------|---------------|
| **Create Order** | frmOrderRequest.frm | MDI Child | Customer order creation with product grid, tax/freight |
| **Create Stock Order** | frmOrderReception.frm | MDI Child | Supplier stock order creation with product grid |
| **Order Approval List** | frmRequestAproval.frm | MDI Child | Customer order search/filter with approval actions |
| **Order Action** | frmActionOrderRequest.frm | MDI Child | Approve/cancel individual customer orders |
| **Reception Approval List** | frmOrderAproval.frm | MDI Child | Supplier order search/filter with approval actions |
| **Reception Action** | frmActionOrderReception.frm | MDI Child | Approve/cancel supplier orders, trigger stock creation |

### UI Components — Inventory Management

| Component | File | Type | Responsibility |
|-----------|------|------|---------------|
| **Inventory Update** | frmAddStockManual.frm | MDI Child | Add new stock entries manually |
| **Inventory Adjust** | frmAdjustStockManual.frm | MDI Child | Modify existing stock entries (price/quantity) |

### Legacy/Unused

| Component | File | Notes |
|-----------|------|-------|
| **Account Profile** | OfrmAccountProfile.frm | References non-existent tables (`tblAccountProfile`), appears from a different app version |

## Component Interactions

```
frmMain (MDI Parent)
 ├── frmOrderRequest ──────────> frmAddProductTo ──> frmSearch
 │                                                     │
 ├── frmOrderReception ────────> frmAddProductTo ──> frmSearch
 │
 ├── frmRequestAproval ────────> frmActionOrderRequest
 │
 ├── frmOrderAproval ──────────> frmActionOrderReception
 │
 ├── frmCustomers ─────────────> frmSearch
 ├── frmProducts ──────────────> frmSearch
 ├── frmProviders ─────────────> frmSearch
 │
 ├── frmAddStockManual
 ├── frmAdjustStockManual
 ├── frmUsersManage
 └── frmAbout
```

All components share:
- `modConnection.CurrentConnection` — Single ADODB connection
- `modConnection.rs` / `rs2` — Global shared recordsets
- `modMain.*` globals — User identity, connection string
- `modFunctions.*` — Validation and utility functions
