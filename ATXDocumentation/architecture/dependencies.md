# Dependencies

[Back to README](../README.md) | [System Overview](system-overview.md) | [Components](components.md) | [Patterns](patterns.md)

## External Dependencies (COM/ActiveX References)

All dependencies are declared in `SKS.vbp` (line references from project file):

| Library | File | Purpose | Status |
|---------|------|---------|--------|
| **OLE Automation** | stdole2.tlb | Standard OLE type library | Legacy, ships with Windows |
| **Microsoft Data Formatting** | MSSTDFMT.DLL | Data format objects | Deprecated |
| **Microsoft Common Controls 6.0** | MSCOMCTL.OCX | Toolbar, ListView, StatusBar, ImageList, TreeView | Deprecated, 32-bit only |
| **Microsoft Common Dialog** | COMDLG32.OCX | Open/Save dialogs, color/font picker | Deprecated, 32-bit only |
| **Microsoft Common Controls-2 6.0** | MSCOMCT2.OCX | DateTimePicker, MonthView, UpDown | Deprecated, 32-bit only |
| **Microsoft Masked Edit** | MSMASK32.OCX | Masked text input | Deprecated, 32-bit only |
| **Microsoft ADO Data Control 6.0** | MSADODC.OCX | ADODC data-bound control | Deprecated |
| **Microsoft ADO 2.5** | msado25.tlb | ActiveX Data Objects for database access | Legacy |
| **Microsoft DataList Controls 6.0** | MSDATLST.OCX | Data-bound list/combo | Deprecated |
| **Microsoft DataGrid Control 6.0** | MSDATGRD.OCX | Data-bound grid control | Deprecated |
| **Microsoft FlexGrid Control 6.0** | MSFLXGRD.OCX | Flexible grid for order line items | Deprecated |
| **Microsoft Chart Control 6.0** | MSCHRT20.OCX | Chart rendering (referenced but unused) | Deprecated |
| **Microsoft Data Binding Collection** | MSBIND.DLL | Data binding support | Deprecated |

## Runtime Dependencies

| Dependency | Purpose | Notes |
|-----------|---------|-------|
| **MSVBVM60.DLL** | VB6 Runtime | Required for execution |
| **SQLite3 ODBC Driver** | Database connectivity | Must be installed on target machine |
| **Orders.db** | SQLite database file | Must be in application directory |
| **Windows OS** | Host platform | 32-bit compatibility required |

## Internal Module Dependencies

```
modMain.bas
 ├── References: modConnection (OpenConnection)
 ├── References: frmMain, frmSplash (form loading)
 └── Defines: Global variables used by ALL forms

modConnection.bas
 ├── References: modMain.ConnectionString
 └── Defines: CurrentConnection, rs, rs2 (used by ALL forms)

modFunctions.bas
 ├── References: modConnection (ExecuteSql, rs)
 └── Defines: Utilities used by most forms
```

### Form-to-Module Dependencies

| Form | modMain | modConnection | modFunctions |
|------|---------|--------------|-------------|
| frmMain | UserFullname, LogStatus | — | — |
| frmLogin | UserFullname, UserLevel, UserId, CurrentUserAdmin | ExecuteSql, rs | — |
| frmCustomers | LogStatus | — | TextBoxEmpty, SearchShow |
| frmProducts | LogStatus | — | LoadCombo, TextBoxEmpty, SearchShow |
| frmProviders | LogStatus | — | TextBoxEmpty, SearchShow |
| frmUsersManage | LogStatus, CurrentUserAdmin | ExecuteSql, ExecuteSql2, rs, rs2 | TextBoxEmpty, LoadCombo |
| frmOrderRequest | LogStatus | ExecuteSql, ExecuteSql2, rs, rs2 | RcrdId, TextBoxEmpty, DoubleValue |
| frmOrderReception | LogStatus | ExecuteSql, ExecuteSql2, rs, rs2 | RcrdId, TextBoxEmpty, DoubleValue |
| frmAddProductTo | — | ExecuteSql, ExecuteSql2, rs, rs2 | — |
| frmAddStockManual | LogStatus | ExecuteSql, ExecuteSql2, rs, rs2 | TextBoxEmpty, DoubleValue, RcrdId |
| frmAdjustStockManual | LogStatus | ExecuteSql, ExecuteSql2, rs, rs2 | TextBoxEmpty, DoubleValue |
| frmOrderAproval | LogStatus | ExecuteSql, rs | — |
| frmActionOrderReception | — | ExecuteSql, ExecuteSql2, rs, rs2 | — |
| frmRequestAproval | LogStatus | ExecuteSql, rs | — |
| frmActionOrderRequest | — | ExecuteSql, ExecuteSql2, rs, rs2 | — |
| frmSearch | — | ExecuteSql, rs | — |

## Database Dependencies (Table Access by Form)

| Form | Tables Read | Tables Written |
|------|-------------|----------------|
| frmLogin | Users | — |
| frmCustomers | Customers | Customers |
| frmProducts | Products, Categories | Products |
| frmProviders | Providers | Providers |
| frmUsersManage | Users, Levels | Users |
| frmOrderRequest | Customers, Products, ProductsByCustomer | OrderRequests, OrderRequestDetails, Products |
| frmOrderReception | Providers, Products, ProductsByProvider | OrderReceptions, OrderReceptionDetails |
| frmAddProductTo | Products, ProductsByCustomer/ProductsByProvider | ProductsByCustomer/ProductsByProvider |
| frmAddStockManual | Products | Stocks, ManualStocks, StockLog, Products |
| frmAdjustStockManual | Products, Stocks | Stocks, ManualStocks, StockLog, Products |
| frmOrderAproval | OrderReceptions, OrderReceptionDetails, Providers, Users | — |
| frmActionOrderReception | OrderReceptions, OrderReceptionDetails | OrderReceptions, Stocks, StockLog, Products |
| frmRequestAproval | OrderRequests, OrderRequestDetails, Customers, Users | — |
| frmActionOrderRequest | OrderRequests, OrderRequestDetails | OrderRequests |
