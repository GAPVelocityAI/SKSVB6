# Program Structure

[Back to README](../README.md) | [Interfaces](interfaces.md) | [Data Models](data-models.md) | [API Reference](api-reference.md) | [Modules](modules.md)

## File Hierarchy

```
SKS.vbp (Project File - VB6)
│
├── Modules (.bas)
│   ├── modMain.bas          (~80 lines)  - Entry point, globals
│   ├── modConnection.bas    (~40 lines)  - Database connection
│   └── modFunctions.bas     (~250 lines) - Shared utilities
│
├── MDI Parent Form
│   └── frmMain.frm          (~120 lines) - Application shell
│
├── Dialog Forms
│   ├── frmSplash.frm        (~30 lines)  - Splash screen
│   ├── frmLogin.frm         (~60 lines)  - Authentication
│   ├── frmAbout.frm         (~40 lines)  - About dialog
│   └── frmSearch.frm        (~150 lines) - Reusable search
│
├── Master Data Forms
│   ├── frmCustomers.frm     (~300 lines) - Customer CRUD
│   ├── frmProducts.frm      (~350 lines) - Product CRUD
│   ├── frmProviders.frm     (~280 lines) - Supplier CRUD
│   └── frmUsersManage.frm   (~250 lines) - User management
│
├── Order Processing Forms
│   ├── frmOrderRequest.frm          (~600 lines) - Create customer order
│   ├── frmOrderReception.frm        (~550 lines) - Create supplier order
│   ├── frmRequestAproval.frm        (~350 lines) - Customer order approval list
│   ├── frmOrderAproval.frm          (~380 lines) - Supplier order approval list
│   ├── frmActionOrderRequest.frm    (~200 lines) - Approve/cancel customer order
│   └── frmActionOrderReception.frm  (~300 lines) - Approve/cancel supplier order
│
├── Inventory Forms
│   ├── frmAddStockManual.frm        (~300 lines) - Manual stock addition
│   └── frmAdjustStockManual.frm     (~350 lines) - Stock adjustment
│
├── Utility Forms
│   └── frmAddProductTo.frm          (~300 lines) - Product-entity assignment
│
├── Legacy Forms
│   └── OfrmAccountProfile.frm       (~200 lines) - Unused account profile
│
├── Binary Resources (.frx)
│   └── 13 .frx files (embedded images for forms)
│
├── Database
│   └── Orders.db (SQLite 3, ~108 KB)
│
├── Media/
│   └── BMP, PNG, JPG, ICO image assets
│
├── Migration Config
│   ├── SKS_WinForms.VBCSln          - Mobilize.NET C#/.NET Framework 4.7
│   ├── SKS_WinForms_NETCore.VBCSln  - Mobilize.NET C#/.NET Core 3.1
│   ├── SKS_Winforms_VBNet.VBCSln    - Mobilize.NET VB.NET/.NET Framework 4.7
│   └── Requirements/                - WebMAP migration options
│
└── External
    └── dotnet-genai-samples/         - Separate .NET GenAI project
```

## Global Variables

### modMain.bas
| Variable | Type | Scope | Purpose |
|----------|------|-------|---------|
| CurrentUserAdmin | Boolean | Public | Whether current user is admin |
| UserFullname | String | Public | Display name of logged-in user |
| UserLevel | String | Public | Access level of current user |
| UserId | String | Public | ID of current user |
| ConnectionString | String | Public | ODBC connection string |
| DetectionType | Integer | Public | Search detection type flag |
| n | Double | Public | General-purpose numeric temp |
| i | Long | Public | General-purpose counter |
| s | String | Public | General-purpose string temp |
| d | Date | Public | General-purpose date temp |
| msg | String | Public | General-purpose message |
| ImgName | String | Public | Image name temp |
| ImgSrc | String | Public | Image source temp |

### modConnection.bas
| Variable | Type | Scope | Purpose |
|----------|------|-------|---------|
| CurrentConnection | ADODB.Connection | Public | Single shared database connection |
| rs | ADODB.Recordset | Public | Primary shared recordset |
| rs2 | ADODB.Recordset | Public | Secondary shared recordset |
