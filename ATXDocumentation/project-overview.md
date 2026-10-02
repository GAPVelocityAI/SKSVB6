# Project Overview — Salmon King Seafood (SKS)

[Back to README](README.md)

## Application Summary
**Salmon King Seafood (SKS)** is a desktop order processing and inventory management system built in Visual Basic 6.0. It manages customer orders, supplier stock receipts, product catalogs, inventory tracking, and user administration through an MDI (Multiple Document Interface) Windows application.

## Technology Stack
| Component | Technology | Version/Details |
|-----------|-----------|-----------------|
| Language | Visual Basic 6.0 | VB6 SP6 |
| Runtime | VB6 Runtime (MSVBVM60.DLL) | Windows-only |
| Database | SQLite 3 | Via ODBC driver |
| Data Access | ADO (ADODB) | msado25.tlb |
| UI Framework | VB6 Forms (MDI) | Native Windows Forms |
| Controls | Microsoft Common Controls | MSCOMCTL.OCX, MSCOMCT2.OCX |
| Grid Controls | MSFlexGrid, DataGrid | MSFLXGRD.OCX, MSDATGRD.OCX |
| Additional OCX | Common Dialog, Masked Edit, Chart, DataList, ADODC | COMDLG32.OCX, MSMASK32.OCX, MSCHRT20.OCX, MSDATLST.OCX, MSADODC.OCX |

## Project Structure
```
SKS.vbp                          # VB6 Project file
├── modMain.bas                  # Entry point (Sub Main), global variables
├── modConnection.bas            # Database connection management (ADODB)
├── modFunctions.bas             # Shared utility functions
├── frmMain.frm                  # MDI parent form (menu, status bar)
├── frmSplash.frm                # Splash screen (3-second timer)
├── frmLogin.frm                 # Login form (currently bypassed)
├── frmAbout.frm                 # About dialog
├── frmCustomers.frm             # Customer CRUD management
├── frmProducts.frm              # Product CRUD management
├── frmProviders.frm             # Supplier/Provider CRUD management
├── frmUsersManage.frm           # User administration
├── frmSearch.frm                # Reusable search dialog
├── frmOrderRequest.frm          # Create customer order
├── frmOrderReception.frm        # Create supplier stock order
├── frmAddProductTo.frm          # Assign products to customers/suppliers
├── frmAddStockManual.frm        # Manual inventory update
├── frmAdjustStockManual.frm     # Manual inventory adjustment
├── frmOrderAproval.frm          # Supplier order approval list
├── frmActionOrderReception.frm  # Approve/cancel supplier orders
├── frmRequestAproval.frm        # Customer order approval list
├── frmActionOrderRequest.frm    # Approve/cancel customer orders
├── OfrmAccountProfile.frm       # Legacy/unused account profile form
├── Orders.db                    # SQLite database
├── Media/                       # Images (BMP, PNG, JPG, ICO)
├── Requirements/                # Mobilize.NET WebMAP migration configs
├── dotnet-genai-samples/        # Separate .NET GenAI sample project
└── *.VBCSln                     # Mobilize.NET solution files
```

## Application Purpose
The system supports a seafood distribution business with these core capabilities:
1. **Customer Management** — Maintain customer records with contact and billing information
2. **Supplier Management** — Maintain supplier/provider records
3. **Product Catalog** — Manage product information including categories, pricing, and stock levels
4. **Order Processing** — Create and approve customer orders with tax/freight calculation
5. **Stock Procurement** — Create and approve supplier stock orders with inventory integration
6. **Inventory Management** — Manual stock updates and adjustments with full audit trail
7. **User Administration** — Manage user accounts with basic admin/non-admin access levels

## Key Business Entities
- **Customers**: Companies placing orders
- **Providers**: Suppliers providing stock
- **Products**: Items with pricing, categories, stock levels
- **Order Requests**: Customer orders (create → approve/cancel)
- **Order Receptions**: Supplier stock orders (create → approve/cancel → inventory update)
- **Stocks**: Physical inventory records
- **Users**: System users with access levels

## Migration Context
The project includes Mobilize.NET WebMAP migration solution files targeting:
- C# / .NET Framework 4.7 (WinForms)
- C# / .NET Core 3.1 (WinForms)
- VB.NET / .NET Framework 4.7 (WinForms)

A separate `dotnet-genai-samples/` directory contains an unrelated Amazon GenAI .NET sample project.
