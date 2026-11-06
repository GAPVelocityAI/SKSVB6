# Salmon King Seafood (SKS)

Salmon King Seafood is a fictional seafood wholesaler. SKS is the Visual Basic 6 desktop application that runs its orders, inventory, and customer records.

This project is a reference application for legacy modernization. It packs the VB6 patterns that show up in real line-of-business systems: an MDI shell, ActiveX controls, ADO data binding, and a shared SQLite database. [GAP Velocity AI](https://www.gapvelocity.ai) uses applications like this one to demonstrate assessment and migration from VB6 to modern .NET, including C# WinForms and Blazor.

![Salmon King Seafood](./Media/NewLogo.png)

## Run the application

SKS is a VB6 executable. Build and run it on Windows with Visual Basic 6.

1. Install Visual Basic 6 and the ActiveX controls listed under [Dependencies](#dependencies).
2. Install the **SQLite3 ODBC Driver**. The application opens `Orders.db` with that driver.
3. Open `SKS.vbp`. The startup object is `Sub Main`.
4. Run the project from the `SKSVB6` folder so the relative database path `Orders.db` resolves.

`SKS.exe` in this folder is a prebuilt copy of the same application. Run it from `SKSVB6` as well, so it can find `Orders.db`.

On startup, `Sub Main` opens the database and signs in as Allan Cantillo, an administrator (`acantillo`). The shell then shows a splash screen for about three seconds and opens **Create Order**. A login form (`frmLogin`) is included and checks the `Users` table, but the startup path does not display it.

## What you can do

The main window is an MDI form titled **Sales Agent**. Every other screen opens from its menu.

| Menu | Command | What it does |
| --- | --- | --- |
| File | Manage Customers | Add, edit, and search customer records |
| File | Manage Suppliers | Maintain supplier contacts and payment terms |
| Orders | Create Order | Build a customer sales order with line items, tax, and freight |
| Orders | Create Invoice | Find an order and turn it into an invoice |
| Orders | Add Stock Order | Create a purchase order against a supplier |
| Orders | Add Stock to Inventory | Receive a stock order into inventory |
| Inventory | Inventory Update | Add quantity for a product |
| Inventory | Inventory Adjust | Correct on-hand quantity and price |
| Maintenance | Manage Products | Maintain the product catalog |
| Maintenance | Manage Users | Create users and assign an access level |
| Help | About | Application name, version, and www.gapvelocity.ai |

Customer, supplier, and product screens use a toolbar (Add, Edit, Save, Delete, Search, Cancel), control arrays, an ImageList, and an ADO Data Control. Text boxes are data-bound to `Orders.db`. Order and inventory screens use labels, text boxes, combo boxes, and an MSFlexGrid, and they write changes with ADO recordsets.

Supporting dialogs open from those screens: product search, add-product, and the order or receipt detail windows.

![Add Stock Order](./Media/SKSPhotos/frmOrderReception.PNG)

![Suppliers](./Media/SKSPhotos/frmProviders.PNG)

![Customers](./Media/SKSPhotos/frmCustomers.PNG)

![Add product](./Media/SKSPhotos/frmAddProductTo.PNG)

![Inventory update](./Media/SKSPhotos/frmAddStockManual.PNG)

![Receive stock](./Media/SKSPhotos/frmReceptionApproval.PNG)

![Create invoice](./Media/SKSPhotos/frmRequestApproval.PNG)

![Search](./Media/SKSPhotos/frmSearch.PNG)

## How it is built

| Piece | Role |
| --- | --- |
| `SKS.vbp` | VB6 project file. Startup is `Sub Main` in `modMain.bas`. |
| `modMain.bas` | Connection string, current user, and status-bar logging |
| `modConnection.bas` | Shared ADO connection and the two `ExecuteSql` helpers |
| `modFunctions.bas` | Shared UI helpers |
| `frmMain.frm` | MDI shell and menus |
| `Orders.db` | SQLite database. Keep it in the working directory. |

Data access uses ADO 2.5 through this ODBC connection string:

```text
DRIVER=SQLite3 ODBC Driver; Database=Orders.db; LongNames=0; Timeout=1000; NoTXN=0; SyncPragma=NORMAL; StepAPI=0;
```

### Dependencies

The project references these libraries. Paths inside `SKS.vbp` point at a specific machine, so VB6 may ask you to re-locate a control the first time you open the project.

- Microsoft ActiveX Data Objects 2.5
- Microsoft Data Formatting Object Library 6.0
- Microsoft Data Binding Collection (`MSBIND.DLL`)
- Windows Common Controls 6.0 (`MSCOMCTL.OCX`) — toolbar, status bar, ImageList, ListView
- Windows Common Controls-2 6.0 (`MSCOMCT2.OCX`)
- Common Dialog Control 6.0 (`COMDLG32.OCX`)
- Masked Edit Control 6.0 (`MSMASK32.OCX`)
- ADO Data Control 6.0 (`MSADODC.OCX`)
- DataList Controls 6.0 (`MSDATLST.OCX`)
- DataGrid Control 6.0 (`MSDATGRD.OCX`)
- FlexGrid Control 6.0 (`MSFLXGRD.OCX`)
- Chart Control 6.0 (`MSCHRT20.OCX`)
- SQLite3 ODBC Driver

### Database

`Orders.db` already contains sample customers, suppliers, products, stock, orders, and users. The same schema is copied under `Requirements/netFramework` and `Requirements/netCore` for migrated builds.

| Area | Tables |
| --- | --- |
| People | `Customers`, `Providers`, `Users`, `Levels` |
| Catalog | `Categories`, `Products`, `ProductsByCustomer`, `ProductsByProvider` |
| Sales | `OrderRequests`, `OrderRequestDetails`, `Invoices`, `InvoiceDetails` |
| Receiving | `OrderReceptions`, `OrderReceptionDetails` |
| Stock | `Stocks`, `StockLog`, `ManualStocks`, `Warehouses` |

## Project layout

```text
SKSVB6/
  SKS.vbp              VB6 project
  SKS.exe              Prebuilt executable
  *.frm / *.bas        Forms and modules
  Orders.db            SQLite database
  Media/               Logos, toolbar icons, and screen captures
  Requirements/        Database copies and .NET project stubs for a migrated build
```

## Modernizing this application

SKS is the kind of VB6 system [GAP Velocity AI](https://www.gapvelocity.ai) assesses and modernizes: desktop UI, COM controls, and a relational data model that need to keep their behavior on .NET. Typical target is C#, Blazor, and Azure.

The `Requirements` folder holds companion files for that path: `Orders.db` plus `App.config` for .NET Framework and .NET, and a .NET Framework project file (`SKS.vbproj` / `SKS.csproj`). `SKS_WinForms.VBCSln`, `SKS_WinForms_NETCore.VBCSln`, and `SKS_Winforms_VBNet.VBCSln` describe WinForms upgrade layouts for this project.
