# UI Components

[Back to README](../README.md)

## UI Framework

The SKS application uses **VB6 native forms** with **ActiveX OCX controls**. The UI follows an MDI (Multiple Document Interface) pattern where `frmMain` is the MDI parent and all functional forms are MDI children.

## Control Types Used

| Control Type | Source | Usage |
|-------------|--------|-------|
| MDIForm | VB6 Native | frmMain (application shell) |
| Form | VB6 Native | All child forms and dialogs |
| TextBox | VB6 Native | Data entry fields |
| ComboBox | VB6 Native | Dropdown selections |
| CommandButton | VB6 Native | Action buttons |
| CheckBox | VB6 Native | Boolean flags (e.g., Discontinued) |
| Label | VB6 Native | Field labels |
| Frame | VB6 Native | Grouping containers |
| Timer | VB6 Native | frmSplash auto-close |
| Toolbar | MSCOMCTL.OCX | CRUD operation toolbars |
| StatusBar | MSCOMCTL.OCX | frmMain status display |
| ListView | MSCOMCTL.OCX | Data lists (users, products, customers, orders) |
| ImageList | MSCOMCTL.OCX | Toolbar button images |
| DTPicker | MSCOMCT2.OCX | Date selection (order dates) |
| MSFlexGrid | MSFLXGRD.OCX | Order line items grid (editable) |
| ADODC | MSADODC.OCX | Data-bound navigation (master data forms) |
| DataGrid | MSDATGRD.OCX | Data display grid |
| MaskedEdit | MSMASK32.OCX | Formatted input |
| CommonDialog | COMDLG32.OCX | File open dialog (OfrmAccountProfile) |

## Form Catalog

### frmMain — MDI Parent
- **Type**: VB.MDIForm
- **Caption**: "Sales Agent"
- **WindowState**: Maximized
- **Controls**: StatusBar (3 panels: message, time, date)
- **Menu Structure**:
  - File → Manage Customers, Manage Suppliers, Exit
  - Orders → Create Order, Create Invoice, Add Stock Order, Add Stock to Inventory
  - Inventory → Inventory Update, Inventory Adjust
  - Maintenance → Manage Products, Manage Users
  - Help → About

### Master Data Forms (ADODC-bound)
**Common Pattern**: ADODC control + Toolbar + control array of TextBoxes + ImageList

| Form | ADODC Source | Toolbar Actions | Field Count |
|------|-------------|----------------|-------------|
| frmCustomers | `Select * from Customers` | Add, Edit, Save, Delete, Search, Cancel | 15 (txtField(0..14)) |
| frmProducts | `Select * from Products` | Add, Edit, Save, Delete, Search, Cancel | 7 (txtField(0..6)) + cmbCategory + Check1 |
| frmProviders | `Select * from Providers` | Add, Edit, Save, Delete, Search, Cancel | 15 (txtField(0..14)) |

### Order Forms (MSFlexGrid-based)
**Common Pattern**: Entity search (ListView) + MSFlexGrid for line items + totals calculation

| Form | Entity | Grid Columns | Totals |
|------|--------|-------------|--------|
| frmOrderRequest | Customer | Product, Description, Unit, UnitPrice, Qty, SalePrice, LineTotal, Discount | SubTotal, SalesTax, FreightCharge, TotalTax, Total |
| frmOrderReception | Provider | Product, Description, Unit, UnitPrice, Qty, SalePrice, LineTotal | SubTotal, SalesTax, FreightCharge, Total |

### Approval Forms (MSFlexGrid listing)
**Common Pattern**: Search filters + MSFlexGrid order list + action buttons

| Form | Filters | Actions |
|------|---------|---------|
| frmOrderAproval | Provider name, Status, Date range, Order ID, Product ID | Approve, Cancel, Info |
| frmRequestAproval | Customer name, Status, Date range, Order ID, Product ID | Approve, Cancel, Info |

### Inventory Forms (ListView + TextBox)
| Form | Controls | Purpose |
|------|----------|---------|
| frmAddStockManual | Product search (ListView), txtValues(0..2), quantity/price fields | Add new stock |
| frmAdjustStockManual | Product search (ListView), Stock history (ListView), modification fields | Modify existing stock |

## Data Binding Pattern

ADODC-based forms bind controls declaratively:
```
Control: txtField(0)
  DataSource = dcCustomers
  DataField = "CompanyName"
```

This creates a two-way binding: changes to the textbox are automatically persisted when the ADODC recordset is updated (via `Recordset.Update` or navigation).

## In-Grid Editing Pattern (Order Forms)

The MSFlexGrid does not natively support editing. The pattern used:
1. User clicks a cell in the grid
2. A floating TextBox (`txtEntry`) is positioned over the cell
3. User types in the TextBox
4. On TextBox LostFocus or Enter, the value is written back to the grid
5. Line totals are recalculated
