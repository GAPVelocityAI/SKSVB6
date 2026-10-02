# API Reference

[Back to README](../README.md) | [Program Structure](program-structure.md) | [Interfaces](interfaces.md) | [Data Models](data-models.md) | [Modules](modules.md)

## Public Module Functions

### modMain.bas

| Function/Sub | Signature | Description |
|-------------|-----------|-------------|
| `Sub Main()` | `Public Sub Main()` | Application entry point. Sets connection string, opens DB, initializes user state, shows MDI parent |
| `LogStatus` | `Public Sub LogStatus(message As String, Optional frm As Form)` | Updates status bar text on specified form or all open forms |
| `ClearLogStatus` | `Public Sub ClearLogStatus(Optional frm As Form)` | Clears status bar on specified form or all open forms |

### modConnection.bas

| Function/Sub | Signature | Description |
|-------------|-----------|-------------|
| `OpenConnection` | `Public Sub OpenConnection()` | Creates new ADODB.Connection and opens with `ConnectionString` |
| `ExecuteSql` | `Public Sub ExecuteSql(Statement As String)` | Opens global `rs` recordset with given SQL (adOpenKeyset, adLockPessimistic) |
| `ExecuteSql2` | `Public Sub ExecuteSql2(Statement As String)` | Opens global `rs2` recordset with given SQL (same parameters as ExecuteSql) |

### modFunctions.bas

| Function/Sub | Signature | Description |
|-------------|-----------|-------------|
| `AppendAND` | `Public Sub AppendAND(ByRef filter As String)` | Appends " AND " to filter string if non-empty (SQL WHERE builder) |
| `AddToCollection` | `Public Function AddToCollection(col As Collection, Item As Variant) As Boolean` | Safely adds item to collection; returns True on success |
| `Exists` | `Public Function Exists(col As Collection, Index As Variant) As Boolean` | Checks if index exists in collection |
| `DoubleValue` | `Public Function DoubleValue(strValue As String) As Double` | Safe string-to-double conversion; returns 0 on error |
| `ValidateTextBoxDouble` | `Public Sub ValidateTextBoxDouble(txBox As TextBox, parentForm As Form)` | Validates textbox contains numeric value |
| `ValidateTextDouble` | `Public Sub ValidateTextDouble(text As String, parentForm As Form)` | Validates string is numeric |
| `SelectAll` | `Public Sub SelectAll(txtBox As TextBox)` | Selects all text in a textbox |
| `UpCase` | `Public Sub UpCase(KeyAscii As Integer)` | Converts keystroke to uppercase |
| `LoadCombo` | `Public Sub LoadCombo(Table As String, combo As ComboBox, field As String, Optional valueField As String)` | Populates combobox from database table |
| `ComboEmpty` | `Public Function ComboEmpty(combo As ComboBox, ...) As Boolean` | Validates combobox has selection; shows message if empty |
| `NoRecords` | `Public Function NoRecords(lstView As ListView, Optional Prompt As String) As Boolean` | Checks if ListView is empty; optionally shows prompt |
| `RcrdId` | `Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As Integer) As String` | Generates next record ID for a table (reads MAX + 1) |
| `SearchShow` | `Public Sub SearchShow(Table As String, fieldToSearch As String, itemToSearch As String)` | Opens frmSearch for specified entity type |
| `TextBoxEmpty` | `Public Function TextBoxEmpty(stext As String, ...) As Boolean` | Validates text is non-empty; shows message if empty |
| `TextBoxNumberEmpty` | `Public Function TextBoxNumberEmpty(textbox As TextBox) As Boolean` | Validates textbox contains a number |
| `ExecErr` | `Public Function ExecErr(Prompt As String, ...) As String` | Displays error message box |

## Form Event Handlers (Key Business Logic)

### frmOrderRequest — Customer Order Creation
| Handler | Trigger | Business Logic |
|---------|---------|---------------|
| `DoSearchCustomer` | Text change in search fields | Builds dynamic SQL filter on Customers table; populates ListView |
| `LoadProductsById` | Customer selection | Loads products assigned to customer via ProductsByCustomer JOIN |
| `SaveEdits` | Cell edit in grid | Updates quantity/price in grid; recalculates line totals |
| `ReCalculateTotals` | Any price/quantity change | Computes subtotal, tax, freight, grand total |
| `cmdSave_Click` | Save button | INSERT into OrderRequests + OrderRequestDetails; UPDATE Products.UnitsOnOrder |
| `cmdAddProducts_Click` | Add Products button | Opens frmAddProductTo for product assignment |

### frmActionOrderReception — Supplier Order Approval
| Handler | Trigger | Business Logic |
|---------|---------|---------------|
| `cmdApprove_Click` | Approve button | UPDATE OrderReceptions status='APPROVED'; INSERT Stocks from details; INSERT StockLog; UPDATE Products.UnitsInStock |
| `cmdCancel_Click` | Cancel button | UPDATE OrderReceptions status='CANCELLED' |

### frmAddStockManual — Manual Inventory Update
| Handler | Trigger | Business Logic |
|---------|---------|---------------|
| `cmdSave_Click` | Save button | INSERT Stocks; INSERT ManualStocks(Action='ADD'); INSERT StockLog; UPDATE Products.UnitsInStock |

### frmAdjustStockManual — Inventory Adjustment
| Handler | Trigger | Business Logic |
|---------|---------|---------------|
| `cmdSave_Click` | Save button | UPDATE Stocks; INSERT ManualStocks(Action='MOD'); INSERT StockLog; UPDATE Products.UnitsInStock with delta |
