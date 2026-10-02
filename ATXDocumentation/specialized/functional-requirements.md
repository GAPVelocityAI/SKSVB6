# Functional Requirements

[Back to README](../README.md) | [Project Overview](../project-overview.md) | [Components](../architecture/components.md) | [Data Models](../reference/data-models.md)

## 1. Customer Order Processing (Order Requests)

### 1.1 Create Customer Order
**Source**: `frmOrderRequest.frm`

**Workflow**:
1. Search and select a customer (by company name, contact first/last name using LIKE filtering)
2. Alternatively open the Customers management form to browse/select
3. System displays selected customer's company name and contact
4. Set "Required by Date" and "Promised by Date" via date pickers (enabled only after customer selection)
5. Add products via the product assignment dialog (loads from `ProductsByCustomer` join table)
6. Enter quantity for each product line (only the quantity column is editable)
7. Set Sales Tax percentage and Freight Charge
8. System auto-calculates subtotal, tax, and total
9. Save order

**Business Rules**:
- A customer must be selected before saving or adding products (buttons disabled until selection)
- Date pickers are disabled until a customer is selected
- Quantity input accepts only digits 0-9 and decimal point
- Freight Charge and Sales Tax accept only digits 0-9
- If the user attempts to change the customer while editing, a confirmation prompt appears to cancel previous edits
- Unsaved data protection: on form close, prompts "Do you want to save the edited data?" (Yes/No/Cancel)

**Calculations**:
- Line Price = Quantity × Unit Price
- SubTotal = Sum of all line prices
- Total Tax = SubTotal × (Sales Tax / 100)
- Total = Freight Charge + SubTotal + Total Tax
- All monetary values formatted as `#,##0.00`

**Data Operations on Save**:
- Inserts into `OrderRequests` with Status = "REQUESTED", CustomerId, EmployeeId (current user), OrderDate (today), RequiredByDate, PromisedByDate, FreightCharge, SalesTaxRate
- Inserts into `OrderRequestDetails` for each product line with Quantity > 0
- Updates `Products.UnitsOnOrder` += Quantity for each product

### 1.2 Approve/Cancel Customer Orders
**Source**: `frmActionOrderRequest.frm`, `frmRequestAproval.frm`

**Status Transitions**:
| From | To | Allowed | Side Effects |
|------|----|---------|-------------|
| REQUESTED | APPROVED | Yes | Records ChangedBy + ChangedDate |
| REQUESTED | CANCELLED | Yes (with confirmation) | Records ChangedBy + ChangedDate |
| APPROVED | CANCELLED | No | Blocked: "Order was already approved" |
| CANCELLED | APPROVED | No | Blocked: "Order was already cancelled" |

**Search/Filter Criteria**:
- Customer company name, contact first/last name (LIKE)
- Order ID (exact match), Product ID (LIKE)
- Date range (From/To, optional)
- Status filter (All or specific status)
- Results grouped by order showing: Date, OrderID, CompanyName, Contact, Received by, Sum(LineTotal), Status

---

## 2. Supplier Stock Procurement (Order Receptions)

### 2.1 Create Supplier Stock Order
**Source**: `frmOrderReception.frm`

**Workflow**:
1. Search and select a supplier (by provider name, contact first/last name using LIKE filtering)
2. Alternatively open the Providers management form to browse/select
3. System displays selected supplier's name and contact
4. Add notes (free text, multiline)
5. Add products via the product assignment dialog (loads from `ProductsByProvider` join table)
6. Enter quantity for each product line
7. Set Sales Tax percentage and Freight Charge
8. System auto-calculates subtotal, tax, and total
9. Save order

**Business Rules**: Same validation as customer orders (supplier required, numeric-only inputs, unsaved data protection, editing conflict handling).

**Calculations**: Identical formulas to customer orders (Line Price, SubTotal, Total Tax, Total).

**Data Operations on Save**:
- Inserts into `OrderReceptions` with Status = "RECEIVED", ProviderId, ReceivedBy (current user), OrderDate (today), Notes, FreightCharge, SalesTaxRate
- Inserts into `OrderReceptionDetails` for each product line with Quantity > 0
- Note: `UnitsOnOrder` update is commented out (unlike customer orders)

### 2.2 Approve/Cancel Supplier Orders
**Source**: `frmActionOrderReception.frm`, `frmOrderAproval.frm`

**Status Transitions**:
| From | To | Allowed | Side Effects |
|------|----|---------|-------------|
| RECEIVED | APPROVED | Yes | Creates Stock entries + updates UnitsInStock |
| RECEIVED | CANCELLED | Yes (with confirmation) | Records ChangedBy + ChangedDate |
| APPROVED | CANCELLED | No | Blocked |
| CANCELLED | APPROVED | No | Blocked |

**On Approval — Inventory Integration**:
1. Updates `OrderReceptions` status to "APPROVED"
2. Creates `Stocks` records for each product in the order (ProductID, Stock, InitialStock, dates, UnitPrice, StockPrice)
3. Creates `StockLog` entries for each stock record (DocType, ProductID, Quantity, StockID)
4. Updates `Products.UnitsInStock` += Sum(Quantity) for each product

### 2.3 Key Differences: Customer Orders vs. Supplier Orders

| Aspect | Customer Order | Supplier Order |
|--------|---------------|----------------|
| Business entity | Customer (sale) | Supplier (purchase) |
| Initial status | REQUESTED | RECEIVED |
| On creation | Updates UnitsOnOrder | No stock update |
| On approval | Status change only | Creates Stock records + updates UnitsInStock |
| Date fields | RequiredByDate, PromisedByDate | Notes field instead |
| Product source | ProductsByCustomer | ProductsByProvider |

---

## 3. Inventory Management

### 3.1 Manual Stock Addition
**Source**: `frmAddStockManual.frm`

**Workflow**:
1. Search product by code and/or name (LIKE filtering)
2. Select product from list (or open Products dialog)
3. System displays product name, unit, quantity per unit, and reference price
4. Enter: Quantity, Price (total stock price), Unit Price
5. System shows calculated stock quantity = Quantity × QuantityPerUnit + unit label
6. Save

**Business Rules**:
- All value fields are locked until a product is selected
- Only digits 0-9 allowed in value fields
- Default values on product selection: Quantity=1, Price=product UnitPrice, UnitPrice=product UnitPrice

**Calculations**:
- When Quantity changes: StockPrice = UnitPrice × Quantity
- When StockPrice changes: UnitPrice = StockPrice / Quantity
- When UnitPrice changes: StockPrice = UnitPrice × Quantity

**Data Operations on Save**:
1. INSERT into `Stocks` (ProductID, Stock, InitialStock, dates, User, UnitPrice, StockPrice)
2. INSERT into `ManualStocks` (StockID, Quantity, Price, User, Date, Action="ADD")
3. INSERT into `StockLog` (User, Date, Quantity, StockPrice, ProductID, StockID, DocType="MANUAL", DocID)
4. UPDATE `Products` SET UnitsInStock += Quantity

**Audit Trail**: Three-table logging — Stocks record, ManualStocks (Action="ADD"), StockLog (DocType="MANUAL")

### 3.2 Manual Stock Adjustment
**Source**: `frmAdjustStockManual.frm`

**Workflow**:
1. Search product by code and/or name
2. Select product from list
3. System loads existing stock entries from `Stocks` table
4. Select a specific stock entry
5. System shows original price and original quantity
6. Modify new price and/or new quantity
7. System shows both current and adjusted stock quantities
8. Save

**Business Rules**:
- **Quantity cap**: Stock quantity can only be decreased, never increased above the original. If the user enters a quantity above the original, it is capped to the original with message: "Cannot pass the original stock, to add more, add a new stock manually"
- **No-change detection**: If both price delta and stock delta are zero, save is skipped with message: "There is no modification of the Stock, the data won't be saved"

**Data Operations on Save**:
1. UPDATE `Stocks` SET StockPrice = newPrice, Stock = newQuantity
2. INSERT into `ManualStocks` (StockID, Quantity=deltaStock, Price=deltaStockPrice, Action="MOD")
3. INSERT into `StockLog` (Quantity=deltaStock, StockPrice=deltaStockPrice, DocType="MANUAL")
4. UPDATE `Products` SET UnitsInStock += deltaStock (delta is negative for reductions)

---

## 4. Master Data Management

### 4.1 Customer Management
**Source**: `frmCustomers.frm`

**Operations**: Add, Edit, Save, Delete, Search, Cancel (via Toolbar)

**Fields**: CompanyName, CompanyOrDepartment, EmailAddress, BillingAddress, PostalCode, City, StateOrProvince, Country/Region, PhoneNumber, Extension, FaxNumber, ContactTitle, ContactFirstName, ContactLastName, Notes

**Business Rules**:
- Delete requires confirmation: "Are you sure you want to delete this record?"
- Search opens the reusable search dialog (any field, LIKE matching)
- No field-level required validation on save
- On form close, sets `CurrentCustomerID` for calling forms to read

### 4.2 Product Management
**Source**: `frmProducts.frm`

**Operations**: Add, Edit, Save, Delete, Search, Cancel (via Toolbar)

**Fields**: ProductID (max 20 chars), ProductName (max 50 chars), ProductDescription (max 255 chars), SerialNumber, CategoryID (combo from Categories table), UnitPrice, QuantityPerUnit, Unit, Discontinued (checkbox)

**Business Rules**:
- On Add: defaults UnitsInStock=0, UnitsOnOrder=0, Discontinued=0
- Category loaded dynamically from `Categories` table
- On form close, sets `CurrentProductID` for calling forms

### 4.3 Supplier Management
**Source**: `frmProviders.frm`

**Operations**: Add, Edit, Save, Delete, Search, Cancel (via Toolbar)

**Fields**: ProviderName, EmailAddress, PostalCode, City, StateOrProvince, Country/Region, PhoneNumber, Extension, FaxNumber, ContactTitle, ContactFirstName, ContactLastName, Notes, PaymentTerms

**Business Rules — Required Fields on Save**:
- ProviderName, PaymentTerms, EmailAddress, PostalCode, City, StateOrProvince, Country/Region, FaxNumber, ContactTitle, ContactFirstName, Notes
- Validated using `TextBoxEmpty()` which shows "You need to fill in all required fields" on failure
- PaymentTerms is unique to suppliers (not present on customers)

### 4.4 Product-Entity Associations
**Source**: `frmAddProductTo.frm`

**Purpose**: Manages many-to-many relationships between Products and Customers or Suppliers.

**Business Rules**:
- Duplicate prevention: checks if product already exists in the entity's list before adding
- Smart delta tracking: if a stored product is removed then re-added, the operations cancel out
- On save, only executes INSERT/DELETE for actual net changes
- Shows "No data to be saved" if no changes are detected

---

## 5. User Administration

### 5.1 User Management
**Source**: `frmUsersManage.frm`

**Operations**: Add, Edit, Delete users (ListView-based, manual SQL)

**Fields**: Username, Password, Full name, User level

**Business Rules**:
- **All fields required**: Username, Password, Full name, and User level must be filled
- **Unique username**: On add, checks if username already exists. On edit with username change, checks new name is not taken
- **Administrator integrity**: The system must always have at least one Administrator account. If removing Administrator role would leave zero admins, the operation is blocked
- **Non-admin restriction**: Non-admin users cannot assign the "Administrator" level to any account
- **Last user protection**: The last remaining user cannot be deleted
- **System integrity check**: On form close, if current user is admin, verifies Users table is not empty

### 5.2 Authentication
**Source**: `frmLogin.frm`

**Workflow**: Username + password authentication against the `Users` table

**Business Rules**:
- On success: sets global state (UserFullname, UserLevel, CurrentUserAdmin)
- `CurrentUserAdmin = True` when UserLevel = "Administrator"
- On failure: "Invalid 'Username' or 'Password', please try again!"
- Note: Login is currently bypassed in `modMain.Main()` with hardcoded credentials for development

### 5.3 User Levels
- Levels loaded dynamically from `Levels` table
- Known elevated level: "Administrator"
- Global `CurrentUserAdmin` boolean controls admin-specific features

---

## 6. Cross-Cutting Functional Requirements

### 6.1 Unsaved Data Protection
All editing forms implement an `editingData` flag that triggers a Yes/No/Cancel prompt on form close:
- **Yes**: Save pending changes
- **No**: Discard changes
- **Cancel**: Abort the close action

Applies to: Order Request, Order Reception, Add Stock, Adjust Stock, Product Associations

### 6.2 Reusable Search
**Source**: `frmSearch.frm`

- Generic search dialog used by Customers, Products, and Providers forms
- Dynamically loads field names from the target table's schema
- Searches using LIKE with user-entered value
- Basic SQL injection mitigation: strips trailing single quotes from search input

### 6.3 Input Validation Patterns
- **Numeric-only fields**: KeyPress handlers allow only digits 0-9 plus navigation keys; invalid input triggers a beep
- **Grid editing**: Only the Quantity column is editable in product grids
- **Double validation**: Two-step via `ValidateTextBoxDouble` then calculation
- **Required field checks**: `TextBoxEmpty()` / `ComboEmpty()` with MsgBox prompt before save operations

### 6.4 Application Startup Flow
1. `Sub Main()` → Open database connection
2. Set user state (currently hardcoded, bypassing login)
3. Show `frmMain` (MDI parent)
4. Show `frmSplash` (modal, 3-second timer)
5. Auto-open `frmOrderRequest`

### 6.5 Menu Structure
- **File**: Manage Customers, Manage Suppliers, Exit
- **Orders**: Create Order, Create Invoice, Add Stock Order, Add Stock to Inventory
- **Inventory**: Inventory Update, Inventory Adjust
- **Maintenance**: Manage Products, Manage Users
- **Help**: About
