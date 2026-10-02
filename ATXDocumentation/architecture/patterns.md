# Architectural Patterns

[Back to README](../README.md) | [System Overview](system-overview.md) | [Components](components.md) | [Dependencies](dependencies.md)

## Identified Patterns

### 1. MDI (Multiple Document Interface) — Application Shell
- `frmMain` serves as the MDI container with menu-driven navigation
- All functional forms are MDI children (`MDIChild = -1 'True`)
- Menu handlers in frmMain instantiate and show child forms
- Status bar in MDI parent provides cross-form messaging via `LogStatus()`

### 2. ADODC Data Binding — Master Data Forms
- Used by: `frmCustomers`, `frmProducts`, `frmProviders`
- Pattern: ADODC control bound to a SQL query → TextBox controls bound to ADODC fields via `DataSource`/`DataField` properties
- Navigation: ADODC provides built-in record navigation (first/prev/next/last)
- CRUD: Toolbar buttons toggle edit/add modes, ADODC handles commits

### 3. Manual SQL with Global Recordsets — Transaction Forms
- Used by: Order forms, inventory forms, user management
- Pattern: Build SQL string → call `ExecuteSql(sql)` → read from global `rs` recordset
- Two global recordsets (`rs`, `rs2`) allow nested queries
- No parameterized queries; all values concatenated into SQL strings

### 4. Form-Level State Machine — Edit Modes
- Master data forms maintain state via private Booleans: `NewMode`, `EditMode`, `CancellingMode`
- Toolbar button availability changes based on current state
- State transitions: Browse → Edit → Save/Cancel → Browse

### 5. Reusable Search Dialog
- `frmSearch` is a general-purpose search component
- Called via `SearchShow(Table, fieldToSearch, itemToSearch)` in modFunctions
- Dynamically loads column names from the target table
- Applies search criteria to the calling form's ADODC recordset

### 6. Module-Based Utility Library
- `modFunctions.bas` provides shared utilities (validation, combo loading, ID generation)
- Functions accessed globally without instantiation (VB6 module pattern)

## Identified Anti-Patterns

### 1. Global Mutable State
- **Location**: modMain.bas, modConnection.bas
- **Issue**: User credentials, connection, and recordsets are global variables shared across all forms
- **Risk**: Any form can corrupt shared state; no encapsulation; impossible to unit test

### 2. SQL Injection via String Concatenation
- **Location**: All forms performing SQL queries
- **Issue**: User input directly concatenated into SQL strings
- **Example**: `"SELECT * FROM Users WHERE username = '" & txtUserName.Text & "'"`
- **Risk**: Complete database compromise

### 3. Monolithic Form Code-Behind
- **Location**: All .frm files
- **Issue**: UI rendering, business logic, data access, and validation all in the same form code
- **Risk**: Cannot reuse business logic; changes to UI require touching data access code

### 4. Shared Recordset Contention
- **Location**: modConnection.bas (`rs`, `rs2`)
- **Issue**: Only two recordsets available for the entire application; nested queries must carefully manage which recordset to use
- **Risk**: One form's query can overwrite another form's data if forms interact concurrently

### 5. Hardcoded Configuration
- **Location**: modMain.bas (`Sub Main`)
- **Issue**: Database connection string and admin bypass are hardcoded in source
- **Risk**: Cannot change configuration without recompilation

### 6. Dead Code
- **Location**: OfrmAccountProfile.frm
- **Issue**: Entire form references tables (`tblAccountProfile`, `tblAccountPosition`) that don't exist in the database
- **Risk**: Confuses maintainers; increases code surface without value
