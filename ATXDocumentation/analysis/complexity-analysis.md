# Complexity Analysis

[Back to README](../README.md) | [Code Metrics](code-metrics.md) | [Dependency Analysis](dependency-analysis.md) | [Security Patterns](security-patterns.md) | [Tech Debt](tech-debt.md)

## Complexity Hotspots

### 1. frmOrderRequest.frm — Highest Complexity
- **~600 lines** with complex business logic
- Dynamic customer search with multi-field filtering
- MSFlexGrid in-cell editing with event-driven price recalculation
- Multi-table INSERT transaction (OrderRequests + OrderRequestDetails + Products update)
- Tax, freight, and total calculation logic
- Product loading via JOIN with ProductsByCustomer
- **Cyclomatic Complexity**: High — multiple nested conditionals in DoSearchCustomer, SaveEdits, cmdSave_Click

### 2. frmOrderReception.frm — High Complexity
- **~550 lines** — near-duplicate of frmOrderRequest for suppliers
- Same grid editing, price calculation, and multi-table save pattern
- Provider search instead of customer search

### 3. frmActionOrderReception.frm — Medium-High Complexity
- **cmdApprove_Click** performs a multi-step transaction:
  1. Updates OrderReceptions status
  2. Reads all OrderReceptionDetails
  3. Creates Stocks records for each detail line
  4. Creates StockLog entries
  5. Updates Products.UnitsInStock with aggregate SUM
- Uses nested recordsets (rs + rs2) requiring careful ordering
- **Risk**: Partial failure leaves inconsistent state (no transaction management)

### 4. frmAdjustStockManual.frm — Medium Complexity
- Stock selection → modification workflow with delta tracking
- Calculates price/quantity differences for audit trail
- Updates three tables (Stocks, ManualStocks, StockLog) plus Products

### 5. frmOrderAproval.frm / frmRequestAproval.frm — Medium Complexity
- Complex multi-table JOIN queries with dynamic filters
- Date range filtering
- Status-based action enabling/disabling
- Grid population from joined resultsets

### 6. modFunctions.bas — Medium Complexity
- `RcrdId()` — ID generation with MAX+1 pattern (race condition risk)
- `LoadCombo()` — Dynamic table/field loading
- `SearchShow()` — Multi-entity search dispatcher

## Concurrency Risks

| Risk | Location | Description |
|------|----------|-------------|
| Shared Recordset Collision | modConnection.bas | `rs` and `rs2` are global; concurrent form operations can overwrite each other's results |
| No Transaction Management | All save operations | Multi-table writes have no rollback capability; partial failures corrupt data |
| Race Condition in ID Generation | modFunctions.RcrdId | MAX+1 ID generation is not atomic; concurrent inserts could generate duplicate IDs |

## Control Flow Complexity

### Order Creation Flow (frmOrderRequest.cmdSave_Click)
```
Validate fields (empty checks)
├── Fail → Show error, exit
└── Pass
    Generate OrderID via RcrdId
    Build INSERT SQL for OrderRequests
    ExecuteSql(INSERT)
    ├── Error → On Error Resume Next (silently continues!)
    └── Success
        Loop through FlexGrid rows
        ├── For each row with quantity > 0
        │   Build INSERT SQL for OrderRequestDetails
        │   ExecuteSql(INSERT)
        │   Build UPDATE SQL for Products.UnitsOnOrder
        │   ExecuteSql(UPDATE)
        └── End loop
    Show success message
    Clear form
```

### Stock Approval Flow (frmActionOrderReception.cmdApprove_Click)
```
Check current status != APPROVED
├── Already approved → Show error, exit
└── Not approved
    UPDATE OrderReceptions SET Status='APPROVED'
    SELECT all OrderReceptionDetails for this OrderID
    Loop through details (using rs)
    ├── For each detail line
    │   INSERT into Stocks (using rs2 for nested query)
    │   Get new StockID
    │   INSERT into StockLog
    └── End loop
    Calculate SUM(Stock) per ProductID
    UPDATE Products.UnitsInStock
    Refresh display
```
