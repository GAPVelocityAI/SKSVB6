# Database Patterns

[Back to README](../README.md)

## Database Overview

- **Engine**: SQLite 3
- **File**: `Orders.db` (~108 KB, 562 lines of schema)
- **Connectivity**: ODBC via "SQLite3 ODBC Driver"
- **Connection String**: `DRIVER=SQLite3 ODBC Driver;Database=<app_path>\Orders.db;`
- **Tables**: 15

## Connection Pattern

```vb
' modConnection.bas
Public CurrentConnection As New ADODB.Connection

Public Sub OpenConnection()
    Set CurrentConnection = New ADODB.Connection
    CurrentConnection.Open ConnectionString   ' from modMain.ConnectionString
End Sub
```

Single global connection opened at startup and shared by all forms. Never explicitly closed during application lifetime.

## Query Patterns

### Pattern 1: ADODC Data Binding (Master Data Forms)
Used by frmCustomers, frmProducts, frmProviders. SQL is set declaratively on the ADODC control:
```
dcCustomers.RecordSource = "Select * from Customers"
dcCustomers.ConnectionString = ConnectionString
```
Navigation, add, edit, delete performed via ADODC methods (`.Recordset.AddNew`, `.Recordset.Delete`, `.Recordset.Update`).

### Pattern 2: ExecuteSql with Global Recordset
Used by order forms, user management, inventory forms:
```vb
ExecuteSql "SELECT * FROM Products WHERE ProductID = " & productId
If Not rs.EOF Then
    productName = rs("ProductName")
End If
```

### Pattern 3: Nested Queries (rs + rs2)
When a query needs to be run inside a loop that's reading from another query:
```vb
ExecuteSql "SELECT * FROM OrderReceptionDetails WHERE OrderID = " & orderId
Do While Not rs.EOF
    ExecuteSql2 "INSERT INTO Stocks ..."   ' Uses rs2 to avoid overwriting rs
    rs.MoveNext
Loop
```

### Pattern 4: Dynamic Filter Building
```vb
Dim filter As String
filter = ""
If txtCompanyName.Text <> "" Then
    AppendAND filter
    filter = filter & "CompanyName like '%" & txtCompanyName.Text & "%'"
End If
' ... more conditions ...
ExecuteSql "SELECT * FROM Customers WHERE " & filter
```

## Notable SQL Queries

### Complex JOIN (frmOrderAproval — Order Approval List)
```sql
SELECT OrderReceptions.OrderID, Providers.ProviderName,
       OrderReceptions.OrderDate, Users.Fullname,
       SUM(OrderReceptionDetails.LineTotal) as Total,
       OrderReceptions.Status
FROM OrderReceptions
INNER JOIN OrderReceptionDetails ON OrderReceptions.OrderID = OrderReceptionDetails.OrderID
INNER JOIN Providers ON OrderReceptions.ProviderId = Providers.ProviderID
INNER JOIN Users ON OrderReceptions.ReceivedBy = Users.Username
WHERE [dynamic filters]
GROUP BY OrderReceptions.OrderID, Providers.ProviderName,
         OrderReceptions.OrderDate, Users.Fullname,
         OrderReceptions.Status
```

### ID Generation (modFunctions.RcrdId)
```sql
SELECT MAX(CAST(CustomerID AS INTEGER)) FROM Customers
-- Returns MAX + 1 as new ID
```

### Stock Balance Update (frmActionOrderReception)
```sql
UPDATE Products SET UnitsInStock = (SELECT SUM(Stock) FROM Stocks WHERE ProductID = X)
WHERE ProductID = X
```

## Data Integrity Concerns

| Issue | Description |
|-------|-------------|
| No foreign keys enforced | SQLite has FK support but it may not be enabled (`PRAGMA foreign_keys = ON`) |
| No transactions | Multi-table operations (order save, stock approval) have no BEGIN/COMMIT/ROLLBACK |
| Race condition in ID gen | MAX+1 is not atomic; concurrent access could produce duplicates |
| No cascading deletes | Deleting a customer does not clean up orders or product assignments |
| Date stored as text | Date values stored as formatted strings, not proper DATE/DATETIME |
