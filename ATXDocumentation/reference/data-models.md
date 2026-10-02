# Data Models

[Back to README](../README.md) | [Program Structure](program-structure.md) | [Interfaces](interfaces.md) | [API Reference](api-reference.md) | [Modules](modules.md)

## Database Schema (SQLite 3 — Orders.db)

### Users
| Column | Type | Notes |
|--------|------|-------|
| Username | TEXT | Primary identifier |
| Password | TEXT | Stored in plaintext |
| Fullname | TEXT | Display name |
| Level | TEXT | Access level (FK to Levels) |

### Levels
| Column | Type | Notes |
|--------|------|-------|
| Level | TEXT | Level name (e.g., "Administrator") |

### Customers
| Column | Type | Notes |
|--------|------|-------|
| CustomerID | TEXT | Primary key (generated via RcrdId) |
| CompanyName | TEXT | |
| CompanyOrDepartment | TEXT | |
| EmailAddress | TEXT | |
| BillingAddress | TEXT | |
| PostalCode | TEXT | |
| City | TEXT | |
| StateOrProvince | TEXT | |
| Country/Region | TEXT | Note: column name contains "/" |
| PhoneNumber | TEXT | |
| Extension | TEXT | |
| FaxNumber | TEXT | |
| ContactTitle | TEXT | |
| ContactFirstName | TEXT | |
| ContactLastName | TEXT | |
| Notes | TEXT | |

### Products
| Column | Type | Notes |
|--------|------|-------|
| ProductID | INTEGER | Primary key |
| ProductName | TEXT | |
| ProductDescription | TEXT | |
| SerialNumber | TEXT | |
| CategoryID | INTEGER | FK to Categories |
| UnitPrice | REAL | |
| QuantityPerUnit | TEXT | |
| Unit | TEXT | |
| UnitsInStock | REAL | Updated by stock operations |
| UnitsOnOrder | REAL | Updated by order operations |
| Discontinued | INTEGER | Boolean (0/1) |
| LeadTime | TEXT | |
| ReorderLevel | REAL | |
| SupplierID | INTEGER | |

### Categories
| Column | Type | Notes |
|--------|------|-------|
| CategoryID | INTEGER | Primary key |
| CategoryName | TEXT | |

### Providers
| Column | Type | Notes |
|--------|------|-------|
| ProviderID | TEXT | Primary key (generated via RcrdId) |
| ProviderName | TEXT | |
| EmailAddress | TEXT | |
| PostalCode | TEXT | |
| City | TEXT | |
| StateOrProvince | TEXT | |
| Country/Region | TEXT | |
| ContactTitle | TEXT | |
| ContactFirstName | TEXT | |
| ContactLastName | TEXT | |
| PhoneNumber | TEXT | |
| Extension | TEXT | |
| FaxNumber | TEXT | |
| Notes | TEXT | |
| PaymentTerms | TEXT | |

### OrderRequests
| Column | Type | Notes |
|--------|------|-------|
| OrderID | INTEGER | Auto-increment PK |
| CustomerID | TEXT | FK to Customers |
| EmployeeID | TEXT | FK to Users |
| OrderDate | TEXT | Date string |
| RequiredByDate | TEXT | |
| PromisedByDate | TEXT | |
| FreightCharge | REAL | |
| SalesTaxRate | REAL | |
| Status | TEXT | PENDING/APPROVED/CANCELLED |
| ChangedBy | TEXT | User who last changed status |
| ChangedDate | TEXT | |
| Notes | TEXT | |
| Ship* | TEXT | Shipping address fields |

### OrderRequestDetails
| Column | Type | Notes |
|--------|------|-------|
| OrderDetailID | INTEGER | Auto-increment PK |
| OrderID | INTEGER | FK to OrderRequests |
| ProductID | INTEGER | FK to Products |
| DateSold | TEXT | |
| Quantity | REAL | |
| UnitPrice | REAL | |
| SalePrice | REAL | |
| LineTotal | REAL | Quantity × SalePrice |
| Discount | REAL | |

### OrderReceptions
| Column | Type | Notes |
|--------|------|-------|
| OrderID | INTEGER | Auto-increment PK |
| ProviderId | TEXT | FK to Providers |
| ReceivedBy | TEXT | FK to Users |
| OrderDate | TEXT | |
| Notes | TEXT | |
| FreightCharge | REAL | |
| SalesTaxRate | REAL | |
| Status | TEXT | PENDING/APPROVED/CANCELLED |
| ChangedBy | TEXT | |
| ChangedDate | TEXT | |

### OrderReceptionDetails
| Column | Type | Notes |
|--------|------|-------|
| OrderID | INTEGER | FK to OrderReceptions |
| ProductID | INTEGER | FK to Products |
| DateSold | TEXT | |
| Quantity | REAL | |
| UnitPrice | REAL | |
| SalePrice | REAL | |
| SalesTax | REAL | |
| LineTotal | REAL | |

### Stocks
| Column | Type | Notes |
|--------|------|-------|
| StockID | INTEGER | Auto-increment PK |
| ProductID | INTEGER | FK to Products |
| Stock | REAL | Current quantity |
| InitialStock | REAL | Original quantity |
| DateStarted | TEXT | |
| DateModified | TEXT | |
| User | TEXT | |
| UnitPrice | REAL | |
| StockPrice | REAL | Total value |

### ManualStocks
| Column | Type | Notes |
|--------|------|-------|
| StockID | INTEGER | FK to Stocks |
| Quantity | REAL | |
| Price | REAL | |
| User | TEXT | |
| Date | TEXT | |
| Action | TEXT | "ADD" or "MOD" |

### StockLog
| Column | Type | Notes |
|--------|------|-------|
| User | TEXT | |
| Date | TEXT | |
| Quantity | REAL | |
| StockPrice | REAL | |
| ProductID | INTEGER | FK to Products |
| StockID | INTEGER | FK to Stocks |
| DocType | TEXT | Source document type |
| DocID | TEXT | Source document ID |

### ProductsByCustomer
| Column | Type | Notes |
|--------|------|-------|
| CustomerId | TEXT | FK to Customers |
| ProductId | INTEGER | FK to Products |

### ProductsByProvider
| Column | Type | Notes |
|--------|------|-------|
| ProviderId | TEXT | FK to Providers |
| ProductId | INTEGER | FK to Products |

## Entity Relationships

```
Users ──────────────────┐
  │                      │
  │ (EmployeeID)         │ (ReceivedBy)
  ▼                      ▼
OrderRequests ◄──── OrderReceptions
  │                      │
  │ (OrderID)            │ (OrderID)
  ▼                      ▼
OrderRequestDetails  OrderReceptionDetails
  │                      │
  │ (ProductID)          │ (ProductID)
  ▼                      ▼
Products ◄───────── Stocks
  │    │                 │
  │    │ (CategoryID)    │ (StockID)
  │    ▼                 ▼
  │  Categories      ManualStocks
  │                      │
  │                      ▼
  │                  StockLog
  │
  ├──── ProductsByCustomer ───► Customers
  └──── ProductsByProvider ───► Providers

Levels ──► Users.Level
```
