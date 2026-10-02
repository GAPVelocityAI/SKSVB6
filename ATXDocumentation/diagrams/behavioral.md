# Behavioral Diagrams

[Back to README](../README.md) | [Structural](structural.md) | [Architecture](architecture.md)

## 1. Customer Order Creation — Sequence Diagram

```
User          frmOrderRequest    modConnection    Orders.db
 │                  │                 │               │
 │  Search Customer │                 │               │
 │─────────────────►│                 │               │
 │                  │  ExecuteSql     │               │
 │                  │  (SELECT FROM   │               │
 │                  │   Customers)    │               │
 │                  │────────────────►│  SQL Query    │
 │                  │                 │──────────────►│
 │                  │                 │◄──────────────│
 │                  │◄────────────────│               │
 │  Select Customer │                 │               │
 │─────────────────►│                 │               │
 │                  │  LoadProductsById               │
 │                  │  (SELECT FROM Products          │
 │                  │   JOIN ProductsByCustomer)       │
 │                  │────────────────►│──────────────►│
 │                  │◄────────────────│◄──────────────│
 │  Edit Quantities │                 │               │
 │─────────────────►│                 │               │
 │                  │ ReCalculateTotals               │
 │                  │ (subtotal,tax,  │               │
 │                  │  freight,total) │               │
 │  Click Save      │                 │               │
 │─────────────────►│                 │               │
 │                  │  Validate fields│               │
 │                  │  RcrdId("OrderRequests")        │
 │                  │────────────────►│──────────────►│
 │                  │◄────────────────│◄──────────────│
 │                  │  INSERT OrderRequests            │
 │                  │────────────────►│──────────────►│
 │                  │  Loop: INSERT OrderRequestDetails│
 │                  │────────────────►│──────────────►│
 │                  │  Loop: UPDATE Products           │
 │                  │────────────────►│──────────────►│
 │                  │◄────────────────│◄──────────────│
 │  Success Message │                 │               │
 │◄─────────────────│                 │               │
```

## 2. Supplier Stock Order Approval — Sequence Diagram

```
User         frmOrderAproval    frmActionOrderRcpt    modConnection    Orders.db
 │                 │                    │                   │              │
 │  Search Orders  │                    │                   │              │
 │────────────────►│                    │                   │              │
 │                 │  ExecuteSql        │                   │              │
 │                 │  (JOIN OrderReceptions,                │              │
 │                 │   Providers, Users)│                   │              │
 │                 │───────────────────────────────────────►│─────────────►│
 │                 │◄───────────────────────────────────────│◄─────────────│
 │  Click Approve  │                    │                   │              │
 │────────────────►│                    │                   │              │
 │                 │  Show(Action=1,    │                   │              │
 │                 │   OrderId=N)       │                   │              │
 │                 │───────────────────►│                   │              │
 │                 │                    │  Load order details│              │
 │                 │                    │──────────────────►│─────────────►│
 │                 │                    │◄──────────────────│◄─────────────│
 │  Confirm        │                    │                   │              │
 │─────────────────────────────────────►│                   │              │
 │                 │                    │  UPDATE status=    │              │
 │                 │                    │  'APPROVED'        │              │
 │                 │                    │──────────────────►│─────────────►│
 │                 │                    │  Loop: INSERT Stocks│             │
 │                 │                    │──────────────────►│─────────────►│
 │                 │                    │  Loop: INSERT StockLog            │
 │                 │                    │──────────────────►│─────────────►│
 │                 │                    │  UPDATE Products.  │              │
 │                 │                    │  UnitsInStock      │              │
 │                 │                    │──────────────────►│─────────────►│
 │  Done           │                    │                   │              │
 │◄─────────────────────────────────────│                   │              │
```

## 3. Manual Inventory Update — Activity Diagram

```
┌─────────────┐
│ Open Form   │
│ (frmAddStock│
│  Manual)    │
└──────┬──────┘
       ▼
┌─────────────┐
│ Search for  │
│ Product     │◄────────────────┐
└──────┬──────┘                 │
       ▼                        │
  ┌─────────┐    No             │
  │Product  │──────────────────►│
  │Found?   │                   │
  └────┬────┘                   │
       │ Yes                    │
       ▼                        │
┌─────────────┐                 │
│ Enter Stock │                 │
│ Quantity    │                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ Enter Unit  │                 │
│ Price       │                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ Auto-calc   │                 │
│ Stock Price │                 │
│ (Qty×Unit)  │                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ Click Save  │                 │
└──────┬──────┘                 │
       ▼                        │
  ┌─────────┐    No             │
  │Fields   │──────────────────►│
  │Valid?   │                   │
  └────┬────┘                   │
       │ Yes                    │
       ▼                        │
┌─────────────┐                 │
│ INSERT      │                 │
│ Stocks      │                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ INSERT      │                 │
│ ManualStocks│                 │
│ (Action=ADD)│                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ INSERT      │                 │
│ StockLog    │                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ UPDATE      │                 │
│ Products.   │                 │
│ UnitsInStock│                 │
└──────┬──────┘                 │
       ▼                        │
┌─────────────┐                 │
│ Success     │                 │
│ Clear Form  │─────────────────┘
└─────────────┘
```

## 4. Order Status State Machine

```
                    ┌──────────┐
                    │          │
         Create ───►│ PENDING  │
                    │          │
                    └────┬─────┘
                         │
              ┌──────────┴──────────┐
              │                     │
              ▼                     ▼
       ┌──────────┐          ┌───────────┐
       │ APPROVED │          │ CANCELLED │
       │          │          │           │
       └──────────┘          └───────────┘
       (Terminal)            (Terminal)

Notes:
- Both OrderRequests and OrderReceptions follow this state machine
- Approved orders cannot be re-approved or cancelled
- Cancelled orders cannot be re-cancelled or approved
- Approval of OrderReceptions triggers stock creation
```
