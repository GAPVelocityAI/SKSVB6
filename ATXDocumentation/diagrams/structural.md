# Structural Diagrams

[Back to README](../README.md) | [Behavioral](behavioral.md) | [Architecture](architecture.md)

## Component Diagram

```
┌────────────────────────────────────────────────────────────────────┐
│                        SKS.exe (VB6 Application)                   │
│                                                                    │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │                    frmMain (MDI Parent)                       │  │
│  │  ┌─────────┐  ┌─────────┐  ┌──────────┐  ┌──────────────┐  │  │
│  │  │  Menu   │  │ Status  │  │ Toolbar  │  │  MDI Client  │  │  │
│  │  │  Bar    │  │  Bar    │  │          │  │  Area         │  │  │
│  │  └─────────┘  └─────────┘  └──────────┘  └──────────────┘  │  │
│  └──────────────────────────────────────────────────────────────┘  │
│                                                                    │
│  ┌────────────────┐  ┌────────────────┐  ┌────────────────────┐   │
│  │ Master Data    │  │ Order Module   │  │ Inventory Module   │   │
│  │                │  │                │  │                    │   │
│  │ frmCustomers   │  │ frmOrderReq    │  │ frmAddStockManual  │   │
│  │ frmProducts    │  │ frmOrderRcpt   │  │ frmAdjustStock     │   │
│  │ frmProviders   │  │ frmReqApproval │  │                    │   │
│  │ frmUsersManage │  │ frmOrdApproval │  └────────────────────┘   │
│  │                │  │ frmActionReq   │                           │
│  └────────────────┘  │ frmActionRcpt  │                           │
│                      └────────────────┘                           │
│  ┌────────────────┐  ┌────────────────┐                           │
│  │ Shared Dialogs │  │ Support        │                           │
│  │                │  │                │                           │
│  │ frmSearch      │  │ frmSplash      │                           │
│  │ frmAddProduct  │  │ frmLogin       │                           │
│  │                │  │ frmAbout       │                           │
│  └────────────────┘  └────────────────┘                           │
│                                                                    │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │                  Shared Modules (.bas)                        │  │
│  │  modMain          modConnection          modFunctions         │  │
│  │  (entry/globals)  (DB connection)        (utilities)          │  │
│  └──────────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────────┘
                              │
                              │ ADODB via ODBC
                              ▼
                    ┌──────────────────┐
                    │    Orders.db     │
                    │   (SQLite 3)     │
                    │   15 tables      │
                    └──────────────────┘
```

## Module Dependency Graph

```
                  ┌──────────┐
                  │ modMain  │
                  │ (entry)  │
                  └────┬─────┘
                       │ calls OpenConnection()
                       ▼
                 ┌─────────────┐
                 │modConnection│
                 │ (data layer)│
                 └────┬────────┘
                      │ used by
          ┌───────────┼───────────┐
          ▼           ▼           ▼
     ┌─────────┐ ┌─────────┐ ┌──────────┐
     │All Forms│ │modFuncs │ │ frmSearch │
     │(17 frm) │ │(helper) │ │(utility) │
     └─────────┘ └─────────┘ └──────────┘
```

## Database Table Relationships

```
┌──────────┐     ┌────────────────┐     ┌───────────┐
│  Levels  │────►│     Users      │     │Categories │
└──────────┘     └───────┬────────┘     └─────┬─────┘
                         │                     │
              ┌──────────┼──────────┐          │
              ▼          ▼          ▼          ▼
      ┌─────────────┐ ┌──────────────┐ ┌──────────┐
      │OrderRequests│ │OrderReceptions│ │ Products │
      └──────┬──────┘ └──────┬───────┘ └────┬─────┘
             │               │               │
             ▼               ▼               │
  ┌──────────────────┐ ┌─────────────────┐   │
  │OrderRequestDetail│ │OrderRcptDetail  │   │
  └──────────────────┘ └─────────────────┘   │
                                             │
                    ┌────────────────────┐    │
                    │      Stocks        │◄───┘
                    └────────┬───────────┘
                             │
                    ┌────────┴───────────┐
                    ▼                    ▼
              ┌──────────┐        ┌──────────┐
              │ManualStks│        │ StockLog │
              └──────────┘        └──────────┘

  ┌───────────┐                   ┌───────────┐
  │ Customers │◄── ProductsByCustomer ──►│ Products │
  └───────────┘                   └───────────┘
  ┌───────────┐                   ┌───────────┐
  │ Providers │◄── ProductsByProvider ──►│ Products │
  └───────────┘                   └───────────┘
```
