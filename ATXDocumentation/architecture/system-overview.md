# System Overview

[Back to README](../README.md) | [Components](components.md) | [Dependencies](dependencies.md) | [Patterns](patterns.md)

## High-Level Architecture

SKS is a **single-tier MDI desktop application** built with Visual Basic 6.0. All layers — UI, business logic, and data access — reside in a single executable (`SKS.exe`) that communicates directly with a local SQLite database via ODBC.

```
┌─────────────────────────────────────────────────┐
│                   SKS.exe (VB6)                 │
│                                                 │
│  ┌─────────────┐  ┌─────────────┐              │
│  │  MDI Parent  │  │   Splash    │              │
│  │  (frmMain)   │  │  (frmSplash)│              │
│  └──────┬───────┘  └─────────────┘              │
│         │                                        │
│  ┌──────┴──────────────────────────────────┐    │
│  │         MDI Child Forms                  │    │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐│    │
│  │  │ Orders   │ │ Master   │ │ Inventory ││    │
│  │  │ Module   │ │ Data     │ │ Module    ││    │
│  │  └──────────┘ └──────────┘ └──────────┘│    │
│  └─────────────────────────────────────────┘    │
│                                                 │
│  ┌─────────────────────────────────────────┐    │
│  │  Shared Modules (.bas)                   │    │
│  │  modMain | modConnection | modFunctions  │    │
│  └──────────────────┬──────────────────────┘    │
│                     │                            │
│  ┌──────────────────┴──────────────────────┐    │
│  │  ADODB (ActiveX Data Objects)            │    │
│  └──────────────────┬──────────────────────┘    │
└─────────────────────┼───────────────────────────┘
                      │ ODBC
              ┌───────┴───────┐
              │  Orders.db    │
              │  (SQLite 3)   │
              └───────────────┘
```

## Technology Stack

| Layer | Technology | Notes |
|-------|-----------|-------|
| Presentation | VB6 MDI Forms, ActiveX Controls | 20 forms total |
| Business Logic | VB6 Modules (.bas) + Form code-behind | Embedded in forms |
| Data Access | ADODB (msado25.tlb) | Global shared recordsets |
| Database | SQLite 3 | Local file `Orders.db`, ~108 KB |
| Connectivity | SQLite3 ODBC Driver | Connection string in modMain.bas |

## Deployment Model
- **Single-user desktop application** running on Windows
- Database file (`Orders.db`) is local to the application directory
- No client-server architecture, no network communication
- Compiled as a 32-bit Windows executable (`SKS.exe`)

## Entry Point Flow
1. `Sub Main()` in modMain.bas executes
2. SQLite ODBC connection string is set
3. `OpenConnection()` establishes database connection
4. Admin credentials are hardcoded (login bypassed)
5. `frmMain` (MDI parent) is loaded and shown
6. `frmSplash` displays for 3 seconds
7. `frmOrderRequest` opens as default child form

## Key Architectural Decisions
- **MDI Container**: All functional forms are MDI children of frmMain
- **Global State**: User identity, database connection, and recordsets are module-level globals
- **ADODC Data Binding**: Master data forms (Customers, Products, Providers) use ADODC for direct data binding
- **Manual SQL**: Order and inventory forms construct SQL manually via string concatenation
- **Shared Search**: A reusable search dialog (frmSearch) serves multiple entity types
