# Architecture Diagrams

[Back to README](../README.md) | [Structural](structural.md) | [Behavioral](behavioral.md)

## System Context Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                      Windows Desktop                             │
│                                                                  │
│  ┌─────────┐    ┌──────────────────────────────────────────┐    │
│  │  User   │───►│          SKS.exe                          │    │
│  │(Desktop)│    │                                           │    │
│  └─────────┘    │  ┌──────────┐  ┌────────┐  ┌──────────┐ │    │
│                 │  │ VB6 Forms│  │ .bas   │  │  ADODB   │ │    │
│                 │  │  (UI)    │──│Modules │──│(Data Acc)│ │    │
│                 │  └──────────┘  └────────┘  └────┬─────┘ │    │
│                 │                                  │       │    │
│                 └──────────────────────────────────┼───────┘    │
│                                                    │            │
│                                           ODBC     │            │
│                                                    ▼            │
│                                          ┌──────────────┐       │
│                                          │  Orders.db   │       │
│                                          │  (SQLite 3)  │       │
│                                          └──────────────┘       │
│                                                                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  Required Runtime Components                              │   │
│  │  MSVBVM60.DLL | SQLite3 ODBC | 13 ActiveX OCX/DLL       │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

**External Systems**: None. The application is entirely self-contained with no network communication, no external APIs, no web services, and no inter-process communication.

## Integration Points

The SKS application has **zero external integrations**. It is a standalone desktop application:

| Integration Type | Status |
|-----------------|--------|
| REST APIs | None |
| SOAP Web Services | None |
| Message Queues | None |
| File Import/Export | None |
| Email | None |
| Printing | None (no print functionality) |
| External Authentication | None |
| Network Database | None (local SQLite only) |

## Service Map — Internal Communication

```
┌──────────────────────────────────────────────────────────────┐
│                     frmMain (MDI Parent)                       │
│                                                               │
│  Menu Events ──► Form.Show (child creation)                   │
│  LogStatus() ◄── Child forms (status updates)                 │
│                                                               │
├───────────────────────────────────────────────────────────────┤
│                                                               │
│  Form-to-Form Communication:                                  │
│                                                               │
│  frmOrderRequest ──[Show]──► frmAddProductTo                  │
│                              (passes: Id, Table, ColumnName)  │
│                                                               │
│  frmOrderAproval ──[Show]──► frmActionOrderReception          │
│                              (passes: Action, OrderId)        │
│                                                               │
│  frmCustomers ───[SearchShow]──► frmSearch                    │
│  frmProducts  ───[SearchShow]──► frmSearch                    │
│  frmProviders ───[SearchShow]──► frmSearch                    │
│                                                               │
├───────────────────────────────────────────────────────────────┤
│                                                               │
│  Shared State Communication (Global Variables):               │
│                                                               │
│  modMain.UserFullname ◄──► All forms (read)                   │
│  modMain.CurrentUserAdmin ◄──► frmUsersManage (read)          │
│  modConnection.rs ◄──► All forms (read/write - DANGEROUS)     │
│  modConnection.rs2 ◄──► All forms (read/write - DANGEROUS)    │
│                                                               │
└───────────────────────────────────────────────────────────────┘
```

## Security Boundaries

```
┌─────────────────────────────────────────────────┐
│            TRUST BOUNDARY: None                  │
│                                                  │
│  User Input ──[NO VALIDATION]──► SQL Queries     │
│                                                  │
│  Login ──[BYPASSED]──► Full Admin Access         │
│                                                  │
│  Database ──[NO ENCRYPTION]──► File System       │
│                                                  │
│  Passwords ──[PLAINTEXT]──► Users Table          │
│                                                  │
│  There are NO security boundaries in this        │
│  application. All data flows are unprotected.    │
└─────────────────────────────────────────────────┘
```

## Deployment Architecture

```
Target Machine (Windows, 32-bit compatible)
│
├── C:\Program Files\SKS\
│   ├── SKS.exe           (Compiled VB6 application)
│   ├── Orders.db         (SQLite database, same directory)
│   └── Media\            (Image assets)
│
├── C:\Windows\System32\  (or SysWOW64 on 64-bit)
│   ├── MSVBVM60.DLL     (VB6 Runtime)
│   ├── MSCOMCTL.OCX      (Registered COM control)
│   ├── MSCOMCT2.OCX      (Registered COM control)
│   ├── MSFLXGRD.OCX      (Registered COM control)
│   ├── MSADODC.OCX       (Registered COM control)
│   ├── COMDLG32.OCX      (Registered COM control)
│   ├── MSMASK32.OCX      (Registered COM control)
│   ├── MSDATLST.OCX      (Registered COM control)
│   ├── MSDATGRD.OCX      (Registered COM control)
│   ├── MSCHRT20.OCX      (Registered COM control)
│   ├── MSSTDFMT.DLL      (Registered COM DLL)
│   └── MSBIND.DLL        (Registered COM DLL)
│
└── SQLite3 ODBC Driver   (Separately installed)
```
