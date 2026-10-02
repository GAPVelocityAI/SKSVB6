# Interfaces and Contracts

[Back to README](../README.md) | [Program Structure](program-structure.md) | [Data Models](data-models.md) | [API Reference](api-reference.md) | [Modules](modules.md)

## Overview

VB6 does not use formal interfaces in the SKS application. Instead, implicit contracts exist through:
1. **Public module functions** — callable from any form
2. **Public form properties** — used for inter-form communication
3. **Global variables** — shared state between all components
4. **ADODC data binding** — implicit contract between form controls and database columns

## Form Public Properties (Inter-Form Contracts)

### frmLogin
| Property | Type | Direction | Contract |
|----------|------|-----------|----------|
| LoginSucceeded | Boolean | Out | Set to True when valid credentials provided |

### frmAddProductTo
| Property | Type | Direction | Contract |
|----------|------|-----------|----------|
| Id | String | In | Entity ID (Customer or Provider) |
| ObjectReferred | String | In | Entity type name for display |
| Table | String | In | Target table ("ProductsByCustomer" or "ProductsByProvider") |
| ColumnName | String | In | Foreign key column name |
| SavedChanges | Boolean | Out | Whether changes were saved |

### frmActionOrderReception
| Property | Type | Direction | Contract |
|----------|------|-----------|----------|
| Action | Integer | In | 1=Approve, 2=Cancel, 3=Info |
| OrderId | Integer | In | Reception order ID to act on |

### frmActionOrderRequest
| Property | Type | Direction | Contract |
|----------|------|-----------|----------|
| Action | Integer | In | 1=Approve, 2=Cancel, 3=Info |
| OrderId | Integer | In | Request order ID to act on |

### frmSearch
| Method | Parameters | Contract |
|--------|-----------|----------|
| Search(Table, fieldToSearch, itemToSearch) | Table name, field, value | Loads search criteria for specified entity type |

## Global State Contract

All forms depend on these globals being properly initialized:

| Global | Initialized By | Used By | Contract |
|--------|---------------|---------|----------|
| ConnectionString | modMain.Sub Main | modConnection.OpenConnection | Must be valid ODBC string before any DB access |
| CurrentConnection | modConnection.OpenConnection | All forms via ExecuteSql | Must be open before any form loads |
| CurrentUserAdmin | modMain.Sub Main / frmLogin | frmUsersManage, frmMain | Must be set before MDI child forms open |
| UserFullname | modMain.Sub Main / frmLogin | frmMain status bar | Used for display only |
| rs / rs2 | modConnection.ExecuteSql/ExecuteSql2 | All forms | Caller must consume results before next call |

## ADODC Binding Contracts

Forms using ADODC establish implicit binding contracts between controls and database columns:

### frmCustomers — ADODC: `dcCustomers`
- Source: `Select * from Customers`
- Bound fields: CompanyName, CompanyOrDepartment, EmailAddress, BillingAddress, PostalCode, City, StateOrProvince, Country/Region, PhoneNumber, Extension, FaxNumber, ContactTitle, ContactFirstName, ContactLastName, Notes

### frmProducts — ADODC: `dcProducts`
- Source: `Select * from Products`
- Bound fields: ProductID, ProductName, ProductDescription, SerialNumber, UnitPrice, QuantityPerUnit, Unit, CategoryID, Discontinued

### frmProviders — ADODC: `dcProviders`
- Source: `Select * from Providers`
- Bound fields: ProviderName, PaymentTerms, PostalCode, StateOrProvince, Country/Region, EmailAddress, City, ContactTitle, ContactFirstName, ContactLastName, PhoneNumber, Extension, FaxNumber, Notes
