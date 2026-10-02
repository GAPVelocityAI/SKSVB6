# Test Specifications

[Back to README](../README.md) | [Component Order](component-order.md) | [Validation Criteria](validation-criteria.md)

## Overview

The original VB6 application has zero automated tests. The following test specifications define what should be validated during and after migration.

## Unit Tests

### Data Access Layer
| Test | Input | Expected Output |
|------|-------|----------------|
| Open database connection | Valid connection string | Connection state = Open |
| Execute parameterized SELECT | Valid SQL with params | Recordset with results |
| Execute SELECT no results | Query matching no rows | Empty recordset |
| Handle invalid SQL | Malformed SQL string | Appropriate exception thrown |
| Generate record ID | Table name | Next sequential ID (MAX+1) |

### Authentication Service
| Test | Input | Expected Output |
|------|-------|----------------|
| Valid login | Correct username/password | Authenticated user with correct role |
| Invalid password | Correct user, wrong password | Authentication failure |
| Non-existent user | Unknown username | Authentication failure |
| SQL injection attempt | `' OR '1'='1` as username | Authentication failure (not bypass) |
| Empty credentials | Empty username/password | Validation error |

### Validation Utilities
| Test | Input | Expected Output |
|------|-------|----------------|
| DoubleValue with valid number | "123.45" | 123.45 |
| DoubleValue with invalid input | "abc" | 0.0 |
| TextBoxEmpty with content | "hello" | False |
| TextBoxEmpty empty | "" | True |
| ComboEmpty with selection | Selected item | False |

## Integration Tests

### Customer CRUD
| Test | Steps | Expected Result |
|------|-------|----------------|
| Create customer | Fill all required fields, save | New record in Customers table |
| Read customer | Open customer form | Customer data displayed |
| Update customer | Modify fields, save | Record updated in database |
| Delete customer | Select customer, delete | Record removed from Customers table |
| Search customer | Enter search criteria | Matching records displayed |

### Product CRUD
| Test | Steps | Expected Result |
|------|-------|----------------|
| Create product | Fill fields with category, save | New record in Products table |
| Category combo | Open products form | Categories loaded from Categories table |
| Toggle discontinued | Check/uncheck flag, save | Discontinued field updated |

### Order Processing
| Test | Steps | Expected Result |
|------|-------|----------------|
| Create order | Select customer, add products, set quantities, save | OrderRequests + OrderRequestDetails created |
| Order totals | Set quantities and tax rate | Subtotal, tax, freight, total calculated correctly |
| Approve order | Select pending order, approve | Status = 'APPROVED' |
| Cancel order | Select pending order, cancel | Status = 'CANCELLED' |
| Prevent re-approval | Select approved order, try approve | Action rejected |

### Stock Management
| Test | Steps | Expected Result |
|------|-------|----------------|
| Approve stock order | Approve reception order | Stocks records created, Products.UnitsInStock updated |
| Manual stock add | Add stock for product | Stocks, ManualStocks(ADD), StockLog created |
| Stock adjustment | Modify existing stock | Stocks updated, ManualStocks(MOD), StockLog created |
| Stock balance | After stock operations | Products.UnitsInStock = SUM of all stock quantities |

## Security Tests
| Test | Steps | Expected Result |
|------|-------|----------------|
| SQL injection in search | Enter `'; DROP TABLE Users; --` | Query safely parameterized; no injection |
| SQL injection in login | Enter `' OR '1'='1` | Authentication fails |
| Password hashing | Create user with password | Password stored as hash, not plaintext |
| Authorization check | Non-admin access to admin features | Access denied |
