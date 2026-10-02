# Validation Criteria

[Back to README](../README.md) | [Component Order](component-order.md) | [Test Specifications](test-specifications.md)

## Migration Validation Criteria

### Functional Parity
| Criteria | Validation Method |
|----------|------------------|
| All 15 database tables accessible | Verify schema matches original; all CRUD operations work |
| Customer CRUD operations | Create, read, update, delete customer records |
| Product CRUD operations | Create, read, update, delete products with categories |
| Provider CRUD operations | Create, read, update, delete supplier records |
| User management | Create, read, update, delete users with access levels |
| Customer order creation | Create order with products, quantities, tax, freight |
| Supplier stock order creation | Create stock order with products and quantities |
| Order approval workflow | Approve and cancel both order types |
| Stock creation on approval | Approving reception order creates stock records |
| Manual inventory update | Add stock entries with audit trail |
| Manual inventory adjustment | Modify existing stock entries with audit trail |
| Product-entity assignment | Assign/remove products to customers and providers |
| Search functionality | Search across customers, products, providers |
| Total calculations | Tax, freight, subtotal, grand total match original behavior |
| ID generation | Record IDs generated correctly and uniquely |

### Security Improvements (Must Be Better Than Original)
| Criteria | Validation Method |
|----------|------------------|
| No SQL injection | All queries use parameterized statements |
| Password hashing | Passwords stored as cryptographic hashes |
| Authentication enabled | Login form active; no hardcoded bypasses |
| Input validation | All user inputs validated before processing |
| Role-based access | Admin vs non-admin access properly enforced |

### Technical Quality
| Criteria | Validation Method |
|----------|------------------|
| Application compiles | Clean build with no errors |
| All tests pass | Unit and integration test suites green |
| No deprecated dependencies | All dependencies are current and supported |
| Proper error handling | Exceptions caught and logged; no silent failures |
| Separation of concerns | UI, business logic, and data access in separate layers |
| No global mutable state | Connection and user state managed via DI/services |

### Data Migration
| Criteria | Validation Method |
|----------|------------------|
| Existing data preserved | All records from original Orders.db accessible |
| Schema integrity | Foreign keys, constraints, NOT NULL enforced |
| Data types correct | Numeric fields, dates, strings properly typed |
| Referential integrity | No orphaned records after migration |

### UI Parity
| Criteria | Validation Method |
|----------|------------------|
| All forms accessible | Every menu item opens the corresponding form |
| CRUD toolbar functionality | Add, Edit, Save, Delete, Search, Cancel work on master data forms |
| Grid editing | In-line quantity/price editing works in order forms |
| Data binding | Form fields display and update database records |
| Status bar | Status messages display correctly |
| Navigation | Record navigation (first/prev/next/last) works on bound forms |
