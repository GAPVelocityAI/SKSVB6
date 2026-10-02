# Component Migration Order

[Back to README](../README.md) | [Test Specifications](test-specifications.md) | [Validation Criteria](validation-criteria.md)

## Recommended Migration Order

Components should be migrated in dependency order — foundation layers first, then forms that depend on them, ending with the most complex forms.

### Phase 1: Foundation Layer
| Order | Component | Rationale |
|-------|----------|-----------|
| 1.1 | Database schema | Define proper schema with constraints, foreign keys, migrations |
| 1.2 | modConnection.bas → Data Access Layer | Replace global recordsets with proper connection management, parameterized queries |
| 1.3 | modMain.bas → Application Configuration | Replace globals with configuration class, dependency injection |
| 1.4 | modFunctions.bas → Utility Services | Migrate validation, ID generation, combo loading to typed service classes |

### Phase 2: Authentication and Security
| Order | Component | Rationale |
|-------|----------|-----------|
| 2.1 | frmLogin.frm → Authentication Service | Implement password hashing, proper authentication flow |
| 2.2 | User management (frmUsersManage.frm) | Depends on auth service; implement RBAC |

### Phase 3: Master Data (Low Complexity)
| Order | Component | Rationale |
|-------|----------|-----------|
| 3.1 | frmCustomers.frm | Simple CRUD; ADODC → data binding migration |
| 3.2 | frmProviders.frm | Nearly identical structure to Customers |
| 3.3 | frmProducts.frm | CRUD with category combo; slightly more complex |
| 3.4 | frmSearch.frm | Reusable search; needed by master data forms |

### Phase 4: Utility Components
| Order | Component | Rationale |
|-------|----------|-----------|
| 4.1 | frmAddProductTo.frm | Product assignment dialog; needed by order forms |
| 4.2 | frmAbout.frm | Simple dialog |
| 4.3 | frmSplash.frm | Simple splash screen |

### Phase 5: Order Processing (High Complexity)
| Order | Component | Rationale |
|-------|----------|-----------|
| 5.1 | frmOrderRequest.frm | Most complex form; extract shared order logic first |
| 5.2 | frmOrderReception.frm | Mirror of OrderRequest for suppliers; reuse shared logic |
| 5.3 | frmRequestAproval.frm | Order listing/search |
| 5.4 | frmOrderAproval.frm | Reception listing/search |
| 5.5 | frmActionOrderRequest.frm | Approve/cancel customer orders |
| 5.6 | frmActionOrderReception.frm | Approve/cancel supplier orders + stock creation |

### Phase 6: Inventory Management
| Order | Component | Rationale |
|-------|----------|-----------|
| 6.1 | frmAddStockManual.frm | Depends on Products; simpler inventory form |
| 6.2 | frmAdjustStockManual.frm | Depends on Stocks; more complex modification logic |

### Phase 7: Integration and Shell
| Order | Component | Rationale |
|-------|----------|-----------|
| 7.1 | frmMain.frm → Main window | MDI parent with menu; wire up all migrated child forms |

### Components to Remove
| Component | Reason |
|-----------|--------|
| OfrmAccountProfile.frm | Dead code; references non-existent tables |

## Migration Dependencies

```
Phase 1 (Foundation)
    └──► Phase 2 (Auth)
         └──► Phase 3 (Master Data)
              └──► Phase 4 (Utilities)
                   └──► Phase 5 (Orders)
                        └──► Phase 6 (Inventory)
                             └──► Phase 7 (Shell)
```
