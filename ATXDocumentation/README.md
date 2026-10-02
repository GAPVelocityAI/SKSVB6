# SKS (Salmon King Seafood) - Comprehensive Codebase Documentation

## Overview
This documentation provides a complete analysis of the **Salmon King Seafood (SKS)** application — a Visual Basic 6 (VB6) MDI desktop application for order processing, inventory management, and customer/supplier relationship management.

- **Language**: Visual Basic 6.0
- **Architecture**: MDI (Multiple Document Interface) Desktop Application
- **Database**: SQLite 3 (via ODBC)
- **Total Source Lines**: ~9,617
- **Copyright**: Artinsoft / Mobilize.NET (2018)

## Documentation Navigation

### Root-Level Documents
| Document | Description |
|----------|-------------|
| [Project Overview](project-overview.md) | High-level project summary, technology stack, and purpose |
| [Technical Debt Report](technical-debt-report.md) | Executive summary of technical debt with AWS transformation recommendation |

### Architecture (`architecture/`)
| Document | Description |
|----------|-------------|
| [System Overview](architecture/system-overview.md) | High-level system architecture and technology stack |
| [Components](architecture/components.md) | Major system components and their responsibilities |
| [Dependencies](architecture/dependencies.md) | Internal and external dependency mapping |
| [Patterns](architecture/patterns.md) | Architectural and design patterns identified |

### Diagrams (`diagrams/`)
| Document | Description |
|----------|-------------|
| [Structural Diagrams](diagrams/structural.md) | Component, class, and module dependency diagrams |
| [Behavioral Diagrams](diagrams/behavioral.md) | Sequence, activity, and state diagrams for business workflows |
| [Architecture Diagrams](diagrams/architecture.md) | System context, integration, and security boundary diagrams |

### Technical Debt (`technical-debt/`)
| Document | Description |
|----------|-------------|
| [Summary](technical-debt/summary.md) | Overview of all technical debt findings |
| [Outdated Components](technical-debt/outdated-components.md) | Analysis of obsolete components and EOL technologies |
| [Maintenance Burden](technical-debt/maintenance-burden.md) | Areas requiring significant maintenance attention |
| [Remediation Plan](technical-debt/remediation-plan.md) | Prioritized action items for debt resolution |

### Reference (`reference/`)
| Document | Description |
|----------|-------------|
| [Program Structure](reference/program-structure.md) | Complete structural hierarchy of the codebase |
| [Interfaces](reference/interfaces.md) | Interface definitions and contracts |
| [Data Models](reference/data-models.md) | Database schema and type definitions |
| [API Reference](reference/api-reference.md) | Catalog of all callable methods and services |
| [Modules](reference/modules.md) | Module organization and dependencies |

### Analysis (`analysis/`)
| Document | Description |
|----------|-------------|
| [Code Metrics](analysis/code-metrics.md) | Complexity measurements and quality indicators |
| [Complexity Analysis](analysis/complexity-analysis.md) | Code complexity hotspots |
| [Dependency Analysis](analysis/dependency-analysis.md) | Internal and external dependency mapping |
| [Security Patterns](analysis/security-patterns.md) | Security implementations and vulnerabilities |
| [Tech Debt](analysis/tech-debt.md) | Comprehensive technical debt assessment |

### Migration (`migration/`)
| Document | Description |
|----------|-------------|
| [Component Order](migration/component-order.md) | Recommended migration order based on dependencies |
| [Test Specifications](migration/test-specifications.md) | Test case specifications for validation |
| [Validation Criteria](migration/validation-criteria.md) | Success criteria for migration validation |

### Specialized (`specialized/`)
| Document | Description |
|----------|-------------|
| [Functional Requirements](specialized/functional-requirements.md) | Business rules, workflows, calculations, and validation logic |
| [Database Patterns](specialized/database-patterns.md) | SQLite database schema and query patterns |
| [UI Components](specialized/ui-components.md) | VB6 form and control documentation |
