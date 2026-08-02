# Architecture Decision Records

These records define the breaking Office API architecture begun on 2026-07-31
and refined by focused decisions. They are normative for the refactor: code and
examples that disagree with an accepted record must be changed or the record
must be superseded explicitly.

The records intentionally capture durable contracts rather than every design
branch discussed during the architecture interview. Detailed choices follow the
same priorities and are tested through executable examples, compile-fail cases,
the [CRUD scenario checklist](../CRUD_Scenario_Checklist.md), and real Microsoft
Office round trips.

| ADR | Decision |
|---|---|
| [0001](0001-priorities-and-api-layers.md) | Product priorities and strict public layers |
| [0002](0002-crate-topology.md) | Dependency direction and target crate split |
| [0003](0003-snapshots-edits-and-patches.md) | Immutable snapshots, tracked edits, patches, and concurrency |
| [0004](0004-semantic-api-design.md) | Ergonomic, typed semantic API conventions |
| [0005](0005-io-memory-and-performance.md) | Positional I/O, budgets, streaming, and measured performance |
| [0006](0006-validation-security-and-compatibility.md) | Preservation, validation, security, and compatibility |
| [0007](0007-office-object-models.md) | Word, spreadsheet, presentation, and drawing models |
| [0008](0008-migration-and-verification.md) | Buildable migration phases and evidence gates |
| [0009](0009-odf-detection-ownership.md) | ODF detection ownership and fuzz boundary |
| [0010](0010-facade-archive-ownership.md) | Archive ownership below the facade |
| [0011](0011-ooxml-physical-package-ownership.md) | Physical OPC package ownership below the OOXML migration host |
| [0012](0012-biff8-formula-reference-types.md) | Checked BIFF8 formula references and panic-free encoding |
| [0013](0013-pptx-notes-deletion.md) | Atomic, package-aware PowerPoint notes deletion |
| [0014](0014-core-properties-reader-ownership.md) | Core-properties reader ownership in the OOXML common crate |

## Decision hierarchy

When two records appear to conflict, apply this order:

1. Correctness, lossless preservation, and safety are non-negotiable.
2. The ordinary facade remains concise, intuitive, and panic-free.
3. Performance decisions require representative measurements.
4. Internal modularity serves the first three goals and never leaks type noise.
5. Production readiness gates any support claim.
