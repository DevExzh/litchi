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
| [0013](0013-pptx-notes-deletion.md) | PowerPoint notes ownership and atomic deletion |
| [0014](0014-core-properties-reader-ownership.md) | Core-properties reader ownership in the OOXML common crate (amended by 0015) |
| [0015](0015-lossless-core-properties-crud.md) | Lossless, schema-typed OOXML core-properties CRUD |
| [0016](0016-biff8-writer-location-types.md) | Checked BIFF8 writer locations beyond ordinary cells |
| [0017](0017-ooxml-producer-template-ownership.md) | Format-owned deterministic OOXML producer templates |
| [0018](0018-xlsx-calculation-chain-ownership.md) | Typed XLSX calculation-chain ownership |
| [0019](0019-docx-web-settings-ownership.md) | Typed DOCX web-settings ownership |
| [0020](0020-pptx-table-style-ownership.md) | Typed PPTX table-style ownership |
| [0021](0021-docx-glossary-ownership.md) | Typed DOCX glossary and building-block ownership |
| [0022](0022-pptx-embedded-font-ownership.md) | Typed PPTX embedded-font ownership |
| [0023](0023-odf-family-crate-split.md) | Dedicated ODF family crates and umbrella facade |
| [0024](0024-current-topology.md) | Current post-migration workspace topology |
| [0025](0025-ograph-chart-area-transactions.md) | Typed OGraph chart-area snapshot transactions |
| [0026](0026-ole-directory-metadata-binding.md) | Typed shared OLE CFB directory metadata |
| [0027](0027-xls-sheet-anchor-ownership.md) | Typed XLS sheet-anchor ownership |
| [0028](0028-iwa-monolith-exit.md) | Ordered exit of the legacy IWA migration host |

The OGraph record retains its original 0025 identity. The later XLS record,
which duplicated that number, is indexed as 0027; its decision text remains
unchanged apart from the corrected identifier.

## Decision hierarchy

When two records appear to conflict, apply this order:

1. Correctness, lossless preservation, and safety are non-negotiable.
2. The ordinary facade remains concise, intuitive, and panic-free.
3. Performance decisions require representative measurements.
4. Internal modularity serves the first three goals and never leaks type noise.
5. Production readiness gates any support claim.

## Current-state terminology

[ADR 0024](0024-current-topology.md) is the authoritative inventory of the
current workspace package topology. Older records may retain `litchi-ooxml` as
historical migration-host terminology; those references are preserved as
implementation evidence and do not denote a current workspace package.
