# ADR 0002: Crate topology and dependency direction

- Status: Accepted
- Date: 2026-07-31

## Decision

The target workspace uses small single-responsibility crates and rejects peer
format dependencies in CI. In the diagram, `A -> B` means that `B` may depend on
the more foundational `A`.

```text
litchi-core
├── litchi-detect
├── litchi-word
├── litchi-slide
└── litchi-sheet

litchi-opc -> litchi-ooxml-common -> litchi-drawingml
                                      ├── litchi-docx
                                      ├── litchi-pptx
                                      ├── litchi-xlsx
                                      └── litchi-xlsb

litchi-cfb -> litchi-ole-common -> litchi-odraw
                                   ├── litchi-doc
                                   ├── litchi-ppt
                                   └── litchi-xls
```

The diagram shows the main direction, not every foundation edge. In particular,
concrete Word, presentation, and spreadsheet crates also depend on their neutral
vocabulary crate. `litchi-drawingml` may depend on `litchi-sheet` for neutral
chart data references; no concrete spreadsheet crate may depend on another.

`litchi-word`, `litchi-slide`, and `litchi-sheet` depend only on `litchi-core`.
They contain selectors, queries, events, detached builders, and semantic values,
not container parsing or concrete document handles. Concrete imported objects
remain canonical in their format crate.

`litchi-odraw` owns only the OfficeArt record grammar, property tables, shape
containers, bounded traversal, and deterministic record writing defined by
`[MS-ODRAW]`. The `OfficeArtClientData` and `OfficeArtClientTextbox` payloads
are explicitly host-application records in `[MS-ODRAW]` section 2.2.14, so DOC,
PPT, and XLS decode those payloads in their concrete crates. Shared shapes
expose the borrowed host payload records without interpreting them. Canonical
types use their module context (`record::Record`, `prop::Props`,
`shape::Shape`) instead of repeating an `Escher` or `OfficeArt` prefix.

Additional focused crates are permitted where the responsibility is real:

- `litchi-math` replaces the current equation-focused `litchi-formula` name.
- `litchi-calc` owns spreadsheet formula parsing, dependency graphs, and pure
  calculation; it has no network or async-runtime dependency.
- `litchi-crypto`, `litchi-sign`, and `litchi-vba` own shared inert security
  capabilities rather than creating OPC/OLE cross-dependencies.
- Runtime adapters such as `litchi-tokio` are separate optional crates.

`litchi-crypto` owns bounded `[MS-OFFCRYPTO]` structures and transformations,
including compound-file DataSpaces metadata and password-derived cipher
contexts. It may depend downward on `litchi-cfb` and `litchi-ole-common`, but
not on either migration host or any concrete document format. Its namespaces
provide short typed names such as `rc4::{Flags, Header, Context, Error}`;
format crates remain responsible for locating native records and mapping
crypto failures into their own error vocabulary. Secret-bearing contexts keep
their material private and zeroizing, and the crate has no async-runtime edge.

The current `litchi-ooxml` and `litchi-ole` monoliths are removed after their
contents migrate. They do not remain as compatibility crates. The umbrella
`litchi` contains no format implementation logic and re-exports canonical types
without creating aliases with redundant prefixes.

## Enforcement

- A checked-in dependency allowlist rejects concrete peer edges, including dev
  and optional dependencies.
- The allowlist inventories every direct `crates/*/Cargo.toml` workspace member.
  Every internal edge is either a canonical downward ceiling or an ordered,
  stale-checked migration-debt entry with a reason and exit condition. Migration
  hosts have no canonical edges: adding an unclassified edge fails, and removing
  a debt edge also fails until its ledger entry is deleted.
- `litchi-core` owns only format-neutral sources, blobs, budgets, execution,
  scalars, selectors, diagnostics, patch envelopes, and content events. It owns
  no ZIP, XML, CFB, format feature, Tokio, Reqwest, or Rayon dependency.
- Container/common crates do not depend on concrete formats.
- Default `litchi` enables DOCX, PPTX, and XLSX. XLSB, legacy formats, crypto,
  signing, VBA parsing, calculation, rendering, and runtime adapters are opt-in.
  Enabling a feature adds capability and never changes existing semantics.
