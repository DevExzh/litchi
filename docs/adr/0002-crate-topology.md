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

litchi-cfb -> litchi-ole-common
litchi-ole-common -> litchi-doc
litchi-ole-common -> litchi-ppt
litchi-ole-common -> litchi-xls
litchi-odraw -> litchi-doc
litchi-odraw -> litchi-ppt
litchi-odraw -> litchi-xls

litchi-cfb -> litchi-sign -> litchi-opc
litchi-cfb -> litchi-ograph
litchi-odraw -> litchi-imgconv
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

`litchi-ole-common::object` owns bounded, inert discovery of DOC/XLS object
storage topology and transactional CFB stream/storage rewrites. It exposes
contextual names such as `object::{Object, Objects, Editor, Limits}`. Semantic
lookup (`Objects::get`) is the primary selector, while checked discovery-order
lookup (`Objects::at`) remains available; neither selector panics. Concrete
host metadata is not modeled in the common crate. Common objects retain those
bytes opaquely, and the owning format crate provides the typed interpretation,
such as `doc::embedded_object::Info` for `[MS-DOC]` `ObjInfo` flags.

Additional focused crates are permitted where the responsibility is real:

- `litchi-math` replaces the current equation-focused `litchi-formula` name.
- `litchi-calc` owns spreadsheet formula parsing, dependency graphs, and pure
  calculation; it has no network or async-runtime dependency.
- `litchi-crypto`, `litchi-sign`, and `litchi-vba` own shared inert security
  capabilities rather than creating OPC/OLE cross-dependencies.
- `litchi-ograph` owns the neutral `[MS-OGRAPH]` chart model, record grammar,
  and standalone compound-package codec. XLS owns workbook tab/Obj integration
  and PPT owns presentation frames and embedded-object integration; PPT never
  depends on the concrete XLS crate.
- Runtime adapters such as `litchi-tokio` are separate optional crates.

`litchi-sign` owns the bounded, trust-neutral signature engine rather than a
format facade. Its root vocabulary is compact (`Signer`, `Policy`, `Coverage`,
`Report`, `Status`, and `Trust`), `xml` owns XMLDSig canonicalization and
verification, and `cfb` owns the compound-file storage adapter. `litchi-opc`
depends downward on that engine and owns only OPC graph selection, relationship
and content-type maintenance, and package-level transaction staging. This
direction prevents a signing/OPC cycle and lets DOC, PPT, and XLS use the same
neutral engine without depending on OOXML. Strict policy accepts only complete
package coverage; compatibility policy may report a typed partial coverage for
real producer signatures that intentionally select a subset. Neither policy
turns partial coverage into an unqualified success.

`litchi-odraw::image` owns the OfficeArt BLIP, FBSE/BStore, delayed-storage,
digest, and bounded writer grammar. Image decoding and conversion remain in
`litchi-imgconv`, which consumes the grammar instead of redefining it. Host
crates retain their native topology: in particular, PPT resolves a picture ID
through the drawing-group FBSE table to a delayed Pictures-stream BLIP instead
of treating that headerless stream as a second BStore.

The optional umbrella image facade depends on both layers directly. It exposes
the grammar as `images::art` and codecs as `images::codec`, so the codec crate
does not become a compatibility tunnel for types it does not own. File helpers
use short contextual names (`images::doc`, `images::ppt`, `images::escher`, and
`images::store`) and return borrowed views whenever the input lifetime permits.

`litchi-ograph` owns only neutral chart records, borrowed and owned package
views, deterministic record encoding, and the standalone compound-package
codec. XLS owns workbook tabs, BIFF objects, and chart-substream mutation; PPT
owns frames and embedded-object integration. The neutral crate does not depend
on either host, expose a runtime lock wrapper, or imply rendering or activation.

`litchi-crypto` owns bounded `[MS-OFFCRYPTO]` structures and transformations,
including compound-file DataSpaces metadata and password-derived cipher
contexts. It may depend downward on `litchi-cfb` and `litchi-ole-common`, but
not on either migration host or any concrete document format. Its namespaces
provide short typed names such as `rc4::{Flags, Header, Context, Error}`;
format crates remain responsible for locating native records and mapping
crypto failures into their own error vocabulary. Secret-bearing contexts keep
their material private and zeroizing, and the crate has no async-runtime edge.

`litchi-vba` owns the inert, bounded `[MS-OVBA]` codec and project model. It
depends downward only on `litchi-cfb` and `litchi-core`; it does not own DOC,
PPT, XLS, OPC, or OOXML package integration and never compiles, interprets, or
executes source. Its contextual namespaces keep the public vocabulary short:
`codec::{encode, decode}`, `dir::{Dir, Module, Kind}`,
`project::{Project, Module, Text}`, and
`build::{Project, Module, Id, Platform, Kind}`. A serialized `Payload` is a
validated, move-first capability rather than an arbitrary byte alias. Callers
can obtain one only by validating an existing compound payload or by consuming
a checked builder; host packages consume it directly instead of accepting an
untyped `Vec<u8>`. This preserves a concise high-level boundary without hiding
the lower-level directory and compression codecs needed by focused tooling.
The crate has no async-runtime edge, public lock wrapper, compatibility facade,
or public type carrying a redundant `Vba` prefix.

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
