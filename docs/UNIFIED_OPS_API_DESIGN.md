# Unified Operations API & Python Bindings — Design Consensus

Status: Agreed (design consensus reached 2026-09-20; revised 2026-10-01 after a
feasibility and coverage review against the workspace and the OfficeCLI
reference implementation; §16 migration assessment added 2026-10-04).
Implementation not yet started.

This document records the agreed design for a format-agnostic unified operations
API (uniform CRUD over Word/Excel/PowerPoint documents) plus Python bindings,
replicating the capability model of OfficeCLI (`3rdparty/OfficeCLI/`) on top of
the existing litchi workspace. It was produced through a point-by-point design
review; each section below is a settled decision.

Relationship to existing ADRs: this plan introduces **no conflict** with any
accepted ADR. It builds a facade-level unification layer *above* the parallel
per-format semantic models mandated by ADR 0007, and it delegates all mutation
to the snapshot/Edit/Commit/Patch machinery mandated by ADR 0003. A new ADR
summarizing this design will be drafted as part of milestone M0.

Implementation is additionally constrained by ADRs this design must comply
with (surfaced by the 2026-10-01 review):

- **ADR 0002 / 0024 (crate topology)** — `litchi-ops` must be registered in the
  ADR 0024 package inventory and pass the CI crate-boundary checker;
  dependencies point downward only.
- **ADR 0010 / 0009 / 0023 / 0028 (detection and archive ownership)** — the
  single `litchi::open()` composes the leaf crates' format probes; the facade
  neither traverses container internals nor re-implements ODF/iWork detection.
- **ADR 0005 / 0031 (I/O, memory, budgets)** — `save()` writes through
  sibling-temp atomic replacement; any concurrency added by the Python
  bindings goes through execution-context budgets, never hidden global pools.
- **ADR 0030 (lazy part decode)** — part payloads live behind fallible
  accessors, which determines the return shape of the §12 `raw(part_path)`
  escape layer.
- **ADR 0006 (security)** — encrypted or signed inputs must be explicitly
  consumed (`Locked<T>` / `Signed<T>` transitions) before they become editable
  snapshots; the flat `File` enum (§4) must be able to express these
  not-yet-editable states, and `open()` never silently strips security state.

---

## 1. Core architecture: unified operation surface, not a canonical AST

The "unified AST" is realized as two things, neither of which is a single
canonical in-memory tree:

- **Unified operation surface** — one set of verbs, addressing syntax, and
  property vocabulary across all formats.
- **Unified read-side projection** — a common node view over any opened
  document.

Internally, each format keeps its own typed semantic model and its own
Edit/Commit/Patch transactions. Operations are dispatched per format (enum
dispatch, static dispatch per ADR 0004 — no boxed trait objects in the facade),
exactly as OfficeCLI routes operations through per-format `IDocumentHandler`
implementations. This preserves fidelity (no lossy canonical-model round trip),
reuses the existing investment in 60+ crates, and complies with ADR 0007.

## 2. Addressing model: typed selectors backbone + OfficeCLI string syntax

- Typed selectors are the backbone of the Rust API, per ADR 0004
  (`Result<Option<_>>` lookups, no raw native IDs in ordinary signatures).
- The string path/selector syntax is an edge parsing layer over the same
  selector types (`FromStr`-style), with typed errors distinguishing *invalid
  path syntax* from *path not found*.
- The string grammar is deliberately aligned with OfficeCLI: 1-based
  element-name paths (`/body/p[3]/r[1]`, `/slide[1]/shape[2]`), stable-identity
  forms (`[@name=...]`, `[@paraId=...]`), CSS-like predicates
  (`[attr=value]`, `:contains()`, `:has()`, `>`, boolean `and`/`or`).
- Spreadsheet addressing is part of the unified grammar: `/Sheet1/A1`,
  `/Sheet1/row[5]`, `/Sheet1/col[B]`, header-predicate row selection
  (`row[Salary>5000 and Region=EMEA]`).
- **Story-dimension addressing is part of the grammar from day one.** Paths
  must reach non-main stories, not only main bodies: `/header[1]/p[1]`,
  `/footnote[2]`, `/comments/comment[3]`, `/slide[2]/notes`,
  `/slideMaster[1]/layout[2]`, chart parts. Without story axes the read
  projection (§5) and the property vocabulary (§6) would silently cover main
  stories only. The M0 parser milestone explicitly includes story axes.
- Implementation home: extension of the existing `litchi-core::selector`
  vocabulary module. That module today holds only the typed vocabulary
  (`Position`, `Selector::{Name, Position, Id}`); no string parsing exists
  anywhere in the workspace, so the grammar is new code on the M0 critical
  path, not a small extension.

## 3. Transaction model: single-op sugar + atomic batch, immutable snapshots

- Underlying state model is unchanged: ADR 0003 immutable `Send + Sync`
  snapshots, Edit transactions, `commit()` returning (new snapshot, reversible
  `Patch`, diagnostics).
- Single-operation convenience methods (`doc.set(...)`, `doc.add(...)`, …) are
  sugar over a one-operation Edit+Commit.
- `doc.batch([...])` / explicit `edit()` carry multi-operation transactions.
  Batches are **atomic by default** (any failure rolls the whole batch back,
  matching OfficeCLI's `atomicRolledBack` semantics); an opt-in `best_effort`
  mode continues per-item and returns per-item results.
- All write methods return a **new document handle**; the old snapshot remains
  readable. Reversible patches provide undo for free.
- **The uniform commit shape is a facade adaptation, not a given.** Only the
  docx main-document commit returns exactly (new snapshot, reversible patch,
  diagnostics); the xlsx and pptx top-level commits carry no `Diagnostics`
  member, and pptx names its Edit type `Transaction`. `litchi-ops` normalizes
  all three to the promised shape.
- **Batches sequence per-operation commits across Edit families.** docx
  mutation alone is split across ~20 scoped transaction families (numbering,
  sections, content controls, glossary, …); the batch executor commits each
  operation against the evolving package state in order. Atomicity is cheap —
  a failure discards the intermediate chain and the caller keeps the old
  snapshot — but every per-op commit pays dependency-closure validation, so
  the executor carries a large-batch performance budget (validated at M2 with
  ≥1k-operation batches).
- Batch item payloads mirror OfficeCLI's full `BatchItem` shape — the 21 JSON
  keys `command` (alias `op`), `path`, `path2`, `parent`, `type`, `from`,
  `to`, `index`, `after`, `before`, `mode`, `depth`, `part`, `xpath`,
  `action`, `xml`, `props`, `selector`, `text`, `dumpVersion` — so OfficeCLI
  `dump` output can be replayed. (`path2` is swap's second element;
  `part`/`xpath`/`action`/`xml` serve the raw layer; `props` accepts both an
  object and the `["key=value"]` array form.)

## 4. Format coverage: read everything, write OOXML first, capabilities queryable

- Unified **read** covers all formats the facade already reads.
- Unified **write** ships first for `docx` / `xlsx` / `pptx` (OfficeCLI parity).
  Other formats return typed `unsupported_capability` errors.
- A capability query API (`doc.capabilities()`) reports the supported
  verbs/elements/properties, so agents probe before operating.
- ODF → iWork → legacy are onboarded in later waves without changing the API
  shape.
- ADR 0001's outstanding debt is repaid: a concrete flat
  `File::{Document, Presentation, Workbook}` enum plus a single
  `litchi::open()` entry point. (ADR 0001 mandates the facade-root `File`
  enum and single entry point; the data-bearing enum style follows ADR 0007;
  the variant names are this design's choice.) The three existing
  facade-internal `*Impl` enums merge into it, and the sheet family's
  `Box<dyn WorkbookTrait>` text-format path (CSV/TSV/…) is replaced with
  concrete variants to satisfy ADR 0004's no-boxed-trait-objects rule.
- `File` also represents inputs that are **not yet editable**: encrypted or
  signed packages surface as `Locked` / `Signed` states that must be consumed
  per ADR 0006 before any edit verb applies.

## 5. Read projection: typed nodes primary, generic node view for export

- Rust consumers get strong typed nodes (`Block`/`Inline` for word processing,
  `cell::View` for sheets — the xlsx cell lookup view enum —, `Shape` for
  slides: the ADR 0007 vocabulary).
- Every node can additionally export a generic node view
  `{path, type, text, format, children}`, shaped like OfficeCLI's
  `DocumentNode`/get/query JSON. This view feeds the Python bindings, batch
  `dump`, and agent consumption.
- The `format` map is flat `key → string` (e.g. `font.bold=true`), and **read
  and write share the same property vocabulary** — what an agent reads is what
  it can write.

## 6. Property vocabulary: a typed Rust registry as the single source of truth

- Each (element type × format) property is declared once in an internal
  registry: `name → {scalar type (bool/enum/length/color/…), constraints,
  typed setter}`.
- The string layer performs unit/color/enum coercion at parse time, following
  OfficeCLI's coercion rules in full: length units EMU/mm/cm/in/pc (pica)/
  pt/px/Q (quarter-millimeter), bare number = EMU — **except line widths,
  where a bare number means pt**; colors: hex (3/4/6/8-digit), the 147 CSS/SVG
  named colors plus `transparent`, `rgb()`/`rgba()`, `hsl()`/`hsla()`, OOXML
  preset colors (`prstClr`), theme/scheme names, scheme+transform suffixes
  (`accent1+lumMod75`), and Excel tint forms.
- Machine-readable schemas (OfficeCLI `schemas/help/*.json`-compatible) are
  **generated from the registry**, driving validation, help, capability
  reporting, Python introspection, and error suggestions — implementation,
  docs, and errors can never drift apart.
- Unknown property names produce typed errors with auto-correction suggestions
  (`unknown property "font.bolld", did you mean "font.bold?"`), critical for
  agent self-healing retries.

## 7. Crate placement: new `litchi-ops` crate, re-exported by the facade

- All unified-layer machinery — path/selector grammar (in
  `litchi-core::selector`), property registry, generic node view, batch
  executor, `File` enum, per-format op handlers — lives in a new `litchi-ops`
  crate (name provisional), registered in the ADR 0024 package inventory and
  enforced by the CI crate-boundary checker (ADR 0002).
- `litchi` depends on it and re-exports, so users still see one entry point
  (`litchi::open()`), satisfying ADR 0001 from the user's perspective while
  keeping the physical layering clean.
- The handler trait is internal to `litchi-ops` (not in `litchi-core`); dispatch
  follows the existing `DocumentImpl`-style enum pattern with static dispatch.
- `litchi-py` (crate directory `crates/pyo3-litchi`) switches its dependency to
  `litchi-ops`; its existing read API stays backward compatible, write API is
  purely additive (the current bindings are read-only, so there is no write
  surface to conflict with).

## 8. Python bindings: dual surface

- Pythonic method surface: `open/close/get/query/set/add/remove/move/swap/
  batch/dump/save/capabilities`, props passed as plain dicts
  (`doc.set(path, {"font.bold": True, "fill": "1F4E79"})` — dotted keys are
  legal in dicts), typed `.pyi` stubs (already present for the read API).
  `close` releases the handle explicitly (OfficeCLI parity; otherwise a no-op
  beyond dropping references). `capabilities()` corresponds to OfficeCLI's
  `help` schema command — same machine-readable payload, better name for a
  library; the mapping is recorded as a parity delta.
- Command-dict channel: `doc.apply({"command": "set", ...})` and
  `doc.batch([BatchItem...])` accept OfficeCLI-compatible payloads; `dump`
  emits the same format for round-trip replay.
- Return values are generic-node-view dicts; errors are Python exceptions
  carrying structured `code` / `message` / `suggestion`.
- Document objects are immutable handles: every mutating call returns a new
  handle. `save(path=None)` atomically writes back to the source path; a path
  argument means save-as (same format only, see §9).

## 9. Save semantics: same-format save; conversion is explicit and reported

- `save` / `save_as` only in the source format (first wave: OOXML ×3; read-only
  formats return `unsupported_capability`). Atomic write-back goes through ADR
  0005's sibling-temp replacement path.
- No implicit cross-format conversion on save. A future `doc.export(target)`
  entry point will be explicit and return a **fidelity report** (what was
  lost/degraded), per ADR 0007's "conversion is explicit and
  fidelity-reported".
- New-document authoring goes through a unified `create(format)`
  (`docx`/`xlsx`/`pptx` first), reusing the existing per-format builders
  (docx `Package::new()`, xlsx `Workbook::create()`, pptx `Package::new()`).

## 10. Rendering is out of scope; semantic views are in

- `view(mode)` supports `outline | text | annotated | stats` — cheap semantic
  projections that are an agent's primary document-comprehension entry point.
- Visual rendering (`html | svg | screenshot | pdf`) is explicitly **not** in
  scope (it is a separate multi-year engine program); calls return typed
  `unsupported_capability`.
- Semantic-level export (e.g. the existing `ToMarkdown`) may later join via the
  explicit `export` + fidelity-report channel. Note that `ToMarkdown` exists
  for word-processing and presentation documents only; a workbook
  implementation is new work. In OfficeCLI, "export" is not a verb either —
  PDF etc. are `view --mode pdf` via exporter plugins; our explicit
  `export(target)` is a deliberate improvement, recorded as a parity delta.

## 11. Deliverables: library + Python + thin CLI; MCP later

- Rust library (`litchi-ops` + facade re-export) and Python bindings
  (`litchi-py`, extended) are the primary deliverables.
- A thin CLI binary ships with the same verbs and `--json` output — nearly free
  once string paths, batch JSON, and generic node views exist — and doubles as
  the acceptance-test harness.
- An MCP server is deferred until the API stabilizes, then implemented as a
  thin wrapper over the CLI (OfficeCLI's pass-through model). OfficeCLI's
  resident-server verbs (`watch`/`unwatch`, `mark`/`unmark`/`get-marks`,
  `goto`) are server concerns with no library analog; they are deferred
  together with the MCP wrapper.

## 12. Raw escape layer: provided, clearly format-specific, rarely needed

- OOXML documents expose `raw(part_path)`, `raw_set(part_path, xpath, op,
  xml)`, `add_part(part_path, content_type, bytes)` (OfficeCLI `add-part`
  parity; returns the relationship id + part path), and `validate()` (schema
  validation, never mutating — ADR 0006), all routed through the same Edit
  transactions and atomic batches. Return types follow ADR 0030's fallible
  part accessors.
- These operations are explicitly documented and typed as format-specific
  escape hatches; they are not part of the generic node view.
- **Semantic coverage breadth is a first-class goal.** High-frequency special
  objects get typed representations and CRUD verbs so agents rarely need the
  escape hatch: charts, SmartArt/diagrams, equations (LaTeX input via
  `litchi-formula`), tables (word-processing tables, spreadsheet ListObjects,
  slide tables), pictures, comments/notes (including threaded comments),
  hyperlinks, text boxes — and the structural plane where the workspace
  already has typed models: styles, numbering/lists, sections and page setup,
  headers/footers, footnotes/endnotes, bookmarks, content controls, worksheet
  lifecycle, cell merge, freeze panes, slide masters/layouts, and document
  metadata (core/app/custom properties). The underlying models already exist
  (`litchi-drawingml` chart/diagram, `geom::Preset`, docx `styles`/
  `numbering`/`section`/`header_footer`/`footnote`/`content_control`, xlsx
  workbook edit model/`sheet_view`, pptx `master_layout`, `litchi-ooxml-common`
  properties). Coverage gaps are closed driven by real agent usage feedback.

## 13. Feature tiers and v1 boundary

- **P0 (v1)**: full basic CRUD verbs; worksheet lifecycle (add/remove/rename/
  move/hide sheets); document metadata (core/app/custom properties); style
  read + paragraph/character style application; numbering/list basics; CSV/TSV
  `import`; formula write + evaluation (via `litchi-eval`); comments/notes;
  hyperlinks; pictures. **Three P0 items are new engines, not facade wiring,
  and are scheduled accordingly**: formatting-aware find/replace (no
  implementation exists anywhere in the workspace today), template `merge`
  (the docx `mail_merge` module owns settings/recipient data only — it
  generates nothing), and `refresh` (TOC/field recalculation — the docx TOC
  writer today emits an "update this field" placeholder, with no entry
  computation).
- **Also in v1** (per the §12 principle, models already exist): typed chart
  CRUD and typed SmartArt CRUD.
- **P1 (immediately after)**: conditional formatting, data validation, named
  ranges, autofilter, sort; sections/page setup, headers/footers,
  footnotes/endnotes, content controls (SDT), slide masters/layouts, cell
  merge/unmerge, freeze panes, print areas/settings, sheet/workbook
  protection, and the signature/encryption open-and-preserve policy.
- **P2 (explicitly later)**: pivot tables + slicers, animations/transitions,
  3D models, audio/video media, tracked-changes accept/reject workflows, OLE
  embedding; cross-document copy/assembly, document merge/split (distinct
  from template merge — see the naming note below), redaction/sanitization,
  streaming/append authoring, semantic document compare.
- **Naming note**: `merge` in this design always means template/mail merge
  (OfficeCLI parity); document merge/split is the separate P2 concern above.
  In OfficeCLI, find/replace is modeled as `find=`/`replace=` props on `set`;
  whether we additionally expose a dedicated verb is an open micro-decision.
- Every capability lands in `capabilities()` and the generated schemas, so
  "what is supported in which version" is always machine-queryable.

## 14. Acceptance: three layers

1. **Parity layer** — OfficeCLI's **152** `schemas/help/*.json` files (46
   docx + 41 xlsx + 34 pptx + 30 `_shared` + 1 meta-schema; the `_shared` set
   includes root-metadata, OLE, equation, chart, picture, table, run, comment,
   and shape schemas) serve as the machine-checkable parity checklist (every
   delta explicitly recorded as wontfix/deferred); its 96
   `examples/{word,excel,ppt}/*.sh` scripts (14 word + 26 excel + 56 ppt) are
   ported as the CLI behavior-test corpus. Because pivot/slicer/animation/
   transition/3D/media/OLE scenarios are P2, the ported corpus goes green in
   waves: M5 requires the v1-scope subset green, with every remaining delta
   recorded.
2. **In-house CRUD gate** — `docs/CRUD_Scenario_Checklist.md` (ADR 0008):
   create/read/update/delete/reorder evidence per opened-document scenario,
   byte-identical exact no-op, atomic failure, reversible patch with inverse,
   deterministic concurrent composition.
3. **Fidelity regression** — open → empty edit → commit → save byte-level
   regression over the full `test-data/` corpus; golden snapshots of the
   property registry and generated schemas to prevent drift; Python-side
   pytest mirroring the same scenarios.

## 15. Milestones

- **M0 — design consolidation + trunk migration**: draft the new ADR; scaffold
  `litchi-ops` and register it in the ADR 0024 inventory; then two tracks —
  **M0a**: merge the existing facade open pipelines and `*Impl` enums into the
  flat `File` enum (including `Locked` / `Signed` states) behind a single
  `litchi::open()`, unify the error enum, normalize eager/source duality, and
  switch all in-tree consumers in the same commit (see §16); **M0b**:
  implement the path/selector parser **including story axes** in
  `litchi-core::selector`. (Everything later depends on the parser.)
- **M1 — unified read**: generic node view, `get/query`,
  `view(outline/text/annotated/stats)`, capability-query skeleton. Agents can
  read every format.
- **M2 — write kernel**: `set/add/remove/move/swap` for OOXML ×3, single-op
  transaction sugar, atomic batch (meeting the ≥1k-operation performance
  budget), property registry + coercions, typed errors with suggestions —
  plus the P0 structural surface: worksheet lifecycle, metadata, style
  application, numbering basics. **First working CRUD loop.**
- **M3 — consumption surfaces**: Python write bindings (dual surface), CLI
  binary, schema export + capabilities introspection.
- **M4 — advanced capabilities**: CSV import, formula evaluation, typed chart
  and SmartArt CRUD — **plus the three new engines** (formatting-aware
  find/replace, template merge, TOC/field refresh), which dominate this
  milestone's cost and may justify splitting it.
- **M5 — acceptance closure**: ported OfficeCLI corpus green for the v1 scope
  (deferred deltas recorded), CRUD-gate evidence, full fidelity regression,
  Python pytest mirror.

Each milestone is independently valuable: after M1 agents can read anything;
after M2 they can modify OOXML; after M3 everything is scriptable.

## 16. Migration from the existing partial unified API

The facade today ships a partial, **read-only** unified API —
`litchi::Document` (~4.3k lines, `crates/litchi/src/document/`),
`litchi::Presentation` (~2.8k lines, `crates/litchi/src/presentation/`),
`litchi::sheet::Workbook` (~2.5k lines, `crates/litchi/src/sheet/`), and
`litchi::iwork::Document` (~2.3k lines, `crates/litchi/src/iwork/`), plus
~1.8k lines of `sheet::text` CSV/TSV tooling. None of the four unified types
exposes any mutation surface (no `edit`/`commit`/`save`/`set_*` anywhere in
the facade). The 2026-10-03 assessment concluded the gap with this design is
**paradigmatic, not incremental** — the migration is an outright replacement
(ADR 0001: "backward compatibility is not a constraint"; ADR 0004: renamed
APIs are removed, not retained as aliases; CI compiles with `-D deprecated`),
executed in continuously buildable, dependency-ordered phases per ADR 0008.

### 16.1 Why replacement, not extension

- The incumbent element types are half-canonical wrapper enums that leak leaf
  payload types per variant (`Paragraph::Docx(crate::docx::Paragraph)`, `Run`,
  `Table`, `Cell`, …). This design's read projection is the ADR 0007
  vocabulary (`Block`/`Inline`/`cell::View`/`Shape`) plus the generic node
  view — a different contract. The ~1.6k lines of wrapper enums
  (`element.rs`, `paragraph.rs`, `run.rs`, `table.rs`, `cell_merge.rs`) are
  deleted, not adapted.
- Format-specific leak methods (`Document::docx_section_inventory`,
  `Cell::v_merge`, xlsx-only `Workbook::sheet()`, pptx-only `slide_catalog`)
  are superseded by `capabilities()` + uniform verbs.
- The error model cannot be evolved: `map_ooxml_error` flattens every leaf
  error to a string (destroying transaction diagnostics), the sheet family
  returns `Box<dyn Error>`, and iwork has its own structured `Error`.
  `litchi-ops` defines its own error enum with per-format `#[from]`
  conversions, feeding the structured `code`/`message`/`suggestion` contract.
- The CSV/TSV `Box<dyn WorkbookTrait>` openers violate ADR 0004 and duplicate
  the `Workbook` API; text formats stop masquerading as a document format and
  move behind the §13 `import` channel.

### 16.2 What carries over (relocate into `litchi-ops`)

- The open/detection pipelines (`detect_*_format`, `finish_source_*_result`
  TOCTOU guards, polyglot precedence, source-backed open paths) are the
  ready-made skeleton of `litchi::open()` — roughly 5k lines across `doc.rs`,
  `prs.rs`, `workbook.rs`.
- The `DocumentImpl` / `PresentationImpl` / `WorkbookImpl` enum-dispatch
  pattern merges into the flat `File` enum (obstacles in §16.3).
- Sidecar state (eager `Metadata`, markdown `OnceLock` caches,
  `Arc<SharedOleFile>` bundles) moves into per-variant wrapper structs or a
  `File` outer struct.
- The iwork `Options` / `SourceLimits` / `SnapshotLimits` shapes are the
  design reference for a unified `OpenOptions`.
- **The mutable owners are already in hand on the primary open paths**:
  `DocxSource` holds `source_backed::Package` (with `edit()`),
  `PptxSource` holds `SourceBackedPresentation`, and the xlsx adapter holds
  `WorkbookModel::{Owned, Source}`. Wiring Edit transactions requires no
  re-open.

### 16.3 Obstacles to the flat `File` enum (identified in M0a)

The following migration issues were resolved through M0a and subsequent waves.

1. Two `Result` types (typed `litchi_core::Error` vs the sheet family's
   `Box<dyn Error>`) unify to the typed error — breaking, accepted per
   ADR 0001.
2. Transaction errors must keep their structure (retire the flattening
   mapper).
3. Per-variant `#[cfg]` feature gates plus nested unix/windows target gates
   carry over mechanically.
4. Divergent open options converge on the non-exhaustive unified `OpenOptions`
   struct, with common controls and per-format fields. Default `open()` covers
   the common case; configured opens use the same options contract (0351).
5. The CSV/TSV trait-object path exits the enum (§16.1).
6. Sidecar state relocation (§16.2).
7. Eager/source duality normalizes per format following the sheet adapter
   pattern (one enum variant, two-state owner inside).
8. Legacy `Doc`/`Ppt` variants retain the leaf's editable owner/source by
   default. `OpenOptions::retain_source(false)` explicitly releases the
   retained source and makes later edits unavailable.

### 16.4 Leaf access after migration

- Pages, Keynote and Numbers use `File` and their immutable leaf `Package`
  owners. The former `litchi::iwork::Document` facade is retired
  ([record 0306](unified-ops/0306-iwork-facade-retirement.md)).
- CSV/TSV reader/writer tooling lives in the independent `litchi-sheet-text`
  crate, re-exported as `litchi::sheet::text` by the `sheet-text` feature.
  Unified `import` consumes that leaf
  ([record 0354](unified-ops/0354-sheet-text-and-structured-keywords.md)).
- All leaf re-export modules (`litchi::docx`, `litchi::xlsx`, …) — leaf APIs
  remain reachable for format-specific work.

### 16.5 Breakage surface and gates

- `crates/pyo3-litchi`: ~30 shallow delegation sites; Python signatures
  unchanged (the read API stays source-compatible).
- `tools/perf-baseline`: ~80 call sites plus frozen baseline route-name
  strings (`"litchi::Document"`, …). Frozen performance records are
  historical evidence and are not rewritten; new routes record under the new
  names, and the Python policy tests referencing route names are updated.
- Facade integration tests assert *semantics* (typed wrong-format refusals,
  read-limit enforcement, selector case-insensitivity) — these behaviors must
  carry over verbatim into `litchi-ops`.
- Docs: `README.md` (~16 sites), `docs/OFFICE_API_GUIDE.md` (dedicated
  Unified API section), the facade `lib.rs` doc comments.
- Gates updated in the same commits: `tools/crate_boundaries.json` (new
  `litchi → litchi-ops` edge), the ADR 0024 inventory,
  `tools/non_iwork_gate.py` (`EXPECTED_BULK_PACKAGES` 46 → 47), and
  `tools/check_iwork_public_api.py` if the iwork surface moves. The rustdoc
  glossary gate now checks registered non-iWork and iWork profiles through
  `tools/check_api_glossary.py`. `tools/final_api_inventory.py` supplies the
  separate inventory of all 68 direct libraries; its fresh final capture
  and dispositions remain required. The §14 registry/schema goldens remain
  part of the gate.

### 16.6 Settled migration decisions

- **Legacy package retention**: ADR 0033 §4 requires each `File` variant,
  including legacy binaries, to retain its editable owner by default.
  The implemented source-release option is explicit (§16.3).
- **Unified `OpenOptions` shape**: the non-exhaustive struct carries common
  controls and per-format policies; options enter through configured open
  and edit entry points (0351).
- **iWork facade timing**: record 0306 retired `litchi::iwork::Document`;
  `File::{Pages, Keynote, Numbers}` is the unified surface (§16.4).

## Open micro-decisions (deferred to implementation)

- Concrete naming: `litchi-ops` crate name, `File` enum name, exact verb names
  (`apply` vs `send`; dedicated find/replace verb vs OfficeCLI-style
  `find=`/`replace=` props on `set`; `merge` vs `template_merge`). Whether
  join/three-way composition of independent edits gets a public verb (the
  machinery exists in the per-format durable layers; §14 currently gates it
  only as an acceptance criterion).
- Error-code closed list finalized against OfficeCLI's **full** taxonomy (an
  earlier draft of this section listed only a subset): `not_found`,
  `path_not_found`, `zero_matches`, `out_of_range`, `range_target_not_found`,
  `duplicate_name`, `invalid_value`, `invalid_input`, `invalid_argument`,
  `invalid_json`, `invalid_xpath`, `invalid_path`, `invalid_selector`,
  `unsupported_property`, `unsupported_path`, `unsupported_element`,
  `unsupported_type`, `unsupported_format`, `unsupported_command`,
  `missing_property`, `missing_argument`, `file_not_found`, `file_locked`,
  `file_exists`, `file_required`, `io_error`, `corrupt_file`,
  `corrupt_batch`, `document_protected`, `validation_error`,
  `output_too_large`, `max_depth_exceeded`, `merge_not_exact`,
  `refresh_failed`, `bare_selector_rejected`, `internal_error`, plus our
  `unsupported_capability`. Advisory/warning codes (`auto_corrected`,
  `text_overflow`, `position_overlap`, `filter_warning`, …) travel as
  diagnostics, not errors.
