# ADR 0033: Unified operations API over the per-format owners

- Status: Accepted (2026-10-04, by the owner's decision recorded in
  [unified-ops record 0006](../unified-ops/0006-owner-accepts-adr-0033.md);
  proposed 2026-10-04)
- Date: 2026-10-04
- Supersedes: nothing. Amends: nothing. It fulfils the outstanding facade
  obligation of [ADR 0001](0001-priorities-and-api-layers.md) (the concrete
  facade-root `File` enum and single entry point) and adds one package to the
  [ADR 0024](0024-current-topology.md) inventory.
- Raised by: the design consensus
  [`docs/UNIFIED_OPS_API_DESIGN.md`](../UNIFIED_OPS_API_DESIGN.md) (agreed
  2026-09-20, revised 2026-10-01, migration assessment 2026-10-04), the
  execution goal [`docs/UNIFIED_OPS_GOAL.md`](../UNIFIED_OPS_GOAL.md) including
  its owner override of 2026-10-04, and unified-ops change record 0004, which
  drafts this record and scaffolds the `litchi-ops` package.

## Context

Litchi has typed per-format owners for every claimed format, each with its own
immutable snapshots and Edit/Commit/Patch transactions (ADR 0003, ADR 0007).
The facade exposes only a partial, read-only unification
(`litchi::Document`, `litchi::Presentation`, `litchi::sheet::Workbook`,
`litchi::iwork::Document`) with wrapper enums that leak leaf payload types,
an error mapper that flattens leaf errors to strings, and a CSV/TSV path behind
`Box<dyn WorkbookTrait>`. ADR 0001's concrete facade-root `File` enum and single
entry point were never delivered.

Agents and developers need one uniform way to open, read, query, and edit any
document Litchi understands, comparable to the OfficeCLI reference
implementation, without a second document model, a lossy canonical round trip,
or a slower leaf path. The design document records the point-by-point
consensus; this record restates its settled decisions as an architecture
decision so they can be accepted, cited, and enforced.

## Decision

### 1. An operation surface, not a canonical AST (design §1)

The unification is a uniform operation surface (one set of verbs, one
addressing syntax, one property vocabulary) plus a uniform read-side
projection. Each format keeps its own typed semantic model and transactions.
Operations dispatch per format through enums and generics; there is no boxed
trait object anywhere in the facade or on a per-operation path.

### 2. Typed selectors, with an OfficeCLI-aligned string grammar (design §2)

Typed selectors are the backbone (`Result<Option<_>>` lookups, no raw native
IDs in ordinary signatures). A string path grammar is an edge parsing layer
over the same types, with typed errors separating `invalid_path` from
`path_not_found`. The grammar follows OfficeCLI: 1-based element paths,
stable-identity predicates, CSS-like predicates, spreadsheet addressing
including header predicates, and story axes (headers, footers, footnotes,
comments, notes, masters, layouts, chart parts) from the first parser
milestone. The grammar lives in `litchi-core::selector`, extending its existing
typed vocabulary.

### 3. Single-op sugar over atomic batches of immutable snapshots (design §3)

The ADR 0003 state model is unchanged. Single-operation verbs are sugar over a
one-operation Edit and commit; batches carry multi-operation transactions and
are atomic by default (any failure leaves the caller's snapshot untouched),
with an opt-in best-effort mode that returns per-item results. Every write
returns a new handle; the old snapshot stays readable; reversible patches are
retained. `litchi-ops` normalizes the per-format commit shapes to (snapshot,
patch, diagnostics). Batch items accept OfficeCLI's full `BatchItem` JSON
shape so `dump` output replays.

### 4. Every claimed format is read and written (design §4, superseded in part)

The owner override of 2026-10-04 supersedes design §4's "write OOXML first;
other formats return `unsupported_capability`", §9's OOXML-only save, the
"write OOXML first" ordering of §13, and the §16.6 legacy-package-retention
decision point:

- Every format the library claims in
  [`docs/FEATURE_MATRIX.md`](../FEATURE_MATRIX.md) (doc, docx, xls, xlsx,
  xlsb, ppt, pptx, rtf, odt, ods, odp, odg, odc, odi, odb, oth, odm, odf
  formula, pages, keynote, numbers) gets both unified read and unified write,
  including `save` and `create`.
- Whole-format `unsupported_capability` is not an acceptable end state.
  `unsupported_capability` is legitimate only at element or property
  granularity, where a format lacks the concept or its leaf owner has no typed
  model yet; every such gap is listed by `capabilities()` and in a delta table
  with a follow-up.
- OOXML is still the first write wave (the OfficeCLI parity target); the other
  formats are in scope for v1.
- Every `File` variant, legacy binary included, retains the leaf's editable
  owner, source-backed where the leaf has one.

`File` is one flat, data-bearing enum (ADR 0001, ADR 0007) behind a single
`litchi::open()`. It also expresses not-yet-editable inputs as `Locked` and
`Signed` states that must be consumed explicitly before an edit verb applies
(ADR 0006). The existing facade `*Impl` enums merge into it, and text formats
(CSV/TSV) leave the enum for an `import` channel.

### 5. Typed nodes first; a generic node view on demand (design §5)

Rust consumers receive the ADR 0007 typed vocabulary (`Block`/`Inline`,
`cell::View`, `Shape`). Every node can export a generic node view
`{path, type, text, format, children}` shaped like OfficeCLI's JSON. The view's
flat `format` map uses the same property names that writes accept.

### 6. One static property registry (design §6)

Each (element type x format) property is declared once: name, scalar type,
constraints, typed setter. The string layer applies OfficeCLI's coercion rules
in full (length units, the line-width exception, colors in every OfficeCLI
form). OfficeCLI-compatible schemas, validation, help, capability reports,
Python introspection, and "did you mean" suggestions are generated from the
registry, guarded by golden snapshots.

### 7. A new `litchi-ops` package, re-exported by the facade (design §7)

All unified-layer machinery (the `File` enum, unified error, node view,
property registry, batch executor, capability reporting, and the internal
per-format handlers) lives in a new workspace package `litchi-ops`; the path
grammar stays in `litchi-core::selector` and is re-exported. The handler trait
is internal to `litchi-ops`, never in `litchi-core`. The `litchi` facade
depends on `litchi-ops` and re-exports its entry points, so users still see one
`litchi::open()`. `litchi-py` (`crates/pyo3-litchi`) switches to the unified
surface; its existing read API stays source compatible and the write API is
additive.

Topology (ADR 0002, ADR 0024):

- Registered now, by the scaffold: package `litchi-ops`, one normal edge
  `litchi-ops -> litchi-core`, runtime-neutral classification, and bulk
  membership of the non-iWork release gate (47 bulk packages).
- Planned, each added in the change that first uses it and not before:
  `litchi-ops -> <format leaf>` for the claimed formats' owner crates (the OOXML,
  OLE, ODF, RTF, and iWork leaves) and the common crates the facade already
  consumes for those leaves; `litchi -> litchi-ops`; and
  `litchi-py -> litchi-ops` if the bindings consume it directly rather than
  through the facade.
- Dependencies point downward only. `litchi-ops` is facade-tier: no leaf or
  common crate may depend on it, and it does not traverse container internals
  or depend directly on `soapberry-zip` (ADR 0010).
- When `litchi-ops` gains iWork leaf edges, its `--all-features` tree will
  contain the iWork packages; the non-iWork release gate must then classify
  `litchi-ops` like the facade (safe and unsafe feature closures) instead of
  as a bulk root. That reclassification is part of the change that adds the
  first iWork edge.
  *Done in record 0296 (owner decision D1, 2026-10-15): `litchi-ops` and
  `litchi-cli` left the bulk slice as the iWork-capable unified layer; see
  the ADR 0024 amendment of that date.*

Format features of `litchi-ops` mirror the facade's per-format features; they
are inert until the corresponding leaf edge exists.

### 8. Python bindings: a dual surface (design §8)

A Pythonic method surface (`open/close/get/query/set/add/remove/move/swap/
batch/dump/save/capabilities`) with dict props using dotted keys, plus an
OfficeCLI-compatible command-dict channel (`apply`, `batch`, `dump`). Handles
are immutable; errors are exceptions carrying `code`, `message`, and
`suggestion`.

### 9. Same-format save; conversion is explicit (design §9)

`save` and `save_as` write the source format only, through ADR 0005's
sibling-temp atomic replacement. Under the owner override this applies to
every claimed format. There is no implicit cross-format conversion; a future
`export(target)` is explicit and returns a fidelity report (ADR 0007).
`create(format)` reuses each leaf's existing builders.

### 10. Semantic views in, rendering out (design §10)

`view(outline | text | annotated | stats)` is in scope. Visual rendering
(`html | svg | screenshot | pdf`) is out of scope and returns a typed
`unsupported_capability`.

### 11. Deliverables (design §11)

The Rust library (`litchi-ops` plus facade re-export) and the Python bindings
are primary. A thin CLI with the same verbs and `--json` output doubles as the
acceptance harness. An MCP server and OfficeCLI's resident-server verbs are
deferred.

### 12. A typed raw escape layer (design §12)

OOXML documents expose `raw`, `raw_set`, `add_part`, and a non-mutating
`validate`, routed through the same Edit transactions and atomic batches,
with return types following ADR 0030's fallible part accessors. These are
documented format-specific escape hatches, not part of the node view. Typed
CRUD for high-frequency special objects (charts, SmartArt, equations, tables,
pictures, comments, hyperlinks, text boxes, and the structural plane) is a
first-class goal so the escape layer is rarely needed.

### 13. Feature tiers (design §13)

P0 (v1): full basic CRUD; worksheet lifecycle; document metadata; style read
and application; numbering basics; CSV/TSV import; formula write and
evaluation; comments; hyperlinks; pictures; and three new engines
(formatting-aware find/replace, template merge, field/TOC refresh). Typed chart
and SmartArt CRUD are also v1. P1 and P2 follow the design's lists. `merge`
means template merge; find/replace is first exposed as `find=`/`replace=`
props on `set`. Every capability is reported by `capabilities()`.

### 14. Acceptance in three layers (design §14)

1. Parity: OfficeCLI's 152 `schemas/help` files are the machine-checkable
   checklist and its 96 example scripts the CLI behavior corpus, with every
   delta recorded as wontfix or deferred.
2. The in-house CRUD gate of
   [`docs/CRUD_Scenario_Checklist.md`](../CRUD_Scenario_Checklist.md)
   (ADR 0008).
3. Fidelity regression: open, empty edit, commit, save, byte-compared over the
   corpus; golden snapshots of the registry and generated schemas; a pytest
   mirror.

### 15. Errors

`litchi-ops` defines one typed error enum with per-format conversions. Every
error maps to exactly one code from the design's closed list (OfficeCLI's full
taxonomy plus `unsupported_capability`) with `message` and an optional
`suggestion`. Leaf errors are never flattened into strings; the facade's
`map_ooxml_error` is retired. Advisory codes travel as diagnostics, not errors.

## Compliance with accepted records

- **ADR 0001**: the flat facade-root `File` enum and single `litchi::open()`
  are delivered; no public operation panics; unsupported editing is a typed
  capability error. Backward compatibility is not a constraint, so the old
  wrapper types are replaced, not extended.
- **ADR 0002 / ADR 0024**: `litchi-ops` is registered in the package inventory,
  the CI crate-boundary policy, and the non-iWork release gate; edges point
  downward only; planned edges are listed in section 7.
- **ADR 0003**: all mutation goes through the leaf snapshot/Edit/Commit/Patch
  machinery; exact no-ops stay exact; failed batches leave the caller's
  snapshot untouched.
- **ADR 0004**: static dispatch only (enums, generics, `impl Trait`); typed
  selectors; replaced APIs are removed, not aliased, under `-D deprecated`.
- **ADR 0005 / ADR 0031**: save uses sibling-temp atomic replacement; no hidden
  global pool, runtime, network, or clock; any parallelism, including
  GIL-released work in Python, runs under an explicit ADR 0031 execution
  context. `litchi-ops` is classified runtime-neutral.
- **ADR 0006**: validation never mutates, save never silently repairs, and
  encrypted or signed inputs surface as `Locked`/`Signed` states that must be
  consumed explicitly; `open()` never strips security state.
- **ADR 0007**: per-format typed models stay parallel; the read projection is
  the ADR 0007 vocabulary plus the generic view; conversion is explicit and
  fidelity-reported.
- **ADR 0008**: every migration phase leaves the workspace buildable and green;
  acceptance uses the CRUD scenario checklist.
- **ADR 0009 / ADR 0010 / ADR 0023 / ADR 0028**: `open()` composes the leaf
  crates' format probes; the unified layer neither traverses container
  internals nor re-implements ODF or iWork detection, and consumes iWork only
  through the iWork crates' public APIs.
- **ADR 0030**: part payloads stay behind fallible accessors; open decodes no
  part beyond what the leaf open decodes; `raw(part_path)` returns the
  fallible accessor shape.

`forbid(unsafe_code)` holds in `litchi-ops`, and no record's
`forbid(unsafe_code)` is weakened elsewhere.

## Performance commitments

The unified layer is judged against the direct leaf APIs it wraps. Every
benchmarked scenario records a paired measurement (leaf leg vs unified leg,
same corpus, build, and machine) in `tools/perf-baseline` under new route
names; frozen route-name records are never rewritten. The budgets are review
triggers, not claims:

- **Open**: one detection pass and one container index, reusing the prepared
  source; no part payload decoded beyond the leaf open. Overhead at most 5% or
  20 µs over the leaf open, whichever is larger.
- **Read / get**: path resolution cost proportional to path depth and the
  scanned axis, never to document size; only touched parts decoded; the
  generic view built on demand, never by typed `get`.
- **Query**: results stream through an iterator with early termination;
  predicates compile once per query.
- **Single-op write**: unified `set` overhead at most 10% over the equivalent
  hand-written leaf Edit and commit on the small-edit scenarios.
- **Batch**: linear in operation count; a 1k-operation homogeneous batch
  within 2x of one hand-written leaf Edit doing the same work, by coalescing
  consecutive same-family operations into one leaf transaction while keeping
  per-item error attribution and all-or-nothing atomicity. Measured at 1, 10,
  100, 1k, and 10k operations.
- **Save**: the leaf save, with compressed-entry passthrough for untouched
  members and no added copies.
- **Allocation discipline**: registry lookup, short-path parsing, and value
  coercion allocate nothing on the success path; registry tables are `static`;
  suggestions run only on the error path.
- **Handles**: `File` and node handles are cheap to clone, `Send + Sync`, small
  (with size assertions), and every write costs in proportion to the changed
  closure.
- **Python**: the GIL is released around open, query, batch, save, and any
  operation that can exceed about 100 µs; result dicts are built directly from
  the projection; batch payloads convert to Rust once.

Measurements follow the performance program's method (before/after, ABBA
where noise warrants, machine-readable results, no single-run claims). A
regression above 5% in an existing performance-program scenario caused by the
facade migration blocks the merge until explained or fixed.

## Consequences

- The facade's read-only unified types, wrapper enums, string-flattening error
  mapper, and CSV/TSV trait-object openers are deleted in the migration
  milestone; every in-tree consumer (facade, `pyo3-litchi`, `perf-baseline`,
  tests, docs) moves in the same change.
- Write handlers are owed for all 21 claimed formats. Missing leaf write APIs
  are added in the owning leaf crate, except in iWork crates, whose gaps are
  recorded as requests for their owners and reported by `capabilities()`.
- The topology gains one facade-tier package. The non-iWork release gate must
  reclassify it when its first iWork edge lands (section 7).
- The scaffold registered with this draft has no public API and no behavior;
  later code that relies on this record's decisions waits for its acceptance,
  per the proposed-records rule in the ADR index.

## Verification

- `python3 tools/check_crate_boundaries.py` and
  `python3 -m unittest tools/test_check_crate_boundaries.py` pass with the
  `litchi-ops -> litchi-core` edge registered.
- `python3 tools/non_iwork_gate.py print` reports 66 workspace packages and 47
  bulk packages; `python3 -m unittest tools.test_non_iwork_gate` passes.
- `cargo check`, warning-denied `cargo clippy`, and `cargo doc` with
  `RUSTDOCFLAGS=-Dwarnings` pass for `litchi-ops`.
- Later milestones add the paired leaf-vs-unified benchmarks, the OfficeCLI
  parity delta table, the CRUD gate evidence, and the registry and schema
  golden snapshots described above.
