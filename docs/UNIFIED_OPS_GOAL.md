You are the principal Rust API engineer responsible for implementing Litchi's
unified operations API and its Python bindings, as agreed in
docs/UNIFIED_OPS_API_DESIGN.md.

Repository:
https://github.com/DevExzh/litchi

Working branch and worktree:
feat/unified-ops-api, checked out at /home/zhuhe/code/litchi-worktrees/unified-ops-api

Base branch:
feat/office-format-completeness (the performance program continues there in
/home/zhuhe/code/litchi; never edit that working copy from this goal)

The design document docs/UNIFIED_OPS_API_DESIGN.md is the specification. Its
sections are settled decisions; this goal adds the execution rules, the
performance and ergonomics bar the implementation must meet, and the definition
of done. Where this goal and the design appear to disagree, the design wins and
the disagreement is recorded as a finding — with one explicit owner override,
below.

OWNER DIRECTION (2026-10-08) — API vocabulary and restart:

  Refactor toward the glossary/default-policy draft supplied at the base
  checkout's docs/adr/0033-unified-api-glossary-and-default-policies.md.
  It is adopted here as ADR 0034 because ADR 0033 already identifies the
  accepted unified-operations design. See record 0016 and docs/API_GLOSSARY.md.
  The glossary governs API shapes and strict default policies; the existing
  design and all-format read/write owner override still govern coverage.
  M0a/M0b have already landed: adapt their consumers with each leaf wave,
  rather than restarting them. iWork ownership constraint 7 remains binding.

OWNER OVERRIDE (2026-10-04) — supersedes design §4, §9, §13 "write OOXML
first", and §16.6 "legacy package retention":

  For every file format the library claims to support, the unified API must
  provide BOTH read and write. Whole-format `unsupported_capability` is not an
  acceptable end state for any claimed format. The claimed formats are those
  listed in docs/FEATURE_MATRIX.md: doc, docx, xls, xlsx, xlsb, ppt, pptx,
  rtf, odt, ods, odp, odg, odc, odi, odb, oth, odm, odf (formula), pages,
  keynote, numbers. `unsupported_capability` remains legitimate only at
  element/property granularity, where a format genuinely lacks the concept or
  the leaf owner has no typed model yet; every such gap is listed in
  `capabilities()` and in a delta table, with a follow-up.

  PRIORITY (owner, 2026-10-04, later the same day): implement in tiers.
    Tier 1 — OOXML: docx, xlsx, pptx, xlsb (OfficeCLI parity target).
    Tier 2 — legacy OLE: doc, xls, ppt.
    Tier 3 — delayed: rtf, ODF (odt/ods/odp/odg/odc/odi/odb/oth/odm/odf), and
             iWork (pages/keynote/numbers).
  Tier 3 is still required for done, but it starts only after tiers 1–2 have
  a working read+write loop. Until then tier-3 formats keep their existing
  read paths and may appear as read-only `File` variants; the `File` shape
  must admit their write variants later without reshaping.

The reference implementation is OfficeCLI. It is not tracked in git, so read
it from the main checkout: /home/zhuhe/code/litchi/3rdparty/OfficeCLI/
(schemas/help/{_shared,docx,pptx,xlsx}/*.json, examples/{word,excel,ppt}/*.sh,
and the handler sources). Do not copy OfficeCLI code; replicate behavior.

========================================================================
MISSION
========================================================================

Give agents and developers one uniform way to open, read, query, and edit any
document Litchi understands:

    let file = litchi::open("report.docx")?;
    let para = file.get("/body/p[3]")?;               // typed node + node view
    let file = file.set("/body/p[3]/r[1]", props! { "font.bold" => true })?;
    file.save()?;                                      // atomic, same format

and the same in Python:

    with litchi.open("report.docx") as doc:
        doc = doc.set("/body/p[3]/r[1]", {"font.bold": True})
        doc.save()

The unified layer must be **thin over the existing per-format owners** (ADR
0007 parallel models, ADR 0003 Edit/Commit/Patch) and **must not make any
existing read or write path materially slower**. It is a dispatch-and-coercion
layer, not a second document model.

========================================================================
NON-NEGOTIABLE CONSTRAINTS
========================================================================

Before editing code, read docs/adr/README.md, every accepted ADR, and the
design document in full. The design's preamble lists the ADRs it must comply
with (0001, 0002/0024, 0003, 0004, 0005/0031, 0006, 0007, 0008, 0009/0010/0023/
0028, 0030); treat them as hard constraints. In addition:

1. Correctness, lossless preservation, bounded resource use, and safety take
   precedence over speed and convenience. Exact no-ops stay exact no-ops;
   failed batches leave the caller's snapshot untouched; validation never
   mutates; save never silently repairs.
2. Static dispatch only in the facade (ADR 0004): enums, generics, and
   `impl Trait`; no `Box<dyn Trait>` on any public or per-operation path.
3. No new `unsafe`; do not weaken `forbid(unsafe_code)` anywhere.
4. Typed errors end to end. Never flatten a leaf error into a `String`
   (retire `map_ooxml_error`); every error maps to exactly one code from the
   design's closed code list with `message` and optional `suggestion`.
5. No ambient runtime: no hidden global thread pool, Tokio, network, or clock.
   Any parallelism (including GIL-released work in Python) runs under an
   explicit ADR 0031 execution context.
6. The library is in early alpha; breaking changes are acceptable and
   expected (ADR 0001). Replaced APIs are removed, not aliased (ADR 0004; CI
   uses `-D deprecated`). Every phase still leaves the workspace buildable and
   green (ADR 0008).
7. iWork crates (`litchi-iwa*`, `litchi-numbers*`, `litchi-pages`,
   `litchi-keynote`) are owned by other concurrent agents: do not modify them.
   The unified layer consumes their public read and edit APIs. Where a write
   verb needs an iWork API that does not exist yet, record the request in
   docs/unified-ops/IWORK_API_REQUESTS.md for those owners and keep the gap
   visible in `capabilities()`; do not patch the iWork crates yourself.
   The same rule applies to any leaf crate whose missing write API would need
   a new leaf engine: prefer adding the API in the leaf crate (allowed for
   non-iWork crates) over working around it in `litchi-ops`.
8. Do not make unsupported performance claims in code, docs, or commits.

========================================================================
PERFORMANCE BAR
========================================================================

The unified layer is judged against the direct leaf APIs it wraps. For every
benchmarked scenario, record a paired measurement: leaf-API leg vs unified-API
leg, same corpus, same build, same machine, in tools/perf-baseline (add routes
under new names; never rewrite frozen route-name records).

Budgets (review triggers, not marketing targets; tighten after measuring
noise):

- **Open**: `litchi::open()` performs one detection pass and one container
  index, reusing the prepared source from detection; no part payload is
  decoded at open beyond what the leaf open already decodes (ADR 0030).
  Unified-vs-leaf open overhead ≤ 5% or ≤ 20 µs, whichever is larger.
- **Read / get**: path resolution cost is proportional to the path's depth and
  the scanned axis, never to total document size, and decodes only the parts
  the path touches. The generic node view is produced **on demand**; typed
  `get` returns borrowed/typed nodes without building `format` maps.
- **Query**: results stream through an iterator with early termination
  (`limit`, `first()`); predicates are compiled once per query, not per node.
- **Single-op write**: unified `set` overhead over the equivalent hand-written
  leaf Edit+Commit ≤ 10% on the small-edit scenarios (one paragraph, one cell,
  one shape property).
- **Batch**: ≥ 1,000-operation batches (design §3) scale linearly in operation
  count. Target: a 1k-op homogeneous batch (e.g. 1k cell sets, 1k run
  property sets) within 2× of a single hand-written leaf Edit doing the same
  work. Achieve this by coalescing consecutive operations that target the
  same Edit family into one leaf transaction while keeping per-item error
  attribution and all-or-nothing atomicity; fall back to per-op commits only
  across family boundaries. Measure 1, 10, 100, 1k, 10k ops.
- **Save**: unified save is the leaf save (sibling-temp atomic replacement,
  compressed-entry passthrough for untouched members); zero added copies.
- **Allocation discipline**: property-registry lookup, path parsing of short
  paths, and enum/length/color coercion allocate nothing on the success path.
  Registry tables are `static` (sorted arrays or `match`), not lazily built
  hash maps. "Did you mean" suggestions (edit distance over the registry) run
  only on the error path.
- **Handle cost**: `File` and node handles are cheap to clone (`Arc`-backed
  snapshots, structural sharing); every write returns a new handle at cost
  proportional to the changed closure. Keep `File` itself small (box or `Arc`
  large variants) and `Send + Sync`; add size assertions.
- **Python**: release the GIL around open, query, batch, save, and any
  operation that can exceed ~100 µs. Build result dicts directly from the
  projection (no intermediate JSON string round trip). `raw()` returns
  `bytes` without an extra copy where PyO3 permits. Batch payloads are
  converted to Rust once.

Use the performance program's existing method (docs/GOAL.md, "DECISION AND
REGRESSION RULES"): before/after, ABBA where noise warrants it, machine-
readable results, no single-run claims. Any regression > 5% in an existing
performance-program scenario caused by the trunk migration (M0a) blocks the
merge until explained or fixed.

========================================================================
ERGONOMICS BAR
========================================================================

Rust:

- One entry point: `litchi::open(path)`, `litchi::open_bytes(bytes)`,
  `litchi::create(Format::Docx)`; `OpenOptions` for limits and security
  consumption. The common case needs no options and no format knowledge.
- Paths accept both forms through one trait: `&str` (parsed, typed error
  distinguishes `invalid_path` from `path_not_found`) and typed `Path`/
  `Selector` values (no parsing cost, no stringly-typed hot loops). Parsed
  paths are reusable values (`Path::parse` once, use many times).
- Properties accept both string keys and typed keys generated from the
  registry (e.g. `prop::font::BOLD`), with values via `Into<PropValue>` so
  `true`, `12.pt()`, `"1F4E79"`, and `Color::rgb(..)` all work. A `props!`
  macro is acceptable sugar, not a requirement.
- Batches: both `file.batch(items)` (OfficeCLI JSON-compatible items) and a
  closure form `file.edit(|tx| { tx.set(..)?; tx.add(..)?; Ok(()) })` that
  commits atomically. `BatchMode::BestEffort` returns per-item results.
- `Locked`/`Signed` states are visible in the type and in `capabilities()`;
  calling an edit verb on them is a typed error that says how to consume the
  state, never a silent strip.
- Everything a user can do is discoverable via `capabilities()` and the
  generated schemas; every public item has rustdoc with a runnable example.

Python:

- `litchi.open/create`, immutable handles, context-manager support, dict
  props with dotted keys, OfficeCLI-compatible `apply`/`batch`/`dump`.
- Exceptions: one base `litchi.LitchiError` with `code`, `message`,
  `suggestion`, plus subclasses for the major families (path, property,
  unsupported, io, corrupt) so `except` clauses can be specific.
- Complete `.pyi` stubs, kept in sync by a test; existing read API stays
  source-compatible (§7).

CLI:

- Same verbs, `--json` output identical to the node view; doubles as the
  OfficeCLI example-corpus harness (§11, §14).

========================================================================
EXECUTION PLAN
========================================================================

Follow the milestones of design §15 in order; each milestone ends with a
green workspace and a progress record.

M0  ADR 0033 (draft summarizing the design; proposed until the owner accepts).
    Scaffold `litchi-ops`; register it in the ADR 0024 inventory,
    tools/crate_boundaries.json, and tools/non_iwork_gate.py.
    M0a: flat `File` enum + single `litchi::open()` + unified error; switch
    every in-tree consumer (facade, pyo3-litchi, perf-baseline, tests, docs)
    in the same change (§16.5). Carry facade integration-test semantics over
    verbatim. Measure open/read scenarios before and after.
    M0b: path/selector grammar in `litchi-core::selector`, including story
    axes, with a fuzz target and a round-trip (parse → display → parse) test.
    M0a and M0b are independent and should run in parallel.
M1  Unified read: node view, `get`/`query`, `view(outline|text|annotated|
    stats)`, capability skeleton. Paired leaf-vs-unified read benchmarks.
M2  Write kernel for docx/xlsx/pptx, property registry + coercions, typed
    errors with suggestions, atomic and best-effort batch with family
    coalescing, P0 structural surface. Batch scaling benchmark (1 → 10k ops).
M2b Write handlers for xlsb (tier 1) and legacy OLE doc/xls/ppt (tier 2).
M6  Delayed tier 3, after M2b: RTF; ODF (odt/ods/odp first, then
    odg/odc/odi/odb/oth/odm/odf); iWork (pages/keynote/numbers, via their
    public edit APIs). Each stream starts from the leaf write-capability
    inventory (record 0002), closes leaf API gaps in the leaf crate where it
    owns them, and reports remaining element-level gaps in `capabilities()`.
M3  Python write bindings, CLI, schema export, capabilities introspection.
M4  CSV import, formula write/eval, typed chart and SmartArt CRUD, then the
    three new engines (find/replace, template merge, field/TOC refresh) —
    split into M4a/M4b if cost warrants.
M5  Acceptance closure (§14): OfficeCLI v1-scope corpus green with every
    delta recorded, CRUD checklist evidence, full-corpus open→empty
    edit→save byte regression, registry/schema golden snapshots, pytest mirror.

Decision points (§16.6 and "Open micro-decisions"): do not block on them.
Adopt the recommended default below, record it in a numbered decision record,
list it in the progress log under "awaiting owner confirmation", and continue.
Revisit only if the owner overrides.

- Legacy package retention: superseded by the owner override. Every `File`
  variant, legacy binary included, retains the leaf's editable owner (source-
  backed where the leaf has one, so retention stays lazy and cheap).
- `OpenOptions` shape: one `OpenOptions` struct with the common knobs (limits,
  security consumption, execution context) plus optional per-format sub-
  options; plain `open()` covers the common case.
- iWork facade timing: iWork joins the unified `File` enum in tier 3 (M6);
  until then `litchi::iwork::Document` stays put (§16.4).
- Verb names: `apply` (not `send`); find/replace as OfficeCLI-style `find=`/
  `replace=` props on `set` first, a dedicated verb only if usage demands it;
  `merge` means template merge.

========================================================================
WORKING RULES
========================================================================

- Act as coordinator: dispatch implementation, review, test, and benchmark
  subagents in parallel; implementing subagents use `model: "opus"`.
- Each implementer works in its own worktree and branch under
  /home/zhuhe/code/litchi-worktrees/ (disk-backed; never the tmpfs
  scratchpad), branched from feat/unified-ops-api, and is merged back by the
  coordinator after review. Implementers do not edit shared logs; they hand
  log sections to the coordinator.
- Keep a progress log at docs/unified-ops/PROGRESS.md and numbered records at
  docs/unified-ops/NNNN-*.md (what changed, evidence, deltas vs OfficeCLI,
  what is not claimed). Performance evidence goes under
  docs/unified-ops/results/.
- Periodically merge feat/office-format-completeness into this branch to
  stay current; resolve conflicts in favour of the base branch's leaf-crate
  changes and adapt the unified layer.
- End each turn with work committed and temporary files cleaned; sweep
  finished worktrees and target directories during long turns, and check free
  disk space before launching a parallel wave.

========================================================================
VERIFICATION GATES (every milestone)
========================================================================

cargo fmt; cargo check across feature combinations; warning-denied clippy;
rustdoc with warnings denied; crate-boundary checker; unit, integration, doc,
compile-fail tests; the new parser fuzz target; pytest for pyo3-litchi;
the paired leaf-vs-unified benchmarks for the milestone's scenarios; and the
OfficeCLI parity checklist delta table (generated, not hand-maintained, from
the registry vs schemas/help).

========================================================================
DEFINITION OF DONE
========================================================================

- `litchi::open()` and the flat `File` enum are the single entry point; the
  old wrapper types are deleted; every in-tree consumer is migrated.
- Every claimed format is readable through `get`/`query`/`view` AND writable
  through `set/add/remove/move/swap`, `batch`, `save`, and `create`; docx/
  xlsx/pptx additionally cover the P0 surface of §13. No claimed format
  returns whole-format `unsupported_capability`; remaining element/property
  gaps are enumerated in `capabilities()` with follow-ups.
- The property registry is the single source of truth for coercion,
  validation, schemas, capabilities, Python introspection, and suggestions,
  guarded by golden snapshots.
- Python and CLI surfaces are complete, typed, and tested.
- §14's three acceptance layers pass for the v1 scope, with every OfficeCLI
  delta recorded as wontfix or deferred.
- Paired benchmarks show the unified layer within the budgets above on every
  measured scenario, with the 1k-op batch budget met; any miss is explained
  in a record with a follow-up.
- No accepted ADR, preservation, security, determinism, or boundary
  requirement regressed.

Start by reading the design, the ADRs, the current facade (crates/litchi/src),
pyo3-litchi, and OfficeCLI's schemas; then write ADR 0033 and launch M0a and
M0b in parallel. Do not ask the owner to choose a starting point.
