# 0355 — W1 DOCX: immutable package, scoped edit families, `Refused`, limits through options

- Status: implemented on `uo2/w1-docx`, cut from `feat/unified-ops-api`
  at `e652864de8`; trunk merged at `88d253579b` (records 0352, 0353) and
  at `4d59677f78` (record 0354) before the final gates (§6).
- Date: 2026-10-20
- Input: record [0351](0351-api-contract.md) §2 (the normative contract),
  §3.1 (the DOCX inventory), §4.3 (the W1 gate registrations) and §5 (the
  W1 routes).
- Owned: `crates/litchi-docx`; the litchi-opc helpers the graph patch
  needs (`OpcPackage::changed_parts`, `PartPayload::same_allocation`,
  `Relationships::same_entries`); the DOCX consumers in litchi-ops
  (`format/docx*`, the DOCX `Classify` arms of `error.rs`, `raw/validate.rs`,
  DOCX tests and examples), the facade, `tools/perf-baseline`, the DOCX
  profiles of `tools/api_glossary.json`, the W1 line of
  `docs/API_GLOSSARY.md` §2a, `tools/unified_bench_matrix.py` (`p0355`),
  this record and `results/0355/`.
- `performance_claim`: measured (§5). No W1 route moves by more than 1.2%
  between the wave's base and head (leaf and unified legs, every route).
  The unified 1k paragraph-text batch misses its 2.0× budget by about 300×
  at base and head alike: a pre-existing litchi-ops cost (follow-up
  0355-P1), not a W1 change.

## 1. Result

| 0351 rule | DOCX before | DOCX after |
|---|---|---|
| §2.1 immutable owner | 85 public `&mut self` methods on `Package` | none on `Package`, `source_backed::Package` or `Document` (gate `immutable_owners`) |
| §2.1.2 outputs on `&self` | `save*`, `to_*`, `write_to` took `&mut self` and flushed a hydrated writer | `&self`; the package holds no unmaterialized state, so repeated output is byte-identical |
| §2.1.3 signature verbs | `sign`/`resign*`/`unsign` mutated in place | `sign*`/`resign*` return `Result<Package>`, `unsign` returns `Package` |
| §2.1.4 history, `custom_props_mut` | `&mut self` history verbs; `custom_props_mut` | `undo_document`/`redo_document`/`publish_document_commit_with_history(&self) -> Result<document::Commit>`; `custom_props_mut` deleted (`metadata::Edit`) |
| §2.1.7 drafts | `writer::Mutable*` (44 + 202 methods) | `builder::{Document, Paragraph, Run, Table, Row, Cell, Style, Theme, …}`; `writer` is private |
| §2.3 scoped families | direct `put_*`/`remove_*`/`set_*`/`apply_*_patch` mutators | 23 families registered with `Edit`/`Patch`/`Commit`/`Diagnostics`/`Error`/`LimitKind` (§2) |
| §2.3 no `Transaction`/`rollback`/`is_changed`/`snapshot` | 70 public names (0351 inventory: 37 commit snapshots, 16 `is_changed`, 16 `Transaction` types, 1 `rollback`) | none; banned in both `docx` profiles |
| §2.4 `UnsafeEdit` | one variant, 130 construction sites | `Error::Refused(Refusal)`, `Error::PatchConflict`, `Error::Internal`; typed refusal enums have `kind()` |
| §2.5 limits | 50 public `*_with_limits` paths (0351 inventory) | none; `SemanticLimits` on `OpenOptions`, kept by the package; record parsers take limits as a required argument |

## 2. What landed

Three commits, each migrating every consumer with its API change:

1. `ec8e157b62` — `UnsafeEdit` replaced (0351 §2.4).
2. `efca0f7c9c` — immutable `Package`, package edit families, builder
   drafts (§2.1, §2.3).
3. `92473427cd` — limits through `SemanticLimits`, the remaining bans,
   the gate registration (§2.5, §4.3).

### Refusals (§2.4)

`litchi_docx::Error::UnsafeEdit` is deleted. Each of its 130 sites maps by
what its reason says:

- a route or content refusal is `Error::Refused(litchi_core::edit::Refusal)`
  with a `RefusalKind` (`Error::refusal()` reads it);
- a patch whose base does not match is `Error::PatchConflict { operation,
  reason }`;
- a missing admission or ledger that the crate itself should have provided
  is `Error::Internal { operation, reason }`;
- a changed write of a signed source keeps the typed
  `OpcError::SignedSourceRequiresExplicitPolicy`.

`document::Refusal`, `document::TransferRefusal` and the source-backed
`paragraph_copy::Refusal` and `tail_append::Refusal` gain `kind()`.
litchi-ops classifies a DOCX refusal by its kind alone (`Refused` →
`refusal.classify()`; `document::Error::Refused { reason }` →
`reason.kind().classify()`). `PatchConflict` is `invalid_input` in every
DOCX arm; the document and metadata arms used to report `io_error`.

### Immutable owner and families (§2.1, §2.3)

- **Outputs.** `save`, `save_with_durability`, `save_plain*`,
  `save_encrypted*`, `save_reencrypted*`, `to_bytes`, `to_plain_bytes`,
  `to_encrypted*`, `to_reencrypted*`, `write_to` and `to_plain_stream` take
  `&self`. The fields that carried unpublished state (`mutable_doc`,
  `raw_edit_committed`, `custom_props_dirty`, `font_embedding`) and
  `WriteRollbackGuard` are gone. A plain write is `OpcPackage::write_to` of
  the snapshot.
- **Graph patch.** `package::graph_patch::GraphPatch` holds the exact
  before and after state of every part, package relationship set and
  owning relationship set an edit changed. Untouched parts share their
  payload allocations with the source, so the difference
  (`OpcPackage::changed_parts`) is found without decoding them. `check`
  refuses a target that does not hold the before state, including an
  unchanged owner whose relationships to a changed part differ (the
  topology the change was made in). `inverse` swaps the states.
- **Family engine.** `package::family::Staged` pairs the source package
  with a working copy and commits through the graph patch; a changed
  commit of a signed source is refused. The `graph_family!` and
  `graph_edit!` macros declare a family's `Patch`, `Commit` (an alias of
  `litchi_core::edit::Commit<Package, Patch>`), `Diagnostics`, `Error`,
  `LimitKind`, `Edit` (`package()`, `reset()`, `commit()`) and
  `Package::apply_x(&Patch)`.
- **Families.** `ribbon`, `vba` (`vba-inspection`), `attached_template`,
  `document_variables`, `opc` (the raw escape; `edit_opc(&self)` returns an
  `Edit` whose commit publishes a `Package`), `task_panes`
  (`edit_task_panes_with_options(task_panes::EditOptions)`), `font`, `web`,
  `glossary`, `custom_xml`, `bibliography`, `mail_merge`, `numbering`,
  `chart`, `styles::effects` (owner-bound typed patch), `revision::authoring`
  (`edit_tracked_revisions`), `revision::conflict` (`edit_conflicts`,
  `edit_story_conflicts`), `content_control`, `settings::extensions`,
  `package::story` and `ink` (`edit_ink_with_options(ink::EditOptions)`,
  `apply_durable_ink(&self) -> Result<Package>`), plus the existing
  `document` and `metadata`. Each exports the six names; `family_methods`
  checks the triad accessors.
- **Primary and history.** `Package::apply` returns `document::Commit`;
  `apply_durable_document(&self)` and the history verbs return
  `Result<document::Commit>` built from `Patch::transition`.
- **Builder.** `Builder { package, document, font_embedding }` owns the
  draft document (`Builder::document()`), font embedding and the alt-chunk
  verbs; `build()` returns the `Package`. The drafts and the authoring
  vocabulary are re-exported from `litchi_docx::builder` without the
  `Mutable` prefix. An opened package never hydrates into the writer.
- **Detached fragments.** The part-level transactions without a package
  (font OpenType, run symbols, collapsed paragraphs) keep only their
  `Snapshot`; `Section::apply_footnote_columns_patch` is deleted. The
  internal drafts are `Draft`/`DraftCommit`/`DraftPatch`, crate-private.

### Limits through options (§2.5)

`litchi_docx::SemanticLimits` (revisions, sections, stories, conflicts,
tracked revisions, content controls, ink, task panes, hyperlinks, story
hyperlinks, story text, paragraph copy, tail append, embedded objects)
has consuming setters named by the noun and `<noun>_value()` getters.
`OpenOptions::semantic_limits(..)` sets it; `Package` and
`source_backed::Package` keep it (`semantic_limits()`,
`with_semantic_limits(&self, ..) -> Self`, which shares the snapshot), and
their plain verbs read it. Every public `*_with_limits` method is deleted
(the bounded forms are crate-private `*_bounded`). Record parsers take
their limits as a required argument (`content_control::Inventory::parse`,
`section::Inventory::parse`, the `Snapshot::from_xml` constructors,
`SourceDrawing::scan`, `validate_read_at(source, read, docx, report)`).
litchi-ops keeps its own `*_with_limits` names (W8) and passes the limits
through `with_semantic_limits`.

### Gate (§4.3)

The `docx-minimal` and `docx-full` profiles of `tools/api_glossary.json`
register:

- `immutable_owners`: `litchi_docx::Package`,
  `litchi_docx::source_backed::Package`, `litchi_docx::Document`;
- bans: `UnsafeEdit`, `rollback`, `Transaction` (exact and as a suffix),
  `is_changed`, `into_snapshot`, the `Mutable` prefix, the `_with_limits`
  suffix, `custom_props_mut`, `into_workbook` (62 bans in each profile);
- families: the 23 of §2 (22 in `docx-minimal`, which has no `vba`).

`docs/API_GLOSSARY.md` §2a: the W1 line is ✅.

## 3. Decisions and deltas

- **Protected documents report `document_protected`.** The publication
  refusal of a document or write-protected package is a `Protected`
  refusal, so the unified code is `document_protected` (was
  `unsupported_capability`). Five corpus files change in
  `results/0289/fidelity.tsv` (re-blessed); the OfficeCLI corpus is
  re-checked in §6.
- **`document::Refusal::kind` keeps the established codes.** Only
  `InvalidPropertyValue` is `InvalidValue`; every other reason is
  `UnsupportedOperation` or `UnsupportedContent`, both
  `unsupported_capability`, as before. Finer kinds (`NotFound` for the
  `*NotFound` reasons, `UnsupportedTarget`, `UnsupportedProperty`) would
  change OfficeCLI-visible codes; follow-up 0355-F3.
- **`unsign` is infallible.** It returns `Package`, not
  `Result<Package>` (0351 §2.1.3): removing a signature from the
  in-memory graph cannot fail.
- **Removed behaviour.** Alt-chunk verbs exist on `Builder` only; an
  opened package no longer edits alt chunks through the writer.
  Per-paragraph revision limits (`revisions_with_limits` on
  `Paragraph`/`Table`/`Row`/`Cell`) are no longer configurable; the
  package's `SemanticLimits::revisions` bounds `Document::revisions()`.
  The Builder-to-family route of content controls and conflicts no longer
  takes a `max_output_bytes` option. The tests of removed behaviour are
  deleted: alt-chunk edits of opened packages, `Builder::package`,
  retrying a failed save of unmaterialized writer state, the detached
  output-limit tests of the hydrated writer, and the pending-facade
  content-control test.
- **Glossary cells.** Following the wave's coordination rule, this record
  flips only the W1 line of §2a. The other cells record 0351 §4.3 assigns
  to W1 (§1 *Immutable owner snapshot* (docx), §2 *New snapshot accessor*,
  *Diagnostics*, *Split commit*, *Patch application*, *No-op* (DOCX
  scoped), *Mutable writer* (DOCX `UnsafeEdit`)) are ready to flip in the
  coordinator's closure.

## 4. Byte identity

- Every DOCX test passes (`cargo test -p litchi-docx --all-features`),
  including the round-trip, no-op and CRUD suites; the edits only renamed
  the APIs they call, except the deletions of §3.
- A no-op family commit returns a package equal to its source
  (`GraphPatch::is_noop`); a rewritten part with equal bytes is no change.
- The bench (§5) asserts that two `to_bytes` calls on one package return
  identical bytes, at base and head, on both legs.
- `corpus_fidelity` (`open_empty_edit_save_is_byte_identical`) reports the
  same no-op, revert and structure outcomes for every file; only the five
  protected-document refusal codes of §3 change.

## 5. Measurements

`docx_w1_contract_bench` (`crates/litchi-ops/examples/`): each pair times
the unified call against the leaf call doing the same work, ABBA within a
run, 31 rounds after 5 of warm-up, medians. The fixture is `create()` with
1,000 paragraphs. The leaf legs use only API present before and after W1,
so the same source builds at the wave's base. Release builds, pinned
(`taskset -c 27`), runs ordered base, head, head, base; build commands and
binary SHA-256s in [`results/0355/binaries.sha256`](results/0355/binaries.sha256),
raw lines in `results/0355/docx-w1-contract-{base,head}.txt`.

| Route | Leaf base → head (µs) | Unified base → head (µs) | Unified/leaf at head |
|---|---|---|---|
| open (1,000 paragraphs) | 65.1 → 65.9 (1.011) | 39.2 → 39.1 (0.997) | 0.59 |
| set paragraph text | 563.7 → 568.2 (1.008) | 65.7 → 65.0 (0.989) | 0.11 |
| add paragraph | 943.6 → 954.2 (1.011) | 127.5 → 126.4 (0.991) | 0.13 |
| set `/` title | 12.1 → 11.9 (0.979) | 12.4 → 12.1 (0.980) | 1.02 |
| `File::edit`, 1k paragraph text sets | 4,811 → 4,805 (1.00) | 1,461,944 → 1,453,840 (0.994) | **302** |
| `File::batch`, 1k paragraph text sets | 4,804 → 4,797 (1.00) | 1,460,676 → 1,456,459 (0.997) | **304** |
| save (`to_bytes`) after one set | 70.8 → 70.1 (0.989) | 68.6 → 68.0 (0.992) | 0.97 |
| two `to_bytes` of one package | 141.8 → 140.1 (0.988) | 138.1 → 136.1 (0.986) | 0.97 |

- **No W1 regression.** Every leg is within 1.2% of its base (budget 1.05).
  The second save of one package costs the same as the first (70 µs each,
  head), no more than one save cost before the change. `edit*()` shares
  the snapshot; no payload is copied.
- **The unified legs of the single-op rows are faster than the leaf legs**
  because the leaf leg builds the document projection on every
  `Package::edit()`, which the unified handle caches.
- **Batch miss (pre-existing, 0355-P1).** A 1k paragraph-text batch costs
  1.46 s through `File::edit` or `File::batch`, against 4.8 ms for one leaf
  edit and about 65 ms for 1,000 separate unified sets. The cost is
  O(paragraphs × items): about 1.5 µs per paragraph per item (100 items on
  1,000 paragraphs: 151 ms; 1,000 on 2,000: 2.94 s). A callgrind probe
  (300 items on 300 paragraphs) puts 58% of instructions in
  `docx::nav::load_body` under `path::resolve::get`: each item's path
  resolution reloads the staged body (`Snapshot::from_shared_xml`,
  `blocks::scan::index`). Base and head are equal, so W1 neither caused
  nor fixed it; the 0337 batch rows (run bold sets) are a different
  route.

Matrix rows: `p0355` in `tools/unified_bench_matrix.py` (the unified/leaf
rows at head and an after/before row for each leg).

## 6. Gates

Run on the final tree, after the trunk merge at `4d59677f78` (trunk tip
`5f6fb75164`), with the wave's `env.sh` (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`). Every command exited 0 except one timing-sensitive
OPC test under host load (below).

| Gate | Command | Result |
|---|---|---|
| format | `cargo fmt --all --check` | clean |
| clippy | `cargo clippy -p litchi-docx -p litchi-opc -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | clean |
| clippy, ops features | `cargo clippy -p litchi-ops --no-default-features [--features docx] --all-targets -- -D warnings`; `cargo clippy -p litchi-docx --no-default-features --all-targets -- -D warnings` | clean |
| clippy, ODS lib | `cargo clippy -p litchi-ods --lib -- -D warnings` | clean |
| docs | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-opc` | clean |
| DOCX tests | `cargo test -p litchi-docx --all-features` | 2,454 passed |
| OPC tests | `cargo test -p litchi-opc --all-features` | 1,189 passed; `source_backed_batch::batch_cancellation_during_a_gated_load_returns_cancelled_and_joins_workers` failed once at load average 11 (exit 101) and passed 5 of 5 re-runs and in the pre-merge run; W1 does not touch the source-backed batch |
| ops tests | `cargo test -p litchi-ops --all-features`, lib and the 167 test targets in chunks of 28 | 2,422 passed |
| facade, CLI | `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 209 and 34 passed |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` | 301 tests OK |
| perf-baseline | `cd tools/perf-baseline && cargo check --locked --all-targets` | clean |
| native-resave | `cd tools/native-resave && cargo check --locked --features docx` | builds (one unused-import warning, also at base) |
| boundaries | `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | pass |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 44 profiles, 2,741 bans, 294 families, 62 immutable owners |
| goldens | `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every litchi-ops feature but `encryption` | no golden changed |
| capability delta | `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI`; `--check` | unchanged |
| OfficeCLI corpus | `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py`; `--check` | no corpus file changed |
| tool tests | `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate`; `tools.test_check_api_glossary` | 1,269 and 27 tests OK |
| matrices | `python3 tools/unified_bench_matrix.py --check`; `python3 tools/check_feature_matrix.py` | pass |

## 7. Follow-ups

- **0355-P1** — the unified DOCX batch reloads the staged body for each
  item's path resolution (§5): cache the body navigation across the items
  of one batch (litchi-ops `format/docx/nav.rs`).
- **0355-W1b** — the source-backed streaming families (variables, section
  layout, sanitize, redact, story hyperlinks, story text, glossary batch,
  SVG, tail append, paragraph copy, alt-chunk publication) still publish
  part-level commits through `publish_*_to_stream`; the bans pass because
  their names changed (`variables::Edit`, `Patch::is_noop`), not because
  they are registered families. Convert them to `edit_x`/`apply_x`
  families with the six names, or record why a streaming publication
  stays.
- **0355-F2** — the family `LimitKind` enums are empty: the families'
  budgets are reported through crate `Error` variants with resource names
  (`InkLimit`, `TaskPanesLimit`, …). Give them typed kinds.
- **0355-F3** — refine `document::Refusal::kind` (`NotFound`,
  `UnsupportedTarget`, `UnsupportedProperty`) with an OfficeCLI corpus
  re-bless.
- **0355-F4** — restore what §3 removed if a consumer needs it: alt-chunk
  edits of opened packages as a family, a revision-limits option on
  `Document`, `max_output_bytes` on the content-control and conflict
  families.
