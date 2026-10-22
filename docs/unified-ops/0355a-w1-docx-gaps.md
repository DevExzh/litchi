# 0355a — W1 DOCX gaps closed: source-backed families, typed limits, restored capabilities, batch budget

- Status: implemented on `uo2/w1-docx` on top of the W1 merge (`3e686506f3`);
  trunk merged at `fcf1f8ff44` (records 0353a, 0356, 0357) before the final
  gates (§6).
- Date: 2026-10-20
- Input: the coordinator review of record [0355](0355-w1-docx-contract.md):
  items F4 (restore removed capabilities), W1b (source-backed streaming
  families follow §2.3), F2 (typed `LimitKind`), F3 (one refusal
  taxonomy), P1 (the 1k paragraph-text batch within 2.0×), and the flaky
  `source_backed_batch` cancellation test.
- `performance_claim`: measured (§5). The unified 1k paragraph-text batch
  is 1.03× (`File::edit`) and 1.05× (`File::batch`) one leaf edit (budget 2.0×; 302× at record 0355). Every
  source-backed family writes byte-identical output to the record-0355
  publisher, at 0.13–1.02× its time.

## 1. The source-backed family shape (for other waves to follow)

This is the shape every source-backed (lazy, `ReadAt`-backed) DOCX edit
family now has; the W2 PPTX `SourceBacked*` editors mirror it.

```text
source_backed::Package::edit_x(&self, ..) -> x::Edit          // O(1), shares the package
x::Edit::{verbs..}(&mut self, ..)                              // staging only
x::Edit::commit(self) -> Result<x::Commit, x::Error>           // validates, writes nothing
x::Commit = litchi_core::edit::Commit<source_backed::Package, x::Patch>
    .package()       // the committed package (a new owner; the source is unchanged)
    .patch()         // x::Patch: is_noop(), inverse()
    .diagnostics()   // litchi_core::edit::Diagnostics: changed()
source_backed::Package::apply_x(&self, &x::Patch) -> Result<x::Commit, x::Error>
source_backed::Package::write_to(&self, impl Write) / save(..)  // the only output
```

How a committed package works (`source_backed/deferred.rs`):

- **Commit validates, output streams.** `commit()` runs every check the
  former `publish_x_to_stream` ran before output (source conflict,
  protection, preservation, signature, encryption, limits) and returns a
  new `Package` holding the pending output: the same byte-preserving source
  splice (`write_part_overlay*`) the old publisher wrote, run when the
  package is written. A commit followed by `write_to`/`save` therefore
  streams exactly as before, with the same output bytes, memory and
  speed; untouched members stay raw-copied references into the source.
- **Reads materialize once.** Reading the committed package (or editing it
  again) writes its output into memory once and serves every later read
  from it: chained edits and reads work, at O(archive) memory for that
  package only. A caller that only saves never pays it.
- **Patch.** A family's `Patch` wraps its detached record patch (the exact
  before and after bytes of what it changed) and, for a commit, the exact
  package pair. `apply_x` on a package that writes exactly the commit's
  before artifact (a share of it, or a reopened copy of its bytes, compared
  by length and SHA-256) returns the exact after package: `inverse()` on a
  reopened output restores the original archive byte for byte, streamed
  through the current package's execution context
  (`SourceBackedPackage::restore_source_artifact_to_stream`, both artifacts
  authenticated again as they are written). Any other package gets the
  record patch's own replay rule: content (the before bytes, selector and
  limits) for most families; whole-artifact binding for the durable
  plain-paragraph copy and removal patches; source lineage for the main
  document, story-text and story-hyperlink patches. `From<record::Patch>`
  replays a detached record patch; `Patch::durable()` (copy, removal,
  paragraph stream) returns the wire-serializable record patch bound to the
  artifact it applies to.
- **Refusals before output.** A changed commit of a signed or encrypted
  source is refused at commit (`OpcError::SignedSourceRequiresExplicitPolicy`),
  not at write. A source that changes between commit and write is refused
  by the write (`SourceChanged`, or the family's artifact-fingerprint
  conflict) before any byte is emitted.
- **Macro.** `source_family!(what, Change[, Error][, diagnostics D])`
  declares `Patch`, `Commit`, `Diagnostics`, `Error`, `commit_change` and
  `apply_change` over a family's `Change` (`is_noop`, `inverse`,
  `publish(&Package) -> Package`); the family adds `Edit` and `LimitKind`.
  A family with richer diagnostics (SVG: operations and selected pictures)
  names its own `D` with a `changed()` accessor.
- **Main document.** `source_backed::Package::apply_document(&document::Patch)`
  returns a `document::Commit<source_backed::Package>` whose patch carries
  the package pair; `Package::edit()?.commit()` does the same.

Families (all under `litchi_docx::source_backed`): `document_variables`,
`section_layout`, `hyperlink_detachment`, `hyperlink_redaction`,
`story_hyperlink_redaction`, `alt_chunk_payload`, `story_text_edit`,
`glossary_text`, `plain_paragraph_copy`, `plain_paragraph_removal`, `svg`,
`tail_append` (`Edit::prepare()` → `Plan` with its proofs → `commit()`)
and `paragraph_stream` (`ParagraphStreamEdit`). The source document edit
already committed to a package; its detached-snapshot stream publishers are
removed. No public `*_to_stream` publisher remains in `litchi-docx`.

A prepared splice can be written more than once: litchi-opc gained
`SourcePartSplicePlan::{into_owned, duplicate}`, so a committed tail append
keeps the plan it prepared (no second scan) and every save streams it.

## 2. What landed

| Commit | Item |
|---|---|
| `39f949aee2` | flaky test: `source_backed_batch::batch_cancellation_during_a_gated_load…` holds its read gate (`arm_holding`) until it releases it; pinned to one CPU it failed 33 of 40 runs before, 0 of 40 after |
| `a7b9cc5e38` | **P1**: direct body-paragraph `text` sets join the coalesced DOCX batch run (§4) |
| `290f12634b` | **F3**: `document::Refusal::kind` classifies every refusal (`NotFound`, `InvalidValue`, `UnsupportedTarget` for section breaks and field interiors, `UnsupportedProperty`, `UnsupportedContent`); no OfficeCLI corpus line changed, so `tools/officecli_dispositions.json` is untouched and `--check` passes |
| `9036be2293` | **F4**: `Package::edit_alt_chunks()` → `alt::Edit` (add, insert, replace, remove, `move_to`) with `apply_alt_chunks` and `alt::LimitKind`; `revision::Reader` (per-element revision budgets, and `SemanticLimits::revisions` on the document reads); `content_control::EditOptions` and `revision::conflict::EditOptions` with `max_output_bytes`; the removed tests restored on the package routes |
| `196986eb9a` | **F2**: a typed `LimitKind` per family; every limit error carries `kind` (crate variants `InkLimit`, `RevisionLimit`, `GlossaryLimit`, `ContentControlLimit`, `ConflictLimit`, `ChartLimit`, `FontLimit`, `BibliographyLimit`, `WebLimit`, `StoryLimit`, `NumberingLimit`, `MailMergeLimit`, `StylesWithEffectsLimit`, `SettingsLimit`, `SectionInventoryLimit`, `ExternalHyperlinkDetachmentLimit`, `ExternalHyperlinkRedactionLimit`, `AltChunkLimit`; module errors of story text, paragraph copy, tail append and the paragraph stream). `Display` keeps the resource string; litchi-ops classifies them `limit_exceeded` |
| `8274a49a7a`, `a17e185665`, `21ce9d4325` | **W1b**: the 13 source-backed families of §1; every `publish_*_to_stream` and `*Publication` removed; exact restore, durable patches, tests and `tools/perf-baseline` migrated; the families registered in both `docx` glossary profiles, with bans on `*_to_stream`, `publish_*` (source-backed) and `*Publication` |
| `3911f37d91`, `fcf1f8ff44` | trunk merges |
| `results` commit | `source_family_bench`, results, matrix rows, this record |

## 3. Decisions and deltas

- **One output.** `Package::write_to`/`save` is the only output of every
  family. A committed package holds its pending splice; reads materialize
  it once. The splice is the one the publisher ran, so the output bytes
  are the same (§5) and an untouched member is still a raw copy.
- **Refusals move to commit.** Protection, signature, encryption, source
  conflicts and limits refuse at `commit()` or `apply_x()`, before any
  package exists, not at write. Tests that asserted "refused, nothing
  written" now assert the refusal at commit. A source that changes after
  commit is still refused by the write before its first byte.
- **Replay rules are the record patch's own.** A family patch replays by
  content where its record patch is content-addressed (variables, section
  layout, sanitize, alt-chunk payload, story text, glossary batch, SVG),
  and keeps the stricter binding where the record patch has one: the
  durable plain-paragraph copy and removal patches stay whole-artifact
  bound (a package with the same main story but another member is
  refused), the main document and story hyperlinks stay lineage bound. A
  detached section-layout patch therefore now replays on a second open of
  the same bytes (it was refused as a foreign source); an incompatible
  section base reports the shared `PatchConflict`, classified as
  `invalid_input`.
- **Exact inverse by artifact identity.** The inverse of a commit applied
  to a reopened copy of its output restores the source archive byte for
  byte, streamed through the reopened package's execution context, with
  its output budget and cancellation (`managed_reopened_inverse_*` tests).
  The two redactions and the stream appends are forward-only: their
  inverse restores only by artifact.
- **Durable patches kept.** `plain_paragraph_copy::Patch::durable()`,
  `plain_paragraph_removal::Patch::durable()` and
  `paragraph_stream::Patch::durable()` return the wire patches the
  publications used to return, bound to the written artifact; the stream's
  candidate identity is the one its write measured (no second hash).
- **Typed errors through the writer.** `Error::ParagraphStream` carries a
  paragraph-stream error through `Package::write_to`, and
  `From<crate::Error>` unwraps it, so a write still reports
  `StreamError::Replay`/`Limit`/`Opc` as before.
- **Clean indexes.** Writing a committed package releases the base's clean
  paragraph indexes first, as the publishers did.
- **SVG.** `svg::Diagnostics` keeps the operation and selected-picture
  counts. The single-picture SVG edit types are gone; a one-selector
  `edit_svg([selector])` (family) or `edit_svg_attachments([selector])`
  (record batch) replaces them.
- **Empty `LimitKind`.** The OOXML-common families (task panes, custom
  XML, ribbon, VBA), the raw OPC escape and SVG (bounded by the package's
  source limits) keep an empty `LimitKind`; their limit errors are typed in
  their own crates (a W4 dependency).
- **perf-baseline.** `tools/perf-baseline/src/docx_family_publication.rs`
  keeps the routes' publication-shaped evidence on the families: one
  commit, one streaming write of the same splice, nothing buffered. The
  section-layout control that expected a foreign-source refusal now checks
  that the content replay writes the commit's bytes.

## 4. P1: the 1k paragraph-text batch

Record 0355 measured `File::edit`/`File::batch` of 1,000 paragraph-text
sets at 1.46 s, 302× one leaf edit: each item resolved its path by
reloading the staged body. The coalesced DOCX batch run (litchi-ops
`format/docx/write/coalesce.rs`) now also takes `text` on a direct
`/body/p[N]` path: a body paragraph's ordinal survives a text replacement,
so the run resolves every text set against its base, tracks the paragraphs
whose text it replaced, and closes itself before a target it cannot admit
(an element path, or a path inside a paragraph it rewrote). The body block
index is cached on source-backed handles too (`nav.rs`). Equivalence:
`tests/docx_batches.rs::coalesced_paragraph_text_sets_match_single_operations`
compares the batch with the same sets applied one by one, on both owners
and both modes, plus the closure edit and the refusals.

## 5. Measurements

Release builds (`CARGO_INCREMENTAL=0`), pinned (`taskset -c 27`), ABBA
(base, head, head, base). Base: the record-0355 tree `3e686506f3`. Head:
this branch after the trunk merge (`fcf1f8ff44`) plus this record's
example. Build commands and SHA-256s in
[`results/0355a/binaries.sha256`](results/0355a/binaries.sha256).

### `docx_w1_contract_bench` (P1)

Each pair times the unified call against the leaf call doing the same
work (31 rounds after 5 warm-ups, medians; µs). Raw lines:
`results/0355a/docx-w1-contract-{base,head}.txt`.

| Route | Leaf base → head | Unified base → head | Unified/leaf at head |
|---|---|---|---|
| open (1,000 paragraphs) | 65.5 → 65.5 (1.000) | 39.1 → 38.8 (0.993) | 0.59 |
| set paragraph text | 567.3 → 563.2 (0.993) | 65.4 → 64.4 (0.984) | 0.11 |
| add paragraph | 956.4 → 936.9 (0.980) | 128.2 → 126.0 (0.983) | 0.13 |
| set `/` title | 12.1 → 11.9 (0.989) | 12.4 → 12.4 (0.998) | 1.03 |
| `File::edit`, 1k paragraph text sets | 4,757 → 4,705 (0.989) | 1,473,766 → 4,862 (0.003) | **1.03** |
| `File::batch`, 1k paragraph text sets | 4,818 → 4,701 (0.976) | 1,473,219 → 4,932 (0.003) | **1.05** |
| `File::batch`, 1k sets, eager owner | 4,789 → 4,729 (0.987) | 1,645,722 → 5,763 (0.004) | **1.22** |
| save (`to_bytes`) after one set | 70.9 → 70.6 (0.996) | 68.8 → 68.5 (0.996) | 0.97 |
| two `to_bytes` of one package | 141.6 → 141.0 (0.996) | 137.5 → 136.9 (0.995) | 0.97 |

- **P1 met.** The 1k paragraph-text batch is 1.03× (`File::edit`) and
  1.05× (`File::batch`) one leaf edit on the source-backed owner, 1.22× on
  the eager owner (budget 2.0×), from 302×.
- **No other route moved**: every leaf leg and every other unified leg is
  within 2.4% of its base, all faster.

### `source_family_bench` (W1b)

`crates/litchi-docx/examples/source_family_bench.rs`: open a source-backed
package over a 2.1 MB archive (1,000 paragraphs, header, settings, one
external hyperlink, a 2 MiB stored image), make one edit, commit, write to
a pre-sized `Vec`; 3 warm-ups, 15 samples, medians. The base variant
([`results/0355a/source_family_bench.base.rs`](results/0355a/source_family_bench.base.rs))
makes the same edits through the record-0355 publishers. Medians of the
two runs per side, µs:

| Scenario | Base (publisher) | Head (family) | Head/base |
|---|---|---|---|
| document paragraph text (`edit()`, control) | 798.5 | 810.4 | 1.015 |
| document paragraph text (detached commit) | 778.7 | 797.2 | 1.024 |
| document variables set | 1,100.7 | 142.3 | 0.129 |
| section layout columns | 4,779.8 | 4,831.8 | 1.011 |
| header story text | 1,103.4 | 151.6 | 0.137 |
| plain paragraph copy | 4,736.9 | 3,772.0 | 0.796 |
| tail append one paragraph | 3,557.2 | 3,569.4 | 1.003 |
| story hyperlink redaction | 4,438.2 | 4,417.9 | 0.995 |

- **Byte identity.** Each scenario's output SHA-256 is the same in all
  four runs, base and head (raw lines).
- **Faster where the publisher hashed.** The variables, header-story and
  plain-copy publishers fingerprinted their whole output to authorize the
  returned inverse (about 1 ms for 2 MB); the family keeps the package
  pair instead and hashes only when an inverse needs a reopened copy's
  identity.
- **The rest is within noise** (≤ 2.4%, budget 1.05): the splice and its
  checks are the same code.

Matrix rows: `p0355a` in `tools/unified_bench_matrix.py`.

## 6. Gates

Run on the final tree after the trunk merge (`fcf1f8ff44`) with the
wave's `env.sh` (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`);
`.coord/w1-docx/scripts/gates.sh` runs each command and records its exit
code. All exited 0; the DOCX test target was re-run after this record's
example was fixed (its first run failed to compile the example).

| Gate | Command | Result |
|---|---|---|
| format | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy -p litchi-docx -p litchi-opc -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| clippy, features | `cargo clippy -p litchi-ops --no-default-features [--features docx] --all-targets -- -D warnings`; `cargo clippy -p litchi-docx --no-default-features --all-targets -- -D warnings`; `cargo clippy -p litchi-ods --lib -- -D warnings` | 0 |
| docs | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-opc` | 0 (one redundant link target in `numbering/catalog.rs`, from trunk, fixed) |
| DOCX tests | `cargo test -p litchi-docx --all-features --no-fail-fast` | 0; 2,466 passed |
| OPC tests | `cargo test -p litchi-opc --all-features --no-fail-fast` | 0; 1,190 passed |
| ops tests | `cargo test -p litchi-ops --all-features`, lib and test targets in chunks of 28 | 0; 2,423 passed |
| facade, CLI | `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 0; 209 and 34 passed |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` | 0; 301 tests |
| perf-baseline, native-resave | `cd tools/perf-baseline && cargo check --locked --all-targets`; `cd tools/native-resave && cargo check --locked --features docx` | 0 (native-resave: one unused-import warning, also at base) |
| boundaries | `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0; 44 profiles, 2,985 bans, 408 families, 64 immutable owners |
| goldens | `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature but `encryption` | 0; no golden changed |
| capability delta | `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI`; `--check` | 0; unchanged |
| OfficeCLI corpus | `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py`; `--check` | 0; no corpus file changed |
| tool tests | `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate`; `tools.test_check_api_glossary` | 0; 1,269 and 27 tests |
| matrices | `python3 tools/unified_bench_matrix.py --check`; `python3 tools/check_feature_matrix.py` | 0 |

`cargo check --workspace --all-targets --all-features` also fails on
`crates/litchi-ods/tests/ods_formula_conditional_native.rs`, which
`include_str!`s an evidence file the 2026-10-05 rewrite purged; that is on
trunk and outside this wave. Target directory: 20 GB at the end of the gate run, 16 GB after removing test executables and examples (release builds in `.coord/w1-docx/rel-{base,head}`, 0.7 GB each).

## 7. Follow-ups

- **0355a-F1** — the empty `LimitKind` of the OOXML-common families (task
  panes, custom XML, ribbon, VBA) waits for typed kinds in
  litchi-ooxml-common and litchi-vba (W4).
- **0355a-F2** — `Package::apply_durable_ink` returns a `Package`, not a
  family `Commit`; the durable wire patches of the other formats do the
  same, so it is left to the wire-patch contract.

## 8. Log entry

- 0355a W1 DOCX gaps: 13 source-backed families commit to a `Package`
  (`edit_x` → `Edit` → `commit()`; `apply_x`; `write_to` the only output);
  every `publish_*_to_stream`/`*Publication` removed and banned.
- Exact inverse by artifact identity, streamed through the current
  context; durable copy/removal/stream patches kept (`Patch::durable`).
- F2 typed `LimitKind` per family; F3 `document::Refusal::kind`, corpus
  unchanged; F4 alt-chunk edits, revision reader budgets,
  `max_output_bytes` on content-control and conflict edits, tests restored.
- P1: the 1k paragraph-text batch is 1.03–1.05× the leaf edit (was 302×);
  family outputs byte-identical to the publishers at 0.13–1.02× their time.
- The flaky OPC batch-cancellation test holds its gate (0/40 failures).
