# 0356 — PPTX on the record 0351 API contract (wave W2)

- Status: implemented on `uo2/w2-pptx`, cut from `feat/unified-ops-api`;
  trunk merged three times (records 0352–0353; 0354 at `5f6fb75164`; 0355
  at `4ff05a303e`), for coordinator review.
- Date: 2026-10-20.
- Input: record [0351](0351-api-contract.md) §2 (rules), §3.2 (the 77 PPTX
  `Package` mutators), §4.3 (the W2 gate), §5 (performance routes); owner
  decision [0350](0350-owner-api-debt-decision.md) items 14–19.
- Owned: `crates/litchi-pptx/**`; `crates/litchi-ops/src/format/pptx*`, the
  PPTX `Classify` arms of `error.rs`, the `pptx*` tests and benches; the PPTX
  hunks of the facade examples and tests, `pyo3-litchi` (none needed),
  `tools/perf-baseline` (PPTX sources); the `pptx-*` glossary profiles; line
  W2 of `docs/API_GLOSSARY.md` §2a; `tools/unified_bench_matrix.py` (the
  0356 parser).
- `performance_claim`: measured (§5). No route regresses: open, `set`, the
  1k batches, `to_bytes` and `to_bytes` twice are within ±1.1% of the base
  build (after/before, A-B-B-A, three runs); every 1k batch route is within
  1.97× of the leaf (cold notes 1.90–1.96×); the leaf hot paths are the
  same code.

## 1. Result

| Rule (0351 §2) | Before (0351 §1 inventory) | After |
|---|---|---|
| Immutable `Package` | 77 `&mut self` methods | none; outputs take `&self`, `unsign(&self) -> Result<Package>`; the gate's `immutable_owners` lists `litchi_pptx::Package` |
| Edit families | 25 commit and 19 edit types; `put_*`/`remove_*`/`apply_*_patch` | 33 registered families on `edit_x`/`apply_x` with `x::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`; every commit publishes a `Package` |
| Errors | `UnsafeEdit` at 224 sites | deleted; `Refused(Refusal)`, `PatchConflict`, `Internal`; typed refusal enums gain `kind()` |
| Limits | 37 public `*_with_limits` | none; `OpenOptions`, `opened::EditOptions`, or a required `Limits` argument |
| Drafts | `writer::Mutable{Presentation, Slide, Shape}` | `builder::{Presentation, Slide, Shape}`; `writer` keeps the streaming writer |
| Gate | `pptx-*`: 2 families, owner `Presentation` | bans `UnsafeEdit`, `Mutable*`, `*_with_limits`, `readonly_recommended`/`ReadonlyRecommended`; 33 families; owners `Package`, `Presentation`; `family_methods` |

Rerunning `tools/api_debt_inventory.py` on the branch leaves, for PPTX:
the five `SourceBacked*` commit and edit types (§6 gap 1), the ten
`&mut self` methods of the `builder::Presentation` draft (drafts keep
`&mut self`, rule 7; the inventory names every `Presentation` an owner),
and 31 `commit_type` rows with an empty detail: the families' `Commit` is
`litchi_core::edit::Commit<Package, Patch>`, whose methods rustdoc does not
list in this crate (the inventory's documented blind spot; the core type is
checked by its own unit tests).

## 2. What landed, by commit

| Commit | Step |
|---|---|
| `56fb8814fe` | errors: `UnsafeEdit` deleted, each site mapped by its reason |
| `85756534af` | families: the `Package` mutators become `edit_x` families on one staged engine |
| `0d356f958e` | immutable `Package` |
| `11079cf30f` | limits through options |
| `76ec884a3b` | drafts move to `builder` |
| `2f192f406f` | gate registration, intra-doc link fixes |
| `b25006ec4f` | `opened::TextEdit` → `TextAssignment` |
| `d68ccf0b72`, `737a27ae82`, `3e4ec2ddd1`, `9d957e89eb` | the detached snapshot transactions become families |
| `802cb5b4eb` | glossary line W2; custom-show examples |
| `90458123a7` | `pptx_contract_bench`, evidence, matrix parser, this record |

Merges of trunk: `b34e21ce60` (records 0352, 0353) and `842c1ca64e`
(record 0354), clean; `9536c3fa7a` (record 0355, W1 DOCX): the §2a W1/W2
lines and the bench-matrix parser list, both kept.

### 2.1 Errors (`56fb8814fe`)

`Error::UnsafeEdit` is gone. A well-formed request a route refuses is
`Error::Refused(litchi_core::edit::Refusal)` with the `RefusalKind` its
reason names; a patch or plan made against another base is the unit
`Error::PatchConflict`; an invariant break is `Error::Internal { operation,
reason }`. The 197 remaining sites were rewritten by reason (the script and
its site table are in the session scratch, not committed); every site whose
old `litchi-ops` code differed from its kind's code keeps the old code by
the kind it was given (`UnsupportedOperation`, `UnsupportedContent`).
`ShapeTransferRefusal`, `SlideCopyRefusal`, `SlideRemovalRefusal` and
`SlideOrderRefusal` gain `kind()`.

`litchi-ops` (`error.rs`, PPTX arm): `PatchConflict` → `invalid_input` (was
`io_error`); `Refused` → its kind; the typed refusals by `kind()`, except
`SlideOrderRefusal::AmbiguousBindings` (`corrupt_file`) and
`SlideRemovalRefusal::FinalSlide` (`invalid_argument`), which keep their
codes; `SignedPackage` keeps its suggestion; `Allocation`/`Internal` →
`internal_error`. Code changes visible to callers: `SlideCopyPlan` and
`ShapeTransfer` refusals were `corrupt_file` and are now classified by kind;
`UnsupportedTarget` is `unsupported_type`. The digest-reuse golden
(`digest_reuse_lptx0003.golden.txt`) changed by exactly the variant name.

### 2.2 Families (`85756534af` and the four residue commits)

One engine, `part_edit::Staged`: a staged clone of the package's OPC
graph, `update(f)` that keeps the staged graph on an error, `reset`, and
`commit -> (Package, PartPatch)`. `PartPatch` captures the exact before and
after state (content type, payload, relationships) of every part the edit
added, changed or removed, plus the root relationships; `apply` checks each
before state (else `PatchConflict`), refuses a changed signed package, and
settles the unsigned change. An unchanged staging publishes the source
package itself (`Package::duplicate`), so a no-op keeps every byte and
signature. The `family!` macro supplies the shared members.

Families (all `Package::edit_x(&self)`, `Package::apply_x(&self, &Patch)`):

- **Package mutators of 0351 §3.2:** `tag`, `tag::shape`, `font`,
  `table::style`, `presentation_properties::math`, `shape::zoom`,
  `shape::classification`, `shape::designer`, `shape::designer::properties`,
  `model3d`, `web` (task panes), `master_layout`,
  `master_layout::placeholder`, `notes`, `change_tracking`,
  `presentation::embedded::ink_actions`,
  `presentation_properties::read_only_recommended` (the module renamed from
  `readonly_recommended`; one spelling),
  `presentation_properties::metadata::designer_tags`, and `opc` (the raw
  clear-OPC escape: `edit_opc().update(|opc| ..)`, which declassifies a
  package with encryption provenance).
- **Detached snapshot transactions (0351 §2.3 rules 2–4):**
  `presentation_properties::metadata::{guides, custom_show, changes,
  structure, events, tracks, color_map}`,
  `presentation::embedded::{content_parts, controls::slide, ole::slide}`,
  `comments::collaboration::{presence, threading}` (the prefixed
  `AuthorPresence*`/`CommentThreading*` types). Each keeps its validated
  value editor as a public `Draft` (`events::ListDraft`, since
  `events::Draft` is the event value) that only `Edit::update(key, |draft|
  ..)` hands out; the old detached `Commit`/`Patch` and the raw
  `apply_*`/`store`/`put_*`/`remove_*` free functions are crate-private, and
  the raw verbs are `Edit` methods. Snapshots stay public read views.
  `color_map` had no package path at all; its family locates the map of a
  slide master, layout or slide part by part name.
- **Plans:** slide copy, slide removal and cross-slide copy stage on the
  primary `opened::Edit` (`stage_slide_copy`, `stage_slide_removal`,
  `stage_cross_slide_copy`); their durable patches replay through
  `Package::apply_slide_removal` and `apply_cross_slide_copy` (both return
  `opened::Commit`).

Deleted, not aliased: every `put_*`/`remove_*`/`apply_*_patch`/
`apply_*_commit`/`*_plan` `Package` method of §3.2, `is_modified`,
`ensure_graph_current`, `authoring_presentation`, `set_font_embedding`,
`release_part_digests`, `encryption_mode`, the structure aliases
`replace_custom_show`/`replace_section`, the unused
`synchronize_after_slide_mutation`, and the forbidden `snapshot()` readers
of `change_tracking::Edit` and `ink_actions::Edit` (now `change_tracking()`
and `ink_actions()`).

### 2.3 Immutable `Package` (`0d356f958e`)

The package lost its mutable state (`mutable_pres`,
`font_embedding_dirty`). `builder::Builder` owns the draft presentation and
the font-embedding switch and publishes once in `build()`
(`publish_authored`). Outputs take `&self`; `to_encrypted`,
`to_reencrypted`, `save_encrypted`, `save_reencrypted` and
`write_reencrypted` use the encryption limits the package was opened with
(`OpenOptions`), and saving no longer changes the retained encryption mode.
`unsign(&self) -> Result<Package>`.

### 2.4 Limits (`11079cf30f`)

The six `SourceBacked{Presentation, PresentationEditor}` and the four
`source::Package` constructors are `open`/`from_read_at`/`from_reader`
with `_with_options` forms; patch decoders (`opened::Patch`,
`slide_removal::Patch`, `cross_slide_copy::Patch`) take a required
`Limits`; validation takes explicit limits; `opened::EditOptions { limits }`
with `Package::edit_with_options` and `opened_presentation_with_options`.
The per-call designer, writer and tag `_with_limits` forms are crate-private;
`Tags::{insert, replace}` use the default limits.

### 2.5 Drafts (`76ec884a3b`)

`writer::{MutablePresentation, MutableSlide, MutableShape}` →
`builder::{Presentation, Slide, Shape}` (the builder already owned them).

## 3. Behaviour changes

- Error codes: §2.1.
- A family patch conflicts on the parts it changes (exact before bytes),
  not on a source revision of the whole snapshot; the detached revision
  check still guards each `update` against its own staged graph. The
  ActiveX revision-determinism test moved into the crate, where that patch
  lives.
- Metadata edits are allowed on a package with encryption provenance (the
  typed families refuse it through `ensure_plain_mutation` as before; the
  raw `opc` family declassifies).
- Saving encrypted no longer changes the retained mode.
- Cross-slide-copy staging captures one more memo hit.

No writer's output changed: the round-trip, CRUD, corpus and no-op tests
pass unchanged, the OfficeCLI corpus is green, batch equivalence holds,
and every bench scenario compares leaf and unified bytes (all equal).

## 4. Consumers migrated

`litchi-ops` (`format/pptx*`, `error.rs`, `raw/validate.rs`, tests,
examples), the facade (`examples/pptx_*`, `core_props_office`,
`office_crud_demo`, `tests/encryption_facade.rs`, `read_limits_facade.rs`),
`tools/perf-baseline` (the PPTX sources, `filesystem/ops.rs`,
`unified-pptx-text-format`), the crate's own examples, tests and benches
(`tests/in_place/mod.rs`: in-place helpers over the public families for
fixtures whose subject is not the family). `litchi-cli` and `pyo3-litchi`
needed no change.

## 5. Measurements

Release, `taskset -c 7`, on a shared 32-core host (load 7–9). Evidence in
[`results/0356/`](results/0356/) (132 KB); binaries and build commands in
`results/0356/binaries.sha256`. "Before" is a `git archive` export of trunk
`5f6fb75164` (this branch's merge base) with this branch's `Cargo.lock` and
`pptx_contract_bench.rs` copied in.

### 5.1 Before/after (`examples/pptx_contract_bench.rs`)

One source, built on both trees (it uses only calls both compile), run
before, after, after, before, three times; the median of each side's six
runs:

| route (fixture) | before µs | after µs | after/before |
|---|---:|---:|---:|
| open, leaf (`45545_Comment.pptx`, 308 KB) | 1,614.8 | 1,626.7 | 1.007 |
| open, unified (same) | 726.8 | 724.0 | 0.996 |
| `to_bytes`, leaf (same) | 12,069.1 | 12,075.2 | 1.001 |
| open + `to_bytes` twice, leaf (same) | 25,798.2 | 25,752.4 | 0.998 |
| open + `to_bytes` twice, unified (same) | 743.9 | 741.5 | 0.997 |
| open, leaf (20-box deck) | 278.7 | 277.0 | 0.994 |
| `to_bytes`, leaf (20-box deck) | 704.3 | 712.3 | 1.011 |
| open + `to_bytes` twice, leaf (20-box deck) | 1,810.7 | 1,820.3 | 1.005 |
| `set` shape text, unified (20-box deck) | 460.5 | 461.3 | 1.002 |
| `set` + `to_bytes`, unified (20-box deck) | 535.9 | 537.5 | 1.003 |
| 1k run-format batch, unified | 13,863.2 | 13,740.6 | 0.991 |
| 1k shape-text batch, unified | 10,274.2 | 10,236.4 | 0.996 |

(The unified `to_bytes` of an unedited handle returns its source bytes,
0.3–7 µs; those rows are in the JSONL files.)

### 5.2 Leaf against unified (`pptx_p0_final_bench`, `pptx_p0_bench`)

Record 0348's routes, three runs each on the branch; the matrix rows
supersede 0348's. 1k and 10k, ratio range of the three runs (0348's worst
in brackets):

| scenario | 1,000 | 10,000 |
|---|---|---|
| `cell_format_batch/eager` cold / warm | 1.27–1.30 (1.27) / 1.09 | 1.25–1.27 / 1.09–1.10 |
| `cell_paragraph_batch/eager` cold / warm | 1.32–1.33 (1.31) / 1.10–1.11 | 1.31–1.35 / 1.11 |
| `cell_format_batch/source` cold / warm | 1.19–1.20 / 1.06–1.07 | 1.18–1.20 / 1.07–1.08 |
| `cell_paragraph_batch/source` cold / warm | 1.21–1.23 / 1.06–1.08 | 1.21 / 1.06–1.08 |
| `notes_format_batch/eager` cold / warm | 1.90–1.96 (1.92) / 1.18–1.24 | 1.95–1.97 (1.97) / 1.23–1.25 |
| `paragraph_text_batch` eager / source | 1.24–1.26 / 1.13–1.14 | 1.28–1.29 / 1.14–1.16 |
| `run_format_batch` eager / source | 1.25–1.26 / 1.14–1.15 | 1.25–1.26 / 1.14–1.15 |
| `shape_text_batch` eager / source | 1.07–1.08 / 1.03–1.05 | 1.08–1.09 / 1.05–1.06 |

Every 1k route is within 1.97× (record 0348's bound) and every 1k and 10k
route within the 2.0 budget. The worst value, 10k cold notes at 1.972 in
one run, is 0348's 1.966 within run-to-run spread; a base-tree run of the
same binary source in the same session measured 1.926 there. The single
operations are 1.00–1.09× as in 0348. Against the one base run, the
unified legs are within ±2% at N ≥ 1,000; the small cold rows (N = 10–100,
a fresh handle's fixed cost) are up to 5% slower in this single
comparison, inside the spread of the cold 1-item rows (1.55–1.79× across
runs of one binary).

## 6. Gaps and follow-ups

1. **0356-P1: the `SourceBacked*` stream editors.**
   `SourceBackedPresentationEditor`'s `SourceBackedSlide{,Batch,Order}` and
   `SourceBackedSvg{,Attachment}` edits commit to a detached snapshot and
   publish to a byte stream (`publish_*_commit_to_stream`), not to a
   `Package`, so they cannot take the family triad without losing the
   bounded-memory route they exist for. They keep their prefixed names. An
   owner decision is needed: delete them in favour of `source::Package`'s
   `opened::Edit` (which `litchi-ops` uses), or define a stream-publishing
   family shape.
2. **0356-P2: raw `&mut OpcPackage` helpers.** About 55 public free
   functions over a raw graph remain (legacy and modern comments, themes,
   VBA, ink, laser traces, media, charts, revision information, and raw
   `put`/`remove` twins of the font, tag, table-style and
   read-only-recommended families). They mutate a caller's `OpcPackage`,
   not a `Package` (rule 1 holds); reached from a package they run inside
   `edit_opc().update(..)`. Rule 3 asks for their publication paths to go;
   the twins can be demoted once their tests and examples move to the
   families, and the comment helpers want comment families.
3. **0356-P3: selector parameters.** The converted families keep their raw
   keys (`&PackURI` slide part names, slide indexes, author IDs) instead of
   `impl Into<slide::Selector>`.
4. **0356-P4: inventory blind spot.** `api_debt_inventory.py` reports a
   family whose `Commit` is the core alias as missing its accessors;
   teaching it the alias would make "no rows left" the done criterion.

## 7. Gates

Logs: `(deleted-coord-artifact)` (`gates-summary.txt`). Commands run from the worktree
with `source (deleted-coord-artifact)`.

All on the tree merged with trunk `4ff05a303e` (record 0355), merge `9536c3fa7a`.

| gate | exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pptx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pptx] --all-targets -- -D warnings` | 0, 0 |
| `cargo clippy -p litchi-ops --features <all but encryption> --all-targets -- -D warnings` | 0 |
| `cargo clippy --workspace --all-features --all-targets --exclude litchi-ods -- -D warnings` (before the last merge) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops` | 0 |
| `cargo test -p litchi-pptx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, 169 `--test` targets in chunks of 20 (executables deleted per chunk) | 0, 0, 0 |
| `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 0; 0 |
| `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (44 profiles, 315 families, 60 owners) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix tools.test_check_api_glossary` | 0 |
| `tools/officecli_capability_delta.py --check`; `tools/officecli_corpus.py --check` | 0; 0 |
| `tools/unified_bench_matrix.py --check` | 0 |
| `tools/check_feature_matrix.py` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` (all features; all but `encryption`), then a clean tree | 0, 0, 0 (no golden changed) |

The OfficeCLI corpus (`corpus_fidelity`, `pptx_corpus_parity`) and the batch-equivalence tests run in the `litchi-ops` chunks. Target directory after cleanup: 16 GB (limit 20 GB).

`ods_formula_conditional_native` (litchi-ods) cannot compile on any tree
since the 2026-10-05 purge of its evidence fixture; the workspace commands
exclude `litchi-ods` for that reason only.
