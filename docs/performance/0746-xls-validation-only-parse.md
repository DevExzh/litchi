# 0746: XLS edit owners validate through a validation-only mode of the complete reader, and their generic commits publish the rendering they already validated

Status: retained, implemented, revised after review. `performance_claim: none`
— this record carries paired ABBA timings on two builds, isolation-pair
instruction and cycle counts, deterministic allocation counts, a three-leg
correctness census and two frozen first-error matrices. They are reported as
evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `009d515bef`; branch `perf/0746-xls-validation-only-parse`; five code
commits: `3a2f233cdc` (deterministic multi-defect refusals), `ec747fbd64` (the
validation-only mode), `fa88d7cc9b` (the validated-render handoff in
`cell_values`), `3eea4a8bec` (comments and visibility owners adopt the mode) and
`b85d3e534c` (the handoff in the comments and visibility commits), recorded in
`31afd6f11a`. An independent review returned "merge after fixes"; the fixes are
four more commits: `8fa1d4b54a` (the occupancy map sized by occupied rows
instead of 256-row bands, plus two documentation fixes), `7f3c6c4b78` (a frozen
worksheet-level multi-defect first-error matrix generated on the base, and a
cross-tab `kept_cell` refusal test), `3d49e03044` (the map's current-row answer
inline, everything else out of line) and `0cd989bb15` (documentation only).
Sections "Result" to "Allocations" measure the reviewed build, `b85d3e534c`;
section "Review fixes" measures the final source against it and the base.

## Result

Every XLS edit owner that opened the complete eager reader, `Workbook::new`,
only to validate a package and read a few facts from it now runs **the same
parser in a validation-only mode**: every record is validated identically, but
no per-cell map is built and no cell is decoded except the few the commit reads
back. The three generic commits (`cell_values`, comments, sheet visibility)
additionally publish the CFB rendering their package publication already
validated, instead of rendering the same editor state a second time.

What the change removes is stable: **43.9–51.0% of the instructions of every
open and commit it touches** (37.0% / 41.0% for `commit_source_backed` on
`54016.xls` / `xls-large`), **26–68% of the allocation calls, 41–71% of the
allocated bytes and 44–83% of the peak live bytes**, with the public reader's
instructions (−0.2%, −0.1%) and allocations (identical to the call) unchanged.
Those are the reviewed build. The final source executes a further 2.6–5.2%
fewer instructions than the reviewed build on every re-measured `54016.xls` and
`xls-large` operation (the public-reader control within ±0.1%), allocates
4.1–9.6% more bytes than it on `54016.xls` (still 38–54% below the base), and no
longer regresses the sparse-band shape the review constructed, on which the
reviewed build was slower than the base. Final source against the base, one
window: `54016.xls` open 10.462 → 4.866 ms (0.465), generic Number commit
23.242 → 10.118 ms (0.436), sparse-band open 60.284 → 27.343 ms (0.453)
(section "Review fixes"). Wall-clock of the reviewed build, paired ABBA on
CPU 20:

| fixture | operation | before p50 | after p50 | paired ratio | pair range | same code, other build |
| --- | --- | ---: | ---: | ---: | --- | ---: |
| `54016.xls` | `cell_values::Snapshot::from_bytes` | 10.640 ms | 7.564 ms | 0.715 | [0.678, 0.743] | 0.539 |
| `54016.xls` | `commit_source_backed_plan` | 11.509 ms | 4.776 ms | **0.415** | [0.400, 0.440] | 0.549 |
| `54016.xls` | `commit_source_backed` | 14.953 ms | 11.546 ms | 0.768 | [0.764, 0.776] | 0.662 |
| `54016.xls` | `commit` (Number edit) | 24.094 ms | 13.276 ms | **0.553** | [0.525, 0.558] | 0.498 |
| `54016.xls` | `commit` (LabelSst to existing SST text) | 24.533 ms | 13.915 ms | 0.569 | [0.557, 0.571] | 0.506 |
| `54016.xls` | `comments::Snapshot::from_bytes` | 11.221 ms | 7.483 ms | 0.678 | [0.610, 0.683] | 0.523 |
| `54016.xls` | `sheet_visibility::Snapshot::from_bytes` | 11.055 ms | 6.993 ms | 0.623 | [0.596, 0.653] | 0.500 |
| `54016.xls` | control: public `Workbook::new` | 10.595 ms | 10.430 ms | 0.979 | [0.950, 1.001] | 0.929 |
| `xls-large` | `cell_values::Snapshot::from_bytes` | 1.568 ms | 1.360 ms | 0.853 | [0.848, 0.922] | 0.505 |
| `xls-large` | `commit_source_backed_plan` | 1.598 ms | 0.692 ms | **0.433** | [0.399, 0.436] | 0.495 |
| `xls-large` | `commit_source_backed` | 2.187 ms | 1.815 ms | 0.836 | [0.825, 0.841] | 0.603 |
| `xls-large` | `commit` (Number edit) | 3.473 ms | 2.212 ms | 0.637 | [0.589, 0.665] | 0.486 |
| `xls-large` | `comments::Snapshot::from_bytes` | 1.536 ms | 1.300 ms | 0.836 | [0.751, 0.870] | 0.480 |
| `xls-large` | `sheet_visibility::Snapshot::from_bytes` | 1.449 ms | 1.228 ms | 0.851 | [0.820, 0.884] | 0.474 |
| `xls-large` | control: public `Workbook::new` | 1.430 ms | 1.388 ms | 0.968 | [0.922, 1.000] | 0.921 |

The pair range is the lowest and highest of the four adjacent-pair ratios
(the first version of this record labelled it a 95% CI; see "Measured"). The
last column is the same comparison run earlier against a build of
`3eea4a8bec`, whose `cell_values` and open code is identical to the reviewed
source. **The two builds execute the same instructions (±0.1%) and differ by up
to 1.7× in cycles** on the validation-only walk, because the walk is now lean
enough for a code-layout-dependent store-forwarding stall to dominate it
(section "Build sensitivity"). Both are reported; the reviewed build's numbers
are the conservative ones. The registered selectors, built with the pinned
toolchain from the reviewed source and compared against a base harness built
with the identical command, agree with the faster layout:
`xls_semantic_one_edit_save` (the generic commit on `xls-large`) **3.535 →
1.643 ms (0.464)**, `xls_visibility_eager_edit_save` **25.47 → 13.33 ms
(0.523)**, `xls_comments_eager_edit_save` 21.54 → 16.67 ms (0.792),
`xls_numeric_eager_rk_mulrk_edit_save` 2.395 → 1.235 ms (0.516).

Every published artifact, refusal and readback is byte-identical to the base
across a 126-fixture census, and every owner's open (its refusal, or a digest of
its decoded cells) across 18,600 mutated packages: the mutation corpus runs
opens only, so published artifacts are compared only in the census. Both, and
the 29-case first-error matrix, give identical output on the final source
(section "Correctness evidence").

## Why this record exists

Change [0620](0620-xls-edit-save-attribution.md) measured the complete eager
parse at 82% of `Snapshot::from_bytes` on `54016.xls` and 55–80% of each
commit; change [0633](0633-xls-commit-single-framing.md) priced it: `add_cell`
33.82% of the open, of which the `BTreeMap` insert 20.17%, and 4.68% more to
drop the `Workbook` the edit owner never reads. 0633 froze three designs around
that parse (fusing its framing with the offset inventory, handing it the
already-extracted stream, replacing its owner) and named the remaining cost:
"the eager `Workbook` is built and discarded … to answer one bit per sheet".
Nothing since 0633 touched `cell_values` or the eager parser (0668, 0684–0690
and 0723–0727 changed the lazy `SourceBackedWorkbook` owner; the checkpoint
family there was rejected and is not retried here).

A fresh frame-pointer profile of the base (`results/change-0746/profiles/`)
confirms it and adds one more term. On `54016.xls` the open is 84.7%
`Workbook::new` (`add_cell` 38.7%, `Cell::from_record_with_formula_context`
10.6%, a `memmove` inside the `BTreeMap` insert 12.25% of all samples) plus 4.9%
dropping it. The generic `commit` is 40.7% that parse, 3.8% its drop, and **49%
two CFB renders**: `put_stream_shared` renders, reopens and recaptures the
candidate and throws the bytes away, then the snapshot constructor's `finish()`
renders the same editor state again. Each `render_copy_through` is about a
quarter of the commit, most of it SHA-256 fingerprint passes over the whole
artifact. The comments and visibility generic commits have the same two renders.

## What changed

Scope: `crates/litchi-xls` only. No public type, public signature, limit or
dependency changed; no `unsafe`; no output byte changed (proved below).

**1. Deterministic multi-defect refusals** (`3a2f233cdc`,
`workbook/codec/semantic/worksheet.rs`, `comments/codec.rs`). The worksheet
walk refused several orphan `PtgExp` formulas, and the comment collector
several OBJs without a NOTE, by naming whichever offender a `std` `HashMap`
visited first. Map iteration order comes from a per-map random state, so two
opens of the same bytes could name different cells; this change's first
mutation differential found it (`Formula at (17, 256)` against
`Formula at (3, 3)` for one input). Both sites now take the minimum offender in
one pass, with no allocation and no change for accepted inputs. They name the
lowest `(row, column)` and the lowest object id, not the first offender in
record order: the maps they scan hold positions and ids, not arrival order, so
the minimum costs at most the rest of the one scan each check already makes,
and only on the refusal path, where naming the first in record order would add
ordering state to the per-record walk and to the collector for every input; and
every offender refuses the sheet with the same error kind, so which one is named
changes the message, never the outcome. Two tests parse
inputs with five orphans and four unmatched objects 32 times each and require
one exact message; both fail on the base. This is an ADR 0006 determinism fix
the differential needed as its oracle: "the same first error" is only well
defined once the first error is a function of the bytes.

**2. The validation-only mode** (`ec747fbd64`). The worksheet walk,
`Workbook::parse_worksheet_records_with_compatibility`, is now generic over a
`CellStore` (`workbook/codec/semantic/cells.rs`). Every check it runs against
earlier cells asks one of two questions — is this `(row, column)` already
occupied, and is the latest record there a `Formula` — and the store answers
them:

- `DecodeEveryCell`, the public reader's store: each validated record is
  decoded into the worksheet's cell map, exactly as before, and that map answers
  both questions.
- `ValidateCells`, the validation-only store: each validated record is counted
  in an occupancy map, and decoded only when its position was asked for. Since
  `8fa1d4b54a` the map holds two 256-bit masks (64 bytes) per occupied row,
  allocated fallibly when the row's first record arrives and found through the
  row most recently stored, an append-only index of the rows first seen in
  ascending order, or — for a row first seen below an earlier one — a hash
  map; columns at 256 or beyond, which the single-cell decoder does not bound
  and the public reader keeps, go to an exact map. The reviewed build kept the
  same masks in 256-row bands, 16 KiB per touched band, which a sparse
  worksheet could make 4 MiB (section "Review fixes").

`add_cell` (duplicate detection and its tracking limit), the string-`Formula`
path, `MulRk`/`MulBlank`, the `ShrFmla` anchor rendering and the post-walk
`Array` checks ("not materialized", "non-Formula cell") read the store. The cell
decode moved into one `decode_cell` both stores call, and
`Cell::from_record_with_formula_context` now returns `Cell` instead of an
`Option<Cell>` that was always `Some`, so "every validated record produces
exactly one cell" is a type, not an inspection (its three lazy-owner callers in
`workbook/source.rs` lose an unreachable `else`).

`Workbook::validation_only(reader, KeptCells)` (`workbook/validation_only.rs`)
runs `OleFile::open`, the XML-map stream, `parse_workbook` and every
package-level check exactly as `Workbook::new` does, under the same default
`OpenOptions`, with the validation-only store per worksheet. It returns a
`ValidationWorkbook` newtype, not a `Workbook`: it exposes the sheet directory,
workbook and worksheet protection, worksheet comments, VBA markers and the
shared-string tables through accessors of the same names, and one cell accessor,
`kept_cell`, which refuses any position it was not asked to keep instead of
answering it as empty.

The four `cell_values` validation owners use it: `Snapshot::open_package` (a
plain open keeps no cell; a source-backed target verification keeps the edited
cells its readback reads), `from_fixed_numeric_package_editor` and
`verify_source_backed_numeric_plan_target` (both keep the edited cells), with
`SourcePolicyFacts`, `require_public_worksheet_coverage`,
`require_unprotected_workbook`, `require_macro_free_workbook`,
`verify_public_numeric_readback` and `retained_shared_string_properties` taking
the newtype. `numeric_readback_cells` lists every staged change's
`(tab, row, column)`, a superset of what the readback reads.

**3. The validated-render handoff** (`fa88d7cc9b` for `cell_values`,
`b85d3e534c` for comments and visibility). The generic commits call the
existing `PackageEditor::put_stream_shared_with_rendered`, which returns the
exact artifact its publication rendered and reopened, and hand it to the
snapshot constructor (a `rendered` argument on each owner's `open_package`,
and the fixed-numeric constructor) in place of `finish()`. This is change
[0730](0730-doc-bounded-render-handoff.md)'s DOC pattern in its simplest form:
the rendering is consumed inside the same call, so there is no retention,
ceiling or release policy to add. Plain opens and the source-backed paths keep
`finish()`, which on an unchanged editor returns the original bytes.

**4. Comments and visibility** (`3eea4a8bec`). `comments::Snapshot::from_bytes`,
`sheet_visibility::Snapshot::from_bytes` and both owners' source-backed
candidate readbacks opened `Workbook::new` only for sheet metadata, protection
and comments; they now use the validation-only mode with no kept cell, reading
comments and protection in the same order as before.

## Why it is sound

**It is the same parser.** Both stores run behind one generic function; every
record is framed, decoded (`CellRecord::parse`,
`parse_formula_preserving_defect`, `parse_mul_rk`, `decode_string_record`, …)
and checked (`validate_cell_xf`, companion claims, `PtgExp` and `ShrFmla`
tracking, `Array` ranges, `DVAL`/`DV`, every collector's `feed_record` and
`finish`) by the same code, in the same order, under the same limits. The two
things the validation-only store skips cannot refuse: the cell decode is an
infallible conversion with no side effect (now typed so), and the map insert is
infallible. The `ShrFmla` and `Array` formula renderings it skips for cells it
does not keep are pure functions of the tokens whose only effect is the text
stored on that cell.

**The two questions get the same answers.** A position is occupied in the
occupancy map exactly when an earlier record there reached `add_cell`, which is
when the public reader's map holds a cell there, because every such record
decodes to exactly one cell at the record's own `(row, column)`. The `Formula`
flag is rewritten on every record at the position, matching the map's
last-record-wins insert, and it is what `Cell::set_array_formula` tests
(`formula_metadata` is present exactly for `Formula` records). Each occupied
row has exactly one entry whichever path finds it: a new row joins the
ascending index only when it is above every row seen so far and the hash map
only otherwise, so a row above the highest indexed row is unseen, and any other
row is in the index or the hash map. The only change of order is inside the
`Array` check: the non-`Formula` refusal is decided before the rendering is
attached rather than after, which is unobservable because the worksheet is
discarded on either refusal.

**Kept cells are the public reader's cells.** A kept position is decoded by the
same `decode_cell`, in record order, receives the same `ShrFmla` and `Array`
enrichment, and sits in the worksheet's own map, so `kept_cell` returns what
`xls_worksheet(i).get_cell(r, c)` returned. `verify_public_numeric_readback`
keeps its refusals in their order; only its final lookup moved behind the
kept-set check, which is unreachable because the kept set is built from the
same change list.

**Coverage keeps its meaning.** Change 0620's frozen design rejected replacing
the readback's owner with the lazy `SourceBackedWorkbook` because
`parsed_worksheet_index()` is a property of the eager parser's success. That
design is not engaged: the validation-only mode *is* the eager parser, the index
is set by the same `parse_workbook` loop on the same `Ok`, and the same
`Err(_) => {}` arm and pivot-view propagation decide it.

**The only new failure is resource exhaustion, and it surfaces as a coverage
refusal.** The occupancy rows, their index and the beyond-grid map are
reserved fallibly: a failed reservation fails that worksheet's parse with
`Error::Allocation`, where the public reader's infallible cell map would abort
the process. That error does not reach the caller. `parse_workbook`
(`workbook/package.rs`) handles a worksheet's parse error in its
`Err(_) => {}` arm, which drops every per-worksheet refusal except the
pivot-view one, exactly as it does for the public reader, so the sheet is not
published and its `parsed_worksheet_index` stays unset. The edit owners
then refuse the package through their coverage checks with
`Error::UnsafeEdit`: the `cell_values` and visibility owners with "worksheet at
tab position N was not published by the complete XLS reader", the comments
owner with "a worksheet substream was not completely parsed". The reviewed
build's documentation and this record's first version said the `Allocation`
error itself was returned; the `validation_only` documentation now describes
this path (`8fa1d4b54a`, `0cd989bb15`). (The kept-cell list is built fallibly
by the owner before the open; a failure there returns `Error::Allocation` to the
caller directly.) The map is proportional to what is present: 64 bytes of
masks and an 8-byte index entry (or one hash entry) per occupied row, with at
most doubling slack, and one entry per occupied position beyond column 255,
never the row range the cells span. On the review's sparse-band
shape that is 18 KiB per worksheet where the reviewed banded map held 4 MiB —
this record's first version called the banded map "far below the per-cell map",
which was false there (about 40 KB); on dense sheets it is the banded map's
64 bytes per row plus the index.

**The handoff publishes the bytes that were validated.** The returned rendering
is the exact artifact the editor rendered, reopened with `OleFile::open`,
checked with the package codec, recaptured and target-discovered; the old path
published a *second* rendering of the recaptured state. Debug and test builds
re-derive that second rendering after every generic commit of the three owners
and assert byte identity, which the whole `litchi-xls` suite exercises;
release-mode tests compare the published artifact with the old put-then-finish
route on five packages for the fixed-numeric and the structural `cell_values`
routes, three comment packages and three visibility changes.

**ADR constraints.** ADR 0003's publication boundary is intact: the candidate
is still materialized, reopened under `Targets::default()` and
`Limits::default()`, validated by the complete reader's parser and read back
through typed values before a `Commit` exists. ADR 0005: nothing is retained
beyond the call (no cache, no memo; the handoff's `Vec` becomes the snapshot's
bytes exactly as `finish()`'s did). ADR 0006: validation still never mutates,
refusals are deterministic in two more places, and output bytes are unchanged.
Change [0652](0652-owner-decisions-for-the-third-wave.md)'s trade-offs 2 and 3:
no check, limit or refusal is weakened, and the benign majority stops paying for
a per-cell model nobody reads. No accepted ADR is amended and no owner decision
is needed.

## Measured

Host: AMD EPYC 9R45, 32 cores without SMT, 123 GiB, Linux 7.0.0-1012-aws,
shared with five other implementers (load average 6–17). Every measured process
is pinned with `taskset -c 20`; nothing of this change was building while
measuring. Binary SHA-256s are in `results/change-0746/binaries.sha256`.

Two instruments:

- **The probe** — change 0620/0633's `xls_edit_probe`, reused with three
  additions: an `alloc-count` feature so the timing binary runs on the system
  allocator, `--generate-xls-large` (reproduces the harness corpus call for call;
  its SHA-256 matches the manifest's `228c6585…`), and `comments-open`,
  `visibility-open` and `reader-open` operations. Built from identical source
  against each leg's `crates/` with rustc 1.98.1 (the host default, both legs),
  release with `debug = 1`, no LTO. It times each phase separately; a commit
  operation's number is its `commit` phase.
- **The harness** — the registered selectors. The after leg is
  `tools/perf-baseline` built from `b85d3e534c` with `cargo build --release
  --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin
  litchi-perf-baseline`; the quoted before leg is the base built with the same
  command (rustc 1.95.0 for both, the pinned toolchain; the harness source is
  unchanged). The runs were first taken against the shared prebuilt base binary;
  after the coordinator reported that the prebuilt binary shifts untouched paths
  by 2.7–3.4% against an identically built one, the whole table was rerun
  against a self-built base, and both are retained.

Design: 4 before and 4 after processes per case in the order
A B B A A B B A, ≥3 warmups and 20 samples (54016.xls), 60 samples
(`xls-large`), 30–60 (selectors); the table reports the median of per-process
p50s, the median of the four adjacent-pair after/before p50 ratios, and the
lowest and highest of those four ratios ("pair range"). The first version of
this record labelled that range a 95% CI because the scripts computed it as a
10,000-resample percentile bootstrap of the median; with four ratios that
bootstrap returns exactly their minimum and maximum (checked for every row of
the three summaries), so it is reported as what it is, not as a confidence
interval. Every process p50/p95/mean and every raw report is in
`latency-probe/`, `latency-harness-selfbuilt/` and `latency-harness/`.

### Paired timing, registered selectors (reviewed source)

Before leg built with the identical command (`latency-harness-selfbuilt/`).
The last two columns are the same design against the prebuilt base binary
(`latency-harness/`) and, for the first run, against a build of `3eea4a8bec`
before the comments and visibility handoff existed
(`superseded-run-3eea4a8bec/`):

| selector | corpus | before p50 | after p50 | paired ratio | pair range | vs prebuilt | first run |
| --- | --- | ---: | ---: | ---: | --- | ---: | ---: |
| `xls_semantic_one_edit_save` | xls-large | 3.5349 ms | 1.6429 ms | **0.464** | [0.458, 0.469] | 0.466 | 0.458 |
| `xls_numeric_eager_rk_mulrk_edit_save` | RK/MulRK | 2.3950 ms | 1.2352 ms | **0.516** | [0.515, 0.516] | 0.515 | 0.516 |
| `xls_numeric_eager_number_edit_save` | comments-heavy | 26.864 ms | 16.113 ms | **0.589** | [0.547, 0.680] | 0.534 | 0.704 |
| `xls_visibility_eager_edit_save` | visibility | 25.470 ms | 13.334 ms | **0.523** | [0.521, 0.525] | 0.526 | 1.003 |
| `xls_visibility_eager_batch_edit_save` | visibility | 25.330 ms | 13.241 ms | **0.521** | [0.519, 0.528] | 0.524 | 1.002 |
| `xls_comments_eager_edit_save` | comments-heavy | 21.537 ms | 16.674 ms | **0.792** | [0.600, 0.797] | 0.771 | 1.005 |
| `xls_comments_eager_batch_edit_save` | comments-heavy | 20.310 ms | 14.346 ms | **0.699** | [0.646, 0.854] | 0.772 | 1.026 |
| `xls_numeric_source_backed_number_edit_save` | comments-heavy | 78.039 ms | 75.071 ms | 0.962 | [0.953, 0.967] | 0.963 | 0.982 |
| `xls_numeric_plan_only_number_edit_save` | comments-heavy | 33.815 ms | 33.789 ms | 0.998 | [0.990, 1.004] | 1.003 | 1.000 |
| `xls_numeric_source_backed_rk_mulrk_edit_save` | RK/MulRK | 0.8247 ms | 0.8249 ms | 1.000 | [0.996, 1.006] | 1.001 | 1.000 |
| `xls_numeric_plan_only_rk_mulrk_edit_save` | RK/MulRK | 0.4025 ms | 0.4036 ms | 1.002 | [1.000, 1.005] | 1.000 | 1.005 |
| `xls_comments_source_backed_edit_save` | comments-heavy | 34.982 ms | 33.632 ms | 0.960 | [0.926, 1.002] | 1.040 | 1.000 |
| `xls_comments_source_backed_batch_edit_save` | comments-heavy | 36.648 ms | 34.081 ms | 0.963 | [0.930, 0.998] | 0.937 | 0.964 |
| `xls_visibility_source_backed_edit_save` | visibility | 10.151 ms | 10.154 ms | 1.000 | [0.999, 1.001] | 0.999 | 0.999 |
| `xls_visibility_source_backed_batch_edit_save` | visibility | 10.216 ms | 10.226 ms | 1.002 | [0.997, 1.003] | 0.999 | 1.000 |
| `xls_semantic_noop_edit_save` | xls-large | 0.0020 ms | 0.0018 ms | 0.923 | [0.538, 1.162] | 1.079 | 1.085 |
| control `xls_semantic_open` | xls-large | 1.3947 ms | 1.3392 ms | 0.960 | [0.954, 0.962] | 0.943 | 0.963 |
| control `xls_semantic_full_cell_scan` | xls-large | 0.0733 ms | 0.0736 ms | 1.005 | [1.002, 1.007] | 0.993 | 0.966 |
| control `xls_owned_source_control_open_one_cell` | `--ole2-file 54016.xls` | 0.5736 ms | 0.5535 ms | 0.962 | [0.921, 0.972] | 0.961 | 0.974 |

The comment and visibility generic commits move between the first run and the
later ones because the first run predates their handoff (`b85d3e534c`). The flat
rows are the corpora change 0620 described: the Number and comments corpus is a
16,995,840-byte CFB whose Workbook stream is 80,946 bytes and the RK/MulRK
corpus carries a 1.6 KB stream, so a source-backed or plan-only commit there is
CFB fingerprinting and copying, and the parse this change shortens is a sliver
of it. The generic commits move on every corpus because the render the handoff
removes is proportional to the whole archive. The `--ole2-file` selectors run
the lazy `SourceBackedWorkbook`, whose only change is three removed unreachable
branches; none reads `54016.xls` through the edit owners, which is why those
paths are timed through the probe.

### Instructions and cycles (isolation pairs)

The probe run twice under `perf stat` with 4 and 24 timed iterations and no
warmup, differenced and divided by 20; commit operations use `--reuse-source`,
so an iteration is the staged edit, the commit and its publication alone
(`instructions/isolation.json`, reviewed build):

| operation | instr/op before | after | change | cycles/op before | after | change |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `54016.xls` open | 154,144,661 | 85,426,375 | −44.6% | 53,896,767 | 35,792,984 | −33.6% |
| `54016.xls` `commit_source_backed_plan` | 153,836,255 | 85,747,297 | −44.3% | 61,342,245 | 26,412,947 | −56.9% |
| `54016.xls` `commit_source_backed` | 185,045,516 | 116,599,467 | −37.0% | 79,337,685 | 55,688,809 | −29.8% |
| `54016.xls` `commit` (Number) | 219,517,267 | 118,840,583 | −45.9% | 110,789,804 | 60,093,272 | −45.8% |
| `54016.xls` `commit` (SST text) | 230,424,097 | 129,180,485 | −43.9% | 112,403,452 | 63,453,876 | −43.5% |
| `54016.xls` comments open | 146,272,446 | 77,848,491 | −46.8% | 52,302,875 | 30,287,252 | −42.1% |
| `54016.xls` visibility open | 140,622,628 | 72,199,896 | −48.7% | 52,928,168 | 30,005,456 | −43.3% |
| `54016.xls` control `Workbook::new` | 139,964,826 | 139,739,210 | **−0.2%** | 44,699,968 | 43,209,843 | −3.3% |
| `xls-large` open | 25,472,291 | 13,473,226 | −47.1% | 7,179,989 | 6,033,228 | −16.0% |
| `xls-large` `commit_source_backed_plan` | 25,198,850 | 13,237,949 | −47.5% | 7,585,278 | 3,685,368 | −51.4% |
| `xls-large` `commit_source_backed` | 29,254,784 | 17,261,772 | −41.0% | 10,356,263 | 8,632,944 | −16.6% |
| `xls-large` `commit` (Number) | 34,771,216 | 17,359,006 | −50.1% | 15,079,813 | 9,784,724 | −35.1% |
| `xls-large` comments open | 24,761,610 | 12,837,017 | −48.2% | 6,750,186 | 5,273,417 | −21.9% |
| `xls-large` visibility open | 23,503,394 | 11,528,024 | −51.0% | 6,866,161 | 6,073,669 | −11.5% |
| `xls-large` control `Workbook::new` | 23,397,826 | 23,382,500 | **−0.1%** | 6,490,953 | 6,198,289 | −4.5% |

The earlier build's instructions agree with these to within 0.1% on every row
(`superseded-run-3eea4a8bec/`); its cycles differ, which is the next section.

**Attribution of the generic commit.** A build at the validation-only commit
(`ec747fbd64`, no handoff) prices the Number commit at 151,281,848 instructions
on `54016.xls` and 22,802,520 on `xls-large` (`attribution/`): of the
219.5 M → 118.8 M instructions removed from one generic commit on `54016.xls`,
the validation-only mode removes 68.24 M (31.1% of the commit) and the handoff
32.44 M (14.8%). In cycles the handoff's share is larger, because the second
render is SHA-256 (`sha256rnds2`) and `memmove`, far more cycles per
instruction than the parse; and in wall time larger again, because its fresh
artifact-sized buffer is faulted in by the kernel, which `cycles:u` does not
count (this is why the comments and visibility selectors' wall times fall
21–48% while their per-sample user cycles fall 2–40%).

### Build sensitivity: same instructions, up to 1.7× the cycles

The first probe run used a build of `3eea4a8bec` and measured the
`xls-large` open at 0.789 ms; the reviewed build measures 1.360 ms, with the same
instructions (13,474,151 against 13,473,226 per open). To separate the build
from the host, four binaries were run in one window, three rounds, same core
(`results/change-0746/layout/same-window-xls-large-open.txt`): the base (1.52–1.64 ms,
2.70 G instructions per process), a build of `ec747fbd64` (1.34–1.37 ms), a
fresh rebuild of `3eea4a8bec` (0.77–0.80 ms) and the reviewed build
(1.25–1.36 ms). The three after builds execute **1.4287 G instructions per
process each** and take **370 M against 594–640 M user cycles**; page faults
vary between 1,286 and 5,754 per process uncorrelated with speed. It is the
code, not the host.

`perf annotate` locates it (`layout/annotate-hot-instructions.txt`): in
both layouts the hottest instruction of the validation-only walk is a 16-byte
reload of the `CellRecord` that `CellRecord::parse` has just returned through
the stack (66% of the walk's samples in the fast build, 31% plus a cluster of
`Result`-tag branches in the slow one) — a store-to-load forwarding stall whose
cost depends on stack offsets and code alignment. The walk is 32% of the fast
build's open and 51% of the slow build's. The public reader has the same copy,
but there the per-cell map dominates and hides it.

So the reviewed source's open speedup on `xls-large` is anywhere from 15% to 50%
depending on the compiler's layout, and its instruction and allocation savings
are the stable facts. (The final source measured 0.444 against the base in one
window, and one binary's cycles per `54016.xls` open measured 29.99 M and
21.98 M four minutes apart with its instructions unchanged; section "Review
fixes".) The harness binaries (pinned rustc 1.95.0) behave like the
faster layout on the one selector that contains this walk
(`xls_semantic_one_edit_save`: 0.458, 0.466 and 0.464 across the three harness
runs). A
follow-up that validates non-kept, non-`Formula` records through the existing
measure-only instantiation (`CellRecord::measure`, change 0576) instead of
building and copying an 88-byte `CellRecord` would remove the stall and make the
walk layout-robust; it is not attempted here.

### Allocations

The probe's counting allocator armed for exactly one measured iteration
(`alloc/`; two runs per leg, identical to the unit on every row, and identical
between the two after builds). A commit row counts that iteration's open, staged
edit, commit and publication together.

`54016.xls`

| operation | allocation calls | allocated bytes | peak live bytes |
| --- | ---: | ---: | ---: |
| `cell_values::Snapshot::from_bytes` | 51,823 → 35,515 (−31.5%) | 26,077,467 → 13,942,942 (−46.5%) | 17,787,131 → 5,651,247 (−68.2%) |
| open + `commit_source_backed_plan` | 99,957 → 67,346 (−32.6%) | 44,560,408 → 20,293,446 (−54.5%) | 18,540,525 → 6,406,649 (−65.4%) |
| open + `commit_source_backed` | 121,659 → 89,048 (−26.8%) | 59,091,760 → 34,824,798 (−41.1%) | 25,628,658 → 13,167,501 (−48.6%) |
| open + `commit` (Number) | 118,272 → 85,494 (−27.7%) | 71,167,056 → 38,849,192 (−45.4%) | 23,289,770 → 13,059,015 (−43.9%) |
| open + `commit` (SST text) | 127,898 → 95,115 (−25.6%) | 74,276,715 → 41,956,763 (−43.5%) | 23,269,070 → 13,092,958 (−43.7%) |
| `comments::Snapshot::from_bytes` | 40,216 → 23,908 (−40.6%) | 20,950,624 → 8,816,099 (−57.9%) | 15,401,290 → 3,208,493 (−79.2%) |
| `sheet_visibility::Snapshot::from_bytes` | 40,214 → 23,906 (−40.6%) | 20,950,426 → 8,815,901 (−57.9%) | 15,401,228 → 3,208,693 (−79.2%) |
| control: public `Workbook::new` | 40,102 → 40,102 (0.0%) | 16,987,030 → 16,987,030 (0.0%) | 14,417,145 → 14,417,145 (0.0%) |

`xls-large`

| operation | allocation calls | allocated bytes | peak live bytes |
| --- | ---: | ---: | ---: |
| `cell_values::Snapshot::from_bytes` | 2,081 → 729 (−65.0%) | 4,676,245 → 2,188,565 (−53.2%) | 3,513,931 → 976,907 (−72.2%) |
| open + `commit_source_backed_plan` | 4,092 → 1,393 (−66.0%) | 7,755,807 → 2,782,535 (−64.1%) | 3,521,479 → 986,463 (−72.0%) |
| open + `commit_source_backed` | 4,394 → 1,695 (−61.4%) | 10,320,166 → 5,346,894 (−48.2%) | 4,605,289 → 2,070,273 (−55.0%) |
| open + `commit` (Number) | 4,476 → 1,671 (−62.7%) | 12,090,171 → 5,649,254 (−53.3%) | 3,995,094 → 1,785,075 (−55.3%) |
| `comments::Snapshot::from_bytes` | 1,994 → 642 (−67.8%) | 3,494,973 → 1,007,293 (−71.2%) | 2,923,450 → 489,177 (−83.3%) |
| `sheet_visibility::Snapshot::from_bytes` | 1,989 → 637 (−68.0%) | 3,494,753 → 1,007,073 (−71.2%) | 2,923,390 → 489,177 (−83.3%) |
| control: public `Workbook::new` | 1,927 → 1,927 (0.0%) | 2,831,642 → 2,831,642 (0.0%) | 2,761,483 → 2,761,483 (0.0%) |

The harness's `litchi-perf-baseline-alloc` was built for both legs, but the
registered XLS selectors emit no operation-scoped allocation fields (checked on
`xls_semantic_one_edit_save` and `xls_semantic_open`), so the probe's counts are
the allocation evidence. These tables are the reviewed build; the final
source's counts, which differ from them on `54016.xls` bytes and `xls-large`
calls, are in "Review fixes".

### Review fixes: the final source re-measured

The review re-verified the equivalence on its own inputs (300,000 multi-defect
worksheets identical between the base and the reviewed head, 920,000
store-against-store comparisons identical, 279 more commits with byte-identical
handoff renderings) and asked for five fixes, applied as new commits without
rewriting history:

1. the sparse-band cost of the banded occupancy map (should-fix):
   `8fa1d4b54a` and `3d49e03044`, measured below;
2. the allocation-failure documentation: `8fa1d4b54a` and `0cd989bb15`
   ("Why it is sound");
3. record numbers: corrected in place (the instruction range and controls in
   "Result", the pair-range wording, the handoff's 32.44 M, the mutation
   corpus's scope);
4. tests: `7f3c6c4b78` ("Correctness evidence");
5. the determinism fix's lowest-offender choice: documented at both code sites
   (`8fa1d4b54a`) and in "What changed".

**The occupancy map.** The reviewed store counted positions in 256-row bands of
two 256-column masks, allocating and zeroing 16 KiB for every band that held
any cell. The review's crafted, valid package — 1,000 worksheets of 256
`Number` cells, one per band (rows 0, 256, …, 65,280, column 0) — made every
owner open slower than the base on the review's host (`cell_values` 47 → 78 ms,
visibility 44 → 77 ms, comments 45 → 79 ms; the worksheet parse 29.5 → 570 µs
per sheet) and held 4 MiB of bands per worksheet. `8fa1d4b54a` gives each
occupied row one 64-byte pair of masks, allocated when its first record arrives
and found as "What changed" describes; reservations stay fallible and
failure-atomic, so a failed reservation leaves the map as it was.
`3d49e03044` answers a position on the current row inline and moves every other
lookup and insert out of line, after the first fix alone added 0.8% to the
dense opens' instructions (below). Unit tests in `cells.rs`: one cell per
256-row band retains at most 2 × 256 × 72 bytes, where the banded map held
4 MiB; 2,048 rows arriving in descending then rotated order, two cells each,
answer exactly like ascending ones and retain at most 2 × 2,048 × 80 bytes; the
grid-edge test also checks one row entry per occupied row and the beyond-grid
count.

**Sparse-band case.** Generated by the probe
(`--generate-sparse-bands-minimal`, `probe/main.rs`): the writer's globals
(1,000 `BoundSheet8` records, its XF table) and, per worksheet, `BOF`,
`DIMENSIONS`, the 256 `Number` records and `EOF`, repointed and packaged as the
only CFB stream: 4,708,352 bytes, SHA-256 `87b0d72437a27918…`, the size the
review reported. The writer's own output for the same cells is 35,208,192 bytes,
because it adds an `INDEX` and one `DBCELL` per 32-row block (2,041 per
worksheet), which then dominate the parse; it was used only for a first look and
is not tabulated.

**Design.** Three probe legs built with the identical command from identical
probe source (`probe/Cargo-{before,banded,after}.toml`): the base, the reviewed
build ("banded", `b85d3e534c`) and the final source (`3d49e03044`;
`0cd989bb15` changes only a comment). Two ABBA runs back to back — base against
final, then banded against final — then isolation pairs of all three, then
allocation counts, in one window (03:13–03:15 UTC, load average 7–10,
`review-fix/run.log`); sparse cases 2 warmups and 10 samples per process,
`54016.xls` 3 and 20, `xls-large` 5 and 60; commit rows are the `commit` phase.

| fixture | operation | base p50 | final p50 | final / base | pair range | banded p50 | final p50 | final / banded | pair range |
| --- | --- | ---: | ---: | ---: | --- | ---: | ---: | ---: | --- |
| sparse bands | `cell_values::Snapshot::from_bytes` | 60.284 ms | 27.343 ms | 0.453 | 0.446–0.463 | 75.581 ms | 26.358 ms | 0.350 | 0.346–0.357 |
| sparse bands | `comments::Snapshot::from_bytes` | 59.496 ms | 26.587 ms | 0.447 | 0.434–0.462 | 75.545 ms | 25.661 ms | 0.339 | 0.336–0.341 |
| sparse bands | `sheet_visibility::Snapshot::from_bytes` | 56.659 ms | 24.587 ms | 0.433 | 0.429–0.444 | 73.915 ms | 23.882 ms | 0.323 | 0.317–0.332 |
| sparse bands | control: public `Workbook::new` | 54.750 ms | 50.363 ms | 0.919 | 0.883–0.920 | 49.925 ms | 49.757 ms | 0.992 | 0.963–1.063 |
| `54016.xls` | `cell_values::Snapshot::from_bytes` | 10.462 ms | 4.866 ms | 0.465 | 0.452–0.478 | 4.998 ms | 4.889 ms | 0.981 | 0.971–0.990 |
| `54016.xls` | `comments::Snapshot::from_bytes` | 10.240 ms | 4.315 ms | 0.421 | 0.405–0.433 | 4.408 ms | 4.313 ms | 0.979 | 0.968–0.983 |
| `54016.xls` | `sheet_visibility::Snapshot::from_bytes` | 9.937 ms | 4.128 ms | 0.415 | 0.410–0.427 | 4.254 ms | 4.135 ms | 0.973 | 0.970–0.974 |
| `54016.xls` | `commit` (Number edit) | 23.242 ms | 10.118 ms | 0.436 | 0.434–0.439 | 10.225 ms | 10.093 ms | 0.988 | 0.980–0.990 |
| `54016.xls` | `commit_source_backed` | 15.890 ms | 8.704 ms | 0.550 | 0.539–0.564 | 8.505 ms | 8.445 ms | 0.993 | 0.993–0.994 |
| `54016.xls` | control: public `Workbook::new` | 10.430 ms | 9.910 ms | 0.946 | 0.860–1.061 | 9.179 ms | 9.089 ms | 0.990 | 0.878–1.077 |
| `xls-large` | `cell_values::Snapshot::from_bytes` | 1.528 ms | 0.680 ms | 0.444 | 0.439–0.449 | 0.685 ms | 0.673 ms | 0.981 | 0.977–0.992 |
| `xls-large` | `commit` (Number edit) | 3.442 ms | 1.566 ms | 0.456 | 0.446–0.462 | 1.566 ms | 1.557 ms | 0.990 | 0.975–0.998 |
| `xls-large` | control: public `Workbook::new` | 1.452 ms | 1.361 ms | 0.933 | 0.896–0.947 | 1.350 ms | 1.331 ms | 0.989 | 0.968–0.991 |

Instructions and cycles per operation (isolation pairs, 4 and 24 iterations,
2 and 8 on the sparse case; `review-fix/isolation.json`):

| fixture | operation | instr base | banded | final | banded vs base | final vs base | final vs banded | cycles base | banded | final |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| sparse bands | `cell_values::Snapshot::from_bytes` | 766.46 M | 1124.76 M | 468.53 M | +46.7% | −38.9% | −58.3% | 267.28 M | 333.85 M | 117.33 M |
| sparse bands | `comments::Snapshot::from_bytes` | 749.40 M | 1156.69 M | 452.36 M | +54.3% | −39.6% | −60.9% | 245.32 M | 327.03 M | 108.15 M |
| sparse bands | `sheet_visibility::Snapshot::from_bytes` | 701.98 M | 1108.91 M | 404.77 M | +58.0% | −42.3% | −63.5% | 259.97 M | 319.36 M | 99.79 M |
| sparse bands | control: public `Workbook::new` | 700.22 M | 689.35 M | 690.14 M | −1.6% | −1.4% | +0.1% | 243.84 M | 220.15 M | 220.00 M |
| `54016.xls` | `cell_values::Snapshot::from_bytes` | 154.14 M | 85.43 M | 82.39 M | −44.6% | −46.6% | −3.6% | 50.94 M | 21.49 M | 21.28 M |
| `54016.xls` | `comments::Snapshot::from_bytes` | 146.27 M | 77.85 M | 74.80 M | −46.8% | −48.9% | −3.9% | 44.98 M | 18.74 M | 18.31 M |
| `54016.xls` | `sheet_visibility::Snapshot::from_bytes` | 140.62 M | 72.20 M | 69.15 M | −48.7% | −50.8% | −4.2% | 48.08 M | 17.91 M | 17.26 M |
| `54016.xls` | `commit` (Number edit) | 219.52 M | 118.84 M | 115.79 M | −45.9% | −47.3% | −2.6% | 107.35 M | 47.94 M | 46.48 M |
| `54016.xls` | `commit_source_backed` | 185.04 M | 116.60 M | 113.59 M | −37.0% | −38.6% | −2.6% | 74.25 M | 41.29 M | 42.04 M |
| `54016.xls` | control: public `Workbook::new` | 139.96 M | 139.74 M | 139.82 M | −0.2% | −0.1% | +0.1% | 46.14 M | 43.19 M | 45.08 M |
| `xls-large` | `cell_values::Snapshot::from_bytes` | 25.47 M | 13.47 M | 12.77 M | −47.1% | −49.9% | −5.2% | 7.02 M | 3.06 M | 2.93 M |
| `xls-large` | `commit` (Number edit) | 34.77 M | 17.36 M | 16.66 M | −50.1% | −52.1% | −4.1% | 14.92 M | 6.80 M | 6.76 M |
| `xls-large` | control: public `Workbook::new` | 23.41 M | 23.37 M | 23.38 M | −0.2% | −0.1% | +0.0% | 6.21 M | 5.94 M | 5.95 M |

Allocation counts, one armed iteration, two runs per leg identical to the unit
(`review-fix/alloc/`); in parentheses the change against the base and against
the reviewed build. The base and reviewed legs reproduce this record's earlier
`alloc/` counts exactly on every operation both measured.

sparse bands

| operation | allocation calls: base → final (vs base; vs banded) | allocated bytes | peak live bytes |
| --- | ---: | ---: | ---: |
| `cell_values::Snapshot::from_bytes` | 102,283 → 74,283 (−27.4%; −77.0%) | 146,984,153 → 104,696,153 (−28.8%; −97.5%) | 108,378,185 → 29,532,617 (−72.8%; −12.4%) |
| `comments::Snapshot::from_bytes` | 93,244 → 65,244 (−30.0%; −79.2%) | 110,386,005 → 68,098,005 (−38.3%; −98.4%) | 89,866,690 → 14,096,485 (−84.3%; −7.2%) |
| `sheet_visibility::Snapshot::from_bytes` | 92,235 → 64,235 (−30.4%; −79.5%) | 110,282,893 → 67,994,893 (−38.3%; −98.4%) | 89,916,189 → 14,096,485 (−84.3%; −7.6%) |
| control: public `Workbook::new` | 91,156 → 91,156 (+0.0%; +0.0%) | 91,140,426 → 91,140,426 (+0.0%; +0.0%) | 85,188,653 → 85,188,653 (+0.0%; +0.0%) |

`54016.xls`

| operation | allocation calls: base → final (vs base; vs banded) | allocated bytes | peak live bytes |
| --- | ---: | ---: | ---: |
| `cell_values::Snapshot::from_bytes` | 51,823 → 35,515 (−31.5%; +0.0%) | 26,077,467 → 14,793,662 (−43.3%; +6.1%) | 17,787,131 → 5,912,879 (−66.8%; +4.6%) |
| open + `commit_source_backed_plan` | 99,957 → 67,346 (−32.6%; +0.0%) | 44,560,408 → 21,994,886 (−50.6%; +8.4%) | 18,540,525 → 6,668,281 (−64.0%; +4.1%) |
| open + `commit_source_backed` | 121,659 → 89,048 (−26.8%; +0.0%) | 59,091,760 → 36,526,238 (−38.2%; +4.9%) | 25,628,658 → 13,167,501 (−48.6%; +0.0%) |
| open + `commit` (Number) | 118,272 → 85,494 (−27.7%; +0.0%) | 71,167,056 → 40,550,632 (−43.0%; +4.4%) | 23,289,770 → 13,320,647 (−42.8%; +2.0%) |
| open + `commit` (SST text) | 127,898 → 95,115 (−25.6%; +0.0%) | 74,276,715 → 43,658,203 (−41.2%; +4.1%) | 23,269,070 → 13,354,590 (−42.6%; +2.0%) |
| `comments::Snapshot::from_bytes` | 40,216 → 23,908 (−40.6%; +0.0%) | 20,950,624 → 9,666,819 (−53.9%; +9.6%) | 15,401,290 → 3,470,125 (−77.5%; +8.2%) |
| `sheet_visibility::Snapshot::from_bytes` | 40,214 → 23,906 (−40.6%; +0.0%) | 20,950,426 → 9,666,621 (−53.9%; +9.6%) | 15,401,228 → 3,470,325 (−77.5%; +8.2%) |
| control: public `Workbook::new` | 40,102 → 40,102 (+0.0%; +0.0%) | 16,987,030 → 16,987,030 (+0.0%; +0.0%) | 14,417,145 → 14,417,145 (+0.0%; +0.0%) |

`xls-large`

| operation | allocation calls: base → final (vs base; vs banded) | allocated bytes | peak live bytes |
| --- | ---: | ---: | ---: |
| `cell_values::Snapshot::from_bytes` | 2,081 → 769 (−63.0%; +5.5%) | 4,676,245 → 2,195,349 (−53.1%; +0.3%) | 3,513,931 → 969,675 (−72.4%; −0.7%) |
| open + `commit_source_backed_plan` | 4,092 → 1,473 (−64.0%; +5.7%) | 7,755,807 → 2,796,103 (−63.9%; +0.5%) | 3,521,479 → 979,231 (−72.2%; −0.7%) |
| open + `commit_source_backed` | 4,394 → 1,775 (−59.6%; +4.7%) | 10,320,166 → 5,360,462 (−48.1%; +0.3%) | 4,605,289 → 2,063,041 (−55.2%; −0.3%) |
| open + `commit` (Number) | 4,476 → 1,751 (−60.9%; +4.8%) | 12,090,171 → 5,662,822 (−53.2%; +0.2%) | 3,995,094 → 1,777,843 (−55.5%; −0.4%) |
| `comments::Snapshot::from_bytes` | 1,994 → 682 (−65.8%; +6.2%) | 3,494,973 → 1,014,077 (−71.0%; +0.7%) | 2,923,450 → 489,177 (−83.3%; +0.0%) |
| `sheet_visibility::Snapshot::from_bytes` | 1,989 → 677 (−66.0%; +6.3%) | 3,494,753 → 1,013,857 (−71.0%; +0.7%) | 2,923,390 → 489,177 (−83.3%; +0.0%) |
| control: public `Workbook::new` | 1,927 → 1,927 (+0.0%; +0.0%) | 2,831,642 → 2,831,642 (+0.0%; +0.0%) | 2,761,483 → 2,761,483 (+0.0%; +0.0%) |

What this shows:

- **Sparse bands.** In this window the reviewed build's three owner opens
  executed 1.47–1.58× the base's instructions and 1.23–1.33× its cycles, and
  allocated 4,270,552,153 bytes per open, 29× the base; its open took
  75.581 ms against the base's 60.284 ms (the two ABBA runs, 40 s apart). The
  final source runs these opens at 0.433–0.453× the base and 0.323–0.350× the
  reviewed build, executes 39–42% fewer instructions than the base, allocates
  29–38% fewer bytes than the base, and retains 18 KiB of occupancy per
  worksheet (256 × 64 bytes of masks and 256 × 8 bytes of index). The 1,000
  worksheets' counts reconcile exactly with the two layouts: per worksheet the
  final source makes 249 fewer allocation calls and allocates 4,165,856 fewer
  bytes than the reviewed build — 256 bands of 16 KiB plus their index's growth
  (263 calls, 4,202,432 bytes) against the two row vectors' growth to 256
  entries (14 calls, 36,576 bytes, every outgrown buffer included).
- **Dense.** The win is kept and slightly extended: the `54016.xls` open is
  0.465× the base and 0.981× the reviewed build, every dense operation is
  0.973–0.993× the reviewed build in wall time and 2.6–5.2% below it in
  instructions. The price is memory on dense sheets. On `54016.xls` the final
  source allocates 4.1–9.6% more bytes than the reviewed build and holds up to
  8.2% more peak live bytes (the open: +850,720 bytes allocated, +261,632 peak),
  because each row now also carries an 8-byte index entry and the row vectors
  grow geometrically — up to 2× slack, with every outgrown buffer counted in
  allocated bytes — where a full 256-row band held exactly 64 bytes per row. On
  `xls-large` it makes 4.7–6.3% more allocation calls (40 more per validating
  open: two growing vectors in each of four worksheets; the commit rows open
  twice) with bytes within +0.7%. Against the base it still allocates 38–54%
  fewer bytes and holds 43–78% less peak on `54016.xls`.
- **Controls.** The public reader is unchanged: instructions within ±0.1% of
  the reviewed build (−0.1% to −1.4% against the base, the `DecodeEveryCell`
  instantiation of the generic walk), allocations identical to the call. Its
  0.919–0.946 wall ratios against the base are the layout effect of "Build
  sensitivity".

**The first fix alone, and how far one window can mislead.** `8fa1d4b54a`
alone executed 0.8% more instructions than the reviewed build on the dense
opens (86.11 M against 85.43 M on the `54016.xls` open,
`review-fix/inline-split-isolation.json`). In a first ABBA window at load
average 20 (`review-fix/first-window/`) it measured 1.048× the reviewed build on
the `54016.xls` open and 1.294× on the `xls-large` open, and an isolation pair
at 03:08 UTC priced its `54016.xls` open at 29.99 M cycles against the reviewed
build's 21.43 M in the same run. Four minutes later, from a byte-identical copy
of the same binary, the same pair measured 21.98 M against 21.43 M, with its
instructions unchanged (86.11 M against 86.09 M). Hardware counters on the
`xls-large` open in between put the first fix's excess in front-end starvation
and branch mispredictions, not in store forwarding
(`first-window/frontend-counters.txt`). `3d49e03044` removed the excess
instructions (82.39 M, −3.6% against the reviewed build); the tables above are
the later, quieter window, and the first window is retained, not quoted. Its
allocation counts equal the final source's on all 12 operations it counted. As
in "Build sensitivity", a cycle or wall difference between two builds of this
walk is not attributable to the code unless the instructions move with it: on
this host the same two binaries showed a 40% and a 2.6% cycle gap four minutes
apart, while the reviewed build measured 21.43 M both times.

### Regression flags (every case above 5%, none hidden in an average)

**Found by the review, fixed.** The reviewed build regressed the sparse-band
shape ("Review fixes"): on the review's host every owner open took 1.66–1.76×
the base's time; here its instructions were 1.47–1.58× and its cycles
1.23–1.33× the base's, and it allocated 29× the base's bytes. The final source
runs that shape at 0.433–0.453× the base. The two fixtures of the reviewed
build's runs (below) do not contain that shape, so none of those runs could
see it.

**Remaining against the reviewed build**, all on dense inputs and all still far
below the base: allocated bytes +4.1–9.6% and peak live bytes up to +8.2% on
`54016.xls`, allocation calls +4.7–6.3% on `xls-large` (explained in "Review
fixes"); in the final window's isolation pairs, cycles +1.8% on
`commit_source_backed` and +4.4% on the public-reader control with instructions
−2.6% and +0.1%, host variation of the size "Review fixes" documents, with wall
ratios of 0.993 and 0.990 for the same two operations. In the final window no
changed operation is slower than the base or the reviewed build in any pair;
the pairs above 1.05 are all the unchanged public-reader control (`54016.xls`
1.061 against the base and 1.077 against the reviewed build, medians 0.946 and
0.990; sparse bands 1.063 against the reviewed build, median 0.992), its
instructions flat.

**The reviewed build's runs.** On the two fixtures measured there, no changed
operation regresses in any run, pair or statistic; every paired ratio of every
probe case in both builds is below 0.93. In the quoted harness run
(identically built before leg) the only pair above 1.05 is
`xls_semantic_noop_edit_save` (one pair 1.162, median 0.923), at 1.6–3.0 µs
per sample: the exact no-op returns before any changed code
(`Transaction::commit`'s early return precedes both the handoff and the
reopen), so its pairs are timer noise in either direction. The other two harness
runs flagged, all at per-pair level and none reproducing in the quoted run:

- against the prebuilt base: `xls_semantic_noop_edit_save` (median 1.079, one
  pair 1.831), `xls_comments_source_backed_edit_save` (median 1.040, two pairs
  1.079; both legs' processes are bimodal at ≈33.7 and ≈36.4 ms and two after
  processes landed in the slow mode; an isolation pair prices one sample at
  −0.12% instructions, and against the self-built base at +0.02% instructions
  and +0.01% cycles) and the control `xls_semantic_full_cell_scan` (one pair
  1.054 at 75 µs);
- the first run: `xls_semantic_noop_edit_save` (median 1.085),
  `xls_comments_eager_batch_edit_save` (median 1.026, one pair 1.080, isolation
  −0.02% instructions and −0.04% cycles, before its handoff existed) and
  `xls_comments_source_backed_batch_edit_save` (one pair 1.072, bimodal).

Controls: the public `Workbook::new` runs 2–8% faster in the after binaries
(`xls_semantic_open` 0.960 against the identically built base) with its
instructions flat (−0.2%, −0.1% in the probe isolation) and its allocations
identical. Not claimed; it is the same layout effect as "Build sensitivity", in
the other direction, on a function that is now compiled as one of two
instantiations.

## Correctness evidence

All three differentials were run on the reviewed build, and again on the
final source (the last list in this section).

**First-error matrix.** Change 0633's 29-case matrix (defects in the globals
BOF, `FilePass`, the SST, every XF, `BoundSheet8`, the worksheet BOF and EOF,
`Number`/`RK`/`BoolErr`/`LabelSst`/`Formula` payloads, stray `String`, framing,
container) through `Snapshot::from_bytes`: **byte-identical on the base, the
determinism-fix-only leg and the final leg** (`correctness/matrix-*.jsonl`).
The unit-test copy, `snapshot_open_first_error_matrix_is_frozen`, passes
unchanged.

**Corpus census.** 0620's corpus binary, extended with the comments owner
(open, a generic replacement, a same-length source-backed replacement), the
visibility owner (open, generic, source-backed) and a digest of the public
reader's every decoded cell and sheet entry: 126 fixtures, 1,319 rows per run
(published-artifact SHA-256 or exact refusal for every path), two runs per leg,
**one SHA-256 (`62b59bc3…`) for all six outputs of the base, fix-only and final
legs**. It publishes on 35 fixtures through both source-backed numeric paths, 40
through the generic Number commit, 51 through the SST-text commit, 9 comment
replacements on each comment path and 62 visibility changes on each visibility
path; the public-reader digest covers the 117 fixtures that reader accepts.

**Mutation differential.** A new scratch binary derives 200 deterministic
record-level mutations of each fixture's Workbook stream (duplicated,
colliding, truncated, dropped, swapped and bit-flipped records, stray
`String`s, out-of-range XFs, repeated companions, columns outside the grid,
globals bit flips), repackages each with every `BoundSheet8` repointed, and
records the outcome of `cell_values::Snapshot::from_bytes` (refusal text, or a
digest of every editable cell), the comments and visibility opens, and the
public reader (refusal, or a digest of every decoded cell): 93 fixtures, 18,600
packages, 74,400 outcomes per leg, 36,584 of them refusals
(`correctness/mutation-summary.md`). **The three legs' outputs have one SHA-256
(`37b966fe…`).**

**In-crate differential tests** (`workbook/validation_only_tests.rs`), run in
debug and release:

- every `.xls` in the repository (117 accepted, 9 refused by the public reader)
  through both modes, comparing the refusal's text and `Debug`, or every non-cell
  fact of the workbook and its worksheets;
- 6,215 kept cells (every accepting fixture, up to 48 per sheet spread over the
  sheet plus its formulas, plus vacant and beyond-grid positions) against the
  public reader's `Debug` of the same cell, and a check that nothing else was
  decoded;
- 5,088 mutated worksheet substreams of ten real fixtures (2,131 refused)
  parsed by both stores under both compatibility profiles, comparing the refusal
  or every non-cell fact and the kept cells — this is where the per-worksheet
  refusals the package level swallows are compared;
- synthetic worksheets reaching each duplicate and `Array` branch, including the
  two refusals only a trailing unresolved string `Formula` can produce ("not
  materialized", "non-Formula cell"), an orphan `PtgExp`, a beyond-grid
  duplicate and `MulRk`/`MulBlank` over a `Number`;
- repackaged workbooks with one mutated worksheet, compared at the package level;
- the three handoff tests and the two determinism tests described above.

After the review (`7f3c6c4b78`, and reruns on the final source):

- **A frozen worksheet-level multi-defect first-error matrix**
  (`worksheet_first_error_matrix_matches_the_base_for_multi_defect_inputs`,
  cases in `validation_only_tests/multi_defect_cases.rs`). The package-level
  differentials cannot see a worksheet's refusal, because `parse_workbook`
  swallows it, so this matrix calls the worksheet walk directly on 27 synthetic
  worksheets. Twenty-five hold two or more defects, or valid duplicates around
  a defect: duplicate cells before an out-of-range XF, truncated records on
  either side of one, reserved style XFs, pending and stray `String`s and a
  string continuation that is not a `Continue`, orphan `PtgExp`s, `Array`
  ranges without dimensions, with duplicate anchors or members, outside the
  dimensions or unmaterialized, a non-`Formula` member, `ShrFmla` without its
  `Formula`, a duplicate anchor claiming a second companion, `MulRk` and
  `MulBlank` with bad ranges or XFs over earlier cells, beyond-grid duplicates,
  `UserSViewEnd` without a begin, and `DVAL` cut short by a cell or by EOF. Two
  the base accepts: valid duplicate, shared and array cells, and a stream that
  ends without an EOF on a pending string `Formula` after duplicates. The
  expected outcomes were generated on the base `009d515bef` by a temporary
  test, run three times with identical output
  (`review-fix/correctness/base-multi-defect-matrix.txt`, generator
  `scripts/base-multi-defect-matrix-generator.diff`, never committed); the test
  asserts them for the public reader's store and for the validation-only store
  with no kept cell and with every position kept.
- **`kept_cell_refuses_a_position_kept_on_a_different_tab`**: a two-tab
  workbook keeps `(1, 1)` on the second tab; `kept_cell` returns the public
  reader's cell there, refuses `(1, 1)` on the first tab and `(2, 2)` on both,
  and the first tab decodes no cell.
- **Reruns on the final source** (`3d49e03044` code, `binaries.sha256`): the
  29-case matrix, the census (two runs) and the 200-round mutation
  differential give the base's SHA-256s (`6f3d68d3…`, `62b59bc3…`,
  `37b966fe…`; `correctness/output-sha256.txt`), and every in-crate test above
  passes at the final head (section "Verification").

## What is not claimed

- No claim is registered; `performance_claim: none`.
- Timing is warm, in-memory, single-CPU, on a host shared with other agents; no
  cold-cache, RSS, physical-I/O, throughput, concurrency or producer claim is
  made. Allocation figures are the probe's counting allocator over one
  iteration, not process RSS.
- The wall-clock speedup of the validation-only walk depends on the compiled
  layout (previous section); only its instruction and allocation reductions are
  stable across builds.
- The real-file results are scoped to `54016.xls`; `xls-large` is the harness's
  generated shape. 33 of the 126 fixtures never reach a `cell_values` snapshot
  (0620's census); those refusals are unchanged, not faster.
- The public reader is not claimed faster: its instructions and allocations are
  unchanged.
- The handoff is proved byte-identical on the corpus census, by debug
  re-derivation across the crate's tests and by the release tests above; a
  divergent second render on some untested package would now publish the
  validated rendering instead, which is the one the reopen checked.
- No change to `SourceBackedWorkbook`, the pivot editor, the OOXML crates or any
  shared substrate.
- The sparse-band fix is measured on one crafted shape. That the final
  occupancy map is proportional to the occupied rows in general rests on its
  structure and its unit tests, not on a sweep of shapes.
- The final source is not claimed faster than the reviewed build in wall time on
  dense inputs: its 0.973–0.993 ratios are inside the window-to-window
  variation "Review fixes" documents. Its 2.6–5.2% instruction reduction is the
  stable fact, and it allocates more bytes than the reviewed build there.

## What is left, with prices from the after profile

- **Overlay fingerprints dominate the generic commit now.** On `54016.xls` its
  single remaining `render_copy_through` is 53% of the commit, and five SHA-256
  passes over the whole artifact (planning ×2, composed-source preflight, write
  preflight, write validation) are about 42% of it. For an owned, immutable
  in-memory source these repeat one another; change `1535b141e5` elided one such
  pass for the plan-only path. A `litchi-cfb` overlay change with its own
  contract argument.
- **The validation-only walk's store-forwarding stall** (previous section), and
  the twenty per-record collector `feed_record` calls, now the largest parts of
  the walk (about 32% of the open).
- **SST decoding** allocates a temporary UTF-16 buffer per string: the
  deallocation under `SharedStringTable::parse_from_records` is about 10% of the
  validation-only open and `String::from_utf16` another 3.7%. Shared with the
  public reader.
- The offset inventory's `push_entry` (11% of the open); `PackageEditor::open`
  still clones the whole input for its `OleFile` cursor, and
  `commit_source_backed` clones the materialized target once more to open it.
- **Dense-sheet occupancy memory.** The final map allocates 4.1–9.6% more bytes
  than the banded one on `54016.xls` (the row index and geometric growth).
  Reserving the row vector up front from the worksheet's `DIMENSIONS` row span,
  capped by what the substream's remaining bytes could hold as cell records so
  that a lying `DIMENSIONS` cannot reserve more than the records present could
  fill, would remove most of that slack without bringing back a cost
  proportional to the row range. Not attempted here.

## Verification

`results/change-0746/gates.txt`, rerun after the review at the final head
`0cd989bb15` with a fresh `CARGO_TARGET_DIR` under `targets/0746`
(`scripts/gates-final.sh`; deleted afterwards): `cargo fmt --all --check`;
`cargo check -p litchi-xls --all-targets` and `-p litchi --features
doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets`; `cargo clippy -p litchi-xls
--lib --no-deps -- -D warnings`; `cargo test -p litchi-xls` (73 suites, 1,492
passed, 1 ignored; 1,488 at `b85d3e534c`, plus the two occupancy tests and the
two review tests); the release-mode differential, handoff and determinism tests
(13 passed, the two review tests included); `cargo test -p litchi --features
doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` (29 suites, 382 passed, 7 ignored);
`RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-xls --no-deps`;
`check_crate_boundaries.py`, `non_iwork_gate.py verify` and
`check_perf_claims.py --mode structural`, all exit 0. `cargo clippy -p
litchi-xls --all-targets -- -D warnings` exits 101 on one pre-existing
`unusual_byte_groupings` error in `tests/xls_query_index_cache.rs:546`, a file
this change does not touch; the same command fails identically on the base
checkout, and passes on this branch with that one lint allowed. The same gates
passed at `b85d3e534c` before the review. The harness is unchanged, so its own
tests and the coverage validator were not rerun.

## Cleanup

`results/change-0746/cleanup.json`. The probe, corpus, matrix and mutation
sources, every script and every raw timing report are retained in the packet;
binaries, `perf.data` files, the target directories (`targets/0746`,
`targets/0746-before`), the `0746-before-src` worktree and the scratch directory
were deleted after the SHA-256s were recorded. The review fixes repeated this:
their evidence is in `review-fix/`, and `targets/0746` (91 GB, two gate builds
of about 43 GB each and the probe builds), the re-created `0746-before-src`
worktree (10 GB) and the scratch contents (0.5 GB) were deleted after the final
gates. The worktree and branch are kept.
