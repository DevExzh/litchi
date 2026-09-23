# 0746: XLS edit owners validate through a validation-only mode of the complete reader, and their generic commits publish the rendering they already validated

Status: retained, implemented. `performance_claim: none` — this record carries
paired ABBA timings on two builds, isolation-pair instruction and cycle counts,
deterministic allocation counts, a three-leg correctness census and a frozen
first-error matrix. They are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `009d515bef`; branch `perf/0746-xls-validation-only-parse`; five code
commits: `3a2f233cdc` (deterministic multi-defect refusals), `ec747fbd64` (the
validation-only mode), `fa88d7cc9b` (the validated-render handoff in
`cell_values`), `3eea4a8bec` (comments and visibility owners adopt the mode) and
`b85d3e534c` (the handoff in the comments and visibility commits).

## Result

Every XLS edit owner that opened the complete eager reader, `Workbook::new`,
only to validate a package and read a few facts from it now runs **the same
parser in a validation-only mode**: every record is validated identically, but
no per-cell map is built and no cell is decoded except the few the commit reads
back. The three generic commits (`cell_values`, comments, sheet visibility)
additionally publish the CFB rendering their package publication already
validated, instead of rendering the same editor state a second time.

What the change removes is stable: **44–51% of the instructions of every open
and commit it touches** (37% for `commit_source_backed`), **26–68% of the
allocation calls, 41–71% of the allocated bytes and 44–83% of the peak live
bytes**, with the public reader's instructions (−0.2%, −0.0%) and allocations
(identical to the call) unchanged. Wall-clock, measured on the shipped source
(`b85d3e534c`), paired ABBA on CPU 20:

| fixture | operation | before p50 | after p50 | paired ratio | 95% CI | same code, other build |
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

The last column is the same comparison run earlier against a build of
`3eea4a8bec`, whose `cell_values` and open code is identical to the shipped
source. **The two builds execute the same instructions (±0.1%) and differ by up
to 1.7× in cycles** on the validation-only walk, because the walk is now lean
enough for a code-layout-dependent store-forwarding stall to dominate it
(section "Build sensitivity"). Both are reported; the shipped build's numbers
are the conservative ones. The registered selectors, built with the pinned
toolchain from the shipped source and compared against a base harness built with
the identical command, agree with the faster layout:
`xls_semantic_one_edit_save` (the generic commit on `xls-large`) **3.535 →
1.643 ms (0.464)**, `xls_visibility_eager_edit_save` **25.47 → 13.33 ms
(0.523)**, `xls_comments_eager_edit_save` 21.54 → 16.67 ms (0.792),
`xls_numeric_eager_rk_mulrk_edit_save` 2.395 → 1.235 ms (0.516).

Every published artifact, refusal and readback is byte-identical to the base
across a 126-fixture census and 18,600 mutated packages (section "Correctness
evidence").

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
one pass, with no allocation and no change for accepted inputs. Two tests parse
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
  in a banded occupancy bitmap (two 256-bit masks per row, 16 KiB per touched
  256-row band, fallibly reserved; an exact map for columns at 256 or beyond,
  which the single-cell decoder does not bound and the public reader keeps), and
  decoded only when its position was asked for.

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
bitmap exactly when an earlier record there reached `add_cell`, which is when
the public reader's map holds a cell there, because every such record decodes to
exactly one cell at the record's own `(row, column)`. The `Formula` flag is
rewritten on every record at the position, matching the map's last-record-wins
insert, and it is what `Cell::set_array_formula` tests (`formula_metadata` is
present exactly for `Formula` records). The only change of order is inside the
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

**The only new refusal is resource exhaustion.** The bitmap bands, the
beyond-grid map and the kept-cell list are reserved fallibly and fail with a
typed `Allocation` error where the public reader's infallible cell map would
abort the process. Their size is bounded by the 65,536-row grid (at most 4 MiB
of bands per worksheet) and by the cell records present, and is far below the
per-cell map they replace.

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
p50s, the median of the four adjacent-pair after/before p50 ratios, and a
10,000-resample percentile bootstrap interval over those four ratios. Every
process p50/p95/mean and every raw report is in `latency-probe/`,
`latency-harness-selfbuilt/` and `latency-harness/`.

### Paired timing, registered selectors (shipped source)

Before leg built with the identical command (`latency-harness-selfbuilt/`).
The last two columns are the same design against the prebuilt base binary
(`latency-harness/`) and, for the first run, against a build of `3eea4a8bec`
before the comments and visibility handoff existed
(`superseded-run-3eea4a8bec/`):

| selector | corpus | before p50 | after p50 | paired ratio | 95% CI | vs prebuilt | first run |
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
(`instructions/isolation.json`, shipped build):

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
the validation-only mode removes 68.2 M (31.1% of the commit) and the handoff
32.5 M (14.8%). In cycles the handoff's share is larger, because the second
render is SHA-256 (`sha256rnds2`) and `memmove`, far more cycles per
instruction than the parse; and in wall time larger again, because its fresh
artifact-sized buffer is faulted in by the kernel, which `cycles:u` does not
count (this is why the comments and visibility selectors' wall times fall
21–48% while their per-sample user cycles fall 2–40%).

### Build sensitivity: same instructions, up to 1.7× the cycles

The first probe run used a build of `3eea4a8bec` and measured the
`xls-large` open at 0.789 ms; the shipped build measures 1.360 ms, with the same
instructions (13,474,151 against 13,473,226 per open). To separate the build
from the host, four binaries were run in one window, three rounds, same core
(`results/change-0746/layout/same-window-xls-large-open.txt`): the base (1.52–1.64 ms,
2.70 G instructions per process), a build of `ec747fbd64` (1.34–1.37 ms), a
fresh rebuild of `3eea4a8bec` (0.77–0.80 ms) and the shipped build
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

So the shipped source's open speedup on `xls-large` is anywhere from 15% to 50%
depending on the compiler's layout, and its instruction and allocation savings
are the stable facts. The harness binaries (pinned rustc 1.95.0) behave like the
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
the allocation evidence.

### Regression flags (every case above 5%, none hidden in an average)

No changed operation regresses in any run, pair or statistic; every paired ratio
of every probe case in both builds is below 0.93. In the quoted harness run
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

All three differentials were run on the shipped build.

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

## Verification

`results/change-0746/gates.txt`, at `b85d3e534c`: `cargo fmt --all --check`;
`cargo check -p litchi-xls --all-targets` and `-p litchi --features
doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets`; `cargo clippy -p litchi-xls
--lib --no-deps -- -D warnings`; `cargo test -p litchi-xls` (73 suites, 1,488
passed, 1 ignored); the release-mode differential, handoff and determinism
tests; `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` (29
suites, 382 passed, 7 ignored);
`RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-xls --no-deps`;
`check_crate_boundaries.py`, `non_iwork_gate.py verify` and
`check_perf_claims.py --mode structural`, all exit 0. `cargo clippy -p
litchi-xls --all-targets -- -D warnings` exits 101 on one pre-existing
`unusual_byte_groupings` error in `tests/xls_query_index_cache.rs:546`, a file
this change does not touch; the same command fails identically on the base
checkout, and passes on this branch with that one lint allowed. The harness is
unchanged, so its own tests and the coverage validator were not rerun.

## Cleanup

`results/change-0746/cleanup.json`. The probe, corpus, matrix and mutation
sources, every script and every raw timing report are retained in the packet;
binaries, `perf.data` files, the target directories (`targets/0746`,
`targets/0746-before`), the `0746-before-src` worktree and the scratch directory
were deleted after the SHA-256s were recorded. The worktree and branch are kept.
