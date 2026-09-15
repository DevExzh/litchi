# XLSB survey: remaining optimization opportunities

Area: `crates/litchi-xlsb` (98,826 lines). HEAD 2fc5fc657, working tree clean. No XLSB
performance record exists at all: every XLSB-specific record (0304-0376, 16 records) carries
`performance_claim: none` — they are CRUD/correctness records, not measurements. This survey
is therefore the first XLSB-specific performance evidence in the program.

## 1. Path map

**Open (eager, the only path for cell CRUD).** `litchi_xlsb::Workbook::new` →
`OpcPackage::from_reader_with_limits` (`crates/litchi-xlsb/src/workbook/package.rs:764-770`)
decompresses and retains every OPC part up front (general OOXML behavior, see 0581 below), then
`Workbook::from_opc_package_with_external_link_limits` (`package.rs:781-800+`) builds a fully
eager `Workbook` (`workbook/model.rs:20-40`) whose fields — `shared_strings: Vec<SharedString>`,
`styles: StylesTable`, `formula_context`, `pivot_cache_definitions`, `structured_tables`,
`chart_sheets`, `sheet_drawings`, `connections` — are all owned, non-lazy state, i.e. this
constructor parses the full non-worksheet-body feature surface at open. `litchi_xlsb::Package`
(`package/mod.rs:224-260` open, `:456-457` save) is a second, thinner eager wrapper
(`pub struct Package(OpcPackage, ..)`, per 0581) used by structural feature transactions
(timeline, slicer, scenarios, xml_maps, connections, external_link).

**Worksheet cell read (no caching, two independent full-worksheet decoders).**
(a) `Workbook::cell_values(idx)` (`package.rs:86-101`) → `cell_values::workbook::read_with_limits`
(`cell_values/workbook.rs:33-39`) → `package.get_part` (cheap, already decompressed) →
`cell_values::worksheet::read_shared` (`cell_values/worksheet.rs:1563-1651`), which loops
`Records::with_limits(&source, limits.raw)` to completion with **no break**, decoding every
supported cell record into `entries: Vec<Entry>` before returning `Snapshot{source, entries,
limits}` (`worksheet.rs:577-580`). Every call rebuilds this from scratch; `Workbook` caches
nothing across `cell_values()` calls.
(b) `Workbook::worksheet(idx)` (`workbook/codec/codec.rs:14-45`) independently re-resolves the
catalog/relationship and calls `read_worksheet` → `host::cells_reader::CellsReader` (crate-private,
`host/cells_reader/mod.rs`) → `sheet::Worksheet{cells: BTreeMap<(u32,u32),Cell>, ..}`
(`sheet.rs:126-219`), built by `while let Some(cell) = cells_reader.next_cell()? { worksheet.add_cell(cell); }`
(`workbook/codec/codec.rs:207-217`). `Workbook.worksheets: Vec<Worksheet>` (`workbook/model.rs:22`)
looks like a cache slot but is never pushed to or indexed anywhere in the crate — dead, annotated
`#[allow(dead_code, reason = "staged host integration")]`.
Neither path can return one cell without materializing every cell of the target worksheet.

**Edit + save (whole-package clone, whole-workbook reparse, on every commit including no-ops).**
`Workbook::apply_cell_values` (`package.rs:137-153`) unconditionally clones `self.package`
(`:143`) before it is known whether the commit changes any bytes, calls
`cell_values::workbook::apply_with_external_link_limits` (`cell_values/workbook.rs:60-95`), and
**unconditionally reparses the result into a new `Workbook`** (`package.rs:150-151`) even when
the inner call took its no-op early return (`workbook.rs:66-70`) and left the clone byte-identical
to `self.package`. For a real edit, the inner function *also* clones again (`workbook.rs:73`),
reparses that clone into a throwaway `Workbook` purely to run `validate_dependencies` against it
(`:76-78`, `:92`, `:97-179`), then discards it and hands back unparsed bytes (`:93`) that the
caller reparses a second time. Detail in ranked opportunity 1.

**Deferred/source-backed path (exists, but only for `.text()`).** `workbook::source::
SourceBackedWorkbook`/`SourceBackedWorksheet` (`workbook/source.rs:1-120`) defer per-part
decompression and never build a queryable cell store — `SourceBackedWorksheet::materialize`
streams straight into a `SequentialTextWriter` (0304). Reachable only via
`litchi::sheet::Workbook::from_bytes(..).text()`; not wired to `cell_values`, `Package`, or the
eager `Workbook`.

**Malformed-input defenses (all present, not proportional to size in a way that weakens them):**
`raw::record::Limits::DEFAULT` caps payload at 64 MiB and string units at 1,048,576
(`raw/record.rs:26-30`); `Cursor::guard` bounds-checks every scalar/string read
(`raw/cursor.rs:66-79`); `entries.try_reserve(1)` before every cell push converts allocation
failure into a typed `Error::Allocation` instead of trusting a declared count
(`cell_values/worksheet.rs:~1631`); `ExternalLinkLimits`, `litchi_opc::ReadLimits`, and
`cell_values::Limits` are all caller-configurable finite ceilings.

## 2. Remaining opportunities, ranked

### 1. `apply_cell_values` clones and reparses the whole workbook even for a proven no-op, and reparses twice for a real edit
**GOAL step 1** (eliminate unnecessary work), touches step 2 (copying).

**Mechanism.** `Workbook::apply_cell_values` (`package.rs:137-153`) always executes `let mut
candidate = self.package.clone()` (`:143`) before checking whether `commit` changes anything, and
always ends with `*self = Self::from_opc_package_with_external_link_limits(candidate, ..)`
(`:150-151`) — a full reparse of the entire eager `Workbook` (shared strings, styles, pivot
caches, structured tables, chart sheets, connections, every worksheet catalog entry). When the
inner `apply_with_external_link_limits` (`cell_values/workbook.rs:60-95`) finds `updated ==
part.blob()` (`:69-70`, a *provable* no-op — the checklist's own "Exact no-op" contract) it
returns immediately without touching `candidate` further, yet the caller still pays for the clone
already taken and still reparses the untouched bytes as if new. For a real edit, the inner
function clones `package` again (`:73`), reparses that clone into `parsed` purely to run
`validate_dependencies` (`:76-78`, `:92`), throws `parsed` away, and writes the still-unparsed
bytes back (`:93`) — which the caller then reparses a **second, independent** time. Net: a no-op
commit pays for 1 clone + 1 full reparse it can prove is unneeded; a real one-cell edit pays for
3 `OpcPackage::clone()` calls and 2 independent full `Workbook` reparses of byte-identical
candidate bytes, one of whose results (`parsed`) is computed and discarded.

**Code evidence:** `crates/litchi-xlsb/src/workbook/package.rs:137-153`;
`crates/litchi-xlsb/src/cell_values/workbook.rs:60-95` (`:66-70` no-op return, `:73` clone,
`:76-78` reparse-to-validate, `:92` `validate_dependencies`, `:93` write-back).

**Record status:** none found (grepped all 16 XLSB records and "reparse"/"candidate.clone" across
`docs/performance/`). Nearest precedent is 0525 (accepted XLSX "unchanged-cell readback
reconstruction" — a conceptually similar whole-snapshot revalidation on commit, accepted as
necessary), but 0525 does one reconstruction and doesn't address an unconditional reparse of a
*proven* no-op; XLSB's shape is measurably heavier and includes that extra no-op case.

**Size — MEASURED**, `xlsb_crud`, `test-data/poi/test-data/spreadsheet/testVarious.xlsb` (22,715
bytes, 17 parts, 1 sheet, 48 cells), release build, `taskset -c 2`, 3 warmup/30 samples, single
leg, no A/A control (raw: `$SCRATCH/xlsb_crud_default.json`):
| case | p50 | vs read-only baseline |
| --- | ---: | --- |
| `selected_worksheet_cell` (read only) | 1.506 ms | baseline |
| `noop_transaction_commit_save` | 2.766 ms | **+1.26 ms (+84%)** for 0 changed bytes |
| `edit_one_existing_scalar_save` | 4.189 ms | +1.42 ms more, for the second clone+reparse+validate |
Instruction counts not captured (no callgrind run this batch). Magnitude at realistic workbook
size (hundreds of parts, thousands of cells) is unknown — no such XLSB fixture exists in this
corpus (see §4).

**Scenarios:** CRUD checklist "Opened-document CRUD scenarios: Exact no-op" (directly — the no-op
branch is unconditionally wasteful today) and "Update each field family". Harness selectors:
`xlsb_crud --case noop_transaction_commit_save|edit_one_existing_scalar_save`.

**ADR/GOAL, risk:** no ADR identified that requires re-deriving the workbook twice; validation
still happens exactly once either way this is fixed. Skipping the reparse on the proven-no-op path
is **low risk** (bytes are provably identical). Reusing `parsed` from `workbook.rs:76-78` as the
new `self` instead of reparsing is **low-medium risk** — needs confirming `parsed`'s fields are
exactly what a "fresh" `from_opc_package_with_external_link_limits` would produce (they should be,
it's the same constructor), and that atomicity ("failure leaves the published snapshot
unchanged") is preserved by only assigning `*self` after `validate_dependencies` succeeds, which
the current control flow already guarantees.

**Falsified if:** a callgrind/perf run shows `validate_dependencies` itself (not the reparse)
dominates the measured delta, or a field of `Workbook` turns out to need re-derivation from the
*caller's* pre-edit state in a way `parsed` cannot supply — unchecked this batch.

### 2. Cell-level XLSB reads inherit the OOXML-wide eager `OpcPackage` open; XLSB already has a working deferred path, just not wired to cell reads
**GOAL step 2** (eliminate unnecessary I/O/decompression/parsing).

**Mechanism:** see Path map. Every `litchi_xlsb::Workbook`/`Package` entry point decompresses and
retains every part regardless of which worksheet is touched.

**Record status: COVERED by 0581** ("`OpcPackage::open` retains the archive plus every
decompressed part," decision "not to implement" — candidates C2/fallible-lazy-`Part::blob`
[950 call sites across 7 crates, **127 in litchi-xlsb**] and C3/route format crates to
`SourceBackedPackage` [blocked: no source-backed type has a general `save()`] are frozen pending
an ADR). 0581 already names `litchi-xlsb`'s `Package::open` (`package/mod.rs:225→260`, confirmed
unchanged at `:224-260` in this tree) and `Package::save` (`:456→457`, confirmed unchanged). This
entry exists so the coordinator does not re-price the same OOXML-wide question inside an
XLSB-scoped record — it is **not new**, except for the measurement and observation below.

**New contribution (this survey):** XLSB is the only OOXML format in the corpus with a *shipped*
deferred path (`SourceBackedWorkbook`, 0304/0316/0326/0330/0333) — proof the pattern works
end-to-end, not just on paper. Measured on the same run as above: `full_text` (source-backed)
p50 **137 μs** vs `open_identify`/`worksheet_catalog`/`selected_worksheet_cell`/
`full_stored_cell_scan` (all through eager `Workbook::new`) p50 **1.50-2.84 ms** — an **11-20×**
gap on a 22.7 KB, 17-part, 1-sheet fixture that is far too small to show 0581's 254× figure.
`open_identify` (facade `detect_file_format_from_bytes` + eager open) at 2.84 ms is roughly double
`open_direct`'s 1.5 ms, consistent with 0569's general "pre-detect-then-open" finding (0569's
disposition there is a documentation fix, not code — see §3).

**Scenarios:** every read scenario that needs less than the whole workbook. Harness selectors:
`open_identify`, `worksheet_catalog`, `selected_worksheet_cell`, `full_stored_cell_scan` vs
`full_text`.

**ADR/GOAL, risk:** exactly 0581's — needs the same proposed ADR; nothing XLSB-specific to solve
first except that `SourceBackedWorkbook` would need a `cell_values`-shaped materialization
(currently it only feeds a sequential text writer), which is its own design question.

**Falsified if:** 0581's own falsification condition (peak-retention multiplier too small on
realistic corpora to clear the admission gate) — unchanged by this survey.

### 3. No partial/indexed single-cell read; both worksheet representations fully materialize before returning any cell
**GOAL step 1/2.**

**Mechanism:** `cell_values::worksheet::read_shared` (`cell_values/worksheet.rs:1563-1651`) has no
early exit; `Snapshot::cell(reference)` then does `unique_index(&entries, reference)`
(`worksheet.rs:1868-1884` area), a scan over the *already fully built* vec that also detects
duplicate coordinates and returns a typed ambiguity error — a correctness guarantee a naive
early-exit would silently break. `sheet::Worksheet::get_cell` (`sheet.rs:126-190`) has the same
shape via a `BTreeMap`, fully populated by `workbook/codec/codec.rs:207-217` before any query.
(Aside, checked and **not a defect**: `sheet::Worksheet::add_cell` silently replaces on duplicate
coordinate — last-write-wins — while `cell_values` treats duplicates as ambiguous and errors; this
divergence is intentional and covered by a same-file test,
`sheet.rs:507-521 cells_iterator_returns_clones_and_observes_replacement`, which is exactly named
for the replacement behavior it asserts.)

**Record status:** none found for XLSB. Nearest precedent: 0574 opportunity 6 (XLS/CFB, "make
whole-container validation proportional to the touched closure" — rejected: "ADR 0005 names
mandatory structural validation as an open-time obligation... Not available") and the program's
established rejection of "early exit from query_cell" for XLS on the same ADR 0005 ground. Neither
tested XLSB; the reasoning is format-agnostic and, for `cell_values` specifically, is reinforced
independently by the duplicate-reference-ambiguity contract above.

**Size:** code-verified (no early exit exists) with certainty. Timing is inconclusive:
`selected_worksheet_cell` (1 cell) vs `full_stored_cell_scan` (48 cells) are statistically
indistinguishable, p50 1.506 ms vs 1.502 ms (0.3% apart, inside the ~4% p50 noise floor) — fully
consistent with full materialization dominating either way, but the 48-cell fixture is too small
to separate that from "the ~1.5 ms open floor swamps any real per-cell-count difference."
Magnitude at realistic scale (thousands of rows) is **unknown**.

**Scenarios:** "Read and snapshot scenarios: Selection is checked". Selectors:
`selected_worksheet_cell` vs `full_stored_cell_scan`.

**ADR/GOAL, risk: HIGH.** Likely blocked by ADR 0005 the way 0574 opp 6 was, and separately
constrained by the duplicate-coordinate contract, which any bounded design must preserve (e.g. via
a cheaper duplicate-detection pass than full `parse_entry` decode — itself an unexplored
sub-design). Needs a frozen design record before implementation, same classification as 0574 opp 6.

**Falsified if:** a realistic-scale XLSB fixture shows the open floor still dominates a full-sheet
scan even at thousands of rows (opportunity moot regardless of the ADR question), or cheap
duplicate detection costs as much as full decoding anyway.

## 3. Looks like an opportunity but is not

- **Pre-detect then open (`detect_file_format_from_bytes` + `open_xlsb_workbook_from_bytes`).**
  Real, measured (§2.2, `open_identify` ≈ 2× `open_direct`), but 0569 already investigated this
  exact shape for the facade generally and concluded the fix is caller-side documentation ("tell
  callers not to pre-detect"), not a code change to the opener. Nothing XLSB-specific to add.
- **General OOXML eager retention (0581).** Real and large, but frozen pending an ADR — see §2.2.
  A reader who hasn't grepped the record set would propose "make `OpcPackage::open` lazy" as if
  novel; it is the most thoroughly-priced open question in the whole program.
- **`sheet::Worksheet::add_cell` silently overwriting duplicate coordinates.** Looks like a
  correctness gap next to `cell_values`'s ambiguity error; confirmed intentional and tested
  (`sheet.rs:507-521`) — not reported further.
- **`entries.try_reserve(1)` per cell instead of a capacity hint from a declared row/cell count.**
  Looks like a missed preallocation; it is deliberate bounded-allocation hardening (trusting a
  declared count before validating content is exactly the pattern ADR 0006 forbids) — removing it
  would weaken malformed-input defenses, not speed up well-formed input materially.

## 4. Measurement blockers

- **Corpus ceiling.** All 15 `.xlsb` fixtures under `test-data/` (verified count) top out at
  22,715 bytes (`testVarious.xlsb`) and 48 stored cells, one worksheet; the other 14 are 7.5-11 KB.
  No fixture with multiple worksheets, thousands of rows, or hundreds of OPC parts exists anywhere
  in the repo (`find . -iname "*.xlsb"` was exhaustive, including `3rdparty/`). Every "size" figure
  above that depends on per-cell or per-part scaling is therefore a floor-dominated measurement,
  not a scaling curve — opportunities 1 and 3's real-world magnitude cannot be established without
  either a synthetic large XLSB fixture or a modelled estimate, neither attempted this batch.
- **No instruction-count evidence.** `xlsb_crud` is wall-clock only (`Instant`, p50/mean/p95/p99,
  no callgrind/perf integration, no allocator-metrics hookup despite `tools/perf-baseline`'s
  `allocation_metrics.rs`/`allocator-metrics` feature existing for other bins). All sizes in this
  report are single-leg timings with the program's standard ~4%/14% p50/p99 noise floor, not
  instruction counts. A callgrind pass over `edit_one_existing_scalar_save` vs
  `noop_transaction_commit_save` would directly attribute opportunity 1's cost and was not run
  this batch (budget).
- **Harness generality confirmed but unexercised.** `xlsb_crud --fixture` works against a second
  fixture (`test-data/ooxml/xlsb/cond_format.xlsb`, verified this batch, exit 0,
  `$SCRATCH/xlsb_crud_condfmt.json`) — the 8-case matrix is not hardcoded to `testVarious.xlsb`,
  so widening the corpus later only requires new/larger fixtures, not harness changes.
- **No correctness defects found.** The one candidate (duplicate-coordinate handling divergence)
  was traced to an intentional, tested design choice, not a bug — see §3.

## Bottom line

XLSB is not a low-return area for lack of waste — it has the same eager-everything shape the rest
of OOXML has (0581) plus its own additional whole-workbook-reparse-on-commit duplication that
0581 doesn't cover — but it **is** a low-return area *for this corpus*: every fixture is small
enough that the fixed cost of `Workbook::new`'s all-features-eager parse dominates every scenario,
burying per-cell and per-part scaling questions in noise. Opportunity 1 (no-op still reparses,
real edit reparses twice) is the one finding here that is both fully code-verified and clearly
actionable without a new ADR; opportunities 2 and 3 are real but respectively frozen (0581) and
high-risk/precedent-blocked (ADR 0005, à la 0574 opp 6).
