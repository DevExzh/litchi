# 0585: resume the cursor's chain walk too, and the XLS shared-string scan stops rewalking it

Status: retained. `performance_claim: none` — this record carries exact,
deterministic chain-link counts across the whole XLS fixture corpus, with a
per-fixture text digest on both sides. **No timing, cycle, allocation or RSS
measurement was taken, and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements candidate 1 of change [0584](0584-ole2-profile-at-head.md) and
closes the open question change [0579](0579-cfb-resumable-chain-walk.md) left
behind. It does not close it the way that record expected.

## The question 0579 left open, and its answer

0579 ended with an explicit non-result:

> **The worksheet scan is untouched.** It uses a `SharedOleStreamCursor` (change
> 0566) and keeps it. […] Whether that path would also be faster on a hint is an
> open question this record does not answer.

**It would not.** `WorksheetScan` (`crates/litchi-xls/src/workbook/source.rs:2114`)
holds one `SharedOleStreamCursor` and touches it in exactly two places —
`read_exact` (`:2239`) and `skip_forward` (`:2393`) — never reconstructing it.
The cursor's own state *is* a resumption position:
`SharedOleStreamCursorState::{Fat,MiniFAT} { sector, within }`
(`crates/litchi-cfb/src/shared.rs:88-93`), committed after every read
(`shared.rs:2769-2771`). Forward movement runs through `normalize_state`
(`shared.rs:2901`), which performs exactly one `next_chain_sector` per sector
boundary crossed. **The scan already walks each chain link once, and a hint
cannot remove work that is done once.**

0579 had in fact measured this from the other direction without drawing the
inference: its rejected cursor leg walked 1,076 links against the hint form's
2,099, "because a cursor's read loop *is* the chain walk and never walks a prefix
separately".

## What the refutation found instead

The cursor removes prefix walks *once it exists*. **Nothing removed the walk that
builds it.** `stream_cursor_at` (`shared.rs:720` before this change) reached its
starting offset through `cursor_chain_sector`, a cold walk from the stream's
first sector, and there was no hinted variant — `stream_cursor_at_hinted`
returned zero matches over the repository. Change 0584 priced it: of a flagship
one-cell query's 4,428 chain links, **2,238 come from `stream_cursor_at`** — more
than the entire globals scan — from two call sites that each restart at sector
zero.

The sharper statement is the asymmetry with the OOXML side. `litchi-xlsx`
memoises its shared-string table once per workbook in a `OnceLock<Box<[Text]>>`
(`crates/litchi-xlsx/src/workbook/model.rs:388`, resolved at `:1464`). The XLS
source-backed path had no equivalent: `SharedStringSstScan`
(`crates/litchi-xls/src/records.rs:990`) stores only offsets, and
`resolve_shared_string` built a **fresh cursor per string cell**, each walking
the `Workbook` allocation chain from its first sector to reach the string table.

## What changed

**`crates/litchi-cfb/src/shared.rs`** gains `SharedOleFile::stream_cursor_at_hinted`,
which is `stream_cursor_at` with a resumable chain position. It reuses the
existing `StreamChainHint`/`ChainResume` machinery from 0579 unchanged — there is
no second hint type — and `stream_cursor_at` now delegates to it with a fresh
hint, exactly as `read_stream_range` delegates to `read_stream_range_hinted`. One
implementation of the lookup, the bounds check and the walk; the two cannot
drift. `cursor_chain_sector` takes a `(sector, walked)` resume pair instead of a
start sector, and is `(start_sector, 0)` whenever no hint applies, which is the
walk it performed before hints existed.

**`crates/litchi-xls/src/workbook/source.rs`** gains a `SharedStringResolver`
holding two per-scan constants that `resolve_shared_string` was rebuilding on
every string cell: the workbook stream path, previously a fresh
`Vec<&str>` allocation per call, and one `StreamChainHint`.

The resolver's hint is **deliberately dedicated to the shared-string region and
not shared with the worksheet cursor.** A hint retains one position, and the
worksheet cursor sits permanently past the string table. One hint serving both
would be discarded as a backward step by every resolve *and* again by every
sheet, saving nothing on either path while adding bookkeeping to both. That
constraint is recorded in the type's own doc comment so a future reader does not
"simplify" it away.

## Why it is sound

Not an appeal to care — an equivalence. A chain walk is the deterministic
recurrence `s_{k+1} = fat[s_k]`, so the sector at ordinal `N` is a function of
`N` alone. A hint records a pair `(M, s_M)` that this same walk produced by
traversing `0..M`; resuming there and traversing `M..N` yields the **identical**
`s_N` a cold walk yields, for every chain, well-formed or not.

- The per-link checks — marker validity, index bounds, ENDOFCHAIN — are all
  inside `next_chain_sector` (`shared.rs:3099-3117`) and still run on every link
  from `M` to `N`. The links from `0` to `M` were checked when the hint recorded
  them.
- Chain **cycles** are not a walk-time concern: they are rejected at
  `OleFile::open` (`crates/litchi-cfb/src/allocation_validation_tests.rs:381`,
  `:407`), so every chain a cursor walks is already acyclic.
- The loop is bounded by `ordinal` before and after, so no termination property
  moves.
- `resume_from` discards a hint from a different reader (`ptr::eq`), a different
  directory entry, a different first sector, a different allocation table, or a
  **later** ordinal. A discarded hint costs exactly what an unhinted call costs.
  A hint can shorten a walk; it can never redirect a read.

## Measured

Environment: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc
1.95.0. Chain-link counts are exact and deterministic, so they were taken from
**one instrumented build carrying two runtime switches** rather than from paired
release builds — see
[`instrumentation.patch.md`](results/change-0585/instrumentation.patch.md) for
the patch, the three legs, and why that is sound for this quantity and would not
be for cycles.

**The setup validated itself before it was believed.** A `pre-0579` control leg
reproduces change 0579's recorded pre-change figure of **5,796** chain links for
an open of `ConditionalFormattingSamples.xls` exactly, and the `0579-only` leg
reproduces that record's post-change **2,099** exactly. The delta below is
`0579-only → this-change`, so 0579's saving is **not** credited to this change.

Scenario: full text extraction, every `.xls` fixture under `test-data/`.

| | `0579-only` | this change | removed |
| --- | ---: | ---: | ---: |
| **text-extraction chain links, 86 distinct fixtures** | **6,469,099** | **2,233,171** | **65.5%** |
| open chain links | 7,532 | 7,532 | **0.0%** |
| text digest mismatches | — | — | **0** |
| fixtures that got worse | — | — | **0** |

Per fixture, the twelve largest movers:

| fixture | `0579-only` | this change | removed |
| --- | ---: | ---: | ---: |
| `54016.xls` | 4,689,582 | 1,671,048 | 64.4% |
| `59858.xls` | 1,327,605 | 432,546 | 67.4% |
| `29942.xls` | 133,489 | 815 | **99.4%** |
| `WithCustomViews.xls` | 94,401 | 55,551 | 41.2% |
| `pivottable_dates_grouping.xls` | 71,409 | 34,589 | 51.6% |
| `FormulaEvalTestData.xls` | 67,330 | 12,127 | 82.0% |
| `15228.xls` | 32,567 | 15,729 | 51.7% |
| `45365-2.xls` | 26,712 | 540 | **98.0%** |
| `external_name.xls` | 5,909 | 1,551 | 73.8% |
| `duprich1.xls` | 2,495 | 101 | 96.0% |
| `41139.xls` | 2,202 | 219 | 90.1% |
| `48968.xls` | 1,089 | 102 | 90.6% |

**32 of the 86 fixtures save nothing at all** — `24207.xls`, `45365.xls`,
`46136-NoWarnings.xls`, `50939.xls`, `53109.xls`, `55341_CellStyleBorder.xls` and
26 others. These are files with one string resolve, or none, where there is no
prior position to resume from. The first cold walk of a scan is never removed;
only the repeats are.

**The open is untouched, to the link.** 7,532 before and 7,532 after across the
corpus. This change operates entirely inside the scan, and the counter proves it
rather than the prose asserting it.

## Correctness

Every fixture's extracted text carries an FNV-1a digest on both legs.
**86 fixtures, 0 mismatches**, plus identical text lengths. 14 distinct fixtures
are refused by the library on both legs with the same error; their refusal is
unchanged.

Eleven tests were added to `litchi-cfb`, each pricing a distinct property:
hinted and unhinted cursors agree across a FAT stream, across a **fragmented**
FAT chain, and across a MiniFAT stream; one hint serves both a hinted read and a
hinted cursor; a hint from another stream, from another reader, and from past the
requested offset are each ignored; a hint changes neither a refusal nor its text,
including **on a malformed chain**; a hint demonstrably resumes rather than
rewalks; and the `n(n-1)/2` versus `n-1` arithmetic is pinned.

Gates: `cargo fmt -p litchi-cfb -p litchi-xls -p litchi-doc -- --check` clean;
`cargo clippy -p litchi-cfb -p litchi-xls -p litchi-doc --all-targets` clean;
`python3 tools/non_iwork_gate.py check`, `clippy` and `doc` all exit 0;
`python3 tools/check_crate_boundaries.py` exits 0;
`litchi-cfb` **347 tests**, `litchi-xls` **1,380 tests**, `litchi-doc`
**1,170 tests** — **2,897 in total, zero failures**. The counts were taken twice,
by the coordinator and by an independent verification pass, and agree exactly.

One gate fails, and it is **pre-existing and unrelated**:
`tools/test_check_crate_boundaries.py` has one failure of 1,104,
`test_iwa_package_edge_is_archive_owned_and_host_edge_cannot_return`, asserting
240 policy edges against 241 found. It is an **iWork** assertion, and iWork is
outside this workstream. It is not assumed to predate the batch — it was
verified to: stashing this batch in full and re-running the suite against a
clean `e927e139c` reproduces the identical single failure. No file this batch
touches is an input to that test; no `Cargo.toml` and nothing under `tools/` is
modified. It is recorded here rather than fixed, because the crate it concerns
belongs to a concurrent workstream.

## Limitations

**No timing was measured.** Chain links are a count of dependent loads, not a
latency. `GOAL_AUDIT.md` records that change 0579 removed 1.24% of instructions
on `54016.xls` and 6.19% of cycles on the same change, so the relationship
between this 65.5% and any wall-clock figure is *unknown in both directions*.
`perf stat`, callgrind, cold-cache, allocation, peak-RSS, range-source,
concurrency and cross-platform measurement all remain outstanding for this
change. They were planned and not completed: the session's measurement agents
were terminated by a rate limit, and the deterministic counter was chosen as the
evidence that could be taken reliably rather than as the evidence one would
prefer.

**`ConditionalFormattingSamples.xls` could not be measured end to end.** The
library refuses its text extraction (`Invalid record 0x0006: shared Formula
metadata requires a leading PtgExp token`) on every leg. Change 0584's
byte-layout model predicted 703,937 shared-string links for a full scan of that
fixture; **that scan is one the library will not perform on that file**, so the
prediction stands as a statement about the file's structure and not about any
reachable scenario. Its open path is still measured and is unchanged at 2,099.

**The corpus prediction was optimistic.** Change 0584's model predicted 74.4%
corpus-wide against the 65.5% measured here. The model counted FAT links only,
assumed the hint is recorded at the construction ordinal and never advanced, and
— on the largest fixture — decoded only 829 of 16,055 shared-string entries
before its naive SST walker stopped, so its backward-step distribution for that
file was drawn from a 5% sample. The measurement supersedes it.

**The per-sheet cursor term is not implemented.** Change 0584 priced
`WorksheetScan::new`'s construction walk at 28,143 links on a 16-sheet fixture,
about 91% of it resumable. Taking it needs a *second* hint with a disjoint
lifetime, for the thrashing reason above. It is left open.

**One-cell reads gain nothing**, and the open figures show why: a one-cell query
performs one sheet-cursor construction and one shared-string resolve, so neither
has a prior position to resume from. The scenario that gains is one that resolves
many strings.

**The scan this sits inside is itself unbounded.** `query_cell` scans to EOF and
never stops when the target cell is found (`source.rs:2765`); on `54016.xls` one
query frames 37,929 records to return one value. That is change 0574's
opportunity 6, rejected on ADR 0005 mandatory-validation grounds. Nothing here
should be read as a bound on the cost of a query.

**The `refs` hoist is not separately priced.** Removing the per-call
`Vec<&str>` allocation from `resolve_shared_string` is in the same change as the
hint, and the chain-link counter cannot see allocations. Its effect is unmeasured.

## Retained evidence

[`results/change-0585/`](results/change-0585/): `chain-legs.tsv`, the raw
three-leg output for every `.xls` fixture; `chain-summary.txt`, folded from it;
`instrumentation.patch.md`, the exact switches and the argument for using one
binary; and change 0584's byte-layout model `sst_walk.py` with its corpus output,
retained here as the prediction this measurement supersedes.
