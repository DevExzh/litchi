# Change 0059: XLS fixed-width numeric inventory carry-forward

Date: 2026-08-12

Status: accepted after review-trigger analysis

Production base: `cad602d26f79379d10558b4e38a5316b701cef77`

## Decision

Carry the private BIFF8 cell-offset inventory through changed `Number`, `RK`,
and `MulRk` commits whose record family and Workbook-stream length remain
fixed. Share every untouched worksheet inventory and immutable workbook-global
resource allocation; clone only an edited worksheet's private inventory.

This is private native XLS work. It adds no public API, dependency, cache,
runtime, lock, unsafe code, persisted index, archive abstraction, or new format
capability. OOXML, RTF, ODF and every iWork/IWA crate are unchanged.

## Problem and attribution

`Snapshot::from_package_editor` reparsed the complete rewritten Workbook
stream after the common object editor had rendered, reopened, recaptured, and
validated the candidate. That parse rebuilt workbook-global metadata and an
offset-bearing `Entry` for all 8,192 cells even when one exact numeric value
field changed and no record offset could move. It was followed by the required
independent complete `Workbook::new` validation.

The frozen control profile attributes `Snapshot::from_package_editor` and its
`parse_worksheet -> push_entry` subtree directly below
`Transaction::commit`. The matched candidate removes that subtree from the
commit stack. Both profiles have zero lost samples; the reports contain 32K
before and 27K after cycle samples.

## Proof boundary and fallbacks

The carry-forward route is selected only when all effective fixed-width
changes:

- retain their exact `Number`, `RK`, or `MulRk` storage family;
- have a numeric target with the source field width;
- have no structural or workbook-resource operation; and
- coexist with at least one effective changed cell.

Before carrying any inventory, the commit proves:

- the final Workbook stream has the exact source length;
- every target field contains the exact IEEE-754 or RK encoding requested by
  the semantic change;
- sorted target ranges do not overlap; and
- every byte outside those ranges is identical to the source Workbook stream.

Any storage conversion, Boolean, error, blank, SST, formula, structural,
sheet-name, authored-resource, or formula-resource edit retains the former full
owner parse. A failed fixed-field proof is a typed refusal, not a weaker
publication.

The complete candidate CFB still passes the common editor's render, reopen,
recapture, protection and unchanged-stream checks. The final snapshot still
opens the complete candidate through the independent public `Workbook` reader,
requires every mutable tab to survive as a worksheet, and reads every changed
numeric cell from that public model before publication. Date-formatted numeric
cells accept either the public float or date-time projection but require exact
`f64` bits. Exact-source patching, stale-source refusal, inverse restoration,
diagnostics, protection policy, non-Workbook stream preservation, no-op
identity, allocation bounds and all structural/resource fallbacks remain.

Focused tests prove exact field-only mutation, resource sharing, sequential
numeric commits, public-reader reopen, sharing of an untouched worksheet
inventory, exact patch/inverse behavior, opaque-stream preservation and
refusal when any byte outside the certified field changes. Existing real XLS
tests retain `MulRk`, protection, other-stream and public reopen coverage.

## Measurement method

Both frozen release binaries use the unchanged standalone harness. SHA-256:

- control: `672deac8e3056a267941cd4c80a8356e72094d8e8e08de25bfd6759a4b807773`;
- candidate: `955dad8ec0d12103b63dc677a2eeab8c9c61351f285225a7fcd6fa4289d9cee1`.

Environment: Rust 1.95.0 / LLVM 22.1.2, Linux 6.8.0-101-generic,
x86-64 AMD EPYC 9575F VM, Rust system allocator, CPU 11 pinned with `taskset`,
and `perf_event_paranoid=1`. The deterministic large XLS contains four sheets
and 8,192 numeric cells in a 163,840-byte CFB. Its archive SHA-256 is
`228c6585a4d26141aebfaf7b08844a2ee445b269d406006a1fdb0484619120fb`;
the 161,040-byte Workbook stream SHA-256 is
`f806d23f52c978f5215b05fd232b055725a2605d52122ea74ce0cec357ea9386`.

The primary used one 500-sample ABBA pair and one 500-sample reverse BAAB pair,
each with 50 warmups. Pooling the four legs gives 2,000 observations per state
and balances the temporal positions. Every iteration verifies the exact
expected output, one changed numeric cell/stream, complete grid reopen,
forward patch and exact inverse outside the timed interval.

## Latency result

| Large XLS one edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 1.582 ms | 1.458 ms | **-7.83%** |
| mean | 1.597 ms | 1.480 ms | **-7.37%** |
| p95 | 1.757 ms | 1.631 ms | **-7.20%** |
| p99 | 1.915 ms | 1.796 ms | **-6.22%** |

The before mean 95% interval is 1.593-1.602 ms; the after interval is
1.474-1.485 ms. The first long ABBA control legs drifted 6.5%, so the reverse
BAAB pair was added rather than accepting the first pool alone. In the reverse
pair, within-state p50 drift was 2.2% before and 0.9% after. The complete pooled
summary is in the
[`primary summary`](../results/xls-inventory-carry-primary-summary.json).

## Guardrails and review trigger

The large guard ABBA pair used 30 warmups and 250 samples per leg:

| Case | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| Open | 1.318 ms | 1.329 ms | +0.83% | -0.05% | -3.27% |
| Full cell scan | 94.941 us | 89.488 us | -5.74% | -9.85% | -13.31% |
| Exact no-op edit/save | 2.849 us | 2.343 us | -17.75% | -17.87% | -18.56% |

The combined process also triggered review for two nanosecond-scale public
reader operations that cannot execute the changed transaction branch. A
separate 1,000-sample-per-state repeat measured one-cell access at 290 -> 230
ns p50 and worksheet listing at 250 -> 261 ns p50 (+4.4%, 11 ns). Listing
mean/p95 moved 9.6%/6.0%, so the trigger is disclosed rather than hidden.

No public reader, worksheet, cell accessor or harness source changed. The
frozen `Worksheet::get_cell` symbol is 0xa5 bytes in both binaries, and its
normalized 62-line instruction stream is identical. The normalized harness
iterator closure is likewise identical modulo anonymous LLVM symbol hashes.
The review is retained in the
[`layout record`](../results/xls-inventory-carry-readonly-layout.txt) and
[`repeat summary`](../results/xls-inventory-carry-readonly-repeat-summary.json).
No production workaround based on one executable's link/allocator placement
was added.

## CPU and memory evidence

Whole-process `perf stat` ABBA legs used 10 warmups and 100 measured commits:

| Counter, A+B totals | Before | After | Delta |
|---|---:|---:|---:|
| Task clock | 3,748.94 ms | 3,044.88 ms | **-18.78%** |
| Cycles | 18.382 billion | 15.012 billion | **-18.33%** |
| Instructions | 49.183 billion | 48.592 billion | -1.20% |
| Branches | 13.492 billion | 13.377 billion | -0.85% |
| Branch misses | 61.072 million | 56.187 million | **-8.00%** |
| Cache references | 4.908 billion | 5.816 billion | +18.49% |
| Cache misses | 68.963 million | 62.951 million | **-8.72%** |
| Page faults | 71,277 | 46,751 | **-34.41%** |

The cache-reference trigger was reviewed: misses fell 8.72%, so the miss ratio
improves alongside lower cycles, task time, latency and memory. CPU migrations
were zero in every leg.

Matched Heaptrack processes over two warmups and 20 samples report allocation
calls 241,291 -> 239,850 (-0.60%), temporary allocations 22,580 -> 21,729
(-3.77%), peak heap 8.12 -> 7.67 MiB (-5.54%), Heaptrack-inclusive RSS
20.39 -> 19.80 MiB (-2.89%), and identical 544-byte runtime/tool leakage.
Four uninstrumented GNU Time processes per state report mean maximum RSS
30,880 -> 30,816 KiB (-0.21%, flat at process resolution).

Raw primary, guard, repeat, counter, profile, Heaptrack and RSS artifacts use
the `xls-inventory-carry` prefix under [`results`](../results/).

## Validation

Passed on the final source:

- focused `cell_values` tests: 26 passed;
- complete `litchi-xls --all-features` unit, integration and doc-test suite:
  987 library tests plus every integration target and doctests passed;
- warning-denied XLS production-library Clippy;
- unchanged performance harness: 33 release tests and warning-denied
  all-target release Clippy;
- warning-denied ODF common all-target Clippy and rustdoc, rechecking the
  deprecation cleanup from `1194fbc7f`;
- formatter, JSON parsing, link-target, whitespace and final-diff checks.

Warning-denied XLS rustdoc remains blocked by four pre-existing unrelated
broken intra-doc links (`Snapshot`, `Transaction`, `Record::topic`, and
`Record::parse`); none is in the changed module documentation.

## Limitations and next work

This shortcut covers only same-family fixed-width numeric fields. It does not
remove the complete public Workbook open, common CFB publication, exact patch
construction, output materialization or verification. General fixed-width
Boolean/string/formula edits and any offset-changing/resource/structural edit
retain the full private owner parse.

Remaining high-return non-iWork work includes ODP snapshot-to-transaction slide
projection reuse and source-backed OOXML page-break publication. RTF has no
new measured owner above the 5% decision threshold at this revision.
