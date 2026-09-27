# 0798 final results review

This review checks the finished packet against `summary.md`, `analysis.json`,
the 30-row `root-audit.json`, the control and census receipts, the retained
quality/build records, the restoration and cleanup witnesses, and the failure
audit. It is a bounded evidence review; it does not rerun any build, test,
capture, profiler, or replay command.

## Evidence reconciliation

The packet contains the frozen schedule: 15 plain control reports from the
before binary and 30 census reports from the after binary, in forward and
reverse case order. All 45 reports contain one sample. The offline analysis
records 45 reports and samples, with both baseline and census semantic/output
parity checks true. The independent audit records 45 reports and samples and
30 census rows, with `native_timing_claim` false. Its two blocks agree on
every retained grouped row, event counter, lifecycle flag, and raw tag byte.

Across the two identical census repeats there are 295,600 iterator instances,
438,800 successful attributes, 734,400 `next` calls, and 295,600 terminal
`None` results. The event totals conserve exactly: 734,400 equals 438,800
successful yields plus 295,600 end yields. All 295,600 instances are fully
consumed and dropped before finish. Clones, error yields, early drops, partial
consumption, never-advanced rows, live-at-finish rows, and saturated counters
are all zero. The exact lexical histogram is retained; its frequencies sum to
295,600 and its weighted attribute count sums to 438,800.

The per-operation table in the main report is correctly expressed for one
repeat; the second repeat is identical. Its rows sum to 147,800 instances and
219,400 attributes, which double to the across-repeat totals above. Large
capture is the useful concentration point: 30,374 one-attribute instances
(49.528%) and 30,816 two-attribute instances (50.249%) account for 99.777%
of its 61,327 instances. This supports examining a first-attribute design for
these fixtures, but it supplies no timing ratio or workflow speed estimate.

The controls and census reports retain exact source, fixture, publication,
readback, and semantic identities against the sealed 0794 records. The
initial independent reader incorrectly included `metrics.elapsed_ns` in that
identity. The retained correction excludes that one protocol-opaque timing
field, keeps the other semantic metrics in the comparison, and leaves the raw
reports unchanged. The corrected root audit and offline analysis agree.

## Protocol and custody

The source and operation scope match the frozen packet: only
`litchi-opc::xml_attributes::CheckedAttributes` is counted, on the probe's
calling thread, during the declared capture, commit, and lifecycle regions.
The records do not cover other helper copies, unchecked or lenient iterators,
empty attribute tails skipped before iterator construction, foreign threads,
or other producers and workloads. Drop scans and diagnostic allocations are
part of the instrumented environment, so elapsed values remain excluded from
interpretation.

The retained quality and both four-gate build records are successful, and the
receipt schedule binds each lane to its copied binary. The restoration witness
matches all 9,196 production file identities at the base revision. The cleanup
witness records removal of the owned target and 990,773,730 logical bytes,
while retaining both binary hashes and sizes; it does not claim that those
bytes were a performance measurement. The failure audit accounts for the
invalid initial feature mapping, the isolated test shadowing error, the
needless-borrow Clippy failure, and the plain-probe setup fixes. No captured
case is silently omitted or retried as a result.

## Interpretation

The census establishes a repeatable distribution for the generated PPTX
operation regions. It does not establish that production callers consume
every checked iterator fully: this packet observed no partial, error, or clone
cases, and skipped empty tails are not represented by zero-count rows. The
isolated helper tests cover those lifecycle and parser cases, but they do not
turn this workflow census into broad caller coverage.

The common one- and two-attribute classes justify a future candidate that
avoids replaying the first attribute while preserving first-error ordering and
bounded hostile-input behavior. Any candidate still needs fresh semantic,
public-workflow, resource, and cross-format gates. The 0794 and 0797
rejections remain in force, production is unchanged, and this evidence makes
no latency, allocation, RSS, instruction, speedup, or adoption claim.
