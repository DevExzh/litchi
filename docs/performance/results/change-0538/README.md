# 0538 XLSX planning allocation measurement enabler

The harness adds allocation observations around `edit_sheets`, after selector
construction and before updates/commit/publication. Existing planning timings
keep their operation boundary. Planning evidence includes the live returned edit
at the endpoint, and uses the existing checked process-global System allocator
observer. Commit and publication remain separate regions. No production parser
change or performance improvement is claimed.

The [plan](plan.json) classifies 0537 as progress and binds the parent revision.
The [source manifest](source-manifest.json) and [ADR bindings](adr-manifest.json)
identify the tested sources and unchanged normative contracts. The production
parser remains at the baseline revision; the 0537 candidate is still a draft.

`run.py quality` runs XLSX harness tests, allocator observer/wrapper tests,
all-target feature checks, warning-denied Clippy and rustdoc, formatting and
crate boundaries, then builds the two smoke binaries. The separate integration
test launches both binaries and checks measured versus unavailable observations.
`run.py capture` retains medium and dense-sparse reports for all eight unmanaged
and managed cell-value selectors, with one warmup and three samples each.
All commands run serially in an owned target directory and retain logs/receipts.
Receipts retain the exact commands and build settings. Reproduce captures with
fresh output paths; the runner refuses to overwrite a sealed bundle or an
existing receipt.

The initial XLSX run passed 42 tests and failed one filesystem fixture write
with OS `QuotaExceeded` in the default `/tmp`. Its failed receipt/log remain
retained. `run.py resume_quality` reran that exact test successfully with
`TMPDIR` inside the owned target directory, then continued the remaining gates.
The empty failed fixture was removed; this environmental recovery did not
change production or harness Rust source.

The [read-only verifier](verify.py) checks all 32 report rows: 48 measured planning
samples and 48 explicit unavailable samples, complete aligned commit/publication
vectors, elapsed phase sums through acquisition-order indexes, checked byte
balances and region peaks, zero failed allocations, managed budget release,
and identical corpus/output/semantic/untouched-member evidence across binaries.
The [verification result](verification.json) records exact test counts.
These debug smoke samples demonstrate measurement availability and correctness;
they are not native performance baselines or a latency comparison.

Compute incremental allocator peak using the same sample's
`region_peak_live_bytes - live_bytes_before`. This process-global observer
includes callbacks from other process threads. It does not measure physical
RSS, unwrapped allocators or allocator-internal realloc overlap. Planning
metrics contain only measured iterations in acquisition order; warmup records
are omitted and normal binaries never fabricate zero numeric observations.

After successful verification, the owned build directory is removed and
`cleanup.json` records its absence. `SHA256SUMS` seals the final bundle;
`python3 -B docs/performance/results/change-0538/verify.py` replays retained
evidence without rebuilding, launching benchmark children, or modifying files.

Next: freeze a fresh matched release baseline and acceptance gates, then test
0537's transient attribute ownership candidate with planning allocation and
instruction evidence plus native, commit/publication, eager-read and semantic
guards. OLE2 and OOXML remain first, ODF is deferred, and iWork is excluded.
