# Final independent review — change 0496

Current source, helper, and formal-data review is **ready for final seal**.

The frozen Rust source remains
`e3732d932356018dff0bd7a17b8eb38e60e177ee7d83992e9a04dbe82fe26c94`. Its
diagnostic field is opt-in, uses wall-clock `Instant` intervals, checks phase
and lifecycle conservation, reports no fabricated instrumentation overhead,
and keeps one full-lifecycle allocator region with no nested phase allocator
regions. The default serialized report omits the optional field. The phase
instrumentation did not remove the existing 0495 output, semantic/media,
source-version, materialization, commit, patch, range/sink, budget, or
allocator oracles.

The driver now validates all full-lifecycle allocator fields for allocator
children, derives and reports `allocation_peak_increment_bytes`, emits
percentiles and within-child CIs, and includes allocator p50/p95/p99 cells in
paired unmanaged comparisons and repeat visibility. Normal children retain an
explicit unavailable marker rather than synthetic zero counters. The claims
scope names phase latency, RSS, and full-lifecycle allocation evidence. The
current focused suite has 14 passing tests, including allocator vector,
derived-peak, paired, and repeat-visibility tests. An independent synthetic
allocator report also produced all expected vectors and percentiles.

The release test, Clippy, and rustdoc gates in `builds/after-final2-*.json`
all exit zero with `source_unchanged: true`. The four normal/allocator binary
records bind the expected before/after revisions and identical shared harness
file hashes. Phase scope strings in the Rust source and Python driver match;
the driver projects only `phase_diagnostics` away before invoking the sealed
0495 validator, retains raw reports, and validates exact terminal artifacts,
source/binary/protocol bindings, chronological order, and private cleanup.

The formal bundle independently passes the data checks: 32 terminal receipts
and all retained artifact hashes match, children follow the frozen order with
no overlap, all 960 rows conserve their seven phase intervals plus residual,
all retained 0495 output/semantic/source/cache/commit/patch oracles are true,
and the 480 allocator rows satisfy the allocator counter, live-byte, and
peak-live conservation equations. The 480 normal rows correctly carry no
allocator sample. Every raw report has the one expected top-level shape and
the one expected output identity.

Recomputing the analysis from raw reports reproduces every child latency and
phase percentile, allocator percentile, paired cell, repeat cell, and adverse
flag. There are 294 paired metric cells and 74 flagged cells: 60 phase
percentile cells, 9 reallocation-count cells, 4 RSS cells, and 1 latency cell.
These are overlapping metric triggers, not 74 independent lifecycle
regressions. The paired allocator p50 cells show reallocation calls changing
from 185 to 228 while total allocation calls change from 22,859 to 9,696 and
allocated bytes from 5,721,334 to 1,623,696; these remain descriptive
full-lifecycle allocator observations. Sixteen repeat-variance records expose
the two-repeat matrix cells without treating repeat variation as host-level
uncertainty.

Publication is the largest named phase at p50, p95, and p99 for all 32
children, with publication-to-lifecycle p50 ratios ranging from about 61% to
87%. The publication interval includes the preallocated `Vec` sink copy and
write counters. Output hashing and semantic/media verification run after the
outer lifecycle clock; this evidence does not attribute those checks to the
publication phase or claim CPU time.

The earlier helper-test receipt is retained as historical development evidence;
the current helper custody is now closed by `helper-tests-final2.json`, which
binds the current hashes `4fb29fd12079330ac6557cebc8632d8ef3a35ed303eb69f4e398907b3b68554c`
and `583ae7772c716c29e9874d8d4de1e33d2c54c52c1e650170eed406cf50b91d6c` and
records 14 passing tests.

The frozen protocol binds the current driver, focused tests, builds, machine,
provenance, and helper receipt. `verification/formal1.json` reports 32
children and 960 samples with status `pass`. `cleanup.json` reports pass,
reverse patch checks, unchanged protected primary files, preserved measured
executables, and removal of the standalone source/build/scratch paths. The
remaining final action is the root-owned evidence seal.

No further Rust or helper semantic blocker is visible after the allocator
analysis fix. The evidence remains descriptive for owned, warm-file, and
short-range providers; it does not establish cold filesystem, native producer,
network, borrowed-lifetime, concurrency, optimization, or historical-flag
causal claims.
