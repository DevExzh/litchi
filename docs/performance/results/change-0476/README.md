# Change 0476: reusable owned ZIP Deflate state

This experiment follows the exact-stack allocation attribution in 0475.
The owned ZIP writer now retains one raw Deflate backend and one 32 KiB
output buffer between successfully finalized members. Reset occurs only
after the descriptor and central record have been published. An unfinished
or failed entry discards its active compressor.

The implementation preserves flate2's pending-output ordering, sync flush,
accepted-input accounting, compressed limits, short writes and independent
member streams. The borrowed writer supplies the fresh-encoder differential
oracle. The owned API continues to use default compression; the reference
tests of other levels do not introduce an owned compression-level API.

`protocol.json` binds the drivers before candidate construction and capture.
The control reuses both authenticated 0474 executables and restores their
original clean source path. The candidate is built from a committed clean
checkout with the same release, frame-pointer and unwind-table settings.
Source manifests, fixture identities, commands, binary hashes and individual
capture receipts remain in this directory.

The main experiment has 24 lanes: two arms, normal/allocator executables,
three PPTX sizes and two reversed repeats. Each lane retains thirty samples
after three warmups. Four candidate pilots are excluded from formal results.
Four large PPTX process-counter lanes and four ten-row shared-transport guard
lanes use ABBA arm order. The guards cover fresh DOCX, XLSX, ODT, ODS and ODP
streaming creation at tiny and large sizes.

Normal elapsed samples, allocator elapsed samples, operation allocation work,
operation peak heap and process RSS have separate scopes. Process counters
also include materialized preflight and observers. This experiment cannot
establish constant total memory, registered latency improvement, native
producer breadth, logical append, arbitrary repackaging or worker scaling.
The full non-iWork goal remains open.

The initial test build failed on a test-only `Debug` bound. A subsequent
direct-write test exposed a reference sink that accepted a prefix instead
of atomically enforcing the existing compressed limit. Both failed attempts
are retained. Independent review also caught eager sink emission in the
initial production draft; pending-output buffering was corrected before
candidate build and measurement. See `source-review.md` and `validation/`.

## Results

All 36 planned captures pass: 720 main samples, four excluded pilot samples,
120 counter-lane samples and 1,200 shared-format guard samples. The large main
operation requests 31,428,173 bytes instead of 6,809,604,013, a 99.538473%
reduction in every allocator sample. Peak operation heap changes from
8,875,092 to 8,875,252 bytes above entry. All allocator samples have zero failed
allocation calls and zero live-byte exit delta. Every archive hash and length
is preserved.

Main normal mean/p50/p95/p99, incremental-peak and process-RSS checks meet the
frozen thresholds. The shared guards retain one tiny XLSX p99 penalty in the
first pair and two candidate tail-drift flags; the second frozen pair does not
reproduce the XLSX penalty. See `guard-review.md` for the exact samples and
counter scheduling limitations. There is no global regression-free, registered
latency or constant-memory claim.

The [change record](../../changes/0476-zip-deflate-state-reuse.md) contains the
allocation and normal-timing tables. `summary.json` derives every comparison
from the retained reports. `rust-validation.json` binds seventeen required
Rust/repository gates plus subsequent evidence checks and all nonzero attempts.
The whole-workspace format
failure is confined to unchanged, out-of-scope Keynote formatting; scoped ZIP
and harness checks pass.

## Portable verification

From this directory, run `python3 -B verify.py` to verify the sealed bundle
without Cargo, perf, or the original temporary source/binary paths. Run
`python3 -B analyze.py --check` for exact summary replay. The verifier checks
the frozen source/build/capture chain through `custody.py`, strict producer
schemas and arithmetic through `report_checks.py`, and exact seal coverage.

`cleanup.json` records removal of only the owned runtime trees and binary
copies. `portable.json` records fresh-copy replay and resealed negative probes.
Shared Cargo caches and the two user-owned untracked files remain intact.

Final live verification passes both binaries per arm and every source hash.
All eleven Python evidence tests pass. Fresh-copy sealed verification, restored
summary replay and five resealed mutations pass after runtime cleanup.
Initial verifier attempts exposed pilot-count and ledger/helper-binding
mismatches; their records remain alongside the successful final verification.
