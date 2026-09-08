# 0470: reuse compaction for empty worksheet web bindings

This experiment uses a bounded success proof accumulated from emitted XML
events to remove the later web-binding traversal for ordinary worksheets.
Extension-bearing or ambiguous XML goes through the unchanged reader at the
existing validation phase. See [source review](source-review.md), the
[frozen protocol](protocol.json), and the [change record](../../changes/0470-xlsx-empty-web-proof.md).

The six-row diagnostic ABBA uses 100 retained samples and five warmups for
one-cell and one-percent ordinary commit/save, each at tiny, medium and
dense-wide shapes. This is below the existing 500-sample minimum for a
registered latency claim. It does not add or authorize a claim. The complete
201-row default guard uses 15 samples and three warmups. Heaptrack uses dense
one-percent commit/save at five samples and one warmup, covering the whole
process: generation, expected output, warmups, verification and teardown as
well as the operation. Instrumented timing and RSS are excluded from normal
comparisons. Rounded heap display values are not exact byte counts.

The control is the authenticated 0469 candidate, revision `7cc58fc1b`, binary
SHA-256 `dbd3f7014b4b3d5203e7c37efe3f739565b4d34a33913e4317720f4a105729da`.
The candidate is revision `f0ab67b55`. Both use Rust 1.98.1, release debug level
1, forced frame pointers and unwind tables, one worker and CPU 2. They build
from the same absolute path, `/tmp/litchi-goal-0468/profile-tree`; before each
capture that clean checkout is switched to the bound role revision. Source
inventories bind 6,992 control files and 6,993 candidate files, plus the same
two compile-time fixtures. The difference is exactly the four reviewed XLSX
implementation/test files. No harness or corpus code changes.

`build.py`, `capture.py`, and `export_heap.py` retain exact commands, environment,
source and binary identities, timestamps and artifact hashes. Heavy work runs
serially under `/tmp/litchi-goal-0470/cpu.lock`. Reproduction requires fresh role
builds from the recorded revisions and source inventories with the recorded
flags and shared absolute build path. Run A1, B1, B2, A2, then A-full/B-full
and A-heap/B-heap. Authenticate binaries and preserve all output in a fresh
bundle; do not overwrite these historical captures.

`analyze.py` uses the canonical repository ABBA and regression comparators.
Only full-guard comparison copies project empty optional top-level source
vectors as absent; both roles must have exactly the same projected paths.
Raw reports, measured vectors, selectors and the frozen policy are unchanged.
Every individual timing/RSS flag requires review; a geometric mean does not
replace it.

`verify.py --live` verifies live source/binary identity and all retained
receipts before cleanup. Flagless `verify.py` verifies the sealed bundle after
the temporary builds and binaries are removed. Portable replay needs this
bundle, `tools/perf_abba_summary.py`, `tools/perf_compare.py`, and the referenced
0469 candidate binding, source binding, build receipt and its stdout/stderr
artifacts. It does not need Rust, a Git checkout, or the temporary binaries.

Correctness receipts cover all XLSX tests and scoped formatting, workspace
all-feature checking, warning-denied XLSX Clippy and rustdoc, and crate
boundaries. The first test compilation rejected redundant qualified paths in
test helpers; its logs remain under `validation/attempt-1-test-qualifications`.
No native Office run, fuzz campaign, physical-cold, remote/range, worker scaling,
or extension-bearing worksheet performance result is claimed. Those limits
and the broader non-iWork goal remain open.

The original sequence records 19.753% fewer whole-process allocation calls,
with the same rounded peak heap display and lower diagnostic ordinary XLSX
latency. The initial normal RSS increase is not consistently reproduced by
`rss-protocol.json`'s second, identical four-process sequence. The 25-row
`guard-protocol.json` follow-up passes strict identity checks but retains PPT
and small CFB-read latency penalties. See the change record for the decision,
all original flags, both RSS sequences and the remaining limits.

`review-summary.json` is the immutable full-guard snapshot hashed by the targeted
guard protocol. The verifier later gained RSS/guard integration checks, changing
its file hash. `final-review-summary.json` records both helper bindings and
recomputes the current result. Verification permits only that helper-binding
metadata to differ from the historical snapshot; every report, numeric result,
flag, oracle, policy and protocol field must match. The original helper version
is not retained separately. Current code independently replays all the results.

The temporary candidate checkout was truncated during the guard B1 preflight.
That preflight refused to start a benchmark. `validation/guard-preflight-repair`
retains the observed state, source diff, restoration and full reauthentication.
The successful A1 was kept; only the pending B1/B2/A2 ran after repair.

All 14 focused evidence tests, the six correctness gates, live source/binary
verification and the 10-claim strict registry audit pass. Portable replay also
requires the local RSS/guard scripts and their receipts, all included here;
there are no extra external tools beyond the dependencies listed above.

Post-cleanup verification passes in a fresh portable copy containing only the
documented evidence and Python dependencies. An initial replay found that the
RSS/guard reviewers compared recorded destinations against the replay location;
they now validate the fixed historical destinations. All 16 focused tests pass,
including relocated replay and destination-tamper rejection. The replay receipt
is retained in `validation/portable.json`; the copy and owned builds are removed.
