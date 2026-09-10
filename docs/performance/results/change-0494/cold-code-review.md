# Independent cold DOCX implementation review

This is a read-only re-review of the frozen
`tools/perf-baseline/src/docx_edit_provider/cold.rs` and the current
`docs/performance/results/change-0494/cold_measure.py` custody driver. I did
not run a build, capture, or Rust test. The source and retained reports are
bound to the frozen cold-source validation manifest
`214d10bba4b2cd40bf99d80f50a508efb98ef1d01091ed776999edfebca3144d`.

The operation ordering is sound. The child performs all source reads needed for
expected values and patch oracles before the final `cold_verified::prepare`,
reserves the logical counter and sink before that probe, opens `FileSource`
inside the timed closure, and finishes allocation accounting before the
post-operation process-I/O snapshot. The timed closure includes source-version
fences, commit diagnostics, raw source/candidate XML identity checks, and
commit/package/document drops. Output hashing, semantic/media verification,
and patch safety checks remain post-clock.

The Rust fixes close the earlier materialization, source-read, and status-order
gaps:

- `prepare_case` fails unless the preflight materialization count is exactly
  one (`cold.rs` around line 425).
- The measured lifecycle requires both `materializations == 1` and equality
  with the preflight count (`cold.rs` around lines 901–912).
- An eligible row requires positive logical source calls, requested bytes, and
  returned bytes, and reports `logical_source_reads_positive` (`cold.rs` around
  lines 552 and 904–912).
- The parent checks the aligned-source identity only for an eligible child;
  ineligible verifier statuses can serialize zero-row evidence without a
  source hash (`cold.rs` around lines 729–746).

The current formal driver closes the corresponding report-validation and
process-custody checks. It recomputes positive source reads from the serialized
counter, requires all materialization counts to equal one, retains cleanup in
the terminal receipt, requires the private run root to be absent, starts each
sample in a new process session, and terminates the full process group on
timeout with a grace period and `SIGKILL` fallback. Each sample launches a
fresh Rust parent, whose measured child has a fresh PID. Build/source/protocol,
helper hashes, argv, environment, and the private `TMPDIR` are bound in the
receipts.

## Retained final evidence

The accepted cold evidence is a passing `cold-formal1` verification with 120
samples and a passing `cold-pilot2` verification with 6 samples. The formal
run contains normal and allocator cells for repeats 1 and 2, 30 samples per
cell; the pilot contains normal and allocator cells with 3 samples per cell.
All 126 retained sample rows are eligible. A read-only audit of the retained
reports found a distinct measured child process for every row, exact-one
materialization, positive logical source reads, exit code zero, no timeout,
unchanged source custody, and an empty private scratch root after cleanup.
The formal and pilot verification receipts are
[`verification/cold-formal1.json`](verification/cold-formal1.json) and
[`verification/cold-pilot2-pilot.json`](verification/cold-pilot2-pilot.json);
the accepted inventory is recorded in
[`accepted-evidence.json`](accepted-evidence.json).

The allocator rows in those reports total 60 formal rows plus 3 pilot rows.
Every one of the 63 rows is marked `measured`, uses the operation-scoped
allocator scope, satisfies live-byte conservation, and satisfies the recorded
peak and reallocation bounds. The current bundle verifier now re-collects the
accepted raw reports and enforces this exact 60/3/63 inventory in
[`verify_bundle.py`](verify_bundle.py:228), with fixture coverage in
[`test_verify_bundle.py`](test_verify_bundle.py:113). This is an independent
audit boundary because the allocator result is checked from raw rows rather
than trusted from the analysis summary.

The final Rust test receipt reports 468 passed, 1 ignored, and 0 failed; the
format, boundary, build, clippy, and rustdoc receipts also exit zero. The
current helper receipt [`validation/final7-helpers.json`](validation/final7-helpers.json)
reports 72 tests passed with exit code zero, and its source snapshot is
unchanged. [`helper-test-custody.json`](helper-test-custody.json) now binds the
current `verify_bundle.py`, `test_verify_bundle.py`, and all other helper
hashes, so the earlier stale-verifier caveat is resolved. The cleanup receipt
[`cleanup.json`](cleanup.json) also passes after removing 4,004,667,392
allocated bytes while retaining the two final executable binaries.

## Decision

The Rust cold path and the formal driver’s primary validation boundaries pass
static re-review, and the retained reports support a descriptive verified-cold
repeat-variance result. No warm-versus-cold timing claim is authorized. The
remaining points below are implementation observations about edge-failure
custody and report-oracle presentation; the final helper and cleanup custody
gates now pass.

## Remaining implementation observations

`cold_measure.py` creates the private run root before source-revision and
source-snapshot checks (`cold_measure.py` around lines 958–975). An exception
before the process `try` block can therefore leave that root without a terminal
receipt or cleanup attempt. Likewise, an exception in `_remove_private` can
stop receipt creation after the child has exited. The normal formal path now
retains and validates cleanup, but a fully fail-closed custody boundary would
put root cleanup in an outer `finally` and make the terminal pass condition
include the cleanup result. This is an edge-failure cleanup issue; it does not
invalidate a normally completed receipt whose private root is absent.

The output verifier is a real post-clock gate: it reopens the produced package,
compares every non-main Part payload with the source, and checks the media
payloads. The `semantic_reopen_verified` and `unchanged_media_preserved`
report fields are nevertheless literal `true` values after that helper
returns (`cold.rs` around lines 434 and 900–943). This is not a current
correctness bypass because the helper can fail the child, but deriving the
report flags from named checks would prevent the report from presenting
tautological-looking oracle fields.

The Rust parent’s `spawn_child` still uses blocking `Command::output()` without
an internal timeout, and its staging-file `Drop` ignores removal errors. The
formal `cold_measure.py` process-group timeout and private-root receipt cover
the retained capture path. A direct invocation remains dependent on its outer
caller for timeout and cleanup custody; the final evidence should state that
boundary explicitly.

The formal and pilot receipts pass against the frozen cold source, and the
final helper receipt and custody record now bind the allocator-audit verifier
bytes. The prior stale-custody finding is closed; the direct-invocation and
literal-report-field observations above remain documented limitations of the
implementation boundary rather than failures of the retained capture path.
