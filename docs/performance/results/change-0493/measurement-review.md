# Change 0493 managed DOCX measurement review

Review date: 2026-09-10.  This is an independent source review of
`tools/perf-baseline/src/docx_managed_read_ahead.rs`.  It does not review or
modify the sealed 0492 pilot.  The review covers the measurement boundaries,
the managed accounting oracles, and the exact versus opt-in comparison that
will feed the production capture.

## Review result

The benchmark is suitable for a bounded exact/control versus managed
forward-start comparison once the capture gate uses the fixed transport
configuration and validates the physical trace oracle.  The source instrument
has no unresolved measurement-boundary blocker.  The results must remain
scoped to this synthetic in-memory provider and this pinned DOCX corpus.

The second prewarm build completed successfully after the benchmark match-arm
fix.  Its source manifest changed during the command because the shared
checkout was still being edited, so that receipt is a build result rather than
a final source-stable production gate.  Run the final build and captures only
after the source manifest stays unchanged for each gated command.

## Comparison and CLI contract

Each sample uses a fresh managed `litchi-docx::source_backed::Package`, the
same explicit finite `ExecutionContext`, the same `ReadLimits`, payload-cache
limits, corpus, source adapter, and executable.  The control passes
`SourceReadPolicy::exact()`, which must report no source-read diagnostics.  The
candidate passes `SourceReadPolicy::forward_start(4096)`, which must report an
enabled 4,096-byte window.  The policy is opt-in; existing exact constructors
and the control arm do not silently enable it.

The capture should invoke the module following this shape, changing only the
policy, transport arm, output path, and source revision as needed:

```text
litchi-perf-baseline docx-managed-read-ahead --policy exact --max-range 65536 --delay-us 0 --samples N --warmup W --source-revision REV40 --output CONTROL.json
litchi-perf-baseline docx-managed-read-ahead --policy forward-start --window-bytes 4096 --max-range 65536 --delay-us 0 --samples N --warmup W --source-revision REV40 --output CANDIDATE.json
```

The delayed transport arm uses `--delay-us 1000`,
`--transfer-bytes-per-second 104857600`, and
`--transfer-delay-policy minimum-service`.  The zero arm uses
`--delay-us 0` and omits the transfer-rate flag.  Keep `--max-range 65536`
fixed for the supplied physical range oracle.  A different range cap is a
different transport experiment and can split physical calls, so its output
must not be compared with the 19-call control or the expected 3-call
candidate.

## Timing and allocation boundaries

The operation clock starts after corpus, source/provider, bounded physical
trace capacity, limits, budget, and execution-context setup.  It includes:

1. managed package open, including construction and memory reservation of the
   production read-ahead window;
2. the cache snapshot immediately after open;
3. document materialization and `extract_text`;
4. source-read and cache diagnostics after text extraction; and
5. explicit document and package drops.

The returned `String` escapes the timed block.  Text hashing and the pinned
text comparison, source-version comparison, physical and transport snapshots,
and post-drop budget checks occur after the clock.  The allocator region wraps
the complete timed lifecycle, while the physical trace vector is preallocated
before both the clock and allocator region.  This keeps observer-vector growth
out of the policy comparison and reports allocator data separately from
latency.

The range adapter and `CountingReadAt` still add instrumentation work to the
timed path.  Both arms use the same stack, so the result measures the policy
comparison under this bounded observer; it is not an uninstrumented provider
latency claim.

## Required row oracles

The capture validator should require all of the following for every warmup and
retained row; the Rust runner also rejects the lifecycle/accounting failures
that it can establish locally:

| Evidence | Control | Candidate |
| --- | --- | --- |
| Text bytes and SHA-256 | 10,000 and pinned digest | identical |
| Source-read diagnostics | `null` / exact policy | enabled, configured 4096 |
| Logical diagnostics requested bytes | not applicable | logical request total, expected 3,966 under the pinned run |
| Diagnostics returned bytes | not applicable | accepted physical fill bytes, expected 5,445 under the pinned run |
| Physical calls/ranges | 19 calls, 3,966 requested/returned bytes, exact pinned ranges | expected 3 calls and 5,445 returned bytes under the pinned run |
| Cache successful loads | 0 immediately after open, 1 after text | same |
| Cache failures and managed flag | 0 and `true` | same |
| InputBytes | cumulative delta equals physical returned bytes | same |
| Memory and Objects | both zero after package/document drop | same |
| Source version | unchanged from before to after | unchanged |

`SourceReadDiagnostics::requested_bytes` is the logical sum of non-empty
forward requests, including cache hits.  Its `returned_bytes` is the accepted
physical bytes from successful fills.  The independent trace below the range
adapter records the provider's requested and returned ranges.  These fields
must remain separate in analysis; the candidate's logical 3,966 bytes and
physical 5,445 bytes are not interchangeable.

The static exact ranges in the report are an oracle, not merely descriptive
metadata.  The capture validator must compare every control row's offset,
requested length, and returned length against that list, and must check the
control call and byte totals.  Candidate call and byte totals should be
checked for the pinned 4 KiB/65,536-byte transport run; if source-freshness or
budget behavior changes those totals, retain the actual diagnostics and mark
the expected-value gate as changed rather than silently treating it as the
old candidate.

## Scope caveats and blockers

The corpus is generated and hash-checked outside the clock.  It is a pinned
synthetic 20-member DOCX with eight media members, not evidence for files
produced by Word, LibreOffice, or another native producer.  The provider is an
in-memory `PptxRangeSource` transport model; fixed delay and transfer pacing
are service-model inputs, not disk or network measurements.

The source-version check proves that the immutable adapter identity did not
change during one sample.  It does not model a live mutable file or network
source race.  The finite execution policy is explicit and shared by both arms
(one worker/task slot, bounded memory/input/output/object/depth/work limits),
but it does not establish worker scaling or concurrent throughput.

Before production evidence is accepted, two gate conditions remain:

1. complete the final source-stable build/capture receipts after all shared
   edits settle; and
2. have the capture validator enforce the exact control range list and the
   diagnostics-versus-physical conservation checks above.

No change to the 0492 artifacts is required or included in this review.
