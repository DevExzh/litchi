# 0362j — Accept the PPT first-commit and replay repair

- Date: 2026-10-21.
- Status: accepted for integration; the final workspace gate remains required.
- Measured source: `c0f346607bb28559ae21fa780c835dc70e87fb0b`, following
  the returned [0362f](0362f-ppt-replay-performance-review.md) and
  [0362g](0362g-ppt-replay-buffer-review.md) candidates.

Replay now consumes its authenticated source buffer instead of cloning it
before rebuilding the target. It derives the parent digest recipe before
consuming the buffer, admits the target length, then resizes and writes the
changed ranges in place. Failure affects only the local buffer; the immutable
receiver remains unchanged. The same-length regression test also verifies
that reconstruction preserves the buffer's allocation.

This integrates the reviewed first-commit implementation and both replay
repairs together: keyed 128-bit block digests, compact changed-block spans,
reuse of the admitted presentation and CFB index, bounded signature discovery,
and the authenticated unsigned-publication memo. Recorded content and length
continue to admit independent-copy replay. Source freshness, signed-source
refusal, exact signed no-op behavior, recipient options and output-length
admission are preserved. Growth, shrinkage, empty targets and partial block
tails remain covered. [Source review](results/0362j/source-review.json)
records the checks; 0362f documents the closed unsigned-publication proof.

The isolated candidate passes **2,110 distinct PPT/CFB tests**, with three
existing ignored tests, and both all-target/all-feature Clippy commands.
Repeated library and focused invocations are not double-counted. All eight
integrated source files match the measured candidate byte-for-byte.
[Candidate artifacts](results/0362j/candidate-artifacts.json) retain the
commands, log hashes, source comparison, normalized lock and protected binary.
These are targeted checks; they do not replace the final integrated gate.

The File classifier also handles the now non-exhaustive PPT slide-order
error. The fallback classifies future unknown variants as `internal_error`;
the existing explicit classifications are unchanged. This consumer migration
was caught by the subsequent workspace check and folded into this unpublished
commit. The all-target/all-feature `litchi-ops` check and all 957 File library
tests pass ([consumer receipts](results/0362j/consumer-validation.json)).
No measured PPT source changed during this correction.

## Native qualification

Five complete ABBA blocks run against each protected reference using the
unchanged [0362g runner](results/0362g/ppt_phase_abba.py). Every first-commit,
steady-commit, publication-loop, forward-replay and inverse-replay block is
faster against both references. Ratios below are candidate/reference medians.

| Metric | Original reference | Matched SipHash reference |
|---|---:|---:|
| First commit | 0.955050 | 0.947592 |
| Steady commit | 0.881166 | 0.909412 |
| Commit plus publication loop | 0.868659 | 0.909847 |
| Forward replay | 0.737669 | 0.927728 |
| Inverse replay | 0.915338 | 0.766309 |
| Process peak RSS | 0.155156 | 1.001682 |
| Logical retained bytes per patch | 0.156494 | 1.000000 |

Original-reference inverse ratios are `0.915338, 0.944477, 0.820246,
0.929131, 0.757960`. Matched-Sip inverse ratios are `0.839294, 0.766309,
0.567397, 0.845127, 0.633917`. All observations remain in the
[measurement receipt](results/0362j/measurements.json), including individual
legs with higher times. Neither earlier candidate is pooled into this result.

The original comparison uses `N=7400, FIRST=3500, REPLAY=42800`; matched Sip
uses `7400, 3300, 42600`. Calibration targets 500 ms with headroom. Minimum
measured phase spans are 600.60 ms and 578.39 ms respectively, exceeding the
150 ms floor. The steady qualifier uses the rigorous median lower bound,
not the wider publication-loop timer. Every inverse output is checked
byte-for-byte. The 728 original-reference and 696 matched-Sip host snapshots
contain no competing build, profiler or recognized native test. Both attempts
completed without an overlap or child failure.

The first-commit regression in 0357b and the replay regressions returned in
0362f/g are closed by this candidate. This decision is limited to this
workload; it makes no new claim about other formats or File routes.

## Memory attribution

The matched-Sip process RSS observation is 0.1682% higher and is retained,
not omitted. A separate heaptrack comparison at the **same full workload**
reports 51,498,571 allocation calls and 117.96 MB peak heap for matched Sip,
versus 48,728,071 calls and 117.90 MB for this candidate. Logical retained
bytes remain 15,384 per patch, including the 24-byte span slice. The original
retains 98,304 bytes per patch.

The smaller `100/10/100` attribution workload likewise reduces calls from
350,101 to 322,971 and reported peak heap from 2.27 MB to 2.20 MB. The source
repair removes a full-buffer allocation and copy rather than adding retained
state. These allocation and live-heap results support memory acceptance;
process RSS also includes resident allocator pages and other mappings and is
reported separately. Rounded heaptrack values do not establish exact byte
equality. Sampled Massif maxima are not substituted for the peak.

[Heap receipts](results/0362j/heap.json) include raw captures, commands,
hashes and process snapshots. Concurrent compilation is recorded for the
full-workload captures; no profiler latency or RSS is used as native evidence.
The common 544-byte process-exit allocation is unchanged.

## Reproduction and remaining work

The protected candidate binary is 1,763,496 bytes with SHA-256
`e7c6db29e3dd65bf0178eeaa5e6b52fa57d2e7bd9871aacf70302dd17867b4aa`.
Build with `cargo build --offline --locked --release -p litchi-ppt --example
scoped_patch_probe`, using the recorded normalized registry lock and one
capped Cargo job. Registry identities and release-profile parity with both
references are retained in [0362f dependency parity](results/0362f/dependency-parity.json).
Only the build manifest's stale `cargo_lock.origin` path was corrected;
the original metadata file and unchanged lock/binary hashes are retained.

Regenerate the reduction with `python3
docs/unified-ops/results/0362j/reduce.py --coord ../.coord --out OUT`.
The receipt records exact native commands; choose a new output directory for
each attempt. Its artifact inventories identify every external raw file by
name, size and SHA-256. No compiled binary or large capture is committed.

The external source manifest `ppt-replay-c0-source-digests.json` is
1,363,016 bytes, SHA-256
`fbd4da237806a2807efd2c4d58fd9bc700313320cc253f726849ab262ec43955`.
Regenerate it with `python3 docs/unified-ops/results/0362f/source_manifest.py
../.coord/ppt-replay-c0-source-20261021 OUTPUT`.

W4/W4b/W5/W8 integration, remaining raw/shared API work, inventory/glossary
closure and the full final workspace gate remain required. Item 22 remains
the only deferral under its existing redaction rules.
