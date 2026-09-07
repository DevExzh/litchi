# Change 0456: direct shared ZIP payload preparation

The ZIP preservation writer retains verified Store/Deflate payload tokens and
shared Store Arcs, allocating only framing and central metadata. Media
publication allocates 4,243,083 bytes instead of 21,057,867 (-79.85%) and peaks
593,892 bytes above operation entry instead of 17,406,388 (-96.59%). Output,
read/write work and managed budgets remain identical. Plain metadata allocation
rises 3,360 bytes. This is not a bounded complete document lifecycle.

Ordinary bytes/media API medians regress 15.411%/16.347% with higher page faults.
All 15 timing flags remain visible. Separate fixed-policy diagnostics improve
API medians 2.626–10.266%; they show allocator-sensitive results without replacing
the original observations. No unconditional latency improvement is claimed.
Read the [change report](../../changes/0456-zip-shared-payload-framing.md) and
[review decision](review.md) before interpreting the numbers.

## Evidence

- [Frozen protocol](formal/protocol.json) and [hypothesis](formal/hypothesis.md):
  24 ABBA lanes, CPU 2, one worker, 30 samples after 3 warmups; 480 ordinary and
  240 separate allocator observations. [Machine identity](machine.json).
- [Measurements](formal/measurements.md), [raw derived data](formal/measurements.json)
  and [operation heap counters](formal/allocation-summary.json). Each lane retains
  raw reports, source/build identities, timing/resource logs and output oracles.
- [Separate hardware counters](formal/profile-summary.json), local instruction
  profiles, and [fixed allocator-policy diagnostics](formal/diagnostic-summary.md)
  with 240 additional observations. Hardware counters include whole-process
  setup/oracles and are not attributed to the timed API.
- [Release commands](formal/run-checks.py): 2,654 tests pass, 7 ignored, plus
  strict lint, documentation, formatting, minimal workspace and boundaries.
  [Fuzz commands](formal/run-fuzz.py) pass 1,000 ASAN iterations; exact initial
  seeds, target, manifest and post-run lockfile/executable identity are retained.
- [Native proof](formal/native-proof.json): pinned LibreOffice QA self-pair output
  matches the previous exact 42,948-byte archive through both providers. Six
  [distinct native pair probes](pair-probe.json) all refuse incompatible shared
  graphs. These add no native-application roundtrip or distinct-pair success.
- [Validation notes](validation-notes.md) preserve the initial compile correction
  and the native expected-SHA transcription failure with a fresh corrected retry.
  No golden output, input fixture or failed measurement was replaced.
- [Remaining work](next-work.md): broader semantic/native/I/O/scaling coverage
  and bounded existing-document append remain open; no taxonomy row is promoted.

## Replay and cleanup

From the repository root:

```sh
python3 -B docs/performance/results/change-0456/formal/verify.py --portable
```

The verifier pins [final bindings](formal/final-bindings.json), checks exact
source/gate/capture identities, rederives JSON and Markdown, replays oracle
negatives and validates every expected transient against the cleanup inventory.
Seven deliberately corrupted reports/counters are rejected. It uses retained
evidence after the executable/native-output/raw-profile files are removed;
it does not rerun the workload or claim a native application reopened the output.

[Actual precleanup verification](precleanup.json) passed before removing only
`/tmp/litchi-goal-0456`: [373 files / 522,511,229 bytes](cleanup.json), individually
hashed in [the inventory](temporary-artifacts.json). Shared Cargo target caches
remain. [Separate-copy replay](portable-verification.json) passes with the owned
task absent. `SHA256SUMS` seals the complete retained bundle; check it from this
bundle directory with `sha256sum -c SHA256SUMS`.

The full non-iWork goal remains open. User-owned `docs/GOAL.md` is untouched.
