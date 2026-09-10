# Rejected XLSX scanner experiment

This publication preserves the rejected V1 scanner experiment at source base
`cf4f383c4c203381880d4c2e8445d2616a53d283`. It contains no production change.
The four normal A1/B1/B2/A2 captures and paired allocator captures remain byte-exact.

| Operation | Normal p50 change | Allocation-call change | Allocated-byte change |
| --- | ---: | ---: | ---: |
| One-percent commit | +56.251093% | +124.553517% | +8.851307% |
| One-percent commit/save | +41.113269% | +119.407861% | +8.469847% |

Positive values mean regressions. Normal figures compare medians of the two
per-process p50 values per side; allocator figures compare medians of 30 raw
operation counters. Instrumented elapsed time is not latency evidence. Control
tail drift prevents an apparent tail improvement from supporting acceptance.
This result does not describe later scanner revisions or the separately accepted
[plain-cell tag optimization](../xlsx-plain-cell-tag-20260910/README.md).

The archive contains all 41 regular evidence files, including raw vectors,
source snapshots and patch, lockfile, original grant/release and rejection
receipts, the executable verifier, and the later V4 pre-measurement boundary.
It retains empty process logs. The archive does not contain release binaries
or the generated corpus ZIP; hashes and reconstruction inputs are retained.

Verify and extract into a new directory:

```sh
sha256sum -c SHA256SUMS
mkdir extracted
tar --zstd -xf evidence.tar.zst -C extracted
PYTHONDONTWRITEBYTECODE=1 python3 -B extracted/rejected-scanner-v1/verify.py
```

The verifier checks file hashes, source/capture bindings, normal p50 values
against raw samples, and allocation medians against raw counter vectors.
`root-verification.json` records a successful root invocation. No new workload
was run for publication. The full performance program remains incomplete.
