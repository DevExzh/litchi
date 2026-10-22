# IWA raw-operation evidence

`snappy-probe.rs` is the shared probe source. Its `before` feature changes
only calls whose argument lists differ at baseline `e63a1fe3534bcaaae2ce367339f7f8275877baef`.
`snappy-probe.py` copies that exact baseline and the current candidate's eight
substrate crates into separate external projects, uses the same dependency
lock, and builds separate native and allocation-instrumented executables.
It records the compiler, source digests, lock hashes and executable hashes.

From the candidate checkout, with the recorded Cargo lock and capped build
environment, use a fresh directory outside the checkout:

```sh
python3 docs/unified-ops/results/0361d/snappy-probe.py build /tmp/iwa-raw-probes
python3 docs/unified-ops/results/0361d/snappy-probe.py run /tmp/iwa-raw-probes --results-dir /tmp/iwa-raw-native --blocks 5 --target-ms 500
```

The run requires a coordinator-approved quiet host. Children are pinned to
one allowed CPU. Each endpoint is independently calibrated; paired ABBA
ratios divide the measured span by its actual iteration count. Every
calibration, timing and allocation result is appended before parsing it.
Measured legs below 150 ms are ineligible. Output bytes/checksums and bounds
must agree, and the probe checks Snappy round trips and exact sink output.
The final runner checks for conflicting build/profiling/probe processes both
before and after each child, preserving a detected overlap before stopping.
Framing uses 512 objects with two 1-KiB messages each; the other workloads
use 1 KiB or 4 MiB, with repeating and deterministic pseudorandom data.

Allocation instrumentation runs separately from native timing. `peak` is
the additional live heap above the retained fixture and warm result;
`retained` is the remaining additional live heap after the operation;
`allocated` counts requested bytes including reallocations. The fixture,
expected output and sink capacity remain live on both sides. These are
operation allocation measurements, not whole-process RSS.

`native-v1.json`, `native-v2.json`, `native-v3.json` and the longer
`native-v3-followup.json` retain the rejected candidates, all paired ratios
and references to complete external observations. `native-v4.json` and
`native-v4-copy-followup.json` are the accepted repair and its longer
large-copy qualification. Each version uses distinct external source/build
directories; follow-ups reuse their version's protected binaries. The
comparison and interpretation are in [0361e](../../0361e-iwork-raw-performance-repair.md).
No executable or raw build directory is committed.
