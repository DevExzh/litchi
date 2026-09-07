# Integration notes

The accepted ADR files match the carried-forward review hashes exactly. This
batch changes only the standalone performance harness and evidence tooling.
Production source and the 0457 bundle remain unchanged.

The first smoke helper requested one sample with no warmup, while the report
oracle intentionally required the formal 30-sample/three-warmup configuration.
That failed receipt and report are retained. The R1 smoke helper uses the formal
counts; it exposed an oracle schema mismatch for normal phase counters. The
Rust serializer omits unavailable counters, while the initial oracle expected
an explicit null. `oracle-initial.py` retains that implementation. The corrected
oracle requires the actual omitted-field schema. Both failures occurred before
binary/protocol binding; no formal capture input was rewritten.

`smoke-r2` passes all four combinations of normal/allocator instrumentation and
lifecycle/phase clocks. `oracle-tests` passes four actual-report controls and
fourteen targeted mutations with specific error-path assertions. Temporary
mutation files are removed. Both binaries and the corrected oracle are then
bound by `binary-binding.json` and the frozen `protocol.json`.

Raw logs and profiler output retain original whitespace and diagnostic text.
They are authenticated evidence, not authored source to reformat. Failed smoke
reports are excluded from every formal timing summary.

The formal matrix, both profiles and both derivations pass without retries.
The complete release harness suite passes 387 tests with one ignored; strict
Clippy, release build, formatting, warning-denied rustdoc and crate boundaries
also pass at the unchanged Rust source epoch. The bundle verifier and finalizer
received a portable path-depth guard before their first run. Source custody
checks require exact file coverage, and protocol custody includes the binary
binding hash. Derived helper changes do not rewrite frozen capture inputs.

The sampled profile has no resolved phase-marker ancestry despite the marker
symbols being present in the bound executable. All 1,656 samples remain
unattributed; no internal-stage or hardware-counter attribution is claimed.
An incidental Python bytecode cache is removed before sealing. The immutable
proof receipts and temporary-artifact inventory document final replay and
cleanup. Shared Cargo targets are retained.
