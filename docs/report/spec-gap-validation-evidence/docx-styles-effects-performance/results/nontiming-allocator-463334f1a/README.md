# Non-timing allocator correctness check

This is a serial unit-test check from frozen profile source checkout
`463334f1a32de687ffb2b3357e55816dcfd3102e` (operator session `17932`). It is
correctness evidence for the existing profile harness, not a latency or memory
benchmark. The production/profile pin is d000 and the retained current smoke
prerequisite is c4ac.

The exact command was:

```text
cargo test --locked --offline --manifest-path docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/profile-harness/Cargo.toml -- --test-threads=1
```

The test used a fresh external Cargo target with `CARGO_INCREMENTAL=0`,
`TMPDIR=/var/tmp`, and Rust flags unset. Both serial unit tests passed:
`prepared_setup_guard_survives_operation_until_explicit_drop` and
`allocator_subphase_peaks_reset_without_erasing_outer_peak`. The latter checks
separate phase peaks, live-byte continuity, outer peak preservation, and child
traffic containment for allocation/reallocation/deallocation counters.

The before process census recorded an unrelated ODS test process (`cargo test
-p litchi-ods --lib --offline data_style::source::tests`, PID 3338814). The
check therefore ran under observed shared workload and supports no quiet-host,
latency, RSS, allocation-volume, or speedup claim. The after census was empty.
The fresh Cargo target was removed after the successful test; its logs and
status are retained here. No profile sample or timed runner was started.
