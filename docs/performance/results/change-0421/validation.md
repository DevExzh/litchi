# Validation receipt

All Cargo builds, tests and captures were serialized using Rust 1.98.1.
The measured candidate is `dc7cb687da2fc629695510b9610c97a0233f9df1`.
Raw command output is preserved byte-for-byte as gzip; `compression.json`
binds original and stored hashes.

| Check | Outcome and boundary | Retained log |
| --- | --- | --- |
| Counter regression tests on old helper | Expected failure: 7 passed, 5 failed | `tests-before.log.gz` |
| Counter tests after repair | 12 passed | `tests-after.log.gz` |
| Harness library and allocator wrapper suite | 266 library tests passed, 1 ignored; 5 wrapper tests passed | `tests-full.log.gz` |
| Final counter tests after adding report identity | 12 passed | `tests-final-counters.log.gz` |
| Python comparator tests | 92 passed; printed negative-fixture diagnostics are expected | `tests-comparator.log.gz` |
| Allocator all-target Clippy | Command exited successfully with existing warnings and three lint exemptions; this is not a warning-denied clean gate | `clippy.log.gz` |
| Allocator library rustdoc, no dependencies | Passed with `RUSTDOCFLAGS=-D warnings` | `rustdoc.log.gz` |
| Strict claim registry | 9 claims passed after correction notices | `checks/claims.log.gz` |
| Crate boundaries | Passed, 64 packages, 240 dependencies, 14 existing iWork debts | `checks/boundaries.log.gz` |
| CRUD coverage index | Passed, 15 categories, 30 selectors | `checks/crud-coverage.log.gz` |
| Actual-report guard probes | Five passed: valid report, missing revision rejected, peak below live rejected, chronological decrease rejected, order restoration | `checks/report-guards.json` |
| Fresh allocator capture | Four processes, 120 measured observations; current baseline only | `capture.json` |
| Normal report identity | One sample, no warmup, allocator marker omitted and counters unavailable; functional check only | `checks/normal-identity/check.json` |
| Portable replay after cleanup | Three checks passed with no original worktree or binary | `checks/portable-replay.json` |

The full Rust suite ran on the repaired counters before the additive Tool field
was introduced. Final counter tests, comparator tests, Clippy, rustdoc, both
clean release builds and the real normal/allocator captures cover the final
revision. Scoped rustfmt and whitespace checks passed. The Clippy command
allowed `chunks_exact_to_as_chunks`, `clone_on_copy`, and `needless_lifetimes`;
other pre-existing warnings remain visible in the retained output.

The synthetic script preflight is separately labeled in
`checks/script-selftest-provenance.md`. It is not a performance measurement.
The first portable export failed because recorded absolute command paths were
compared with relocated artifact paths. Its failure receipt is retained; the
corrected verifier preserves historical path bindings and validates relocated
artifact hashes. The repeated export and post-cleanup replay pass.

This batch changes benchmark instrumentation and evidence interpretation,
not document behavior. No new fuzz or native Office fixture run was needed
for these counter and report-identity changes. Broader repository lint debt
and the remaining performance program are not declared complete.

Final review confirmed the historical-root path binding. It also identified
that the original portable driver copied current-checkout validators without
pinning them. The bundle now includes those four modules and verifies their
manifest hashes before exporting. A bundle-only replay and a mutated-validator
rejection check cover this correction.
