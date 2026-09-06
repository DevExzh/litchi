# Validation and scope

Root-owned builds, tests, fuzzing and measurement processes run serially. Rust
1.98.1, four Cargo jobs, no incremental compilation, release profiles and one
test thread are recorded by `check.py`. Source manifests cover all workspace and
standalone Rust/TOML/lock files before and after each command. No Rust edit is
made while an owned CPU job is live. All retained draft receipts remain immutable.

The initial integration passed strict lint, 469 OPC tests, 848 PPTX tests and 381
harness tests. Review then found that releasing writer overhead also released
all retained metadata charges for empty captures. A new regression test was run
against the old source and failed with zero bytes charged instead of 4096. The
corrected implementation pins a fixed overhead charge and charges each new
publication separately. A later refinement sizes this bound for long source Part
names/content types as well. The final test uses an explicit 32 KiB member-name
limit to admit its 16 KiB name; the default 4096-byte read limit correctly refused
that first test draft. A compiler-denied unnecessary qualification of `size_of`
was also corrected. These are disclosed draft failures, not hidden passes.

The final source is rebuilt and tested independently. Historical full-suite
passes are not substituted for the final source receipts. The retained draft
applier is guarded against running on any different source and records the
pre-qualification/pre-limit draft; it is not an installation or rerun command.

The performance protocol freezes two binaries, two providers, two corpora and two
repeats before measurement. Each formal report has 3 warmups and 30 samples, with
opposite build/corpus order in the second repeat. Pilots check execution/oracle
compatibility and are excluded from formal estimates. API time sums open,
planning and publication; corpus construction, gates, diagnostics, output hashing
and final handle drops remain outside that clock. Whole-process perf and RSS
include that untimed work. Bootstrap intervals describe conditional per-process
sample variation, not independent-machine uncertainty. Two process repeats and
all >5% paired/repeat flags remain visible.

The range adapter is a caller-controlled simulation, with 64 KiB maximum returns,
200 microseconds per call and 25 MiB/s separate-sleep transfer pacing. It is not
an ambient HTTP implementation, a measured network, cold-disk evidence or a
parallel/shared-bandwidth scaling result. Source compressed buffers remain pinned
until plan drop. Existing staged decoded copies and aggregate staging budgets
remain unchanged; this is an explicit retention tradeoff to examine alongside
latency, RSS and reserved-memory gauges.

The bounded OPC fuzz target exercises cold/warm capture, retained-handle clones
and independent publication authorization with sixteen retained ZIP/OPC/native
input seeds. The 1,000-run ASan/coverage smoke is not exhaustive fuzz coverage.
Native-input substrate tests do not constitute a new Office-application roundtrip.
No registry/default or representative CRUD/native coverage count is promoted.

Final R3 source passes 471 OPC, 848 PPTX and 381 harness tests (1,700 total),
strict lint, docs, workspace, format and boundary checks; instrumented fuzz passes
1,000 runs. Primary performance retains 480 samples, separate confirmation 120,
eight one-sample pilots and four diagnostic profile reports. All final command
source manifests match the exact candidate build epoch. Both binaries identify
the same parent revision but have distinct source manifests and binary hashes.

Portable replay succeeds before and after cleanup. Seventeen precleanup and
nineteen post-cleanup corruptions are rejected, including changed source copies,
counts, hashes, exact-output gates, confirmation/profile summaries and cleanup
claims. Cleanup removes exactly two benchmark executables (115,094,176 bytes)
and 1,622 inventoried fuzz files (1,004,954,192 bytes). Shared build targets,
unrelated temporaries and user-owned docs/GOAL.md remain untouched. Historical
draft receipts and all primary outliers remain in the sealed bundle.
