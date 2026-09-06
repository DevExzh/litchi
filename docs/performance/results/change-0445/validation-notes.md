# Validation

The full standalone release harness passes **373 tests**, zero failed, one
existing ignored. Focused existing OPC topology integration passes **37 tests**.
The previous batch's full OPC suite (458 passed, one ignored) remains sealed;
production sources have not changed in this batch. The newly added end-to-end
test compares plain/observed output and deterministic sink data and verifies
that every plain source metric is unavailable with no values.

Workspace no-iWork and standalone all-feature/all-target release checks pass.
Standalone rustdoc passes with warnings denied, explicit three-file rustfmt
passes, and the boundary audit passes. Strict harness Clippy still has inherited
debt; the before/after diagnostic comparison admits zero new diagnostics. The
baseline lint receipt/log/source table is imported unchanged from sealed 0444;
its source hashes match HEAD before these three harness files changed. Initial
new test needless-borrow diagnostics were corrected; that failed attempt and
source draft remain retained. This is not a clean strict-lint claim.

All six exported source/output ZIP archives are byte-identical to 0444. Python
independently checks payload formulas, CRC-readable archives, content types,
relationships, untouched local/central records (offset changes allowed), order
and comments. Typed no-op/duplicate/missing/stale/short-read/partial-sink/output
and Part-count gates remain Rust checks, not claims of independent Python replay.
Twelve pilots pass across both source modes and allocator modes. All **53**
archive/report mutations reject. No source metrics are invented for plain mode.

The accepted ADR tree and user-owned GOAL.md are unchanged. No production API,
dependency, unsafe code, scheduler or runtime I/O changed. Both selectors remain
opt-in: 438 registered selectors, 36 defaults. The representative coverage index
is unchanged at 15 categories/33 mappings/10 measured/23 correctness-only. This
remains synthetic OPC package topology, not semantic Office owner creation,
native application compatibility, cold/range evidence or scaling.

The initial profile summary retained self rows. Inclusive rows were then added
for production-path attribution. Replay caught unstable ordering of equal-weight
rows because frame sets have process-specific iteration order. Sorting by
descending period and symbol fixes determinism; earlier scripts/derived outputs
and the failed verifier receipt remain in draft history/checks. Raw profiles and
measurements were unchanged. Separate-process summary replay now passes.

The sealed exported bundle verifies before and after cleanup. Five pre-cleanup
and six post-cleanup custody/argv/seal/cleanup mutations reject. Cleanup removed
only `/tmp/litchi-goal-0445-binaries` (916,784,336 regular-file bytes), preserving
both Cargo target directory identities and GOAL.md's pinned digest. The final
bundle verifies without the copied binaries, Git, Cargo, perf or workload execution.
