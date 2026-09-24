# Replay and provenance

Baseline: `ae9bf7bfa11eed50e6560be166eaa18987cda2b5`.
Candidate production change: `crates/litchi-ods/src/model/dde/transaction.rs`.
Candidate test change: `crates/litchi-ods/tests/ods_dde_transactions.rs`.

Both benchmarks used isolated checkouts of the baseline. The candidate overlays
only the two changed ODS files; unrelated shared-workspace OPC/XLSX changes are
excluded. Both retained harness sources are identical. Each source manifest
has 16 entries; the candidate test manifest separately has 2 entries.
Root verified both manifests against baseline Git objects/current scoped source,
compared all nine gated ODS implementation/test files with the candidate checkout,
and checked both executable hashes before cleanup. Root replayed
`candidate/transaction.diff` against the baseline and verified both output hashes.
The candidate also passed 47 isolated DDE tests after timing.

For either side, build its retained harness in a separate baseline checkout,
copying that side's `harness` directory to the identical repository-relative
location. For the candidate, first apply the retained patch:

```sh
git apply --unidiff-zero /absolute/evidence/candidate/transaction.diff
CARGO_TARGET_DIR=/absolute/replay-target cargo build --release --locked --offline \
  --manifest-path docs/report/spec-gap-validation-evidence/ods-dde-catalog-performance/candidate/harness/Cargo.toml
```

Use `baseline/harness/Cargo.toml` for baseline. Commands and run intervals are
retained beside each side's raw CSV. Baseline timing was split into nine
transaction lanes and nine control lanes (`control-commands.txt`); candidate
commands cover all 18 lanes. Both sides use three warmups and 15 iterations.
`perf-stat-commit-many-links.stderr` records three whole-process counter runs
per side; these include parsing and setup outside the commit timer.

Root full gates ran in the shared working tree. The isolated candidate tests
independently verify the changes against committed dependency sources. The
public authoring runner was replayed from the previously committed harness;
its temporary output exactly matched the retained preoptimization XML and
passed the ODF schema before deletion. Owned targets, worktrees, and temporary
XML were removed after final verification. Final evidence retains source,
locks, raw measurements, commands, reviews, and validation receipts only.
