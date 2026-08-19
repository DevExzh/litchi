# Change 0193: ODS source-backed edit protection cache

Date: 2026-08-18

## Decision

Compute the ODS source-backed edit protection facts
(`structure_protected`, protected worksheet names) at most once per
`SourceBackedSpreadsheet` owner instead of reparsing the protection domain of
`content.xml` and `styles.xml` on every `edit_cells()` call.

The facts are a pure function of the retained, immutable `content_xml` and
`styles_xml` projections, which are bound to the owner's captured
`SourceVersion`. The cache is a private `OnceLock` on the owner, mirroring the
existing `cell_locator` laziness pattern. Only successful parses are cached;
a failed parse re-runs the same validation and returns the same error on the
next call, so no error result is cached and no refusal is weakened. The owner
stays `Send + Sync`; a concurrent first use may compute the parse twice, with
one result published.

This changes no public API, no source abstraction, no cache/eviction policy,
no writer or publication path, and no validation boundary. Single-transaction
lifecycles (open, one edit, commit, publish) still pay the one protection
parse on their first `edit_cells()`.

## Mechanism and invariants

Before this change, `SourceCellSnapshot::edit` called
`protection::source_edit_protection(content_xml, styles_xml)` directly. That
helper runs a complete protection-domain parse plus snapshot validation of
both XML projections. The snapshot borrows the owner, whose projections are
immutable, so repeated transactions on one owner repeated identical work.

`SourceBackedSpreadsheet::edit_protection` now serves the facts through a
`OnceLock<(bool, Vec<String>)>`: a cached hit returns the stored value, a miss
runs the unchanged fallible parse and publishes it with `get_or_init`.
`SourceCellSnapshot::edit` clones the cached tuple into the transaction (an
empty `Vec` clone allocates nothing) and keeps both surrounding
`check_source` fences. The protection codec, snapshot validation, refusal
rules, and error texts are untouched.

A focused test proves the cache is empty at open, populated after the first
edit, reused by pointer identity on a second edit, performs zero source reads
in either case, and preserves the repeated-row refusal on both the populating
and the cached path. The existing protected-spreadsheet and exact no-op
integration tests exercise the cached path on a second transaction of the
same owner.

## New harness selector

`ods_source_backed_repeated_edit` (opt-in; the default 36 cases / 198 records
are unchanged; the selectable matrix is now 341 names). Over the fixed
16,790,689-byte two-sheet media-rich ODS corpus it prepares one
`SourceBackedSpreadsheet` owner outside the timer, then times four sequential
transactions. Each transaction runs `edit_cells()`, one `set_cell`,
`commit()`, and publication to a fresh bounded `WindowedHashingSink`, editing
a distinct cell (flat indices 0, 683, 1365, 2047). Per-sample `total_ns`,
`stage_ns`, `commit_ns`, and `publication_ns` vectors sum the four
transactions per phase.

Untimed gates: per-variant eager-oracle byte-exact output, source-backed
byte-exact cross-check, semantic reopen of the edited and an untouched cell,
`Pictures/*` byte identity, bounded sink invariants, source immutability, and
one instrumented replay whose logical `ReadAt` counters (20 preparation reads
for 7,785 bytes; 2,388 replay reads for 67,172,960 bytes covering the four
16.8 MB raw-copy publications) are identical between the control and
candidate implementations.

## Matched release timing

Two frozen release binaries differ only in the litchi-ods edit-protection
path; both contain the identical new selector. Control SHA-256
`07b060794de42146b45197e4c54e8bf9282d27ff1ca7abd01538d64ca27f899a`,
candidate SHA-256
`c1ad10735d02d84528da0ca962c5019d5acf6a6276f85d577debd80a09c576ab`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate, A2
control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its ceiling.

Four-transaction totals:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 9.31% | 10.57% | 1.14% | -0.26% | accept |
| mean | 9.55% | 10.42% | 0.85% | -0.11% | accept |
| p95 | 10.68% | 9.76% | -0.93% | 0.08% | accept |
| p99 | 9.90% | 9.84% | -0.59% | -0.53% | accept |

Stage phase (edit begin plus one staged cell, four transactions):

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 71.59% | 71.18% | -0.57% | 0.88% | accept |
| mean | 71.61% | 71.08% | -0.78% | 1.08% | accept |
| p95 | 71.60% | 70.36% | -2.15% | 2.13% | accept |
| p99 | 68.98% | 67.87% | -3.45% | -0.00% | accept |

Stage p50 falls from 6.93 ms/6.89 ms to 1.97 ms/1.99 ms: three of the four
protection parses are removed per sample. Commit-phase and publication-phase
statistics are rejected as neutral (paired directions disagree inside noise);
no regression trigger fired — every rejected statistic's paired directions
straddle zero within 3.2%.

The accepted result is scoped to this corpus, host, build, and selector:
repeated source-backed ODS edit transactions on one retained owner. The
single-transaction `ods_source_backed_one_edit_save` and
`ods_source_backed_one_percent_edit_save` lifecycles are unchanged by design
(their one `edit_cells()` still pays the first parse). No allocation, RSS,
physical-I/O, cold-cache, scaling, producer, broad ODF, or iWork claim is
made.

## Verification

```text
cargo test --locked -p litchi-ods --all-targets
RUSTFLAGS='-D warnings -D deprecated' cargo clippy --locked \
  -p litchi-ods --lib --test source_cell_transactions -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-ods --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
cargo test --release --manifest-path tools/perf-baseline/Cargo.toml
```

The litchi-ods suite passes 337/337 including the new laziness/sharing test.
Scoped strict Clippy, rustdoc, formatting, and crate-boundary checks pass.
Unrelated pre-existing strict-Clippy failures in untouched litchi-ods test
files (`facade_round_trip.rs`, `tracked_changes*.rs`) reproduce identically
without this change and are outside its scope. The harness release suite
passes 139/139 including the new selector test and the updated 341-case
enumeration.

Artifacts:

- [machine-readable summary](../results/ods-repeated-edit-0193-summary.json)
- [artifact manifest](../results/ods-repeated-edit-0193-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in the manifest
