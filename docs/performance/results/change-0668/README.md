# Change 0668 evidence packet

The XLS follow-on for queue rows 11 and 12: measure-only `MulRk`/`MulBlank`
queries and removal of the source-open shared-string locator pointer vector.
The retained cross-query index, snapshot chain hint, and framing fusion remain
deferred under their existing ADR and coverage gates.

## Provenance

| | |
|---|---|
| base commit | `5fa92d7ce` |
| branch | `perf/0668-xls-query-residues` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0668` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-5fa92d7ce` |
| host | AMD EPYC, Linux; shared multi-agent host |
| toolchain | rustc 1.95.0, cargo 1.95.0 |
| test profile | `CARGO_BUILD_JOBS=2 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0` |
| timing | no wall-clock run; no timing claim or claim-registry entry |

The before and after allocation legs use the same `sst_scan_allocations`
integration test and the same checked-in fixtures. The before checkout was
read-only after its run; its temporary generated `Cargo.lock` was removed.

## Contents

| path | what it is |
|---|---|
| [`../../0668-xls-query-residues.md`](../../0668-xls-query-residues.md) | the retained change record, scope, evidence, gates, and limitations |
| [`decision.json`](decision.json) | machine-readable decision and evidence record |
| [`log-sections.md`](log-sections.md) | the four paragraphs for the coordinator's aggregate logs |

## Replaying the evidence

From the after worktree:

```sh
CARGO_BUILD_JOBS=2 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
cargo test -p litchi-xls --test sst_scan_allocations -- --nocapture

CARGO_BUILD_JOBS=2 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
cargo test -p litchi-xls records::sst_measure_tests --lib -- --nocapture
```

Run the same allocation command from the base checkout for the before leg.
The full validation command is:

```sh
CARGO_BUILD_JOBS=2 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
cargo test -p litchi-xls
```

The allocator output is deterministic for the selected fixtures. The source
open's 28-record and 13-record SST groups lose exactly one allocation of 448
and 208 bytes respectively; the pinned corpus index and the full XLS suite
provide the semantic and refusal-order checks. No noisy timing result is
retained for this change.
