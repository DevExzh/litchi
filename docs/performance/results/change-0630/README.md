# Evidence: change 0630, the first wave on the 0587 queue, closed

Change record: [`0630-queue-refresh-after-the-first-wave.md`](../../0630-queue-refresh-after-the-first-wave.md).

Disposition: retained, coordination record. `performance_claim: none`. This
packet holds the coordinator's own artifacts for the wave; every measurement
the record quotes lives in the packet of the change that took it
(`results/change-0588/` through `results/change-0631/`), and nothing here
re-measures or aggregates them.

## Contents

| Path | What it is |
| --- | --- |
| `merge-log.txt` | The wave's commits on the main branch, oldest first, from the 0587 base to the final head, each with the branch commit it was cherry-picked from (the `-x` trailer). |
| `integration-gate.log` | `cargo fmt --all --check`, `cargo clippy --all-targets --locked` and `cargo test --locked` for the fourteen in-scope crates on the merged head at `1946b964e`, mid-wave, run in the shared working copy while agents were building elsewhere. |
| `integration-gate-final.log` | The same gates on the final head `8b07b45dc`: fmt clean, clippy clean, 436 test binaries, all passing, exit 0. |
| `log-sections.md` | The four paragraphs the coordinator inserted at the top of `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md` for this change. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree area when the wave closed, and what was kept. |

## Provenance

Base: `08d968f8e6a0c0d1b1a04b0f0a4b17c1c4fd2c0e` (change 0587). Final head:
`8b07b45dc`. Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws;
rustc 1.95.0. Each change was implemented by one agent in its own git worktree
and branch under a disk-backed directory outside the repository, with a shared
read-only detached checkout of the then-current base for its before leg; the
coordinator cherry-picked each branch in completion order, inserted its log
sections from its packet, appended its own attribution trailer where the
agent's commit carried only its model's, and amended the commit. The agents'
branches (`perf/0588-…` through `perf/0631-…`) are retained; the `-x` trailers
in `merge-log.txt` bind each main-branch commit to its branch commit.

## What this packet does not establish

No number in the record is the coordinator's. The two integration-gate runs
establish that the merged combination of the wave's changes builds, lints and
passes the in-scope crates' test suites at the two heads named; they do not
re-run any agent's measurement, oracle or differential, and they were run with
default features, which is why the facade test change 0629 fixed was invisible
to them. The gate list gaps 0619 and 0629 found (the harness's own suite, and a
feature-bearing `cargo test -p litchi`) remain open.
