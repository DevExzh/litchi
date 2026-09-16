# Evidence packet for change 0651

Record: [`docs/performance/0651-queue-refresh-after-the-second-wave.md`](../../0651-queue-refresh-after-the-second-wave.md).

This is the coordination packet for the second wave (changes 0632 to 0650).
It holds what the coordinator produced: the merge log, the integration-gate
transcript on the merged head, two verification logs, the four log paragraphs,
the decision and cleanup records and, after the rebase, the commit map. Every
measurement of the wave lives in the merged records' own packets,
`results/change-0632/` through `results/change-0650/`, exactly as their agents
committed them.

## Contents

| file | what it is |
| --- | --- |
| `merge-log.txt` | The nineteen cherry-picks in completion order: subject, the commit on the main branch, and the branch commit it was picked from. The one conflict (0648 against 0641, an import list) is annotated on its line. |
| `integration-gate.log` | On the merged head `f0fc20178`: `cargo fmt --all --check`; `cargo clippy --all-targets --locked`, `cargo test --locked` and `cargo doc --no-deps --locked` for the fourteen in-scope crates (`litchi-core`, `litchi-cfb`, `soapberry-zip`, `litchi-opc`, `litchi-ooxml-common`, `litchi-ole-common`, `litchi-xls`, `litchi-doc`, `litchi-ppt`, `litchi-xlsx`, `litchi-docx`, `litchi-pptx`, `litchi-xlsb`, `litchi`); `non_iwork_gate.py facade-format-tests` and `harness-tests` (the two modes 0639 added to CI); `check_perf_claims.py --mode strict`, `check_report_claim_classification.py`, `validate_crud_coverage_index.py` and `non_iwork_gate.py verify`. Result: every step exit 0 — fmt clean, clippy clean, 439 test binaries all passing, doc clean, both CI modes green, 10 claims validated (strict), 167 REPORT rows classified, the coverage index and the non-iWork gate verified. |
| `xls-merged-0641-0648.log` | `cargo test -p litchi-xls --locked` on the head that first combined 0641 and 0648 (`ba9b5d8a1`), run because the two changes met in one file; exit 0, no failure. |
| `odt-polyglot-local.log` | The facade's `docx,odt` feature combination on the pre-rebase head: the three polyglot-detection tests that are stale on this branch and fail with the behaviour the code already has, which the upstream commit `cb2a1a2d4` rewrites and the rebase takes. |
| `log-sections.md` | The four paragraphs the coordinator inserted at the top of `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md` for this change. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the worktree area, the branch list and the session scratchpad when the wave closed, and what was kept. |
| `integration-gate-rebased.log` | The same gate as `integration-gate.log`, on the rebased head `9790cea68` (after the 144 commits were replayed onto the upstream branch and the dead upstream test helper removed): every step exit 0, 442 test binaries all passing. |
| `rebased-targeted-tests.log` | On the rebased head: the two XLSX test files the rebase resolved (`source_backed_cell_values`, `source_backed_row_visibility`, 74 and 17 tests passing) and the three facade polyglot-detection tests the upstream commit rewrote, run under `--features docx,odt` and passing. |
| `rebase-commit-map.txt` | Added by the follow-up commit after the rebase onto the upstream `feat/office-format-completeness`: every pre-rebase commit hash of this branch beside its post-rebase hash and subject, so the hashes the wave's records cite stay resolvable. |

## Provenance

- Base of the wave: `c7326f680` (0630's merge); 0648 branched from `9f28ea621`
  (0636's merge) because it builds on 0636's cursor window.
- Merged head the gate ran on: `f0fc20178` (0650's merge), before this record's
  own commit.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
  The gate was run in the shared working copy with no agent building.
- All hashes in this packet and in the records 0632 to 0650 are pre-rebase;
  `rebase-commit-map.txt` maps them.

## What this packet does not contain

No number in the record is the coordinator's. The integration gate establishes
that the merged combination of the wave's changes builds, lints, documents and
passes the in-scope crates' test suites and the two CI modes at the head named;
it does not re-run any agent's measurement, oracle or differential.
