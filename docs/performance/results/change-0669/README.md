# Evidence packet — change 0669

Change record: [`0669-xlsb-edit-residues.md`](../../0669-xlsb-edit-residues.md).

Disposition: retained. `performance_claim: none`. This packet records the
remaining XLSB publication and resource-relationship work from queue row 13 of
change 0651. No benchmark was run here; the shared `apply_workbook_structure`
selector belongs to 0674 and is described as a handoff in the change record.

| path | contents |
| --- | --- |
| `decision.json` | Machine-readable disposition, scope, evidence, and known measurement gap. |
| `log-sections.md` | Four standalone paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md`, and `ADR_COMPLIANCE.md`. |
| `gates.txt` | Formatter, check, clippy, library-test, and integration-test results. |

## Evidence boundary

The implementation evidence is the source diff and the tests named in
`gates.txt`. The two resource tests exercise an existing part with a missing
workbook relationship for styles and shared strings. The facade tests exercise
successful sparkline and cell-watch publication, compare the resulting derived
projection with a fresh parse, and verify stale commits remain atomic. The
full `litchi-xlsb` library and integration suites pass.

The packet deliberately contains no timing, allocation, RSS, callgrind, or
release-binary result. A speed claim would require the selector owned by 0674
and a paired before/after run with its own floor.

## Provenance

| field | value |
| --- | --- |
| base commit | `5fa92d7ce` |
| branch | `perf/0669-xlsb-edit-residues` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0669` |
| production scope | `crates/litchi-xlsb/src/cell_values/{root.rs,resources.rs}`, `cell_watches/workbook.rs`, `sparkline/workbook.rs`, `workbook/package.rs` |
| tests | `cell_values/resources.rs`, `cell_watches/tests.rs`, `sparkline/worksheet_tests.rs` |
| build profile | `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`, jobs 2 |
| cleanup | no preexisting Cargo target directory was cleaned |
