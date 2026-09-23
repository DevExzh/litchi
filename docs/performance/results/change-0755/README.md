# Evidence for change 0755

Record: [`../../0755-pptx-nested-text-run-panic.md`](../../0755-pptx-nested-text-run-panic.md).
Fix `ad2c490ee3` on `perf/0743-pptx-semantic-text-and-edit-path`; the failure is
present at `009d515bef`. `performance_claim: none`.

| path | what |
| --- | --- |
| `counters/probe-counters.json` | user instructions and cycles per operation, probe at `b82c81ceef` (before the fix, leg `a`) and `ad2c490ee3` (leg `b`), two iteration counts differenced, ABBA, two blocks; the probe and its driver are change 0743's (`../change-0743/probe/`, `probe_counters.py`) |
| `scan_nested_text_runs.py`, `scan-result.txt` | the corpus scan: slide-like parts of the repository's PPTX fixtures with a nested `a:t` (none) |
| `gates.txt` | the gate run at `ad2c490ee3`, shared with change 0743 |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

The regression and property tests live in
`crates/litchi-pptx/src/opened/text_run_tests.rs`. The reproduction without
the fix was run by reversing the three production files' diff in the working
tree, running `no_text_verb_panics_on_mutated_text_run_markup` (it failed at
variant 1,050 with the panic at `opened/xml.rs:840`), and re-applying the diff.

Binary digests and cleanup are recorded with change 0743
(`../change-0743/binaries.sha256`, `../change-0743/cleanup.json`); the pre-fix
probe is the `pc` entry of that series.
