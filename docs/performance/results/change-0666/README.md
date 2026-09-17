# Evidence: change 0666, worksheet MCE rewrite proof residue

Change record: [0666-mce-rewrite-residue](../../0666-mce-rewrite-residue.md).

Disposition: retained, implemented in `litchi-xlsx`; no registered
performance claim. The change continues decision 1 of change 0652 and closes
the safe worksheet portion of change 0651's MCE rows after changes 0653, 0664
and 0649.

## Contents

| Path | What it contains |
| --- | --- |
| `summary.txt` | The deterministic corpus census, preservation floor, focused tests and the adjacent-consumer disposition. |
| `measurements.tsv` | Machine-readable counts from the focused real-worksheet test and the independent ZIP census. |
| `gates.txt` | The focused formatting, compile and test gate tails. |
| `decision.json` | The retained change decision, evidence, costs and known gaps. |
| `cleanup.json` | Scratch and measurement retention decisions for this packet. |
| `log-sections.md` | Four coordinator-ready sections for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`. |

The real-worksheet test uses the repository's 95 XLSX fixtures and compares
each selected worksheet through the authoritative preprocessing path and the
admitted source path. It prints 205 worksheet parts, 71 borrowed, 30 rewritten
candidates, 25 rewritten candidates completing the validator-backed shared
traversal and 104 admission fallbacks. The independent ZIP census records 130
`mc:Ignorable`, 104 `dyDescent`, three other MCE-directive and zero
`AlternateContent` worksheet markers. The zero-difference floor is the
assertion that no value or exact error changed on any of those 205 parts.

The packet deliberately retains no release benchmark, profiler output,
allocation trace, generated archive or external fixture. This change records a
correctness and admission measurement; `performance_claim: none`.

## Provenance

Base commit: `5fa92d7ce`. Branch: `perf/0666-mce-rewrite-residue`. The focused
run used `CARGO_BUILD_JOBS=2` and the repository debug target. No accepted ADR
was amended, and no shared rollup was edited from this branch.
