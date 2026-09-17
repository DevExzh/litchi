# Evidence: change 0670, DOCX parser residues and correctness follow-ups

Change record: [`0670-docx-parser-residues.md`](../../0670-docx-parser-residues.md).

Disposition: retained. `performance_claim: none`. This packet reports
deterministic allocation counts and correctness witnesses; it does not add a
claim-registry entry.

## Contents

| Path | What it contains |
| --- | --- |
| `decision.json` | Decision, scope, evidence, costs, gaps, and provenance. |
| `gates.txt` | Formatting, clippy, tests, and the one pre-existing failing test verified against the untouched base. |
| `log-sections.md` | Four coordinator paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md`, and `ADR_COMPLIANCE.md`. |
| `follow-ups.md` | The frozen paragraph-index memo design, body-final `sectPr` decision, and tail-reader boundary. |
| `measurements/allocations.txt` | Reps 4 to 20 counting-allocator isolation for eager/source controls and sink text at 24, 200, and 10,000 paragraphs. |
| `measurements/bom-managed.txt` | Managed source-backed BOM range and edit witness. |
| `measurements/spec-witness.txt` | Extracted local ECMA-376 Strict schema lines supporting the body-final and universal-measure decisions. |

The probe was built in release mode against the immutable 5fa92d7ce base and
this worktree. Development and test profiles used debug info level zero and
two build jobs; the release profile retained its measurement fidelity. Existing
Cargo targets were reused and not cleaned.

The one broad test failure is pre-existing: the untouched 5fa92d7ce checkout
fails `supported_settings_mce_fallback_is_admitted_and_output_limited` in
`source_backed_tail_append.rs` with the same `None` result. The changed branch
passes the full library suite, the focused managed BOM test, and clippy.
