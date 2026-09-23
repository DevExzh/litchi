# Ready-to-paste log sections for change 0754

## HOTSPOTS.md

## 0754 — the DOCX semantic edit and full-text paths without their three largest avoidable costs

[0754](0754-docx-semantic-edit-and-text-path.md) removes what profile r2
attributed on the ordinary DOCX semantic routes. The layout scan behind
`edit_document` is also every snapshot's admission check, so it stays eager and
was made cheaper instead: a plain `Reader` with the shared binding tracker,
borrowed events, and namespace resolution only where a verdict reads it
(69.9 M → 33.8 M instructions and 70,040 → 31 allocations per scan of the
1 MB large main part; 65% of what remains is quick-xml's tokenizer). A
one-paragraph edit's commit answers its source gate with 0747's pair audit and
hands the proven candidate to the eager writer as an
`xml_minifier::audit::VerifiedSource` bound to the exact allocation, so the main
part is audited completely once instead of twice (100.7 M → 52.3 M audit
instructions per edit). The namespace tracker caches recent prefixes, keeps a
default-namespace stack, and past 32 bindings an ordered prefix index, which
also removes a pre-existing quadratic case (a 1 MB part with ~30,600 in-scope
declarations: 430 ms → 14 ms per `Document::text`). Remaining: the commit's one
complete source audit (~2 ms on this part), multi-window proofs for scattered
edits (the one-percent route still audits its candidate completely), the
source-backed route's own gate-plus-original duplicate that 0750 noted, and
the text copy helpers' byte-by-byte scans.

## REPORT.md

## 0754 — DOCX semantic edit and text path

[0754](0754-docx-semantic-edit-and-text-path.md) is retained with
`performance_claim: none`. Base `63ec6a5027`; commits `5bb8bd1050`,
`3ab983e5cf`, `fde01f68e1`, `2c467b3e66`. Measured ABBA on core 12 with both
legs built by the identical command, 12 processes per leg, on the large
semantic corpus: `docx_semantic_noop_edit_save` 3.837 → 1.505 ms (−60.70%,
95% CI [−60.85%, −60.10%]), `docx_semantic_one_edit_save` 8.942 → 4.652 ms
(−47.98%), `docx_semantic_one_percent_edit_save` 9.259 → 6.970 ms (−25.03%),
`docx_semantic_full_text` 3.126 → 1.950 ms (−37.69%); medium −56.30%, −38.73%,
−20.39%, −37.02%; `docx_source_backed_one_edit_save` −6.28% in a 32-process
rerun (its process distribution is bimodal on both legs, as 0747 saw). The
controls are flat (`docx_ordinary_save_lifecycle` +0.04%,
`pptx_semantic_full_text` +0.12%, `xlsx_first_cell` −1.36%). Every output
digest is identical, and a base-versus-branch differential over 719 DOCX
inputs (fixtures and mutations, every edit route), 719 raw main parts and 78
PPTX fixtures differs nowhere. Additive API: `VerifiedSource` and
`Part::set_blob_verified`.

## GOAL_AUDIT.md

## 0754 — DOCX semantic edit and text path

[0754](0754-docx-semantic-edit-and-text-path.md) removes unnecessary parsing
and validation work (GOAL.md optimization steps 1–2) without moving a refusal:
the admission scan keeps its refusal set, error text and timing (the previous
scanner is a test oracle over fixtures, mutations, limits and managed budgets);
the audit skipped at publication was run by the same auditor under the same
limits on exactly the published allocation, and every other byte is still
audited by the writer, which stays the last line of defence (ADR 0006 does not
move); exact no-ops still share their source allocation (ADR 0003). The
namespace lookup's worst case falls from linear in every declaration in scope
to logarithmic with no hashing (ADR 0005), and 0652's trade-offs hold: the
benign path is the one made cheaper, and nothing is removed from the
malicious minority's defences. `performance_claim: none`; all gates pass on
the final commit (7,830 tests, 0 failures).
