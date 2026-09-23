# Log sections for change 0748

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself.

---

## For `HOTSPOTS.md`

## 0748 — sealed owned CFB overlay plans hash their artifact once

[0748](0748-cfb-overlay-fingerprint-reuse.md) removes the SHA-256 passes that
0746 priced at about 42% of the XLS generic commit and 0745 listed as a
follow-up. A CFB same-length overlay plan re-hashed the whole artifact (source
and target) at the composed-view preflight, around and during every write, and
twice at planning for a generic source, to prove digests that cannot change
when the source is an owned immutable allocation. A plan over sealed owned
bytes (`open_owned`, new `open_owned_vec`) now computes both digests once, at
planning; generic `ReadAt` sources keep every pass. `render_copy_through` (the
object editor's copy-through, which opened its immutable original generically:
six passes per render) and the XLS visibility source-backed commit now open
sealed. On this base the XLS generic commit goes from 24 whole-artifact digests
to 4: `54016.xls` commit 23.68 → 14.26 ms (0.600), `xls_semantic_one_edit_save`
3.38 → 1.80 ms (0.533), `xls_visibility_eager_edit_save` 25.3 → 5.47 ms
(0.216), `xls_numeric_eager_rk_mulrk_edit_save` 0.213,
`xls_numeric_source_backed_number_edit_save` 0.391, plan-only publication
0.045; generic, DOC, PPT and gated controls are flat with identical
instructions. Remaining: the one planning pass (the target digest could fork
from the source hasher at the first changed byte, a median 35–50% into real
edits); DOC's owned opens are still generic.

---

## For `REPORT.md`

## 0748 — sealed owned CFB overlay plans hash their artifact once

[0748](0748-cfb-overlay-fingerprint-reuse.md) retains three commits
(`litchi-cfb`, `litchi-ole-common` + `litchi-xls`, harness evidence) on base
`ab29ac6291`: sealed owned plans compute their digests once and skip the
composed-view preflight and the emission hash; `SharedOleFile::open_owned_vec`
is added; `render_copy_through` and the visibility source-backed commit open
sealed; `OverlayOperationShape` and the harness evidence (`v2`) report the new
owned contract, and the ABBA validator keeps v1 rows on the v1 contract.
Paired ABBA on CPU 16 (8 processes per case): probe `54016.xls` generic commit
0.600, `xls-large` 0.546; `xls_semantic_one_edit_save` 0.533,
`xls_visibility_eager_edit_save` 0.216, `xls_numeric_eager_rk_mulrk_edit_save`
0.213, `xls_visibility_source_backed_edit_save` 0.220,
`xls_comments_source_backed_edit_save` 0.605,
`xls_numeric_source_backed_number_edit_save` 0.391, plan-only Number 0.531,
`cfb_file_owned_same_length_overlay_atomic_save` 0.758; controls 0.978–1.010
with identical instructions; the generic `cfb_file` control's p95 tail is
within its own A/A noise. Allocations fall by exactly the removed fingerprint
buffers. Outputs, refusals and fingerprint values are byte-identical to the
base over a 126-fixture census (1,319 lines). `performance_claim: none`.

---

## For `GOAL_AUDIT.md`

## 0748 — freshness proofs stay where the source can change

[0748](0748-cfb-overlay-fingerprint-reuse.md) removes only proofs over bytes that
cannot change: a plan over an owned `Arc<[u8]>`/`Arc<Vec<u8>>` retained by the
CFB reader (typed, crate-private provenance; no flag or wrapper can claim it)
computes its digests once, while every generic `ReadAt` keeps planning's
confirming scan, the view preflight, the write fences, the emission hash and
the save fences, each still pinned by a mutating-source test. Recorded
fingerprint values, composed-view versions and published bytes are unchanged
(census, tests); the composed reopen, owner validation and read-back all run.
Breaking changes, stated: sealed plans report zero preflight and zero hashed
emission bytes in `OverlayOperationShape`, the harness evidence moves to `v2`,
and sealed plans can no longer return fingerprint-changed refusals that could
never fire. ADR 0003/0005/0006 unchanged; 0652 trade-offs 1–3 applied. No
claim registered.
