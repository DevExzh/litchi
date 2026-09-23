# Ready-to-paste log sections for change 0747

## HOTSPOTS.md

## 0747 — replaced-Part replacement audit proved from its original's audit

[0747](0747-xlsx-publication-audit-reuse.md) retires the duplicate half of the
source-backed publication's XML audit pair. For a local edit the replacement
repeats the original byte for byte outside one element, and the replacement's
`verify_source` re-scanned those bytes moments after the original's audit had
accepted them in the same parser state. The measured one-cell worksheet pair
differs in one byte of 63,294 / 462,568. `xml_minifier::audit::verify_source_replacement`
keeps the pair's exact verdict and error. It rescans only the replaced
element's new bytes from the original's state, with every budget re-totalled,
and falls back to the complete scan on any doubt. The one-cell edit/save's audit
instructions fall 47.6% / 47.8%, publication time 15.5%, and the whole
operation 6.63% / 6.32% (medium / dense-sparse). The original-bytes audit
remains a complete scan: planning does not subsume it (six witnesses). Fusing
it into planning's traversal is the next opportunity, bounded by the tokenizer
at about 3%. Multi-edit commits need writer-declared windows, since 0705's audit
share of one-percent publication is unchanged.

## REPORT.md

## 0747 — XLSX publication audit reuse

[0747](0747-xlsx-publication-audit-reuse.md) is retained with
`performance_claim: none`. Base `009d515bef`; commits `1ccfe6b354` and
`6b49ce999a`. All seven paired source-policy audit sites in `litchi-opc` call
one `verify_source_replacement`, whose verdict, failing side and error value
equal the two `verify_source` calls. Generated differential campaigns covering
17,753,101 window proofs, 214,761 of them over two separated edits, showed no
difference. So did an independent review of about 595 million cases reported by
the coordinator. The invariant the window relies on (no source-policy check
reads a token's position or other non-local state) is documented at the
auditor's token loop, and debug builds re-derive every window proof from a
complete audit. With both legs built by the same command (a first
campaign against the prebuilt base showed layout-induced control shifts and is
retained), `xlsx_source_backed_cell_values_one_edit_save` moves 4.448 → 4.165 ms
(−6.63%, CI [−7.39%, −5.55%]) on medium and 28.783 → 27.005 ms (−6.32%, CI
[−6.83%, −5.84%]) on dense-sparse. The one-percent route, eager and ordinary
save, docx and pptx controls are within noise (worst median paired change +0.36%, pptx).
Every output digest is identical. No published byte, refusal or error changes.

## GOAL_AUDIT.md

## 0747 — audit duplicates in source-backed XLSX publication

[0747](0747-xlsx-publication-audit-reuse.md) enumerates every check on the
source-backed publication path and asks, for each XML audit, whether an equal or
stronger check already ran over the identical bytes in the same operation. For
the original bytes the answer is no. `litchi-xlsx` never calls the auditor, and
witnesses show planning admits CDATA outside the root, a missing attribute
separator, an invalid `xml:space` and an over-budget attribute count. Two of
these are refused only by the original's audit, so that audit stays complete,
and the existing `SourceXmlPart` route is no substitute. For the replacement,
the writer can turn an audited original into a refused one, so its verdict must
still be established. It now is, by scanning only the edited element in the
original's parser state (ADR 0006 unchanged, 0652 decision 2 honoured, 0705's
rejection of readback-based dropping kept). Gates, the differential campaigns,
fail-closed refusals with empty sinks and identical outputs are recorded.
