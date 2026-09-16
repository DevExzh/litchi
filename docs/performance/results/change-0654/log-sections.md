Four log paragraphs for change 0654, for the coordinator to merge into the
shared program logs. Each block below is written in the style of the newest
section of the file it names.

## For `HOTSPOTS.md`

## 0654 — the queue's first row is cleared: 297 of 321 real packages now publish, and the eager route is the next one

Record: [0654](0654-opc-original-bytes-audit-loosened.md).

**Row 2 of [0651](0651-queue-refresh-after-the-second-wave.md)'s queue —
[0602](0602-xlsx-real-producer-admission-design.md)'s D0, priced by
[0613](0613-opc-original-audit-memo.md), blocked since
[0587](0587-remaining-opportunity-survey.md) — is implemented.** The publication
audit of a replaced Part's *original* bytes ran `verify_authored`, which asserts
this repository's compact output contract on bytes this repository did not
write. The census is exact: **1,378 of the 1,403 XML members** of 0602's
95-package corpus and **6,831 of the 6,981 members** of the 321-package OOXML
fixture corpus were refused for that reason alone, and **every one of the 95 and
318 of the 321 packages** now has every XML member accepted. End to end through
`write_part_overlay_to_stream`, **297 packages move from refused to published**,
13 publish with the **identical output SHA-256**, and **6,781 untouched members
across 310 published packages are byte-identical to their source member with 0
mismatches and 0 member-set changes**. The mechanism is one new policy —
`xml_minifier::audit::verify_source`, which keeps UTF-8, well-formedness, the
single document element, the DOCTYPE refusal and all six finite budgets and
asserts no compactness — reached from **seven** call sites in
`source_backed.rs` through a new `validate_source_part_xml` that builds the
identical `OpcError::XmlPublication`; the seven paired replacement sites and the
four litchi-generated-XML sites are untouched, and across the whole corpus the
`verify_authored` verdict moves on **0 of 6,981 members**. The audit's price is
re-measured and re-split now that the halves are separate symbols: **27.41% of
source-backed publication instructions for the original, 27.86% for the
replacement, 55.27% for the pair**, reproducing 0613's 27.59% and 55.06%. **The
next site is named with its number**: the eager `PackageWriter` route audits a
Part whenever `is_xml_part(…) && !is_exact_source_xml(part)`, has no
original/replacement pair to key on, and is **identical on both legs** over
change [0610](0610-opc-lazy-part-decode-design.md)'s `xlsx-hide` route — 59 of
180 `.xlsx` fixtures still refuse with `NotCompact`, 57 naming
`/xl/worksheets/sheet1.xml`, whose blob the retained witness measures as
**byte-identical to the source member**. `performance_claim: none`; the audit
costs +0.54% instructions and four of five timed scenarios are faster, neither
registered. OLE2/OOXML remain active; ODF is deferred until completion and iWork
excluded. [Record and limitations](0654-opc-original-bytes-audit-loosened.md);
[retained evidence](results/change-0654/README.md).

## For `GOAL_AUDIT.md`

## 0654 — preservation by default, applied to the bytes a producer wrote

Record: [0654](0654-opc-original-bytes-audit-loosened.md).

`docs/GOAL.md` puts correctness and lossless preservation above speed, and ADR
0006 promised that "lexical details are retained when possible". Publication
did not keep that promise: a package whose producer indents its XML — which is
every third-party producer, 94 of 95 by 0602's count — could be opened, read and
republished unchanged, but the moment one Part was replaced the *untouched*
spelling of the replaced Part's previous bytes was audited against litchi's own
byte-minimal output contract and the save was refused. This change separates the
two contracts and writes the separation into ADR 0006 with a dated note naming
0652 decision 2, keeping the amendment to the three sentences that change. What
the goal actually asks for is preserved on every axis that matters: the
published bytes of an untouched member are still the source's own (6,781
compared, 0 mismatches), the replacement is still audited in full (0 verdict
changes over 6,981 members), an exact byte no-op still bypasses both audits and
copies the source artifact verbatim even when its payload is malformed, and
every refusal that is not a compactness verdict still fires before any archive
byte is emitted — twelve malformed shapes, DOCTYPE, encoding and all six finite
budgets, each with a test. Two refusals the compactness verdict had been
shadowing now surface on two fixtures (trailing bytes outside the located
archive; no canonical UTF-8 source member on a package that writes
`xl\sharedstrings.xml` with backslashes), and the record states that rather than
averaging it away, with a base witness for the class. The goal's own trade-off
rule decided two questions the other way: change 0650's byte-order-mark defect
is **reported, not fixed**, because it is a `Malformed` refusal that the
streaming auditor pins with a test and moving it is a second contract change
0652 does not authorize; and the eager writer's audit is **left alone**, because
that site cannot tell an untouched source blob from an authored one and
loosening it there would stop enforcing compactness on litchi's own output.
`performance_claim: none`. OLE2/OOXML remain active; ODF is deferred until
completion and iWork excluded.
[Record](0654-opc-original-bytes-audit-loosened.md);
[retained evidence](results/change-0654/README.md).

## For `REPORT.md`

## 0654 — the original-bytes audit is loosened, and its price is re-measured on the way

Record: [0654](0654-opc-original-bytes-audit-loosened.md).

Two crates, five source files, +531/−60 tracked. `xml_minifier::audit` gains a
private three-constant `Policy` in place of one bool and one added public
function, `verify_source`; `litchi_opc::source_backed` gains one private
`validate_source_part_xml` and points its seven original-bytes call sites at it.
No public item is removed or changed, no `unsafe`, no dependency, no widened
limit. Counts first. **Corpus admission**: 1,403 XML members over 0602's 95
packages, accepted by the original-bytes audit **25 → 1,403**; 6,981 members
over 321 OOXML fixtures, **133 → 6,964**, with 17 still refused (16
byte-order-marked, 1 not UTF-8) and every one keeping its exact error text.
**Publication**: 321 fixtures through `write_part_overlay_to_stream` on both
legs — 297 refused → published, 13 published → published with the identical
SHA-256, 9 → 9 refused (7 identical text, 2 reporting a pre-existing shadowed
refusal), 2 never open; 6,781 untouched members compared with 0 mismatches.
**Instructions**, callgrind isolation pairs at `--samples 1`/`--samples 3` on
`xlsx_source_backed_cell_values_one_edit_save`, differenced and halved:
publication 187,808,737 → 188,636,450 Ir per iteration; the audit pair
103,698,462 → 104,258,468 (**+0.54%**), now split as **51,711,971 original
(27.41% of publication)** and 52,546,497 replacement (27.86%), 55.27% together
— which is the first direct measurement of the split, since before the change
both halves are one symbol. Audit call counts are unchanged at 8 per iteration.
The instruction cost is disclosed: removing compactness verdicts removes
branches, not loops, and the policy plumbing plus the loss of specialization now
that `verify_with_policy` has two callers costs slightly more than the verdicts
saved; a const-generic policy would recover 0.30% of publication and is declined
as speculative complexity. **Timing**, 30 samples per leg, order A1 B1 B2 A2 A3
A4 on CPU 9 in one window, five scenarios over three selectors: `one_edit
/medium` −2.14% (floor 0.03%), `one_edit/dense-sparse` −0.55% (0.85%),
`batch/dense-sparse` −0.83% (0.53%), `opc_source_overlay_one_part_save
/few-large` −0.46% (0.17%), and **`batch/medium` +1.20% against a 0.88% floor**,
reported rather than hidden in a mean and far below the 5% trigger. All six legs
on all five scenarios produced one output SHA-256 each. Gates: fmt, clippy, doc
clean; `litchi-opc` + `xml-minifier` 769 passed / 0 failed / 2 ignored;
`litchi-xlsx` + `litchi-docx` + `litchi-pptx` 3,658 / 0 / 33; the `litchi` facade
under `docx,xlsx,pptx,xls` 265 / 0 / 7; the harness's own suite 540 / 0 / 1;
the non-iWork gate. `performance_claim: none`. OLE2/OOXML remain active; ODF is deferred until
completion and iWork excluded.
[Record](0654-opc-original-bytes-audit-loosened.md);
[retained evidence](results/change-0654/README.md).

## For `ADR_COMPLIANCE.md`

## 0654 — compliant; ADR 0006 is amended by this record, as change 0652 assigned

Record: [0654](0654-opc-original-bytes-audit-loosened.md).

**ADR 0006 is the boundary that moves, and it moves exactly as far as change
0652 decision 2 authorizes.** The ADR never stated the compactness contract —
the file contains no occurrence of "compact", "whitespace" or "minif" — so the
contract on original bytes was implied rather than written, and the paragraph it
contradicted is the `Preserve` clause's *"lexical details are retained when
possible"*. Two sentences are added there, under a dated note naming 0652
decision 2 and this record, saying that the compact output contract binds bytes
this library authors or regenerates and is not asserted against the original
bytes of a Part a package already holds, which are audited for encoding,
well-formedness, a single document element, the absence of a DTD or DOCTYPE and
the finite budgets. Line 43's *"malformed known payloads fail before
publication"* is cited in the note as the surviving rule; no other ADR text
changed, and no ADR status changed. **The rest of ADR 0006 is preserved and
proved so**: preservation by default (6,781 untouched members byte-identical
across 310 published packages, 0 mismatches, 0 member-set changes; 13 packages
publishing the identical SHA-256 on both legs); validation never mutates (both
helpers take a slice and return a `Report` or an error); determinism (30
identical digests across six timing legs on five scenarios); and fail-closed
publication (both new `litchi-opc` tests assert `output.is_empty()` on every
kept refusal). **ADR 0003's typed-refusal rule holds**: no error variant is
added, removed or relocated, and both helpers construct the identical
`OpcError::XmlPublication { part, source }` from the identical
`xml_minifier::audit::Error`. **ADR 0005's rules are untouched**: no retained
state, no hidden pool, no ambient I/O, no archive type, raw lock or executor in
a public signature; the one added public item is a free function over a slice
and a `Limits` value. **0652's trade-off 2 decided two questions the safe way,
and the record says which**: change 0650's byte-order-mark defect is reported
rather than fixed, because it is a `Malformed` refusal that
`reader_matches_slice_bom_rejection` deliberately pins in the streaming auditor
and moving it is a contract change decision 2 does not cover (it costs 0 of the
95 packages and 3 of the 321); and `PackageWriter::validate_authored_xml`
(`pkgwriter.rs:203`) is left alone, because it audits a Part on
`is_xml_part(…) && !is_exact_source_xml(part)` with no original/replacement pair,
so pointing it at the source policy would stop enforcing compactness on
genuinely authored parts — a measured 59 of 180 `.xlsx` fixtures stay refused on
change 0610's `xlsx-hide` route, and the witness shows the refused Part's blob is
byte-identical to its source member. **One compliance cost is stated rather than
elided**: removing a refusal lets whatever stood behind it surface, and on two
fixtures it does — a trailing-bytes refusal and a canonical-member refusal, both
at untouched code strictly after the audit pair and both computed from raw
member names, with a retained base witness for the first class and a source
argument for the second. `performance_claim: none`. OLE2/OOXML remain active;
ODF is deferred until completion and iWork excluded.
