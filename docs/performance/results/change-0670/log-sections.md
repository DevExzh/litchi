# Log sections for change 0670

## For `HOTSPOTS.md`

## 0670 — DOCX parser residues and correctness follow-ups

Record: [0670](../../0670-docx-parser-residues.md).

0651 row 14's safe slice-backed residue is implemented: chart, font,
header/footer, modern-comments, SmartArt, validation, and the semantic text
sink use borrowed quick-xml events, and the sink reuses one bounded paragraph
`String`. The guarded source-backed tail and paragraph-copy readers remain on
their existing buffered form pending a separate refusal-order proof. Row 16's
managed transaction BOM ranges now add the three consumed prefix bytes to every
retained event span and restore the mark through whole-document compaction;
the section codec admits the Strict schema's universal twip measures. The
package-level paragraph memo remains a priced design only, and final `sectPr`
placement remains schema-driven. Allocation evidence is deterministic and
`performance_claim: none`; root 0677's independent publication-audit fix is a
prerequisite for an end-to-end marked managed edit.

[Record and limitations](../../0670-docx-parser-residues.md); [retained
evidence](README.md).

## For `GOAL_AUDIT.md`

## 0670 — source ranges and Strict units are corrected without widening refusal contracts

Record: [0670](../../0670-docx-parser-residues.md).

The change follows ADR 0003's source-checked, typed-refusal behavior and ADR
0006's preserve-by-default rule. A leading BOM is part of the retained source,
so quick-xml's three-byte position skew is corrected before ranges are stored,
and compaction returns the original mark. Universal section measures are
admitted only through the schema-supported unit set and existing integer
twip/domain bounds. Existing parser ceilings, execution checks, and guarded
source-backed refusal order are retained. The body-final `sectPr` refusal is
not relaxed: local ECMA-376 `CT_Body` requires that element to be the optional
last child. No speculative paragraph cache is introduced; its managed memory
admission and generation lifetime are priced in a design memo first. The one
broad integration failure reproduces on the base commit and is recorded as
pre-existing.

[Record and limitations](../../0670-docx-parser-residues.md); [retained
evidence](README.md).

## For `REPORT.md`

## 0670 — the remaining borrowed parser work and a three-byte transaction witness

Record: [0670](../../0670-docx-parser-residues.md).

The DOCX parser's slice readers no longer allocate an event buffer for each
token, and the text sink keeps one paragraph buffer across the operation. At
10,000 paragraphs, the retained allocation probe records eager sink work moving
from 10,045 allocations and 503,555 bytes to 46 and 3,605; source sink work
moves from 10,046 and 503,651 to 47 and 3,701. Eager/source text controls are
unchanged. A managed BOM-marked document now reads and edits both paragraphs,
keeps `EF BB BF` in the committed XML, and releases managed memory. Strict
`612pt`/`792pt` section values round to the same twips as their integer forms,
and the reader preserves the original XML bytes. These are evidence counts and
correctness witnesses, not a registered performance claim. The body-final
section rule and the guarded buffered readers remain explicit boundaries.

[Record and limitations](../../0670-docx-parser-residues.md); [retained
evidence](README.md).

## For `ADR_COMPLIANCE.md`

## 0670 — compliant; preserve source bytes and keep the refusal fences

Record: [0670](../../0670-docx-parser-residues.md).

ADR 0003 is preserved because borrowed events are consumed within the parse,
all existing limit and execution checks remain in order, and the managed
transaction publishes only after source ranges and edits validate. ADR 0005 is
preserved because the sink drops its per-event buffer and retains only one
bounded paragraph capacity; the proposed paragraph index is withheld until a
typed `DocumentIndexAdmission` budget exists. ADR 0006 is preserved because a
producer-written BOM is retained through range correction and compaction, and
unmodified source bytes remain the source of truth. The Strict universal-unit
extension follows the local ECMA-376 schema and preserves integer section
domains. No unsafe code, public executor or archive type, weakened malformed
input defense, or body-final insertion policy was added. The independent
xml-minifier audit fix remains root 0677's separate prerequisite.

[Record and limitations](../../0670-docx-parser-residues.md); [retained
evidence](README.md).
