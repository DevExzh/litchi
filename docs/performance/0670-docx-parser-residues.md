# 0670: DOCX parser residues, strict section measures, and managed BOM ranges

Status: retained, correctness and bounded-allocation follow-on. `performance_claim:
none`; the allocation counts and witnesses in the evidence packet are reported
to review the change and are not registered as a speed or throughput claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This change completes the safe part of 0651 row 14 and the justified part of
row 16 from the decisions in [0652](0652-owner-decisions-for-the-third-wave.md).
It is based on `5fa92d7ce` and keeps the parser's limits, refusal order, source
preservation, and public contracts intact.

## What changed

The slice-backed DOCX readers in the chart, font, header/footer, modern
comments, SmartArt, and validation codecs now use quick-xml's borrowed
`read_event()` and `read_resolved_event()` forms. Their reusable event buffers
and the corresponding clear operations are gone. A borrowed event is consumed
before the next event is read, so no event escapes the parse. The guarded
`BufRead`/source-backed loops in `source_backed/tail_append.rs` and
`source_backed/paragraph_copy.rs` remain on `read_event_into`; their bounded
token window, source freshness checks, cancellation checks, and refusal order
need a separate proof. All existing depth, node, event, byte, namespace, and
error checks remain in their existing order.

The semantic text sink now reuses one paragraph `String` for the duration of a
write operation. It clears and moves that bounded buffer into each paragraph,
then takes it back after the sink accepts the paragraph. The existing
`try_reserve`, paragraph-byte, document-byte, paragraph-count, run-count, and
execution checks are unchanged. The allocation witness at
[`measurements/allocations.txt`](results/change-0670/measurements/allocations.txt)
shows the operation-shaped allocation floor: at 10,000 paragraphs eager sink
work moves from 10,045 allocations / 503,555 bytes to 46 / 3,605, while the
eager and source text controls stay at 51 allocations / 1,641,483 bytes. These
are deterministic evidence counts, not a registered claim.

The managed document transaction now accounts for a leading UTF-8 byte-order
mark when converting quick-xml's `buffer_position()` values into ranges over
the retained source. It adds three bytes to both event boundaries when the
source begins `EF BB BF`, with checked offset arithmetic. Its whole-document
compaction fallback strips the mark only while parsing and restores it at the
front of the compacted result. The source-backed witness reads two paragraphs,
edits the first, keeps the second, retains the mark in the committed XML, and
returns managed memory to zero; it is recorded in
[`measurements/bom-managed.txt`](results/change-0670/measurements/bom-managed.txt).
The independent publication-audit slice fix is root's 0677 commit
`029b17b22`; this branch does not claim marked managed publication without that
prerequisite.

Section parsing and writing now admit the universal-measure member of
`ST_TwipsMeasure`, including `pt`, `mm`, `cm`, `in`, `pc`, and `pi`, while
retaining the existing signed and page-size bounds. Integer twips take the
existing path; universal values are parsed with the shared DrawingML
`Coordinate` lexical grammar and rounded to integer twips before the existing
domain conversion. Strict page dimensions, margins, header/footer distances,
column spacing, column widths, and document-grid line pitch are covered. The
reader test parses a `612pt` by `792pt` section and verifies the original
source bytes remain exact; the writer test verifies the same twip projection.
The local primary citation is the ECMA-376 archive
`3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip`, nested
`OfficeOpenXML-XMLSchema-Strict.zip`: `wml.xsd:3233-3238` defines `CT_Body`
as block-level children followed by one optional final `sectPr`, and
`shared-commonSimpleTypes.xsd:57-59` plus `:95-103` define
`ST_TwipsMeasure`'s positive universal measure and its `mm`, `cm`, `in`,
`pt`, `pc`, and `pi` suffixes. The extracted witness is retained in
[`measurements/spec-witness.txt`](results/change-0670/measurements/spec-witness.txt).

## What remains frozen

The body-final `w:sectPr` rule remains a refusal when another body child follows
it. The local `CT_Body` schema makes that placement invalid, and changing it
would change where an appended paragraph is placed; no insertion policy was
invented in this follow-on. The reader/editor disagreement documented by 0650
therefore remains explicit.

The package-level paragraph-index memo remains design only. The frozen design
in [`follow-ups.md`](results/change-0670/follow-ups.md) scopes one visible main
XML generation, stores checked ranges and metadata in a shared `Arc`, reserves
a typed `DocumentIndexAdmission` on managed owners before parser work, and
releases that reservation with the generation. Its actual range layout,
metadata price, and first/repeated-query measurement must be supplied before
state is added. No speculative cache or performance claim is part of 0670.

## Validation and limits

`cargo fmt --all --check`, `cargo clippy -p litchi-docx --all-targets --
-D warnings`, the 964-test `litchi-docx` library suite, the focused managed BOM
integration test, and rustdoc are recorded in
[`gates.txt`](results/change-0670/gates.txt). The broad `cargo test -p
litchi-docx --tests` run has one failure,
`source_backed_tail_append::supported_settings_mce_fallback_is_admitted_and_output_limited`;
the untouched `5fa92d7ce` checkout produces the same `None` result, so it is a
pre-existing failure and not attributed to this change. The broad rerun with
that test skipped passes all remaining suites. No timing, RSS, instruction,
throughput, or claim-registry result is asserted.

## Retained evidence

[`results/change-0670/`](results/change-0670/README.md) contains the decision
record, four log sections, gate tails, follow-ups, allocation counts, managed
BOM witness, and exact local schema citation.
