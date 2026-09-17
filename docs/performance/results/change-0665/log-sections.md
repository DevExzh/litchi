Four log paragraphs for change 0665, for the coordinator to merge into the
shared program logs. Each block below is written in the style of the newest
section of the file it names.

## For `HOTSPOTS.md`

## 0665 — the eager OPC route stops refusing producer spelling, and 58 of 180 real XLSX edits publish

Record: [0665](0665-opc-eager-writer-publication-audit.md).

**Row 2 of [0651](0651-queue-refresh-after-the-second-wave.md)'s queue is
closed on the eager route.** Decision 2 of [0652](0652-owner-decisions-for-the-third-wave.md)
quotes the owner as *"Original byte contract: loose the audit, accept
non-compact XMLs"* and speaks about original part bytes; 0654 applied that
profile to source-backed originals. This record carries the same
`verify_source` profile through eager planned payloads, including source-spliced
replacements, as an explicit implementation interpretation informed by 0654
and the source-backed replacement precedent in 0657. It is recorded as an
interpretation, not as an additional owner decision. The corpus oracle is
exact: on 0610's `xlsx-hide` route, **59 `NotCompact` refusals over 180 real
`.xlsx` fixtures become 0**, **58 newly publish**, and the 59th surfaces the
pre-existing `PreservationUnavailable` that the compactness verdict had hidden;
the three eager routes over 321 OOXML fixtures retain their no-op and compact
edit digests, while the non-compact edit route moves **0 → 310** published
fixtures. The DOCX census's default policy moves from **1 accepted / 54
fallbacks to 54 accepted / 1 fallback** over 55 openable fixtures, with the
one BOM refusal retained under change 0650's witnesses. Structural, encoding,
DOCTYPE, attribute, `xml:space` and finite-budget checks remain typed and
pre-output; the writer's own manifest and relationship serializations retain a
debug compactness assertion. `performance_claim: none`; instruction deltas
−0.26% to +0.01% and paired medians −1.75% to +1.30% are evidence only.
OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.
[Record and limitations](0665-opc-eager-writer-publication-audit.md);
[retained evidence](results/change-0665/README.md).

## For `GOAL_AUDIT.md`

## 0665 — preservation reaches producer-formatted eager members, with structural refusal boundaries intact

Record: [0665](0665-opc-eager-writer-publication-audit.md).

`docs/GOAL.md` makes correctness and lossless preservation precede speed. The
eager writer had applied the repository's compact-output spelling check to
payloads it did not author, so a producer's indentation, declaration boundary,
line ending or attribute layout could turn a valid edit into a typed
publication refusal. The implementation now separates the profiles: eager
planned payloads use `verify_source`, and the provenance decision is an
`Arc::ptr_eq` proof against the ingress allocation rather than an equality test
that a caller could escape by replacing a payload with equal bytes. This is
the record's implementation interpretation of 0652 decision 2, grounded in
0654 and 0657; it does not assert a new owner decision for authored payloads.
The source-backed replacement route is already covered by 0657 on the target
branch.

The preservation oracle keeps the goal's byte guarantees: 319 eager exact
no-op digests, 310 compact edit digests and 33 previously published
`xlsx-hide` digests are unchanged; all 1,022 XML members compared across the
91 after-leg XLSX publications are byte-identical where untouched, with no
member additions or removals. Kept structural and safety checks are exercised
through the public streaming writer with zero sink bytes on every refusal
family. The 63-fixture DOCX census leaves open, no-op, exact-no-op, opt-in and
reopen aspects unchanged; every changed artifact reopens with the same
paragraph, table and block-control counts. The remaining BOM publication and
managed-transaction defects are the retained change-0650 witnesses in
`results/change-0650/follow-ups.md`, not silently folded into this change.
`performance_claim: none`; OLE2/OOXML remain active, ODF is deferred and iWork
excluded. [Record](0665-opc-eager-writer-publication-audit.md);
[retained evidence](results/change-0665/README.md).

## For `REPORT.md`

## 0665 — eager OPC publication accepts producer XML without dropping safety checks

Record: [0665](0665-opc-eager-writer-publication-audit.md).

`litchi-opc` replaces its eager `validate_authored_xml` boundary with a
structural `audit_published_xml` and a debug-only `audit_authored_xml` for
library-generated manifest and relationship XML. `OpcPackage` now proves an
untouched source XML payload by allocation identity, and DOCX's preservation
gate asks the same source auditor as the writer. The public surface and typed
error identity stay unchanged. Three new OPC tests publish the four formerly
rejected non-compact spellings, assert generated members remain compact and
drive eleven structural/encoding/DOCTYPE/attribute/limit refusal families
through `write_to_stream` with zero output; moved tests retain the structural
fallback. The retained corpus packet contains 321-fixture eager routes,
180-fixture `xlsx-hide` and member identity checks, a 63-fixture DOCX census,
callgrind isolation pairs, paired timing and all gate tails. `cargo fmt`,
clippy, OPC/DOCX tests and docs, XLSX/PPTX consumers, the feature-bearing
facade and the non-iWork gate passed with `CARGO_BUILD_JOBS=2` for the costly
builds. The timing window reports a +1.22% PPTX p50 against a 0.04% floor and
does not explain or register it; no performance claim is made. The 0650 BOM
and managed-transaction follow-ups remain explicitly scoped out.
[Record](0665-opc-eager-writer-publication-audit.md);
[retained evidence](results/change-0665/README.md).

## For `ADR_COMPLIANCE.md`

## 0665 — compliant; ADR 0006 records the eager-route interpretation of decision 2

Record: [0665](0665-opc-eager-writer-publication-audit.md).

ADR 0006 is amended by this record under 0652 decision 2. The owner quote
names original bytes; 0654 established the source profile for source-backed
originals and 0657 established it for source-backed replacements. 0665 records
the eager route's use of that same profile for planned payloads, including
source-spliced replacements, as an implementation interpretation rather than a
new owner decision. Compactness remains checked by debug assertion and tests
where this repository authors manifest and relationship XML; it is not a
publication refusal for package payloads. UTF-8, well-formedness, one root,
character-data placement, DTD/DOCTYPE rejection, attribute grammar,
`xml:space`, finite budgets, deterministic error identity and zero-output
failure ordering remain intact. ADR 0003's atomic, typed and source-checked
publication boundary is unchanged; ADR 0005's limits, explicit providers,
absence of ambient execution and no public archive implementation leakage are
unchanged. `Arc::ptr_eq` tightens provenance without changing the preservation
planner's byte comparison. The BOM publication-audit and managed transaction
offset residues, plus the body-final-section and Strict measure questions,
remain the exact change-0650 witnesses and are not resolved here. The
xml-minifier BOM/offset implementation follow-up is 0677; the managed DOCX
transaction witness is assigned to 0670.
`performance_claim: none`; OLE2/OOXML remain active, ODF is deferred and iWork
excluded. [Record](0665-opc-eager-writer-publication-audit.md);
[retained evidence](results/change-0665/README.md).
