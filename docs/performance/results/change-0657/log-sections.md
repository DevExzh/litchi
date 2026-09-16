# Log paragraphs for change 0657

Four blocks, one for each log the coordinator merges. Each is written to sit
above that file's newest section. Every figure is from the packet: `census/` for the counts, `bench/summary.txt`
and `bench/counts.txt` for the measurements.

## For `HOTSPOTS.md`

## 0657 — the XLSX value editor's hot path was its own door, and the door was a name list

Change [0602](0602-xlsx-real-producer-admission-design.md) measured change
0525's reduced readback at **15.4-43.5% of a plan-and-commit p50** and the
complete candidate parse it removes at **32.9-85.3% of the commit that pays
it** — and then measured the population that receives any of it at **0 of 95**
real packages. The mechanism was not slow; it was unreachable. The reason was
five allow-lists in `cell_values`: sixteen worksheet element names, eight
workbook names, a per-element attribute list, and three relationship type
lists. Every one of them was a proxy for a property the rewrite already has —
`rewrite_value_only_with_provenance` composes the `<dimension>` `ref`
attribute and the `<sheetData>` span and copies every other byte with one
`extend_from_slice` — so an element outside those spans could be admitted
without the editor learning anything about it. Replacing the lists with that
property takes the corpus from **0 of 95 packages and 0 of 208 worksheet parts
admitted to 10 and 15**, and the whole remaining refusal set is four messages:
**64 shared-string workbooks, 8 pivot caches, 2 Strict-namespace shared-string
workbooks**, and the single-worksheet door. The gate behind it was not in `litchi-xlsx` at all. On the first
measured leg **0 of 95 still published**, all ten newly admitted packages
stopping in `litchi-opc`'s compactness audit at the newline their producer
wrote after the XML declaration — change 0602's D0. Change
[0654](0654-opc-original-bytes-audit-loosened.md) loosened that audit for the
*original* bytes and its finding 5 recorded that the value editor's
*replacement* then raised the byte-identical refusal, because the replacement
is a splice carrying the source's own formatting. This change closes it: the
seven paired replacement-audit sites audit the replacement with `verify_source`
too, and the corpus publishes: **9 of the 10 admitted packages complete a plan,
commit and publication**, against 0 on both earlier legs, each changing exactly
two members and reopening with the edited cell readable and no neighbour lost.
[Change and limitations](0657-xlsx-value-editor-d4-admission.md);
[evidence](results/change-0657/README.md).

## For `GOAL_AUDIT.md`

## 0657 — the admission surface was the gap, and four dependency guards were already written for it

Change [0587](0587-remaining-opportunity-survey.md)'s item XLSX-2 asked for
0525's readback to be widened to real worksheets and named two gates;
change 0602 found that every real file stops at the first of eight and that one
of the two named gates cannot fire at all. This change closes that: the
**vocabulary** gates and the **relationship** gates are replaced by the
dependency rule the owner accepted as decision 7 of change
[0652](0652-owner-decisions-for-the-third-wave.md), and the corpus goes from 0
admitted packages to 10. The audit finding worth recording is what was found
behind the door. Four of the rule's verdicts — a protected sheet, a data
validation covering the edited address, a shared or array formula covering it,
a merged range covering it but not anchoring it — were **already implemented**
in `raw/worksheet/edit/validation.rs::validate_actions` and had never been
reachable through this door, because the vocabulary refused the elements that
declare them before the scan ran. The 23 witnesses in
`cell_values/admission_tests.rs` are the first tests to reach them. The
remaining gaps are named rather than closed: shared strings (0602's D2, **66 of
95**), pivot caches, tables and query tables are refused because their meaning
depends on the value being replaced, and the publication audit (0602's D0)
still refuses every real package, so this change makes producer files
**editable, not saveable**. The gate that decides whether this work has any
user-visible effect is therefore change 0654's, not this one's; re-running
`results/change-0657/probe` over the corpus after 0654 merges is the check that
should be recorded.

## For `REPORT.md`

## 0657 — the value-only editor's allow-lists become one property, and 10 of 95 real packages are admitted

`cell_values/validation.rs` no longer carries a vocabulary. It classifies each
element as **modelled** — it lies in a span the rewrite composes, which is the
worksheet's `<dimension>` `ref`, its `<sheetData>`, and the workbook's
`<sheets>` catalog and `<calcPr>` — or **copied**, and admits every copied
element, its subtree, its namespace, its attributes and its character data
unread. Attributes need no list in either half, because `write_tag` re-emits
the source tag of every element the writer touches and replaces only `@r`, `@t`
and `@s` of an edited `<c>` and the `ref` of `<dimension>`; the one attribute
still refused is a relationship reference inside `<sheetData>`, which a removed
cell record would orphan. The three relationship gates admit every unfamiliar
type and refuse, by name, the set whose target stores a copy of or a name
derived from the edited value: shared strings, pivot caches, tables and query
tables, plus every external target. The two validator loops are merged into
one, so the borrowed traversal of change 0603 and the authoritative pass cannot
disagree. Deterministic census over change 0602's 95-fixture corpus, both legs,
`taskset -c 12`: packages admitted through `edit_many` **0 → 10**, through
`snapshot` **0 → 3**, worksheet parts **0 of 208 → 15 of 208**, twelve fixtures
changing verdict and **none regressing**. Packages completing a publication
cycle: **0 → 0 → 9**. The middle leg is the evidence that the admission rule
alone was not enough: all ten stopped in `litchi-opc`'s compactness audit, nine
at byte 55 of their own `xl/workbook.xml`, which the same audit rejects in the
*source*. With the replacement audited by source rules, **9 of 9 published
packages differ in exactly `xl/workbook.xml` and their edited worksheet part**,
0 members added, 0 removed, every other member byte-identical, the edited part
differing in exactly one contiguous range, 9 of 9 reopening with the edited
cell readable and the same stored-extent cell count as the source, and 5 of 5
republishing byte-identically across two processes. The tenth fails on a
non-canonical relationship member, which is not a compactness refusal. Instructions and paired timing over change 0601's producer-shaped corpora, both
legs at `ceb0571a8`, pinned to CPU 12: planning **16,312,015 → 16,217,854 Ir**
per operation (−0.577%), one-edit-save +0.038%, and the changed module's
inclusive share −0.71% and +0.34%. Paired p50 with a **2.31% A/A floor** and a
pre-declared per-leg admission rule (a leg counts only if p95/p50 ≤ 1.05):
medium planning **−3.19%**, dense planning −2.40%, medium one-edit-save
**−3.40%**, dense one-edit-save −2.61%, all past the floor, while the two
**marker-free controls stay flat at −0.25% and −0.58%** — the contrast that
distinguishes the effect from binary layout. Nothing got worse by more than 5%
at p50. The effect does not appear in the instruction counts, so it is reported
as evidence and not registered. The publication change is separately attributed
at **100,172 Ir per operation** across every audit context, 0.26% of a whole
one-edit-save. Gates: `cargo fmt --all --check`; `cargo clippy -p
litchi-xlsx --all-targets` clean; `cargo test -p litchi-xlsx` (1,021 lib tests,
23 of them new witnesses); `cargo doc -p litchi-xlsx --no-deps`; `cargo test -p
litchi --features docx,xlsx,pptx,xls`; `cargo test` in `tools/perf-baseline`;
`python3 tools/non_iwork_gate.py verify`. No public API changed.

## For `ADR_COMPLIANCE.md`

## 0657 — preservation proved by construction instead of by refusal

ADR 0006's preservation default is what this change is *about*, on both of its
halves. The old allow-lists refused a package because it carried `docProps/app.xml`, a printer
setting, a `<pageMargins>` or an `mc:Ignorable` — every one of which
`litchi-opc` already transfers verbatim and the worksheet rewrite already
copies with one `extend_from_slice`. Refusing them was refusing to preserve
what was already preserved. The rule that replaces them is a preservation proof:
an element is admitted precisely because the rewrite reproduces its bytes and
never reads them, and the corpus evidence is the check — **9 of 9 published
packages touch exactly the two members the topology plan names**, 0 added, 0
removed, every other member byte-identical, and all nine reopen with the edited
cell readable and the same cell count as their source. The publication half is
the same argument one layer down: a source-backed *replacement* is a splice of
the part's own bytes, so auditing it as this library's authored output refused
the producer's own indentation. Compactness stops being a publication refusal
on that route and stays a property of this library's serializers; every other
audit — UTF-8, well-formedness, one document element, no DTD, every budget —
holds on both sides, and a refusal still precedes any output. ADR 0005's mandatory validation is untouched: the complete traversal
still runs over the complete output before any snapshot is published and over
the source on every planning path, only its vocabulary widens, no traversal is
skipped and no boundary moves. ADR 0003's readback gate is *strengthened*: the
assertion that a readback worksheet part carries no relationship becomes the
assertion that it carries the same relationships as the source, which is what
makes a relationship-bearing worksheet checkable at all. Two defences the old
vocabulary was maintaining as a side effect are made explicit rather than lost:
a copied subtree is bounded by the raw parser's own `MAX_XML_DEPTH`, and an
element whose prefix no declaration binds stays refused even where it would be
copied. No `unsafe`, no new dependency, no weakened limit, no ambient I/O, no public
API item added or changed, and no refusal traded for a partial result: three
refusal *points* move — cell metadata to `validate_scalar_cells`, misplaced merge
markup to the raw parser, a worksheet relationship to a per-type message — and
each has a witness naming the new point. Change 0541's six frozen
error-precedence properties are unaffected, because widening a vocabulary
changes which inputs refuse, never which error wins when two compete.
