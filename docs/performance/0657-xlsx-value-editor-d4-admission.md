# 0657: the XLSX value-only editor's allow-lists become a dependency rule, and real producer packages are admitted and republished

Status: retained, implemented in `litchi-xlsx` and `litchi-opc`.
`performance_claim: none` — the counts below are reported as evidence, not
registered as a claim. The change is **not** value-identical: it deliberately
widens what the value-only editor admits, under decision 7 of change 0652, and
it stops the source-backed publication audit refusing a replacement for
carrying its own source's formatting, under decision 2 as change 0654's
finding 5 left it. Every input admitted before is admitted after, with the same
output bytes; every input refused after is refused with a typed error named
here.

Base `ceb0571a8`, the merged head after changes 0654, 0658, 0660, 0664 and
0655; branch `perf/0657-xlsx-value-editor-d4-admission`.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was changed

`crates/litchi-xlsx/src/cell_values/validation.rs` — the sixteen-name worksheet
vocabulary, the eight-name workbook vocabulary and the per-element attribute
allow-lists are replaced by the property they were a proxy for. The file now
classifies each element as **modelled** (it lies in a span the rewrite composes)
or **copied** (the rewrite reproduces its bytes and never reads them), and
admits every copied element, its whole subtree, its namespace and its
attributes. The two validator loops — the authoritative one and change 0603's
borrowed observer — are merged into one, so the policy is stated once.

`crates/litchi-xlsx/src/cell_values/snapshot.rs` — the package-root
relationship allow-list becomes change 0602's closure proof; the workbook and
worksheet relationship gates admit every unfamiliar type and refuse the
value-dependent set by name; the readback assertion that a worksheet part
carries *no* relationship becomes the assertion that it carries the *same*
relationships as the source. One new guard, `require_unmerged_target`.

`crates/litchi-xlsx/src/cell_values/source.rs` — that guard is called on the
five staging paths that write into an existing cell.

`crates/litchi-xlsx/src/raw/worksheet/{mod,model}.rs` — `MAX_XML_DEPTH` becomes
`pub(crate)` and is re-exported, so the validator can carry the parser's own
bound instead of inheriting one from a shallow vocabulary. No value changed.

`crates/litchi-xlsx/src/row_visibility/{rewrite,mod}.rs` — the row-visibility
editor asks the same dependency question about its own edit, *which rows are
hidden*, and its answers differ from the value editor's in both directions. A
sheet protection, an autofilter, a sort state and a custom sheet view are
refused by its own scan, because each carries hidden-row state in exactly the
attribute this module owns and the shared vocabulary used to refuse them on its
behalf; so is markup-compatibility content, because its cell store is built
from the preprocessed bytes while its rewrite is lexical over the source. A
worksheet *relationship*, which that module's documentation also claimed to
refuse, is now admitted: the rewrite copies every byte but the row tags, and a
hyperlink, drawing, comment or printer-settings part is anchored by address, so
nothing about it depends on a row's visibility. The module doc is rewritten to
state the rule rather than the old list.

`tools/perf-baseline/src/lib.rs` — the row-visibility corpus's five semantic
refusal gates become four refusals and one *admission*: the relationship gate
now verifies that the variant is admitted. Its evidence field is renamed to
`relationship_source_admission_verified` so the name says what it checks.

`crates/litchi-opc/src/source_backed.rs` — the **seven paired replacement audit
sites** change 0654 left on the authored contract now audit the replacement
with `validate_source_part_xml` (change 0654's `verify_source`), as they
already audit the original. See *The publication half* below. The two sites
that audit this library's canonical relationship XML and the two that audit a
Part a plan *adds* keep `validate_overlay_xml` unchanged: those compose a whole
stream with no source to inherit from.

`tools/perf-baseline/src/producer_shape.rs` — change 0601's producer-shape
corpus builder proved its shape by asserting that the value-only editor
*refuses* each of five producer facts, and returned a corpus-build error if one
was admitted. Four of the five are admitted now, so on this branch the `read`
variant could not be built at all. The census records the editor's verdict
instead of asserting it — an admitted fact carries `admitted by the value-only
editor` as its message — and the evidence schema is unchanged. This is the
clearest single measure of what the change does: of the five producer facts
0601 found the editor refusing, **four are now admitted and one still refuses**,
and the one that refuses is the shared-string relationship.

Documentation: `cell_values`'s module doc now carries the verdict table;
`validation.rs`'s module doc states the rule and its proof;
`validate_source_part_xml`'s doc states the two-sided contract; the facts
builder's decline comment no longer claims the validator refuses the containers
it declines on.

## Authority

Decision 7 of change [0652](0652-owner-decisions-for-the-third-wave.md), the
owner's words:

> "Admission-surface widening: accept 0602's D4 — treat other elements or
> attributes as unfamiliar ones and just copy them; replace the naive whitelist
> with the D4's dependency rule."

0652 states what that authorizes: "the value-only editor's admission surface
widens: unfamiliar elements and attributes are copied through unchanged, and
D4's dependency rule decides which constructs an edit must understand", and
what this record must prove: "the dependency rule stated and tested per
construct it names; the real-producer corpora of 0601 admitted and republished
with the edited cell changed and nothing else; every construct still refused,
listed with its reason".

**What it does not authorize, and this change therefore does not do.** D2 of
change 0602 — capturing the shared-string part and materializing its table — is
not an unfamiliar element or attribute copied through; it is a new construct the
editor would have to *model*, with its own resource falsification (0602's F3).
Decision 7 does not name it, so a `t="s"` worksheet stays refused, by the same
gate and with the same message as before. That is the single largest remaining
refusal, 66 of the 95 real packages, and it is stated as a limitation rather
than smuggled in. D0 (the publication original-bytes audit, decision 2, change
0654 in this wave) is likewise untouched here and is why no real package
publishes yet; section *Measured* gives that number before and after.

## The publication half

Change [0654](0654-opc-original-bytes-audit-loosened.md) implemented decision
2 for the *original* bytes of a replaced part, and its finding 5 recorded what
that left standing: "the XLSX value editor's replacement carries the source's
formatting verbatim, so the *replacement* audit now raises the byte-identical
refusal at the same offset that the *original* audit used to raise". Admitting
a real producer package to the editor is worth nothing while that holds — 0602
said so as its D0 and this record's first measurement leg confirmed it, with
**0 of 95 packages completing a publication cycle** even after the admission
surface widened.

The reading this wave takes of decision 2 is that **compactness is not a
publication refusal on any route**: it is a property of this library's own
serializers, asserted by their own tests. The source-backed replacement is the
clearest case. It is a *splice*: the value-only rewrite copies the worksheet's
envelope, every unedited row and every unedited cell record from the source and
authors compact bytes only for the `<dimension>` tag and the edited `<c>`
records, and `calculation_properties`' span editor does the same to the
workbook's `<calcPr>`. There is no sense in which the assembled part is this
library's authored output, and auditing it as though it were refuses the
producer's indentation — which is to say, every real third-party package.

So the seven paired sites audit both sides with `verify_source`. **Every other
check stays, on both sides**: UTF-8, well-formed XML, exactly one document
element, no DTD or DOCTYPE, every `xml_minifier::audit::Limits` budget, and the
same `OpcError::XmlPublication { part, source }` for each. A refusal still
precedes any output: the fail-closed witnesses below assert an empty archive.
Nothing about the eager `PackageWriter` route is touched — that is change
0665's — and no `xml-minifier` helper changed, so this change and 0665 do not
share a file.

## Breaking changes

**No public API item changed.** No public type, function, trait, constant,
error variant or feature was added, removed or altered. `EditBlock` gains no
variant: the merge refusal reuses `EditBlock::CoveredMerge`, which
`require_insertable_absence` already returned. Every changed item in
`litchi-xlsx` is `pub(super)`, `pub(crate)` or private, and the `litchi-opc`
change swaps which of two private helpers seven call sites use.

**One published contract moves**, and it is a behavioural change callers can
observe:

* `SourceBackedPackage`'s topology and overlay publication routes no longer
  refuse a Part or relationship replacement because its XML is not compact.
  They refuse everything else they refused before. Two tests that pinned the
  old behaviour are rewritten to pin the new one and are named under
  *Correctness evidence*.

## The dependency rule

A value-only rewrite composes exactly three spans, and copies everything else
byte for byte from the source:

1. the worksheet's `<dimension>` `ref` attribute, rewritten in place by
   `write_tag`, which preserves every other attribute of that tag;
2. the worksheet's `<sheetData>` element, where an unedited row is copied
   `source[row.span]`, an unedited cell `source[cell.span]`, and an edited cell
   is re-tagged from its own source tag with only `@r`, `@t` and `@s` replaced;
3. the workbook's `<calcPr>`, spliced by `calculation_properties::rewrite`,
   which is a span editor and preserves unknown attributes and foreign siblings.

So the rule has two halves.

**Vocabulary.** Inside a composed span the editor must model what it meets:
`worksheet > (dimension | sheetData)`, `sheetData > row > c > (f | v | is)`,
`is > (t | r)`, `r > t`, and `workbook > sheets > sheet`. An unmodelled name
there is refused with the message it was refused with before. Outside those
spans — a child of `<worksheet>` or of `<workbook>` that is not one of the five
names above, and everything under it — the element is *unfamiliar*: admitted
whatever it is called, in whatever namespace, with whatever attributes and
whatever character data. `<rPr>`, `<rPh>` and `<phoneticPr>` inside a cell's
`<is>` are copied on the same terms, because the worksheet parser collects
inline text only from `is > t` and `is > r > t`, so nothing under them can
change a value.

Attributes need no list at all, in either half, because the writer re-emits the
source tag of every element it touches and replaces only the attributes it owns.
The one attribute refusal that remains is a relationship reference inside
`<sheetData>` — `CT_Row` and `CT_Cell` have no `r:id` in either dialect, and a
removed cell record would take a planted one with it and orphan the
relationship the rewrite is contracted to preserve. This is 0602's D3
adversarial gate, now enforced directly rather than as a side effect of the
namespace rule.

**Constructs.** What an edit must understand is not a name but a dependency:
does the construct store a copy of, or a name derived from, the value being
replaced? The verdict per construct, each with a witness in
`cell_values/admission_tests.rs`:

| construct | verdict | why, and where it is decided |
| --- | --- | --- |
| shared strings (`t="s"` + the `sharedStrings` part) | **refused** | the value lives in a part the editor does not model (0602's D2); `validate_workbook_relationships`, message unchanged |
| pivot caches (`pivotCacheDefinition`, `pivotCacheRecords`) | **refused** | the cache stores a copy of the source cells |
| tables and query tables (`table`, `queryTable` worksheet relationships) | **refused** | a table column is named after its header cell's text |
| cell metadata (`cm`, `vm` on `<c>`) | **refused** | it describes the value being replaced; admitted by the vocabulary, refused by `validate_scalar_cells` — a *moved* refusal point, witnessed |
| an external relationship, at any of the three levels | **refused** | no closure can be proved over a target outside the package |
| a relationship reference inside `<sheetData>` | **refused** | a removed cell record would orphan it |
| a protected sheet (`<sheetProtection>`) | **edit refused** | `EditBlock::ProtectedSheet`, in the raw layout scan |
| a data validation covering the address | **edit refused** | `EditBlock::DataValidation`, likewise |
| a shared or array formula covering the address | **edit refused** | `EditBlock::GroupFormula`, in the raw scan and in `require_insertable_absence` |
| a merged range covering the address | **edit refused unless the address anchors it** | `EditBlock::CoveredMerge`, now also at staging via `require_unmerged_target` |
| a cell whose payload sits inside MCE markup | **edit refused** | `EditBlock::MarkupCompatibility` |
| the calculation chain | **adjusted** | dropped atomically with its workbook relationship, as before |
| a formula's cached `<v>` elsewhere in the sheet | **adjusted** | `calcPr` invalidation forces a full recalculation on load, as before |
| merged ranges, conditional formatting, data validations, hyperlinks, autofilter, sort state, ignored errors, cell watches, page setup and margins, breaks, drawings, OLE objects, controls, `<extLst>` and vendor payloads | **admitted** | each addresses cells by coordinate; a value change does not move it |
| defined names, external links, workbook protection, file version, book views, `<extLst>` | **admitted** | copied with the workbook's head and tail |
| `docProps/app.xml`, `docProps/core.xml`, thumbnails, custom and vendor metadata | **admitted** | never named by the topology plan; `litchi-opc` transfers them verbatim |
| a worksheet's printer settings, drawing, legacy drawing, comments | **admitted** | the part is never named by the plan and its `r:id` lies outside `<sheetData>` |

Four of those verdicts — protection, data validation, group formula, covered
merge — were **already implemented** in the raw editor's layout scan
(`raw/worksheet/edit/validation.rs::validate_actions`) and were simply
unreachable: the value-only vocabulary refused the elements that declare them
before the scan ever ran. This change makes them live, and the witnesses in
`admission_tests.rs` are the first tests to reach them through the value-only
door.

## Why it is sound

**Preservation is by construction, not by refusal.** The head before
`<sheetData>` and the tail after `</sheetData>` are one `extend_from_slice` each
from the source; an unedited row and an unedited cell likewise. Admitting a name
there costs nothing but the promise not to interpret it, which the code cannot
break because it never looks. That is ADR 0006's preservation default satisfied
directly; the old allow-list refused to preserve what was already preserved.

**Validation is not weakened, and two defences are strengthened.** ADR 0005's
mandatory complete traversal still runs, complete, over the complete output
(`snapshot.rs` validates the candidate before any snapshot is published), and
over the source on every planning path. The malformed-attribute scan
(`attributes().with_checks(true)`) still runs on every element, including
copied ones, so a duplicated or unquoted attribute is still refused with its
existing message. The root-element, dialect, DOCTYPE and text-context refusals
are unchanged inside composed spans.

Two exposures the widening opens are closed in the same change, because the old
vocabulary was closing them as a side effect. A copied subtree may nest as
deeply as the input says, where the old vocabulary admitted at most five
levels, so the validator now carries the raw parser's own `MAX_XML_DEPTH` bound
(256) and refuses past it — the parser refuses on exactly the same input, so
this cannot refuse anything the pipeline would otherwise accept, and it stops
the element stack growing with a hostile input before the parser is reached.
And an element whose prefix no declaration binds stays refused everywhere,
inside a copied subtree as well as in a composed span, with its existing
message: a copied element may belong to any namespace, but this module does not
admit namespace-malformed input merely because it would be copied. Both have
witnesses.

**Error identity.** No refusal message changed at the package, workbook or
element gates for an input that was refused before and is refused after. Three
refusal *points* moved, each with a witness: `cm`/`vm` now refuse at
`validate_scalar_cells` ("cell edits refuse unknown cells and cell metadata")
instead of at the attribute list; a `<mergeCells>` placed before `<sheetData>`
now refuses in the raw parser ("worksheet mergeCells appears before sheetData")
instead of at the element list; a worksheet relationship that is refused now
names its type (`value-only edits refuse worksheet relationship '<type>'`)
rather than refusing all of them at once. Change 0541's six frozen
error-precedence properties are untouched: this change moves which inputs
refuse, never which error wins when two compete, and the one place where order
could have moved — the attribute scan, which previously ran after the element
and parent checks — keeps that order exactly.

**The rewrite's own guards are now reachable, and the fast route still declines
to reach them.** `rewrite_from_facts` (changes 0622 and 0635) does not run
`validate_actions`. It is safe because `FactsBuilder` declines any direct
worksheet child other than `dimension`, `sheetFormatPr`, `cols`, `sheetData` and
`sheetViews`, so every worksheet that declares a protection, a data validation
or a merged range takes the complete scan and reaches the guards. Its comment
said this was because the validator refused those containers; the comment is
corrected to say what is now true.

**No `unsafe`, no new dependency, no new ambient I/O, no relocated limit.** The
only new per-element work is one `memchr` per attribute inside `<sheetData>`;
the per-attribute name matching the old allow-lists did is removed.

**The row-visibility editor gets the same rule, not the same answers.** Until
this change the shared vocabulary enforced that module's documented refusals
for it by refusing every out-of-`sheetData` element. Applying the dependency
rule to *its* edit — which rows are hidden — moves its admission in both
directions. A protected sheet forbids the edit; an autofilter, a sort state and
a custom sheet view each carry hidden-row state in exactly the `hidden`
attribute that module owns; markup-compatibility content makes its preprocessed
and lexical views disagree. Those four are refused by its own scan now, because
nothing else refuses them. A worksheet relationship is the other direction: its
documentation claimed to refuse one, but the rewrite copies every byte except
the row tags and a hyperlink, drawing, comment or printer-settings part is
anchored by address, so a hidden row does not move it. It is admitted, its
documentation is rewritten to state the rule, and the harness gate that pinned
the old refusal now pins the admission. This is decision 7's rule applied to a
different edit; decision 7 widened the *value-only editor's* admission surface,
and this is what that implies for the editor that shares its closure.

**The safer path was taken twice.** A merged range's covered cells could have
been left editable — a value written there is stored, not corrupted, and Excel
writes such cells itself. They are refused instead, because the editor already
refused an *insert* there and a value the caller cannot see is not what the
caller asked for. And `<tableParts>` could have been admitted with only the
header row refused, which needs the table part parsed; the whole construct is
refused instead until the editor models it.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Every measured process pinned with `taskset -c 12` while seven
other agents built and measured on the host.

The census legs were taken against the shared read-only checkout of
`70d7768cc` and are not re-taken at the merged head. Change 0654 changes no
admission gate, and its own finding 5 records that it leaves the value editor's
publication refused; that is cited rather than re-derived, as the coordinator
directed. The timing legs below were re-taken against `ceb0571a8` on both
sides, because changes 0658 and 0664 landed in `litchi-xlsx` and the harness
between the first measurement and this record.

### Admission and publication, counted (deterministic)

The corpus is change 0602's: the 95 `.xlsx` files under `test-data/ooxml/xlsx`
and `test-data/office-interop`, 208 worksheet parts. Three legs were taken with
one probe binary each, all pinned to CPU 12: **before**, the base checkout;
**after-1**, the admission rule alone; **after-2**, the admission rule with the
publication audit. The `test-data` trees were verified byte-identical across
checkouts by SHA-256 before any leg ran.

| | before | after-1 | after-2 |
| --- | ---: | ---: | ---: |
| packages admitted through `edit_many` | **0** of 95 | **10** | 10 |
| packages admitted through `snapshot`/`edit` | **0** of 95 | **3** | 3 |
| worksheet parts admitted | **0** of 208 | **15** | 15 |
| packages completing plan, commit **and publication** | **0** of 95 | **0** | **9** |

Admission is identical in after-1 and after-2, checked package for package
across both doors: **0 differences on all 95**. Twelve fixtures changed verdict
between before and after, every one refused → admitted; **none regressed**.

The one admitted package that still does not publish is
`NewlineInFormulas.xlsx`, and not for a compactness reason: its
`xl/_rels/workbook.xml.rels` is not in `litchi-opc`'s canonical relationship
form, so the overlay is unavailable. That refusal predates this change and is
untouched by it.

What the other 85 do at the multi-sheet door, with the refused relationship
set read structurally from each package's `xl/_rels/workbook.xml.rels` rather
than from the message:

| count | outcome |
| ---: | --- |
| 62 | **refused**: a shared-string relationship |
| 12 | **refused**: a pivot-cache definition *and* a shared-string relationship |
| 9 | not a refusal: the package's first worksheet stores no value cell, so the probe could stage no edit through that door |
| 2 | fail in the XLSX reader before the editor is reached |

A measurement note that outlived three legs and is worth recording: the
*message* of a relationship refusal is run-dependent, because both
`validate_package_relationships` (before) and `validate_workbook_relationships`
(after) walk `Relationships::iter`, a hash-seeded `HashMap`, and name whichever
refused relationship they reach first. Two runs of the same after-2 binary
reported 64/8 and 62/10 for the shared-string/pivot-cache pair — the same 74
packages either way. The table above is the deterministic structural reading;
the packet keeps both, and a future record that wants a stable message here
should sort that walk.

### Republication: what changed in each published package

For each of the nine, the published archive was compared against the source
member by member on uncompressed bytes.

* **9 of 9 changed only the expected members**, zero violations: exactly
  `xl/workbook.xml` and the edited worksheet part. **0 members added, 0
  removed.** Every other member — `docProps`, styles, theme, drawings, charts,
  hyperlink and printer-setting parts, every `_rels` part — is byte-identical.
* **9 of 9 edited worksheets differ in exactly one contiguous range.** On
  `SimpleMultiCell.xlsx`: 1,079 → 1,084 bytes, a common prefix of 446 and a
  common suffix of 632, with `1` becoming `424242` between them.
* **9 of 9 re-open.** Each published archive was reopened with
  `SourceBackedWorkbook::open`, the edited cell read back as the value written,
  and the worksheet's stored-extent cell count compared with the source's:
  identical on all nine, so no neighbouring cell was lost or created.
* **5 of 5 republish byte-identically** across two separate processes.

Three things this evidence does **not** establish, stated rather than implied.
No published package carries an `xl/calcChain.xml`, so the chain-removal branch
is exercised by the unit witness and not by the corpus — the only admitted
package that has one is the package that fails on non-canonical relationships.
Four of the nine show the writer re-ordering the *edited* cell's attributes
(`r="A1" s="1"` → `s="1" r="A1"`) and two of those dropping an explicit
`t="n"`: `write_cell`'s pre-existing behaviour on the cell it rewrites, in code
this change does not touch, semantically identical, and visible for the first
time only because real packages are now admitted. And the corpus is this
corpus: nothing is claimed for producer files outside it.

### Instructions, counted (deterministic)

Callgrind isolation pairs on change 0601's producer-shaped corpora, `--samples
2` against `42`, differenced and divided by 40, so the open, the process start
and the harness cancel and the residue is one operation. Both legs at
`ceb0571a8`; the after binary is this branch's commit `3e95c7fa2`.

| selector | before Ir/op | after Ir/op | ratio |
| --- | ---: | ---: | ---: |
| `xlsx_producer_medium_source_planning` | 16,312,015 | 16,217,854 | **0.99423** (−0.577%) |
| `xlsx_producer_medium_source_one_edit_save` | 38,859,858 | 38,874,462 | **1.00038** (+0.038%) |

Inclusive Ir per operation of the changed module, `cell_values::validation`:
13,353,931 → 13,258,587 on planning (**−0.71%**), 18,725,254 → 18,788,653 on
one-edit-save (**+0.34%**). The vocabulary is doing slightly less work per
element than the allow-lists it replaced — one `matches!` over (parent, name)
pairs and one attribute pass, where there were a name match, a parent match and
a per-attribute name match — and the difference is under a percent either way.

The publication change was attributed separately. Before, the paired sites
audited the source with `verify_source` (2,904,051 Ir/op) and the replacement
with `verify_authored` (2,949,780). After, both go through `verify_source`
(5,808,154 together) and `verify_authored` leaves the profile entirely. The
paired sites are **45,677 Ir/op cheaper**, and **100,172 Ir/op** summed over
every audit context — real, and **0.26%** of the 38.86M a whole one-edit-save
costs. It very nearly cancels the validation module's growth over the same
interval, which is why the whole operation lands at +0.038%.

### Paired timing

20 warm-ups then 30 samples per leg, order A1 B1 B2 A2, every process pinned to
CPU 12, four windows, both legs built `--release --locked` and staged outside
their target directories. Two rounds were run. **Round A is the one quoted**:
`tools/perf-baseline` is byte-identical in both its legs, so the only
difference between them is this change's production code. Round B put the
harness census fix in the after leg only and is reported as a control.

A per-leg admission rule was declared before the analysis and is the reason
these numbers are worth more than this record's first attempt: **a leg counts
only if its p95/p50 is at most 1.05**, and a paired delta counts only when both
its legs are clean. An undisturbed leg on this host sits at about 1.01. The
rule exists because an A/A floor taken at the end of a window can come back
clean while a leg *earlier* in the same window was hit — one observed window
held a source leg at +92% beside floor legs that looked fine. Thirteen of
Round B's 96 legs were excluded this way, all from one window.

Round A, mean paired p50 delta over the clean pairs, in both directions
(negative means the after leg is faster). **A/A floor: 2.31%**, over 37 clean
same-binary comparisons in the same windows.

| selector | pairs | A→B | B→A | against the floor |
| --- | ---: | ---: | ---: | --- |
| `xlsx_producer_medium_source_planning` | 6 | **−3.19%** | +3.30% | beats it |
| `xlsx_producer_dense_source_planning` | 5 | **−2.40%** | +2.48% | beats it |
| `xlsx_producer_medium_source_one_edit_save` | 6 | **−3.40%** | +3.52% | beats it |
| `xlsx_producer_dense_source_one_edit_save` | 7 | **−2.61%** | +2.69% | beats it |
| `xlsx_producer_medium_control_planning` | 7 | −0.25% | +0.26% | flat, a ninth of it |
| `xlsx_producer_dense_control_planning` | 8 | −0.58% | +0.61% | flat, a quarter of it |

**No scenario got worse by more than 5% at p50 in either round.** Round B, with
a floor of 3.50%, keeps the sign on all four source selectors (−1.33% to
−2.22%) but drifts its controls to +2.55% and +1.17%, which is the layout
artefact expected when only one leg's harness changes; everything there is
inside its wider floor.

What this measurement does and does not say. The four source selectors improve
past the floor while the two **marker-free controls** — the same grids with the
producer signature removed, which exercise the same publication path and the
same writers but a much smaller validator surface — stay at a quarter of it or
less. That contrast is what distinguishes a real effect from binary layout, and
it is why the result is reported at all: an earlier window at a different base
moved the controls as much as the subjects, and nothing was claimed from it.
The effect does **not** appear in the instruction counts (−0.577% and +0.038%),
so it is cache and branch behaviour the counts cannot see. It is reported as
evidence and **not registered as a claim**: the deltas are two to four times a
2.31% floor on a host shared with seven other agents, which is enough to
report and not enough to register.

## Correctness evidence

* **23 new witness tests** in `crates/litchi-xlsx/src/cell_values/admission_tests.rs`,
  one per construct of the rule above, each building the smallest package that
  carries exactly that construct: compatibility markers and `x14ac:dyDescent`;
  eleven out-of-`sheetData` children at once; a foreign `<extLst>` payload; a
  relationship-bearing worksheet whose `r:id` survives; `docProps` and an
  external-link workbook relationship; unfamiliar workbook children and defined
  names; inline rich text that reads back concatenated with its `<rPr>` copied
  verbatim; the merged-range anchor admitted and its covered cell refused; array
  and shared formula members refused; the calculation chain dropped; shared
  strings, pivot caches, tables, query tables and external relationships refused
  by name; cell metadata refused at its new point; a planted `r:id` inside
  `<sheetData>` refused; an unknown element inside a cell record refused;
  misplaced merge markup refused by the raw parser; a protected sheet and a
  covering data validation refusing the edit; a 300-level copied subtree
  refused by the depth bound and an unbound prefix refused inside one; an exact
  no-op reproducing the producer worksheet byte for byte; and a published
  producer-shaped package whose unnamed members are transferred verbatim.
* **The frozen differential loop is preserved.** `validation_borrow_tests.rs`
  still compares the borrowed traversal against an owned reference
  implementation on every fixture, including a 1,100-prefix truncation sweep;
  the reference now carries the same copied-subtree state, so the two cannot
  diverge on the new rule either.
* **Corpus oracles unchanged.** `marker_admission_matches_the_authoritative_path_on_every_real_worksheet`
  (change 0603) and `change_0622_oracle_over_test_data_worksheets` (the facts
  oracle) both run over every real worksheet part and both pass: the fused
  traversal's verdicts and the fact route's agreement with the scan are
  unchanged.
* **The frozen error-order matrix was re-run and repaired in place.** Change
  0541's six properties are unchanged, but five of its fixtures used
  constructs this change admits — `<mergeCells>` after `<sheetData>`, an
  unfamiliar attribute on a `<c>`, `cm`/`vm` — as their *validator error*
  marker. Each is re-pointed at a construct the validator still owns (an
  unmodelled element inside a cell record, a relationship reference inside
  `<sheetData>`) so the property stays tested, and the three refusals that
  moved are pinned at their new point: `cm="1"` at the scalar-cell gate,
  `vm="2147483648"` at the raw parser's Office limit, misplaced merge markup at
  the raw parser. The "late validation error overrides MCE and raw errors"
  fixture keeps the validator error last in the byte stream by moving the
  processing instruction ahead of `<sheetData>`.
* **Publication.** `litchi-xlsx`'s exact-output test for a producer-formatted
  worksheet is inverted: where it asserted an exact `Debug` string for the
  compactness refusal, it now asserts that the package publishes, that the
  producer's `\n  ` indentation survives verbatim, that an unedited neighbour
  cell is byte-identical and that the edited value is present. Beside it, a new
  test asserts the replacement audit still fails closed on an unterminated
  element, two document elements and a DOCTYPE, with an empty archive each
  time. In `litchi-opc`, `ordinary_authored_xml_keeps_the_compactness_gate`
  becomes `the_authored_classifier_still_separates_compact_from_formatted_xml`:
  the classifier's three verdicts are unchanged, the formatted replacement now
  publishes, and the malformed one is still refused with no archive.
* **Gates.** `cargo fmt --all --check`; `cargo clippy -p litchi-xlsx -p
  litchi-opc -p xml-minifier --all-targets`, clean; `cargo test -p litchi-xlsx`,
  every target; `cargo test --workspace --lib --bins --tests`, **1,055 test
  targets, no failures**; `cargo doc -p litchi-xlsx -p litchi-opc --no-deps`;
  `cargo test -p litchi --features docx,xlsx,pptx,xls`; `python3
  tools/non_iwork_gate.py verify`. Tails in `results/change-0657/gates.txt`.
  Two notes recorded there rather than hidden. `cargo test --workspace` without
  `--lib --bins --tests` additionally builds examples and fails on a
  pre-existing type mismatch in `crates/litchi-iwa/examples/`, which
  `git diff ceb0571a8 -- crates/litchi-iwa` shows is untouched by this change.
  And `tools/perf-baseline`'s own suite ran at **526 passed, 1 failed** before
  this change's harness fixes — the single failure being the row-visibility
  semantic gate this change turns from a refusal into an admission — and both
  groups this change touches pass after them (`xlsx_row_visibility_matched_controls`
  1 of 1, `producer_shape` 10 of 10); a complete re-run of that 527-test suite
  was still in progress at commit time on a host running eight agents.

## Validation preserved

Nothing was weakened. Every traversal ADR 0005 mandates still runs, complete,
over the complete output; the malformed-XML, depth, dialect, DOCTYPE,
duplicate-attribute and unbound-namespace refusals are unchanged; no limit
moved; no defence was removed. What changed is the *vocabulary* the validator
admits, and every newly admitted element is one the rewrite already copied
verbatim and never interpreted. Where the rule could not prove that — inside
`<sheetData>`, and for the value-dependent constructs — it refuses, with a
typed error, before any partial result reaches a caller.

## Limitations

* **Nine packages of ninety-five is the whole population that saves today.**
  The other 86 are refused at the editor door, overwhelmingly on shared
  strings. Nothing here says a shipped file of any other shape round-trips.
* **The tenth admitted package still does not publish**, on a non-canonical
  relationship member. That is a `litchi-opc` question this change does not
  touch and does not design.
* **Shared strings are still refused**, which is 66 of the 95. Admitting them
  is 0602's D2 and needs its own record: the table has to be captured, bounded
  and lazily materialized, index identity has to become an invariant, and
  0602's F3 (materialization costing more than the readback it enables) has to
  be measured on a large table. Decision 7 does not authorize it.
* **Pivot caches, tables and query tables are refused wholesale**, not
  per-address. A table whose header row is untouched by the edit could in
  principle be admitted; that needs the table part modelled.
* The single-worksheet requirement on the `edit`/`snapshot` door (0602's D5) is
  untouched.
* The admission counts are of this corpus on this host. No claim is made about
  producer files outside it.
* **Compactness is no longer enforced at publication on the source-backed
  route.** It remains a property of this library's serializers, asserted by
  their own tests, and this record does not add a new test that would catch a
  future writer emitting non-compact XML into a source-backed replacement. The
  eager `PackageWriter` route is change 0665's and is untouched.
* No claim is registered. The paired timings are descriptive, taken beside an
  A/A floor in a window shared with seven other agents.

## Retained evidence

[`results/change-0657/`](results/change-0657/README.md).
