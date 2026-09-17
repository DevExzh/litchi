# 0665: the eager package writer stops refusing a member because its XML is not compact, and 58 of 180 real `.xlsx` fixtures plus 53 of 55 DOCX fixtures publish work they could not publish before

Status: retained, a contract change in `litchi-opc` and the gate it unblocks in
`litchi-docx`. `performance_claim: none` — the corpus censuses, the callgrind
isolation pairs and the paired wall-clock medians below are reported as
evidence, not registered as claims. This change removes a refusal and lets an
existing policy reach its default; it is not a speedup and none is claimed.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

Base `5af158123` (the branch after change [0660](0660-docx-compaction-policy.md));
branch `perf/0665-opc-eager-writer-source-provenance`. This implements the
eager-route follow-up to decision 2 of change
[0652](0652-owner-decisions-for-the-third-wave.md), the original-byte movement
that [0654](0654-opc-original-bytes-audit-loosened.md) landed on the
source-backed route and named as its highest open item, and the one function
change 0660 named when it landed. Decision 2 speaks about original part bytes;
this record explicitly interprets the same `verify_source` profile for eager
planned payloads, including source-spliced replacements, so the eager and
source-backed publication boundaries do not disagree. That extension is an
implementation interpretation recorded here, informed by 0654 and the
source-backed replacement precedent in 0657, not an additional owner decision.

## What was changed

Four source files in two crates, one ADR, and the tests that pinned the old
contract.

* `crates/litchi-opc/src/pkgwriter.rs` — `PackageWriter::validate_authored_xml`
  is replaced by two private helpers. **`audit_published_xml`** runs
  `xml_minifier::audit::verify_source` and constructs the identical
  `OpcError::XmlPublication { part, source }`; it is what the part-payload site
  (`:213`) calls. **`audit_authored_xml`** adds a `debug_assert!` that the bytes
  satisfy `verify_authored` and then delegates to it; the **six** sites that
  serialize a member *here* call it — the manifest and the package
  relationships in `PublicationPlan::from_package` (`:192`, `:207`) and in
  `materialize_pristine` (`:248`, `:256`), and each part's relationships
  (`:224`, `:268`). `PlannedPart::authored_xml` becomes `audit_payload` and its
  meaning is documented: whether the plan parses this payload again, not
  whether it is held to a spelling.
* `crates/litchi-opc/src/package.rs` — `is_exact_source_xml` becomes
  **`holds_original_source_xml`** and decides provenance by `Arc::ptr_eq`
  against the payload the package retained at ingress, instead of a pointer
  test with a whole-part byte comparison behind it. A mutation drops the proof,
  because `Part::set_blob` and `set_blob_shared` replace the part's `Arc` and
  `add_part`/`remove_part` drop the retained entry; this is change
  [0647](0647-opc-get-or-add-noop-reuse-design.md)'s "the map changed iff the
  capture was dropped" invariant applied to payloads, and the proof shape
  change [0593](0593-opc-publication-pristine-members.md) gave part
  relationships.
* `crates/litchi-opc/src/error.rs` — the `XmlPublication` doc comment now
  states one contract for every member instead of two.
* `crates/litchi-docx/src/document/transaction.rs` —
  `publication_accepts_preserved_xml` runs `verify_source`, and the
  `carries_character_data_outside_the_root` short-circuit is **removed**: once
  whitespace at depth zero is accepted, that necessary condition is no longer
  necessary and would have sent 53 publishable documents down the fallback.
  This is the single function change 0660 named.
* `docs/adr/0006-validation-security-and-compatibility.md` — the `Preserve`
  paragraph is rewritten with a second dated amendment note; see *ADR 0006*.

Tests: `litchi-opc` gains three
(`publishes_every_noncompact_spelling_the_authored_contract_refuses`,
`every_member_this_writer_authors_is_compact`,
`every_structural_encoding_doctype_and_budget_refusal_survives_with_no_output`)
and two that asserted a compactness refusal now assert a structural one;
`litchi-docx` moves five that pinned the fallback.

## Authority

Change 0652, decision 2, quotes the owner:

> "Original byte contract: loose the audit, accept non-compact XMLs."

0654 implemented it for the *original* half of the source-backed route's
original/replacement pair and stopped there, because the eager route has no
such pair: `pkgwriter.rs` audits a `Part` on
`is_xml_part(..) && !is_exact_source_xml(part)` and cannot tell an untouched
source blob from an authored one. This record therefore adopts the same
`verify_source` profile for every eager planned payload. That includes an
authored or source-spliced replacement, and is an implementation interpretation
of decision 2 rather than a claim that the owner issued a second decision for
those payloads. It follows 0654's structural/source split and the replacement
side of the source-backed route already moved by 0657: compactness is a
quality property of this repository's serializers, not a publication refusal,
while every non-compactness check of the audit stays. The writer's own
manifest and relationship serializations retain `verify_authored` as a debug
quality assertion. The structural half is kept whole and asserted by a new
test because 0652's standing trade-off 2 ("correctness and safety is the
primary consideration") still governs this interpretation.

### The premise this change had to correct first

0654's *Limitations* reported the eager defect as the authored contract being
applied "to bytes litchi did not author", with a witness showing the refused
Part of `dataValidity.xlsx` byte-identical to the source member (1,730 bytes,
`623a3d06…`). That witness was taken on a **freshly opened** package, not on
the package the route publishes. Measured on the route itself — the retained
diagnostic over all 180 `.xlsx` fixtures — **all 59** refused parts differ from
the bytes they were decoded from: 51 are 16 bytes shorter, 6 are 19 bytes
shorter and 2 are 16 bytes longer, because hiding a tab rewrites
`tabSelected` in the worksheets. They are **spliced**: litchi edited the
source's bytes in place and preserved the source's lexical form around the
edit, and the audit then charged the whole part for the producer's newline
after the XML declaration (`FormattingWhitespace` at byte 55 in 57 of the 59).
Not one of the 59 was an untouched part, and an untouched part is not audited
on this route at all — so the provenance signal alone could not have moved a
single one of them. What moves them is dropping the compactness contract.

## Breaking changes

**No public item was added, removed or re-signed.** Both audit helpers are
private, `holds_original_source_xml` is crate-private, and
`carries_character_data_outside_the_root` was private. What breaks is
behaviour, and it breaks in two places:

| item | before | after |
| --- | --- | --- |
| `litchi_opc::PackageWriter::{write, write_to_stream, to_bytes}` | refused a package with `OpcError::XmlPublication { source: NotCompact(..) }` when any XML member the plan audited was not byte-minimal | never raises `NotCompact`; every other refusal of that error is unchanged, with the same variant, part name and text |
| `litchi_docx::document::Edit::commit` under the default `CompactionPolicy::PreserveUnmodified` | fell back to whole-document compaction on 54 of 55 openable DOCX fixtures, because the gate asked `verify_authored` | preserves on 54 of 55; **the published main-document bytes change** on 25 of the 26 fixtures that admit the one-edit route and on 53 of 63 on the insert route |
| `litchi_opc::Part` implementors | a payload replaced with bytes *equal* to the source's escaped the publication audit | the audit is decided by allocation provenance, so a replaced payload is audited even when its bytes match; on the corpus this changes nothing (measured below) |

The DOCX row is the one a caller sees in its output. It is not a new
capability: it is change 0660's documented default finally reaching the
documents it was written for, and `CompactionPolicy::WholeDocument` still
publishes byte for byte what every committed edit published before 0660.

## Why it is sound

**No defence was removed.** `verify_source` performs every structural,
encoding, DOCTYPE and finite-budget check `verify_authored` performs; only the
four compactness verdicts (`FormattingWhitespace`, `AmbiguousWhitespace`,
`AttributeSeparation`, `WhitespaceBeforeClose`) are dropped, and 0654 already
proved by construction that `verify_source` returns `NotCompact` on no input at
all. The new test
`every_structural_encoding_doctype_and_budget_refusal_survives_with_no_output`
drives eleven refusal families — unclosed root, two document elements,
mismatched end tag, character data outside the root, CDATA outside the root, a
DOCTYPE, an attribute with no value, an invalid `xml:space`, an empty document,
a token over the budget, and invalid UTF-8 — through the public
`PackageWriter::write_to_stream` door, and asserts for each that the refusal is
`OpcError::XmlPublication` naming the part, that it is **not** a compactness
verdict, and that the sink received **zero bytes and zero writes**.

**A refusal still precedes every output byte.** The whole `PublicationPlan` is
built, and every audit in it runs, before a sink sees anything; the diff does
not move a single audit relative to emission.

**Compactness is still checked where it is a quality contract.** The members this
writer serializes itself go through `audit_authored_xml`, whose `debug_assert!`
holds them to `verify_authored` in every debug build and therefore in every
test run — 441 `litchi-opc` unit tests and 960 `litchi-docx` ones exercised it
without firing. `every_member_this_writer_authors_is_compact` asserts it
explicitly on the manifest, the package relationships and a part's
relationships of a package whose *payload* is deliberately non-compact, and
asserts that the package publishes anyway.

**Exact no-ops stay exact.** The eager exact no-op over the whole 321-package
OOXML fixture corpus publishes **319 identical SHA-256 digests** on both legs
and keeps its 2 open refusals with identical text; the DOCX census's
`noop_save` and `exact_noop` aspects are unchanged on all 63 fixtures.

**The provenance signal is now a proof rather than an inspection.** An
equality test cannot distinguish the source's bytes from a caller's copy of
them, so a caller that mutated a part and wrote the source's bytes back escaped
the audit. `Arc::ptr_eq` against the payload retained at ingress cannot be
satisfied that way, because every route that replaces a payload installs a
different allocation. The preservation planner's own `source_blob_retained`
keeps its byte comparison deliberately: there an equal payload publishes equal
bytes either way, and change [0610](0610-opc-lazy-part-decode-design.md) — whose
lazy decode change 0661 is implementing in this wave — depends on that
comparison staying where it is.

**Correctness over performance, twice.** First, the audit is still *skipped*
for a payload the package still holds in the allocation it decoded. Auditing
every XML member instead would make the eager route match the source-backed
one, but it costs a whole-part parse per member per publication and it would
newly refuse 4 of the 321 fixtures — three of them only because of the
byte-order-mark defect change [0650](0650-docx-editor-byte-order-mark-admission.md)
froze. That is a refusal added for benign input, which decision 2 does not
authorize in either direction; it is reported in *Limitations*, not taken.
Second, the DOCX gate keeps its whole-document audit rather than being dropped,
because one fixture in this corpus still needs the fallback.

**ADR 0006** now says that the compact output contract binds this library's own
serializers, is verified by test and debug assertion, and is not a publication
refusal, under a second dated amendment note naming 0652 decision 2 and this
record. The note lists every check that does *not* move and names the one route
that still asserts the compact contract at publication — the source-backed
replacement audit, which change 0657 owns.

## Measured

Host AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`, valgrind 3.26.0;
seven other agents built and measured on the same host throughout. Both harness
legs built `--release --locked`; every measured process pinned to **CPU 20**
with `taskset`. Binaries timed, staged outside any Cargo target directory:
before `cdf9344719613f692765038a0497c75a51cb8122209bafca6f0d2ab1d8784201`,
after `d8fd315842c2ffa379d0747d1f841ad752081ccf2b0942f9dcb82bb25195f953`.
Deterministic counts first; timing last.

### Route outcomes over the corpus (measured)

Four routes, every fixture, both legs, by the retained probe.

| route | fixtures | before → after |
| --- | ---: | --- |
| eager exact no-op (`OpcPackage::from_vec` → `PackageWriter::to_bytes`) | 321 | 319 published → 319 published, **319 identical digests**; 2 open refusals, identical text |
| eager open-edit-save, **compact** authored payload | 321 | 310 published → 310 published, **310 identical digests**; 11 refusals, identical text |
| eager open-edit-save, **non-compact** authored payload | 321 | **0 published → 310 published**; 11 refusals, 9 with identical text and 2 surfacing a shadowed refusal |
| 0610's `xlsx-hide` (`Workbook::from_bytes` → `edit()` → hide → `commit()` → `to_bytes()`) | 180 | **33 published → 91 published**; the 33 keep **identical digests**; 147 refusals → 89 |

**The 0610 counts moved exactly as the brief predicted: 59 `NotCompact` → 0.**
58 of the 59 now publish; the 59th, `poi/.../49609.xlsx`, surfaces a
*pre-existing* refusal the compactness verdict had been shadowing
(`PreservationUnavailable`, "source ZIP framing or opaque members cannot be
preserved after this mutation"). That is measured at the base, not modelled:
on the **before** leg the same fixture is refused with the identical message by
the eager edit-save route with a compact payload, where no compactness verdict
can apply. `poi/.../bug62513.pptx` is the same case on the non-compact
edit-save route, with the same base witness. Residual `NotCompact` refusals on
the after leg, all four routes: **0**.

**Untouched members come back byte for byte.** Over the 91 packages the
`xlsx-hide` route publishes on the after leg, every XML part of the published
artifact was compared with the same part of the source: **1,022 parts
byte-identical, 0 parts added, 0 removed**. The parts that differ are exactly
the ones the route edits — `/xl/workbook.xml` in all 90 changed packages, and
the two worksheets whose `tabSelected` moves in the 57 that now publish.

### The DOCX census, 63 fixtures, eleven aspects, both legs (measured)

Change 0660's probe, re-run unchanged except for one added aspect,
`gate_source`, which reports the verdict the flipped gate reaches.

| aspect | unchanged | moved |
| --- | ---: | ---: |
| `open`, `source`, `gate`, `gate_source`, `noop_save`, `exact_noop`, `one_edit_optin`, `one_edit_reopen` | 63 each | **0** |
| `managed_edit` | 10 | 53 |
| `one_edit` | 38 | 25 |
| `one_edit_span` | 38 | 25 |

The gate's verdict moves from 1 compact / 54 non-compact to **54 accepts / 1
refuses**. The one that still refuses is
`libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx`, whose
main part carries a byte-order mark: `malformed XML at byte 0: invalid XML
declaration boundary`, the defect 0650 froze. It takes the whole-document
fallback and publishes exactly what it published before, which is why the
fallback is kept.

**The preserved-versus-compacted spans.** On the 26 fixtures that admit the
one-edit route, the source bytes the commit rewrites fall from **203,203 to
1,007 (−99.5%)**, and the median rewritten span from **981.5 bytes to 0**.

| fixture | source bytes | source bytes rewritten, before → after |
| --- | ---: | --- |
| `drawing.docx` | 288,070 | 176,552 → **0** |
| `OPCCompliance_CoreProperties_AlternateTimezones.docx` | 6,031 | 2,759 → **0** |
| `alt-chunk.docx` | 2,096 | 2,041 → 425 |
| `NumberingWithOutOfOrderId.docx` | 1,983 | 1,915 → 320 |
| `alt-chunk-html.docx` | 1,453 | 1,415 → 62 |
| `Numbering.docx` | 5,411 | 1,340 → **0** |

`one_edit_reopen` is unchanged on all 63 fixtures: every published artifact
reopens with the identical paragraph, table and block-control counts and the
same number of paragraphs carrying the edit. The published main part grows in
total across those 25 fixtures, from 351,114 to 351,634 bytes against a source
of 351,815 — preservation keeps the producer's whitespace, so it is the honest
direction, and it is disclosed rather than presented as a size win.

### Instruction counts and the audit's share of publication (measured)

Callgrind isolation pairs: profile the harness at `--samples 1` and
`--samples 3` of the same selector, difference the inclusive totals and halve,
so corpus construction, process start-up and the harness's own verification
cancel out. One iteration is one lifecycle of the selector.

| selector | whole iteration Ir, before → after | delta |
| --- | ---: | ---: |
| `docx_semantic_one_edit_save` | 483,906,287 → 482,906,261 | **−0.21%** |
| `xlsx_eager_cell_values_one_edit_save` | 1,717,661,569 → 1,713,208,847 | **−0.26%** |
| `pptx_eager_batch_edit_save` | 9,886,980,793 → 9,887,775,778 | **+0.01%** |

| selector | publication Ir (before → after) | the audit's Ir | the audit's share of publication |
| --- | ---: | ---: | ---: |
| `docx_semantic_one_edit_save` (`write_to_stream`) | 74,796,114 → 74,238,772 | 45,224,412 → 44,665,067 | **60.46% → 60.16%** |
| `xlsx_eager_cell_values_one_edit_save` (`write_to_stream`) | 135,855,044 → 135,027,996 | 52,545,853 → 51,720,746 | **38.68% → 38.30%** |

**The audited set does not move.** Per iteration the publication audit runs
**9 → 9** times on the DOCX selector, **4 → 4** on the XLSX one and
**217 → 217** on the PPTX one. That is the measurement that the provenance
signal's tightening — `Arc::ptr_eq` where a byte comparison used to stand
behind the pointer test — changes which parts are audited on none of these
corpora, and `verify_with_policy` keeps its 12, 4 and 217 calls as well.

The audit is **0.5–1.6% cheaper in instructions**, not free: `verify_source`
skips four lexical checks per token but still parses every byte, so removing
compactness removes branches, not loops. Change 0654 measured the opposite sign
(+0.54%) on the source-backed route because it *added* a policy argument to a
shared function; here the same function loses a caller profile rather than
gaining one.

**The PPTX audit symbol is not separable on this selector** and is reported as
such rather than quoted: its isolation pair is negative (−28.5M Ir for 217
calls), because that selector's corpus construction itself publishes packages
and does not cancel between `--samples 1` and `--samples 3`. Only the
whole-iteration figure is quoted for it.

**`docx_semantic_one_edit_save` cannot show the census's movement**, and this is
stated rather than elided. The harness's DOCX semantic corpus is generated by
this repository and is therefore compact, so change 0660's gate already accepted
it on the base leg — 0660's own counts show its after leg taking the preserving
route with `compact_changed_document_xml` at ~0 Ir. What this change moves on
that selector is only which policy the gate's own audit runs under. The
movement this record is about is on **real producer fixtures**, and the census
above is where it is measured.

### Paired wall clock, both directions, with the floor (measured)

Order A1 B1 B2 A2 A3 A4 in one window on CPU 20; `--warmup 5 --samples 40` per
run, so 80 samples per pooled leg. Before medians pool A1 and A2, after medians
pool B1 and B2; A3 against A4 is the A/A floor and B1 against B2 the B/B floor.

| selector / shape | before p50 (ns) | after p50 (ns) | after − before | before − after | A/A floor | B/B floor |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_one_edit_save` / `large` | 8,799,090 | 8,786,426 | −0.14% | +0.14% | −1.30% | +0.88% |
| `docx_semantic_one_edit_save` / `medium` | 230,061 | 226,042 | −1.75% | +1.78% | −1.89% | −0.65% |
| `docx_semantic_one_edit_save` / `tiny` | 71,880 | 71,161 | −1.00% | +1.01% | −0.62% | −1.16% |
| `xlsx_eager_cell_values_one_edit_save` / `dense-sparse` | 40,539,280 | 40,762,189 | **+0.55%** | −0.55% | −0.28% | +0.25% |
| `xlsx_eager_cell_values_one_edit_save` / `medium` | 6,680,362 | 6,767,365 | **+1.30%** | −1.29% | −1.82% | +1.33% |
| `pptx_eager_batch_edit_save` / `media-rich` | 320,147,179 | 324,055,429 | **+1.22%** | −1.21% | +0.04% | +0.01% |

Three scenarios are faster and three slower; every one of the six is smaller in
magnitude than the 5% review trigger, and four of the six are smaller than the
A/A floor measured beside them. **The PPTX +1.22% is not**: that window's floor
was 0.04%, so the regression is real at p50 and is reported rather than hidden
in a mean. It is the one scenario whose publication audits 217 members per
iteration, and its whole-iteration instruction count moved +0.01%, so the wall
clock is not tracking removed or added work; the record does not claim to have
explained it. Per-leg p50, mean, p95, p99, min and max for all six legs are in
`timing/summary.json`.

**All six legs on all three selectors produced exactly one output SHA-256 per
scenario**, so publication is byte-identical across the legs the harness
digests.

## Correctness evidence

* **Three new `litchi-opc` tests.**
  `every_structural_encoding_doctype_and_budget_refusal_survives_with_no_output`
  drives eleven refusal families through the public streaming door and asserts
  the typed refusal, that it is not a compactness verdict, and a sink that
  received zero bytes and zero writes.
  `publishes_every_noncompact_spelling_the_authored_contract_refuses` publishes
  each of change 0654's four compactness spellings and reopens the archive to
  assert the member came back byte for byte.
  `every_member_this_writer_authors_is_compact` audits the manifest, the
  package relationships and a part's relationships with `verify_authored` and
  asserts the package publishes even though its payload is not compact.
* **Two moved `litchi-opc` tests.**
  `refuses_malformed_authored_xml_bytes_before_publication` (was
  `refuses_arbitrary_authored_xml_bytes_before_publication`) and
  `publication_plan_failure_leaves_sequential_sink_untouched` now use an
  unclosed root rather than an ambiguous-whitespace payload, so they assert the
  refusal that survives.
  `source_xml_provenance_requires_the_ingress_allocation` proves that a
  newly allocated payload is audited even when its bytes equal the source,
  while the untouched ingress allocation retains the proof.
* **Five moved `litchi-docx` tests.**
  `the_gate_accepts_every_noncompact_spelling_and_still_refuses_structure`
  replaces the cheap gate's test: it pins the four accepted spellings, four kept
  structural refusals, and — over the whole DOCX fixture corpus — that the gate
  and `verify_source` agree on every fixture and that at most one fixture still
  takes the fallback.
  `a_source_the_publication_audit_refuses_takes_the_whole_document_route` now
  uses a token over the auditor's `TokenBytes` budget, a refusal that survives.
  `both_policies_agree_across_the_docx_fixture_corpus` asserts the fallback
  population is now **empty** and every fixture preserves.
  `rich_owner_edits_preserve_the_source_spelling_and_stay_durable_and_reversible`
  and `nested_controls_and_tables_are_durable_exact_and_path_checked` assert the
  source's indentation and line breaks **survive** where they asserted the
  commit removed them; both still assert the full retained-markup list, the
  durable round trip and the exact inverse.
* **Corpus differentials:** 321 fixtures × 3 eager routes × 2 legs; 180
  fixtures × the `xlsx-hide` route × 2 legs; 1,022 published parts compared
  member by member against their source; 63 DOCX fixtures × 11 aspects × 2
  legs. Counts above; raw outputs in the packet.
* **Gates**, all in the change worktree: `cargo fmt --all --check` clean;
  `cargo clippy -p litchi-opc -p litchi-docx --all-targets` clean (workspace
  lints deny); `cargo test -p litchi-opc -p litchi-docx`; `cargo doc -p
  litchi-opc -p litchi-docx --no-deps` clean; `cargo test -p litchi-xlsx -p
  litchi-pptx`; `cargo test -p litchi --features docx,xlsx,pptx,xls`;
  `cargo test` in `tools/perf-baseline`; `python3 tools/non_iwork_gate.py
  verify` clean. Tails in
  [`results/change-0665/gates.txt`](results/change-0665/gates.txt).

## Validation preserved

No validation was removed, relocated or weakened other than the compact-output
contract at the eager publication boundary, which this record interprets from
decision 2's original-byte movement and the source-backed precedent. No
`unsafe` (both crates keep
`#![forbid(unsafe_code)]`), no new dependency, no widened limit, no weakened
malformed-input defence, no hidden global pool, no ambient I/O, no public
leakage of an archive type, raw lock or executor. Every finite budget keeps its
value and its `Resource`, proved by the budget case of the new refusal test and
by the DOCX fallback witness. Validation still does not mutate: both audit
helpers take a slice and return `()` or a typed error. Determinism is untouched:
319 identical no-op digests, 310 identical edit-save digests, 33 identical
`xlsx-hide` digests, and one output digest per scenario across six timing legs.

## Limitations

* **What is not claimed.** No speedup and no regression is registered. The
  instruction counts move −0.26% to +0.01% and the paired medians −1.75% to
  +1.30%; none is a claim.
* **The PPTX regression is unexplained.** +1.22% at p50 against a 0.04% floor,
  with a +0.01% instruction count. Reported, not chased.
* **An untouched XML part is still not audited on the eager route.** The
  source-backed route audits the same bytes with `verify_source` after change
  0654; this route skips them, as it did before. Closing that asymmetry costs a
  whole-part parse per member per publication and would newly refuse **4 of the
  321 fixtures** — the three byte-order-marked packages, refused only by the
  defect change 0650 froze, and `alt-chunk-header.docx`'s non-UTF-8
  `customXml/item3.xml`. Adding a refusal for benign input is not what decision
  2 authorizes, so the asymmetry is **frozen and reported with its price**, not
  taken. The three BOM witnesses and the managed-transaction residue are the
  retained follow-ups from change
  [0650](0650-docx-editor-byte-order-mark-admission.md), especially
  `results/change-0650/follow-ups.md` and its
  `census/variants-{before,after}.txt`; this change does not silently
  reclassify or fix them. The XML-minifier BOM/offset audit is queued for
  follow-up 0677; the managed DOCX transaction witness remains with 0670.
* **A byte-order-marked main document still takes the DOCX fallback.** 1 of the
  55 openable DOCX fixtures. It is 0650's documented publication-audit
  follow-up, not this change's, and it is why the gate keeps its audit rather
  than being deleted. Its XML-minifier implementation follow-up is 0677. The
  separate managed-transaction witness remains the same typed refusal on both
  legs and is tracked by 0670.
* **Two shadowed refusals now surface** on the eager routes, exactly as change
  0654 found on the source-backed one: `poi/.../49609.xlsx` and
  `poi/.../bug62513.pptx` report `PreservationUnavailable` where the
  compactness verdict used to fire first. Unlike 0654's case this is
  **measured at the base, not modelled**: both fixtures are refused with the
  identical message by the before leg on the eager edit-save route with a
  compact payload, where no compactness verdict can apply.
* **The DOCX published bytes change.** 53 of 63 fixtures on the insert route
  and 25 of 26 on the one-edit route publish different main-document bytes than
  the base. Every one reopens with identical paragraph, table and
  block-control counts, and `CompactionPolicy::WholeDocument` still reproduces
  the base's bytes, but a caller that digests its output will see the change.
* **The source-backed replacement audit is untouched by this branch.** A
  replacement payload is held to the source profile by the 0657 landing on the
  target branch; this branch changed none of those call sites. The eager
  extension is recorded as an interpretation here so a source-spliced payload
  does not regress to the authored compactness refusal.
* **`docx_semantic_one_edit_save` cannot show the movement** (above); the
  harness has no real-producer DOCX corpus.
* **Instruction counts rank work, not latency** (0579). The 38.30% and 60.16%
  are instruction shares of publication, not time shares of a save, and
  callgrind's software SHA-256 (6.4×, 0649) and per-byte `rep movsb` (35×,
  0604) inflate the hashing and copying around them.
* **The corpus is this repository's fixtures**: 321 OOXML packages, 180 of them
  `.xlsx`, and 63 DOCX-family fixtures under `test-data`. Nothing is claimed
  about packages outside it.

## Retained evidence

[`results/change-0665/README.md`](results/change-0665/README.md) — the four
route differentials and their summary, the member-by-member comparison, the
DOCX census on both legs with its diff and summary, the twelve callgrind
profiles and the derived counts, the six-leg timing report, both probes'
sources, every script, the gate tails and the provenance table.
