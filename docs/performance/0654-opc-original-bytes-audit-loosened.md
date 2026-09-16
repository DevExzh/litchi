# 0654: publication stops requiring the *original* bytes of a replaced part to be compact, and 297 of 321 real OOXML packages that it refused now publish with every untouched member byte-identical

Status: retained, a contract change in `litchi-opc` and `xml-minifier`.
`performance_claim: none` — the callgrind isolation pairs, the corpus censuses
and the paired wall-clock medians below are reported as evidence, not
registered as claims. This change removes a refusal; it is not a speedup and
none is claimed.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

Base `70d7768cc`; branch `perf/0654-opc-original-bytes-audit-loosened`. This
record implements decision 2 of change
[0652](0652-owner-decisions-for-the-third-wave.md), which is row 2 of
[0651](0651-queue-refresh-after-the-second-wave.md)'s queue, designed as D0 by
[0602](0602-xlsx-real-producer-admission-design.md) and priced by
[0613](0613-opc-original-audit-memo.md).

## What was changed

Four source files in two crates, one ADR, and two test files.

* `crates/xml-minifier/src/audit.rs` — a private `Policy` value replaces the
  single `reject_ambiguous_space` flag that `verify_with_policy` and
  `verify_reader_with_policy` carried. It has two fields:
  `reject_ambiguous_space`, unchanged, and `require_compact`, new. Three
  constants name the three profiles: `AUTHORED` (the old `true`), `COMPACT`
  (the old `false`) and `SOURCE` (new). `check_start`, `check_end`,
  `check_declaration` and `check_attribute_layout` take the compactness flag
  and return their `Error::NotCompact` verdicts only when it is set. One new
  public function, `verify_source`, runs the slice auditor under
  `Policy::SOURCE`. The streaming auditor builds `require_compact: true`
  unconditionally and is unchanged in behaviour.
* `crates/xml-minifier/src/lib.rs` — six lines of module documentation
  naming the new policy and its scope.
* `crates/litchi-opc/src/source_backed.rs` — one new private helper,
  `validate_source_part_xml`, beside the existing `validate_overlay_xml`. It
  calls `verify_source` and constructs the identical
  `OpcError::XmlPublication { part, source }`. The **seven** call sites that
  audit bytes the package already holds now call it: `7785` (the topology
  path), `8220` (the single-Part overlay door), `8333` and `8336` (the
  single-Part-with-relationship-removals door, part and `.rels`), `8530` (the
  bounded multi-Part `.rels` loop), `8554` (the bounded multi-Part part loop)
  and `8802` (the `write_changed_overlays*` family). Every paired replacement
  site still calls `validate_overlay_xml`, as do the two sites that audit
  litchi's canonical relationship XML (`7284`, `7495`) and the two that audit
  topology additions (`7931`, `7965`).
* `crates/litchi-opc/src/error.rs` — the `XmlPublication` doc comment, which
  said "Authored or changed XML", now states the two contracts separately.
* `docs/adr/0006-validation-security-and-compatibility.md` — two sentences
  added to the `Preserve` paragraph with a dated amendment note; see
  *ADR 0006* below.

Tests: `crates/xml-minifier/tests/audit.rs` gains four tests and
`crates/litchi-opc/tests/source_xml_publication.rs` three. The tracked diff is
**+531 / −60** across the seven files.

## Authority

Change 0652, decision 2, quotes the owner:

> "Original byte contract: loose the audit, accept non-compact XMLs."

and reads it as:

> the publication audit of original part bytes stops refusing an original part
> because its XML is not compact (whitespace, declarations, line endings,
> attribute order); the 94 of 95 real packages it refused become publishable

with four things still to prove: *which checks remain (identity of the
original bytes with the planned bytes stays; structural checks that protect
the archive stay) and which refusals are removed, each with a witness; the
95-package corpus republished byte-identically for untouched parts; ADR 0006's
wording amended where it states compactness.* Each is answered below.

This supersedes the single sentence of change
[0528](changes/0528-xlsx-publication-attribution.md) that had held the
contract in place — *"Both original and replacement OPC audits must stay; XLSX
semantic validation is not a substitute for their different contract"* — and
it supersedes it narrowly: **both audits still run**, on the same bytes, at
the same call sites, with the same error type. Only the *contract* the
original half asserts changes.

### What the audit checked, and why each check existed

The audit entered in commit `5e0a5df71` (change
[0037](changes/0037-opc-source-backed-one-part-publication.md)) in the same
hunk as the replacement audit, with an empty commit body and one boundary
bullet: *"Existing and replacement XML are audited before output when the
selected content type is XML."* No record states why the *existing* half was
included. Its checks fall into three families.

1. **Identity of the original bytes with the planned bytes.** This is not in
   the auditor at all: it is the byte comparison that precedes it
   (`source_backed.rs:7772`, and the exact-payload no-op test at `8201`),
   whose comment says a caller *"may intentionally publish an exact no-op
   replacement for a source whose Part payload is malformed; the contract is
   byte-preserving in that case and must not turn the no-op into a parser
   failure."* **Unchanged**, and pinned by a new test.
2. **Structural properties that protect the published archive.** UTF-8
   decoding; well-formed XML as quick-xml parses it; exactly one document
   element; no character data or CDATA outside it; no DTD or DOCTYPE; the
   attribute grammar (a name, an `=`, a quoted and terminated value); a valid
   `xml:space` value; and the six finite budgets (`Bytes`, `Depth`, `Events`,
   `Attributes`, `TokenBytes`, `TextBytes`) with their bounded allocation
   failure. **All kept**, each with a witness.
3. **Compactness.** The four `xml_minifier::audit::Kind` verdicts:
   `FormattingWhitespace`, `AmbiguousWhitespace`, `AttributeSeparation` and
   `WhitespaceBeforeClose`. These exist to hold *this repository's own output*
   to a byte-minimal spelling — `docs/CRUD_Scenario_Checklist.md:41` states it
   as "**every generated XML part** is byte-minimal", `xml-minifier`'s own
   module doc calls it "the repository's compact XML output contract", and
   0602's D0 put it plainly: *"The audit exists to keep litchi's own authored
   output compact; applying it to bytes litchi did not author refuses every
   real producer."* **Removed for original bytes only.**

### Refusals removed, each with a witness

Every removed refusal is a compactness verdict, and every one has both a
constructed witness in `crates/xml-minifier/tests/audit.rs`
(`source_policy_accepts_every_noncompact_spelling_the_authored_contract_refuses`)
and a real-fixture witness in the corpus census.

| `Kind` removed | constructed witness | corpus witness | members |
| --- | --- | --- | ---: |
| `FormattingWhitespace` | `<?xml …?>\r\n<root>\n  <child/>\n</root>` | the newline a producer writes after the XML declaration — 0602's byte 55 | 6,788 |
| `AmbiguousWhitespace` | `<root> <child/></root>` | `xl/theme/theme.xml` of `tdf66377.xlsx` at byte 148 | 2 |
| `AttributeSeparation` | `<root a="1"\n      b="2"/>`, `<root a = "1"/>` | `_rels/.rels` of `tdf66377.xlsx` at byte 152 | 9 |
| `WhitespaceBeforeClose` | `<root a="1" />`, `<root a="1" ></root >` | `xl/workbook.xml` of `tdf167689_xmlMaps_and_xmlColumnPr.xlsx` at byte 275 | 32 |

Counts are XML members of the 321-package OOXML fixture corpus that the
authored contract refuses and the source policy accepts. Nothing else is
removed: `source_policy_never_reports_a_compactness_violation` asserts that
`verify_source` returns `Error::NotCompact` on no input at all.

### Refusals kept, each with a witness

`source_policy_keeps_every_structural_encoding_doctype_and_limit_refusal`
covers twelve malformed shapes (unclosed root, two document elements,
mismatched end tag, unexpected end element, character data outside the root,
CDATA outside the root, an attribute with no value, an unquoted value, an
unterminated value, an invalid `xml:space`, markup after an end-tag name, and
an empty document), the DOCTYPE refusal, an invalid-UTF-8 refusal, and each of
the six finite budgets narrowed to one below what the input needs — every one
returning the same `Error` variant it returned before, and the budget
refusals naming their own `Resource`.
`original_bytes_keep_every_structural_doctype_encoding_and_limit_refusal`
repeats five of those through the public publication door and asserts that
each still produces `OpcError::XmlPublication { part: "/word/document.xml", … }`
**and emits no archive byte**. On the corpus, 17 of 6,981 XML members stay
refused: 16 with `malformed XML at byte 0: invalid XML declaration boundary`
(the three byte-order-marked packages, below) and one with `XML is not UTF-8
at byte 0` (`customXml/item3.xml` of `alt-chunk-header.docx`). Every one keeps
the identical error text.

## Breaking changes

**None.** `xml_minifier::audit::verify_source` is an added public function;
`Policy` is private; `litchi_opc`'s helper is private; no existing signature,
type, variant or `Display` string changes. The behavioural contract does move,
as decision 2 authorizes, and that movement is what this record is about.

## Why it is sound

**The published bytes for an untouched part are still the source's own
bytes.** This change touches no writer, no plan and no member. Over the 310
packages that publish on the after leg, 6,781 untouched members were compared
byte for byte against their source member: **zero mismatches and zero
member-set changes**. The 13 packages that published on both legs published
the **identical SHA-256** on both.

**The part is still where the plan says, and the archive framing is still
consistent.** Neither the topology plan, the changed-overlay lookup, the
canonical-member check, the trailing-bytes check nor the writer is reached by
the diff; the audit is a pure function of one payload slice and returns before
any of them.

**Error identity is preserved for every refusal that remains.** The same
`OpcError::XmlPublication { part, source }` is constructed by both helpers
from the same `xml_minifier::audit::Error`, and a refusal still precedes any
output: both new `litchi-opc` tests assert `output.is_empty()`.

**The authored contract does not move.** Across all 6,981 XML members of the
321-package corpus, the `verify_authored` verdict is **identical on both
legs — 0 changes**. `ordinary_authored_xml_keeps_the_compactness_gate`, which
predates this change, still refuses a pretty-printed *replacement* and still
emits no archive. `crates/litchi-xlsx`'s
`public_multi_edit_preserves_formatting_whitespace_publication_refusal`, which
asserts an exact `Debug` string for a non-compact worksheet, **still passes
unchanged**: the XLSX value editor's replacement carries the source's
formatting verbatim, so the *replacement* audit now raises the byte-identical
refusal at the same offset that the *original* audit used to raise.

**Exact no-ops stay exact.** `an_exact_no_op_still_precedes_both_audits`
publishes a package whose Part payload is `<document>` — malformed — as an
exact byte no-op and asserts the output equals the source artifact byte for
byte.

**Correctness over performance, twice.** First, the byte-order-mark defect
change [0650](0650-docx-editor-byte-order-mark-admission.md) froze as its
question 3 is **reported, not fixed**: `verify_with_policy` slices
`input.get(start..end)` with `Reader::buffer_position` offsets that do not
count a leading BOM, so a marked document is refused with `Malformed { offset:
0 }`. That refusal is not a compactness verdict, the streaming auditor mirrors
it deliberately and `reader_matches_slice_bom_rejection` pins it, so removing
it is a second contract movement that decision 2 does not authorize. It costs
nothing here: **0 of the 95** packages and **3 of the 321** carry a
byte-order-marked XML member, and none of the three is in 0602's corpus.
Second, the eager writer's own audit is left alone; see *Limitations*.

**ADR 0006** now scopes the contract. Two sentences were added to the
`Preserve` paragraph — the compact output contract binds bytes this library
authors or regenerates and is not asserted against the original bytes of a
part a package already holds, which are audited for encoding,
well-formedness, a single document element, the absence of a DTD or DOCTYPE
and the finite budgets — under a dated note naming 0652 decision 2 and this
record. That paragraph is where the contradiction lived: it already promised
that "lexical details are retained when possible". Nothing else in ADR 0006
states or implies compactness (the file contains no occurrence of "compact",
"whitespace" or "minif"), and line 43's *"malformed known payloads fail before
publication"* is cited in the note as the surviving rule. No other ADR text
changed.

## Measured

Host AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`. Both harness legs built
`--release --locked`; every measured process pinned to CPU 9 with `taskset`.
Binaries timed (staged outside any Cargo target directory):
`44cc4e2875bb20f6fd5c054a967a052ee95b6b6a0ca5063492ffecd2a603da25` (before),
`48ecaf8088f60cb89ee1f8660d4204fe31f2ab372969658f279012752eca0aad` (after).
Deterministic counts first; timing last.

### Corpus counts — what the contract admits (measured)

Every XML member of every fixture, audited with both policies, by the retained
probe. One iteration of the probe is one document.

| | 0602's 95-package corpus | whole OOXML fixture corpus |
| --- | ---: | ---: |
| packages | 95 | 321 |
| XML members | 1,403 | 6,981 |
| members the original-bytes audit accepted **before** | 25 | 133 |
| members it accepts **after** | 1,403 | 6,964 |
| members whose refusal was removed | 1,378 | 6,831 |
| members still refused (identical text) | 0 | 17 |
| packages with every XML member accepted, before → after | 0 → **95** | 0 → **318** |
| `verify_authored` verdicts that moved | **0** | **0** |

### End-to-end publication over the whole corpus (measured)

Each of the 321 fixtures was opened with `SourceBackedPackage::from_path` and
its lexicographically first XML Part replaced through the public
`write_part_overlay_to_stream` door with one fixed compact payload.

| before → after | packages |
| --- | ---: |
| refused → **published** | **297** |
| published → published, **identical output SHA-256** | 13 |
| refused → refused | 9 |
| does not open → does not open | 2 |

* **Untouched-member byte identity:** 6,781 members compared across the 310
  published packages — **0 byte mismatches, 0 member-set changes** — and every
  replaced member carries exactly the payload the probe wrote.
* **Of the 9 still refused, 7 keep the identical error text** (`signed OPC
  source requires an explicit signature edit policy`). **Two report a
  different, pre-existing refusal that the compactness verdict had been
  shadowing**, and this is stated rather than elided:
  `poi/.../bug62513.pptx` now reports *"source ZIP archive has trailing bytes
  outside its located archive"*, and `poi/.../49609.xlsx` now reports
  *"selected Part does not have one canonical UTF-8 source member"* — that
  fixture writes every member with a backslash separator
  (`xl\sharedstrings.xml`) and a lower-cased `[content_types].xml`. Both
  checks are at `source_backed.rs:9708` and `:9803`, strictly after the audit
  pair, untouched by the diff, and computed from `entry.raw_name_bytes()` with
  no reference to XML. The retained witness proves the first at the base: a
  synthetic package whose original bytes **are** compact, so no compactness
  verdict can apply, is refused with the identical message by **both** legs,
  while its clean twin publishes the identical digest on both.

### Instruction shares — 0613's method, reproduced and re-split (measured)

Selector `xlsx_source_backed_cell_values_one_edit_save`; one iteration
publishes two corpus shapes, so one iteration is two source-backed
publications. Callgrind isolation pairs at `--samples 1` and `--samples 3`,
differenced and halved.

| per iteration | before Ir | after Ir | delta |
| --- | ---: | ---: | ---: |
| publication (`write_topology_to_stream`) | 187,808,737.0 | 188,636,450.5 | +827,713.5 (+0.44%) |
| both audits | 103,698,462.5 | 104,258,467.5 | +560,005.0 (+0.54%) |
| — audit of the **original** bytes | — | 51,711,971.0 | |
| — audit of the replacement bytes | — | 52,546,496.5 | |
| whole program (`--samples 1` total) | 24,849,999,153 | 24,861,665,774 | +11,666,621 (+0.047%) |

| share of publication Ir | before | after |
| --- | ---: | ---: |
| audit of the **original** bytes | (not separable) | **27.41%** |
| audit of the replacement bytes | (not separable) | 27.86% |
| both audits | **55.21%** | **55.27%** |

Before the change both halves are one symbol, so only their sum is separable;
after it they are two, which is what makes the split a measurement rather than
a model. The sum reproduces 0613's 55.06%/55.19% on a different base, and the
original half reproduces 0613's 27.59% at 27.41%.

**Call counts are unchanged.** Per iteration the source-backed publication
performs **8** audits before and **8** after — after the change, 4 through
`verify_source` and 4 through `verify_authored`. The eager control,
`PackageWriter::validate_authored_xml`, is 62 calls in all four profiles.

The audit costs **0.54% more instructions**, not fewer. Removing the
compactness verdicts removes branches, not loops; what the profile shows is
the policy plumbing — two extra arguments and a `require_compact` test per
lexical check — plus the loss of specialization now that `verify_with_policy`
has two callers with different policies. It is disclosed rather than tuned
away: a const-generic policy would recover it and is speculative complexity at
0.30% of publication.

### Paired wall clock, both directions, with the floor (measured)

Order A1 B1 B2 A2 A3 A4; 30 samples per leg after 3 warm-up iterations; all on
CPU 9 in one window. Before medians pool A1 and A2, after medians pool B1 and
B2; A3 against A4 is the floor.

| selector / shape | before p50 (ns) | after p50 (ns) | after − before | before − after | A/A floor |
| --- | ---: | ---: | ---: | ---: | ---: |
| `xlsx_source_backed_cell_values_one_edit_save` / `medium` | 4,607,687 | 4,509,021 | −2.14% | +2.19% | 0.03% |
| `xlsx_source_backed_cell_values_one_edit_save` / `dense-sparse` | 29,147,285 | 28,987,459 | −0.55% | +0.55% | 0.85% |
| `xlsx_source_backed_cell_values_batch_edit_save` / `medium` | 16,811,184 | 17,012,772 | **+1.20%** | −1.18% | 0.88% |
| `xlsx_source_backed_cell_values_batch_edit_save` / `dense-sparse` | 32,283,794 | 32,017,198 | −0.83% | +0.83% | 0.53% |
| `opc_source_overlay_one_part_save` / `few-large` | 63,301,025 | 63,008,928 | −0.46% | +0.46% | 0.17% |

Four of five scenarios are faster and one is slower; the slower one, +1.20%
against a 0.88% floor, is **reported, not hidden in a mean**, and is far below
the 5% review trigger. Per-leg p50, mean, p95 and p99 for all six legs on all
five scenarios are in `timing/summary.json`. **All six legs on all five
scenarios produced one output SHA-256 each**, so publication is byte-identical.

## Correctness evidence

* **Seven new tests.** Four in `crates/xml-minifier/tests/audit.rs` (the
  removed-refusal witnesses, the "no compactness verdict ever" assertion, the
  kept-refusal witnesses including every budget and the byte-order mark, and
  an assertion that the authored and default contracts are unchanged) and
  three in `crates/litchi-opc/tests/source_xml_publication.rs` (a non-compact
  original publishes and leaves every other member byte-exact, on three
  spellings; every structural refusal survives through the public door with no
  archive emitted; an exact no-op still precedes both audits).
* **Corpus differential:** 6,981 XML members × 2 policies × 2 legs; 321
  packages published on 2 legs; 6,781 untouched members compared. Counts above.
* **Gates**, all on the candidate in the worktree:
  `cargo fmt --all --check` clean; `cargo clippy -p litchi-opc -p xml-minifier
  --all-targets` clean (workspace lints deny); `cargo test -p litchi-opc -p
  xml-minifier` — **769 passed, 0 failed, 2 ignored**; `cargo doc -p litchi-opc
  -p xml-minifier --no-deps` clean; `cargo test -p litchi-xlsx -p litchi-docx
  -p litchi-pptx` — **3,658 passed, 0 failed, 33 ignored**; `cargo test -p
  litchi --features docx,xlsx,pptx,xls` — **265 passed, 0 failed, 7 ignored**;
  `cargo test` in `tools/perf-baseline` — **540 passed, 0 failed, 1 ignored**,
  with neither of the two flaky allocator tests the briefing names tripping;
  `python3 tools/non_iwork_gate.py verify` clean. Every gate exited 0; tails in
  [`results/change-0654/gates.txt`](results/change-0654/gates.txt).

## Validation preserved

No validation was removed, relocated or weakened other than the compactness
contract on original bytes, which decision 2 authorizes. No `unsafe` (both
crates keep `#![forbid(unsafe_code)]`), no new dependency, no widened limit,
no weakened malformed-input defence, no hidden global pool, no ambient I/O, no
public leakage of an archive type, raw lock or executor. Every finite budget
keeps its value and its `Resource`. Validation still does not mutate: both
helpers take a slice and return a `Report` or an error. Determinism is
untouched, proved by 13 identical output digests across legs and 30 identical
digests across six timing legs on five scenarios.

## Limitations

* **What is not claimed.** No speedup and no regression. The audit is 0.54%
  more instructions and four of five timed scenarios are faster; neither is
  registered.
* **The eager save route is untouched, and still refuses 59 of 180 `.xlsx`
  fixtures.** `PackageWriter::validate_authored_xml` (`pkgwriter.rs:203`) is a
  second audit site with the same contract and a different mechanism: it
  audits a Part when `is_xml_part(…) && !package.is_exact_source_xml(part)`,
  with no original/replacement pair to distinguish. Change 0610's documented
  `xlsx-hide` route (`Workbook::from_bytes` → `edit()` → hide → `commit()` →
  `to_bytes()`) is **identical on both legs**: 33 publish with identical
  digests, 147 refuse with identical text, 59 of them `NotCompact`, 57 naming
  `/xl/worksheets/sheet1.xml`. The retained witness measures why: for
  `dataValidity.xlsx` that Part's blob is **byte-identical to the source
  member** (`623a3d06…`, 1,730 bytes both sides), so the eager writer is
  applying the authored contract to bytes litchi did not author. Fixing it
  needs a provenance signal that site does not have — substituting
  `verify_source` there would stop enforcing compactness on genuinely authored
  parts, which decision 2 does not authorize. **Frozen and reported**, with
  the number.
* **Two shadowed refusals now surface** (above). They are pre-existing and
  their code is untouched; the witness proves the class at the base but only
  for the trailing-bytes check. That the *specific* refusal `49609.xlsx` now
  reports would have fired at the base on a compact original is **modelled**
  from the source, not measured, because no member of that fixture is compact.
* **A byte-order-marked original is still refused**, by the defect 0650 froze.
  3 of 321 fixtures, 0 of 95.
* **Cross-package precompressed transfers are untouched.**
  `try_add_precompressed_part`'s decoded payload (`source_backed.rs:7965`) is a
  donor package's original bytes audited with `verify_authored`. It is an
  addition, not a replacement, on the cross-package copy contract that changes
  0598/0646 own; it is reported, not moved.
* **Instruction counts rank work, not latency** (0579). The 27.41% is an
  instruction share of publication, not a time share of a save, and callgrind's
  software SHA-256 (6.4×, 0649) and per-byte `rep movsb` (35×, 0604) inflate
  the hashing and copying around it.
* **The A/A floor** in this window was 0.03%–0.88% at p50 across the five
  scenarios, below the host's usual 4%; no p99 floor is reported.
* **The corpus is this repository's fixtures.** 0602's 95 packages and the 321
  OOXML fixtures under `test-data`; nothing is claimed about packages outside it.

## Retained evidence

[`results/change-0654/README.md`](results/change-0654/README.md) — the two
corpus censuses and their summaries, the whole-corpus publication differential
and its untouched-member comparison, the `xlsx-hide` differential, the base
witnesses, the four callgrind profiles and the derived arithmetic, the six-leg
timing report, the probe source, every script, the gate tails and the
provenance table.
