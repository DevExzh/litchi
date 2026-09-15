# 0603: the fused XLSX traversal admits declaration-only compatibility markers, and completes on no real worksheet

Status: retained, partially implemented. `performance_claim: none` — the counts
and paired medians below are reported as evidence, not registered as a claim.
The implemented subset is **value-identical**: on every worksheet part of every
real `.xlsx` fixture, and on twelve adversarial synthetics, the admitted path
returns the same `Store` or the same typed error as the authoritative
validate-then-preprocess-then-parse path. The two widenings survey item XML-3
also asked for — worksheets that *use* `x14ac:dyDescent`, and worksheets that
carry an `mc:Ignorable` directive — are **designed and then withheld**, because
both require hiding from the value-only validator the attributes it is
contracted to refuse, which moves where a refusal happens.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements rank 13 of change [0587](0587-remaining-opportunity-survey.md)
(item **XML-3**, "the fused planning traversal of 0546 excludes every real Excel
worksheet"). The mechanism the survey named is real and the fix is worth more
than the survey estimated — up to **−61% of a whole plan-and-commit operation**
and **−60% p50** — but only on the derived shapes change
[0602](0602-xlsx-real-producer-admission-design.md) had to invent, because the
traversal completes on **0 of 207** real worksheet parts before this change and
**0 of 207** after it.

## What was changed

`crates/litchi-xlsx/src/raw/worksheet/mod.rs`

* `source_stream_eligible` is replaced by `source_stream_admission`, which
  returns *how* the markup-compatibility preprocessor would treat the source
  rather than only whether the source is marker-free:
  * `SourceAdmission::Borrowed` — the MCE namespace does not occur, so
    `process_ooxml` returns `Cow::Borrowed` and the fallback's parser sees
    exactly these bytes. This is change 0546's existing admission.
  * `SourceAdmission::Rewritten` — the MCE namespace occurs, so the
    preprocessor re-tokenizes and re-emits the part. Newly admitted.
  The `!contains(content, x14ac::NAMESPACE)` gate is removed outright: with no
  MCE namespace the preprocessor is the identity, so an x14ac *declaration*
  cannot change a single parser event. The `dyDescent` and `AlternateContent`
  gates and the 8 MiB, UTF-8 and MCE input/output gates are unchanged.
* A new `MceRewriteEquivalence` runs, event by event, only in the `Rewritten`
  mode. It carries the proof that the rewrite would deliver the parser the same
  events the source does, and any event it cannot prove ends the traversal in
  the established `SourceParseAttempt::ProvisionalFailed`, which repeats the
  authoritative passes. A rejected proof therefore costs one fallback and can
  never change a result.

`crates/litchi-xlsx/src/raw/worksheet/codec.rs` threads the admission into the
shared traversal and consults the proof before the validator observer, so a
source that fails it stops at the offending event.
`cell_values/validation.rs` and `cell_values/snapshot.rs` pass the admission
through. No public API, no limit, no error type and no output byte changed.

## Why it is sound

The fallback validates the **source** (`validation::worksheet_xml(raw)`) and
then parses the **preprocessed** bytes. Admission changes only the second half:
the validator's view is byte-for-byte what it was. So the whole question is
whether the preprocessor's rewrite is invisible to `raw::worksheet`'s parser.

For a part that mentions the MCE namespace but carries no MCE element and no MCE
attribute, the rewrite differs from its input in exactly four ways, and the
worksheet parser observes none of them:

1. **It re-declares every in-scope namespace on every emitted start tag**
   (`mce/codec.rs` `write_start` → `for_each_effective`; this is the emission
   change 0588 designed, measured and then froze as load-bearing). The parser
   resolves names through the reader and never reads an `xmlns` attribute, and
   re-declaring a binding already in scope resolves identically.
2. **It expands `<a/>` into `<a></a>`.** `Parser::transition` answers
   `Event::Empty` with the same `start` and `finish` pair that `Event::Start`
   and `Event::End` run, and `finish(Context::Worksheet)` is `Ok(())`, so an
   empty root agrees too.
3. **It drops character data, CDATA, comments and references outside the root**
   (`visible(&stack)` is false on an empty stack) — in practice the newline
   after the XML declaration. The parser has no text target and no leaf context
   there and ignores all four.
4. **It copies text, CDATA, comments and references inside the root verbatim**
   and normalizes then re-escapes attribute values with exactly the
   `decoded_and_normalized_value(XmlVersion::Explicit1_0, ..)` normalization
   `unqualified_attribute_value` applies when the parser reads one.

What is left are the refusals the rewrite *adds* — inputs a bare `NsReader`
parses and the preprocessor rejects. `MceRewriteEquivalence` refuses admission
for each, so no typed refusal is traded for a result:

| rewrite refusal | proof condition |
| --- | --- |
| `DTD and processing instructions are rejected` | no `Event::PI`, no `Event::DocType` |
| `late XML declaration` | `Event::Decl` only before the root opens |
| `custom entity` | `Event::GeneralRef` only for a character reference or `amp lt gt apos quot` |
| `invalid namespace` | every `xmlns:p` has a non-empty value and an NCName prefix |
| `unbound prefix`, `invalid QName`, `unknown MCE attribute`, `MustUnderstand` | every element and attribute name is an NCName with no prefix, except `xml:` attributes, whose prefix is bound by definition |
| attribute decode failures (`unrecognized entity`, `unterminated entity`, `invalid char ref`) | no `&` in any attribute value — the preprocessor decodes every attribute, the parser only the ones it reads |
| duplicate or malformed attributes | the proof iterates `attributes().with_checks(true)` and refuses the error |
| `output bytes` | a running upper bound on the rewrite's length stays inside `Limits::default().max_output_bytes` |
| `quick_xml` `TooManyDeclarations` on the *rewritten* tags | at most 256 declarations, counted cumulatively so the count bounds the in-scope set |

The remaining preprocessor refusals cannot be reached: `input bytes` (256 MiB)
and `depth` (256) are at or above the traversal's own 8 MiB and
`MAX_XML_DEPTH = 256`; `directive tokens` and `choices` need MCE attributes and
`AlternateContent`, both refused; `multiple roots` and `unterminated XML` make
`Parser::transition` fail, which is already a provisional failure.

`dyDescent` stays refused because the fallback captures those values in a
separate x14ac pass and hands them to the parser as `x14ac::Values`, which the
shared reader does not do; admitting such a part would silently drop the row and
worksheet descents. `AlternateContent` stays refused because the preprocessor
selects one branch and drops the rest. Error precedence (change 0541's guards)
is unchanged in both directions: every failure inside the traversal — reader,
proof, validator observer, parser transition, event cap — returns
`ProvisionalFailed`, and the caller then runs the authoritative validator and
parser in their historical order.

ADR 0005's mandatory validation is untouched: complete value-only validation
still runs over the source on every path. ADR 0006's preservation contract is
untouched: this is a read-side admission decision, nothing is published, and the
preprocessed stream that change 0588 found two writers publish is produced by
exactly the same code on exactly the same inputs as before. No `unsafe`, no
weakened limit, no new ambient I/O, no new public type.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process pinned to CPU 24 while seven other
agents ran. Both legs are the same probe source (change 0602's
`xlsx-admission-probe`) built `--release` against the before checkout
(`6c4c1469b`) and against this branch.

### Admission, counted (deterministic)

`test-data/ooxml/xlsx` plus `test-data/office-interop`: 95 files, **207
worksheet parts** (`census/markers.tsv`).

| | before | after |
| --- | ---: | ---: |
| parts admitted to the shared traversal | 71 | **102** |
| — `Borrowed` (no MCE namespace) | 71 | 71 |
| — `Rewritten` (MCE namespace declared) | 0 | **31** |
| parts refused (all of them for `dyDescent`) | 136 | 105 |
| parts on which the traversal **completes** | **0** | **0** |

Two counts settle the survey's framing. **0 of 207** parts declare the x14ac
namespace without also declaring the MCE namespace, so widening (a) alone — the
item's first half — admits nothing at all; and **0 of 207** parts carry an
`AlternateContent` element, so that gate refuses nothing in this corpus. The
decisive count is the last row: change 0546's fused traversal completes on no
real fixture worksheet before this change and none after it, because every one
of them is refused by the value-only element or attribute allow-list, or exceeds
the 131,072-event provisional cap, long before the marker gates matter. This is
change 0602's finding one layer down: that record measured that the editor
admits no real *package*; this one measures that it admits no real *worksheet*
either.

On change 0602's projections — the largest real-producer geometry the editor can
process, with the producer's `<sheetData>` byte-untouched — **7 of 16** worksheet
parts newly complete the traversal, against 0 before
(`census/admission-census-derived.txt`).

### Instructions (callgrind isolation pairs, N=1 and N=4, M=3, `--separate-callers=1`)

Inclusive Ir for one complete plan-and-commit operation
(`cg/attribution.txt`, `cg/codec-share.txt`):

| fixture | operation before | operation after | Δ | planning Δ |
| --- | ---: | ---: | ---: | ---: |
| `FormatConditionTests` | 11,865,305 | 8,605,157 | **−27.48%** | −39.73% |
| `dataValidationTableRange` | 56,687,954 | 33,595,897 | **−40.73%** | −53.14% |
| `sheet-state-show` | 267,542,005 | 112,075,548 | **−58.11%** | −73.15% |
| `MatrixFormulaEvalTestData` | 96,998,758 | 38,004,903 | **−60.82%** | −81.37% |
| `no_drawing_patriarch` (marker-free control) | 4,106,291,095 | 4,100,975,500 | −0.13% | −0.18% |

**The preprocessing share this removes**, per admitted fixture — inclusive Ir of
`process_markup_compatibility` over the whole operation, planning and commit
together:

| fixture | codec before | codec after | removed | as a share of the operation |
| --- | ---: | ---: | ---: | ---: |
| `FormatConditionTests` | 6,303,903 (53.13%) | 4,140,642 | 2,163,261 | **18.23%** |
| `dataValidationTableRange` | 29,980,454 (52.89%) | 13,921,216 | 16,059,238 | **28.33%** |
| `sheet-state-show` | 138,329,292 (51.70%) | 20,339,394 | 117,989,898 | **44.10%** |
| `MatrixFormulaEvalTestData` | 48,198,051 (49.69%) | 4,597,299 | 43,600,752 | **44.95%** |
| `no_drawing_patriarch` | 166,834,244 (4.06%) | 166,834,244 | 0 | 0 |

The residue is the commit's own preprocessing of the rewritten candidate, which
this change does not touch; the control's figure is *byte-identical* across the
two legs, which is what makes the four differences attributable. This is change
0588's withdrawn rewrite seen from the other side: because the per-element
namespace re-declaration is load-bearing and stays, the only way to stop paying
it on a worksheet is not to run the preprocessor on that worksheet.

### Paired timing (5 warmup, 40 samples per leg, 30 for the largest; order A1 B1 B2 A2)

`bench/summary.txt`. A is the before leg, B the after leg. S1..S4 are four
further before legs, giving the A/A floor in the same window.

| fixture | A1→B1 p50 | A2→B2 p50 | B1→A1 | B2→A2 | A/A floor (worst pair) |
| --- | ---: | ---: | ---: | ---: | ---: |
| `FormatConditionTests` | −28.69% | −27.96% | +40.23% | +38.82% | 4.28% |
| `dataValidationTableRange` | −38.81% | −41.46% | +63.43% | +70.81% | 1.13% |
| `sheet-state-show` | −55.16% | −56.18% | +122.99% | +128.23% | 0.78% |
| `MatrixFormulaEvalTestData` | −60.46% | −59.92% | +152.88% | +149.47% | 3.44% |
| `no_drawing_patriarch` (control) | −1.00% | +0.60% | +1.01% | −0.60% | 1.90% |

The marker-free control takes the identical code path in both legs — its
traversal fails the provisional event cap and falls back before any marker gate
is consulted — and its p50 moves by at most 1.00% in either direction, inside
its own 1.90% A/A floor. Four further interleaved legs (A3 B3 B4 A4,
`bench/legs/ndp-[AB][34].txt`) put the before median at 253.28 ms against
252.61 ms after, **−0.26%** over eight legs. The control's A1 leg carries a
contaminated tail (mean 290 ms, p99 748 ms, against a 256 ms p50) from the
seven other agents active on the host; its p50 is reported and its tail is not.
No scenario regressed above the 5% review trigger, and the floor is quoted per
fixture because it reached 4.28% on one `FormatConditionTests` pair and 3.44%
on one `MatrixFormulaEvalTestData` pair in this window.

## Correctness evidence

Three tests were added to
`crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs`, all driving the
public `SourceBackedEditor` or the two paths directly:

* `marker_admission_matches_the_authoritative_path_on_every_real_worksheet` —
  the differential oracle. For **every worksheet part of every `.xlsx` fixture**
  under `test-data/ooxml/xlsx` it compares the admitted path against
  `validation::worksheet_xml` followed by `raw::worksheet::parse`, and requires
  either the same `Store` or the same exact error message. It asserts that at
  least one part took the borrowed traversal and at least one exercised the new
  rewrite proof, so the oracle cannot pass vacuously.
* `rewrite_only_refusals_survive_marker_admission` — twelve adversarial
  synthetics, one per rewrite-only refusal, compared the same way and reported
  together rather than one at a time. With the proof disabled, **5 of the 12**
  lose a typed refusal — processing instruction (leading and interior), late XML
  declaration, `xmlns:p=""`, and an unrecognized entity inside an attribute
  value the parser never reads. That negative control is retained verbatim in
  `tests/negative-control.txt`; the other seven are refused anyway by the
  value-only validator or by the parser, and the proof covers them in depth.
* `declaration_only_markers_reach_the_shared_traversal` — pins the admission
  verdict for the four shapes (declaration-only MCE → `Rewritten`; x14ac
  declaration alone → `Borrowed`; `dyDescent` → refused; `AlternateContent` →
  refused) and runs a complete no-op value edit through the editor on the two
  admitted ones, asserting the source bytes are unchanged.

Gates, all in the worktree, all passing (`results/change-0603/gates.txt`):
`cargo fmt --all --check`; `cargo clippy -p litchi-xlsx --all-targets`;
`cargo doc -p litchi-xlsx --no-deps`; `cargo test -p litchi-xlsx` —
**1,297 tests, 0 failures**, including change 0541's public planning
error-order guards and every existing MCE, x14ac and `AlternateContent`
worksheet test.

## Validation preserved

Complete value-only validation still runs over the **source** bytes on every
path, before any store is built, exactly as before: the admitted traversal feeds
the same `Validator` the same events the standalone scan would see, and the
fallback runs the standalone scan. No allow-list, no limit, no refusal message
and no error type changed. The traversal's own resource policy — 8 MiB of
shared source, 131,072 provisional events, `MAX_XML_DEPTH`, the MCE input and
output bounds — is unchanged, and the new proof only narrows admission.

## Limitations

* **Nothing is claimed for any real producer file as shipped.** The traversal
  completes on 0 of 207 real worksheet parts before and after; every measured
  figure is from change 0602's projections, whose package envelope was stripped
  to the editor's admission surface. The producer's cell geometry is kept; the
  envelope the editor refuses is not the producer's file.
* The two widenings the survey also asked for are **not implemented**. A
  worksheet that carries `mc:Ignorable` or `x14ac:dyDescent` is refused today by
  `validate_attributes` (`cell_values/validation.rs`, "value-only edits refuse
  attribute '…'") on the source bytes, in the fallback exactly as in the fused
  path — change 0602 met the same gate as its G6. Making the traversal "see the
  events the fallback's parser would after the codec pass" means dropping those
  attributes *before* the validator, which would admit worksheets the editor
  refuses today. That is a widening of the editor's admission surface, not a
  performance change, and it belongs with change 0602's design D1–D4 and its
  litchi-opc prerequisite; it is frozen here, not attempted.
* Instruction counts rank work, not latency. Callgrind prices `rep movsb` per
  byte, so the rewrite's copy share is an upper bound; no SHA-256 runs on this
  path.
* Not measured: allocation counts, peak RSS, cold cache, range sources,
  concurrency, publication cost, any non-`cell_values` consumer (none exists —
  the shared traversal has exactly one call site), and every platform other than
  this host. The p95 and p99 of the control's A1 leg are contaminated by other
  work on the host and are reported but not used.
* The proof is conservative in four places that could be tightened if a corpus
  ever demanded it: any `&` in any attribute value, any prefixed element name,
  more than 256 cumulative namespace declarations, and a sixfold escaping
  allowance in the output bound. Each only sends a source to the fallback.

## Retained evidence

[`results/change-0603/README.md`](results/change-0603/README.md) — the admission
census and its script, the callgrind attribution and codec-share extraction for both legs, all
44 timing legs and their statistics, the fixture-derivation driver and hashes,
the disabled-proof negative control, and the gate tails.
