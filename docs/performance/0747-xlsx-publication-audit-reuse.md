# 0747: a replaced XML Part's replacement is proved from its original's audit outside the one element an edit replaced

Status: retained, implemented in `xml-minifier` and `litchi-opc`.
`performance_claim: none` — the paired medians, instruction counts and unit
timings below are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `009d515bef`; branch `perf/0747-xlsx-publication-audit-reuse`; candidate
commits `1ccfe6b354` and `6b49ce999a`, and a review follow-up commit (see
*Review follow-up*) that changes no release code path. The coordinator's task: eliminate XML
audit work in source-backed XLSX cell edit/save publication that re-proves a
property already proven within the same operation over the identical bytes,
without weakening the audit of any published byte.

**Result.** The audit of a replaced Part's *original* bytes has no duplicate in
the operation and stays a complete scan; six witness tests show planning does
not subsume it. The audit of the *replacement* re-scanned, byte for byte and in
the same parser state, what the original's audit had just accepted: the
measured worksheet pair differs in one byte of 63,294 (medium) and 462,568
(dense-sparse). A new `xml_minifier::audit::verify_source_replacement` returns
exactly the verdict and error of the two `verify_source` calls it replaces and
scans the replacement again only inside the one element the edit replaced.
Measured with both legs built alike, the one-cell source-backed edit/save
is 6.63% faster on the medium corpus (95% CI [−7.39%, −5.55%]) and 6.32% on
dense-sparse ([−6.83%, −5.84%]). Its publication phase is 15.5% faster on both,
and its XML audits execute 47.6% / 47.8% fewer instructions. Every output byte
is unchanged, and no control case's median paired change is adverse by more
than 0.36%.

## What was changed

* `crates/xml-minifier/src/audit.rs`
  * The slice auditor's token loop moves, unchanged, from `verify_with_policy`
    into `scan`, driven by `verify_observed` with a read-only `Observer`.
    `verify`, `verify_authored` and `verify_source` pass the unit observer,
    which answers "not listening" at compile time, so their path is the old one.
  * New public `verify_source_replacement(original, replacement, limits) ->
    Result<ReplacementProof, ReplacementError>`, `ReplacementProof`
    (`Identical`, `Window { original, replacement }`, `Complete`) and
    `ReplacementError` (`Original(Error)`, `Replacement(Error)`,
    `into_error()`), all `#[non_exhaustive]`.
  * Private `WindowSearch` (the observer that finds the window while the
    original is audited), `Window::prove`, `audit_window`, `State::within`,
    `Counters`, `LastToken` and block-wise `common_prefix_len` /
    `common_suffix_len`.
* `crates/xml-minifier/src/lib.rs`: one sentence of module documentation.
* `crates/litchi-opc/src/source_backed.rs`: the private
  `validate_source_part_xml` becomes `validate_source_part_replacement_xml`,
  which audits the pair with one call and maps either side's error to the
  unchanged `OpcError::XmlPublication { part, source }`. All seven paired
  sites that 0657 put on the source policy call it: the topology writer the
  XLSX editor uses, the single-Part overlay, the single- and multi-Part
  overlays with external relationship removals (Part and relationship
  pairs), and the multi-Part overlay.
* Tests: `crates/xml-minifier/tests/source_replacement.rs` (14 after the
  review follow-up), `crates/litchi-opc/tests/source_replacement_audit.rs` (4),
  `crates/litchi-xlsx/src/cell_values/publication_audit_tests.rs` (6).

**Breaking changes: none.** The `xml-minifier` API is additive; the
`litchi-opc` and `litchi-xlsx` changes are private or test-only. No refusal is
added or removed, no error value or offset changes, and no published byte
changes: every output digest in every measured process is identical between the
two legs.

## Authority

ADR 0006 requires publication to audit every XML member it is about to write
for encoding, well-formedness, one document element, no DTD and its finite
budgets; the replacement's verdict is still established for every byte it
publishes, so ADR 0006's text does not move. Change 0652 decision 2 kept the
structural checks of the original bytes when it removed compactness from them;
they are kept here, as a complete scan. 0652's standing trade-offs 2 and 3 set
the shape: the safer path is unchanged (every failure takes the complete scan),
and the benign common case, a local edit of a valid part, is the one made
cheaper. 0705 rejected dropping an audit because readback passed; nothing here
relies on readback or on the writer. ADR 0005 (bounded, measured) and ADR 0003
(atomic, fail-closed publication) are unaffected: the new state is at most one
entry per open element, bounded by the depth limit, and a refusal still emits
no byte.

## Every audit on this publication path

The measured route is `SourceBackedEditor::publish_multi_commit_to_stream`,
which builds a `SourceTopologyPlan` and calls
`SourceBackedPackage::write_topology_to_stream`. For a one-cell edit the plan
replaces two Parts: the edited worksheet, and the workbook, whose `<calcPr>`
the commit invalidates. The gdb capture of one measured iteration (below) sees
exactly four `verify_source` calls, in this order: workbook original (424
bytes), workbook replacement (527), worksheet original (63,294 medium /
462,568 dense-sparse), worksheet replacement (same lengths; one byte differs,
`<v>0</v>` becoming `<v>1</v>`). The one-percent route replaces the workbook and
all four worksheets: ten calls, as 0705 counted.

| # | check | over which bytes | policy and limits | what relies on it |
| --- | --- | --- | --- | --- |
| 1 | `read_part` of each replaced Part | the original member's compressed record | ZIP framing, declared size, CRC, `ReadLimits` | the exact-no-op comparison and the XML audit read these decoded bytes; the writer re-reads raw records itself |
| 2 | exact no-op comparison | original vs replacement | byte equality | an equal Part is not published as changed and is never audited (0654's *an exact no-op still precedes both audits*) |
| 3 | **audit of the original** | the replaced Part's decoded original bytes | `verify_source`, `Limits::default()` (32 MiB, depth 256, 1,000,000 events, 250,000 attributes, 4 MiB token, 16 MiB text) | nothing downstream reads them as XML: they are not published. The audit is a fail-closed refusal of a malformed source Part; change 0652 decision 2 kept its structural checks when it removed compactness from it |
| 4 | **audit of the replacement** | the bytes about to be written | same | ADR 0006: publication audits every XML member it is about to write; the published archive's readers rely on it |
| 5 | authored audits of generated relationship XML and added Parts | library-authored bytes | `verify_authored` | not reached by the measured corpus: removing a calculation chain regenerates the workbook's relationship member, and the corpus has no chain |
| 6 | topology checks | the plan and the catalog | Part-name, relationship, signature, encryption, limit rules | the writer's placement |
| 7 | preservation writer | raw ZIP records | preservation index, trailing bytes, one canonical member per replaced Part, source-unchanged sink fence | the emitted archive |

Rows 3 and 4 are the XML audits. Both ran on every replaced XML Part; 0654
measured them as 27.41% and 27.86% of publication instructions, and 0705 as
54.56% / 58.13% together.

Planning and commit run different checks over the same bytes: the value-only
validator and the raw worksheet parser over the original worksheet (one shared
traversal when the worksheet is at most 8 MiB and 131,072 events), the same
validator and the catalog parser over the original workbook, and the
specialised readback parser over the replacement. They prove what the value
edit needs, not the publication contract.

## The original audit is not a duplicate of any planning check

Nothing in planning or commit calls `xml_minifier` at all: `litchi-xlsx` does
not depend on it. So the question is whether planning's own checks are equal to
or stronger than `verify_source` over the identical bytes. They are not, and
the answer does not depend on reading code: each row marked *no* has a
constructed witness in
`crates/litchi-xlsx/src/cell_values/publication_audit_tests.rs`, an input that
planning and the commit admit and publication refuses, through the public
multi-sheet door the harness uses, with the exact diagnostic and an empty sink.

| `verify_source` check | planning's corresponding check | equal or stronger? | witness |
| --- | --- | --- | --- |
| whole-payload UTF-8 | worksheet: `source_stream_admission` / `parse` decode the whole part | worksheets only, except markup-compatibility-rewritten fallbacks | — |
| well-formed, matched end names | quick-xml `NsReader`, `check_end_names` | yes | — |
| one document element | raw worksheet parser | yes | — |
| no character data outside the root | validator refuses non-whitespace text there | yes | — |
| no CDATA outside the root | validator admits whitespace-only CDATA there | **no** | worksheet and workbook tests |
| no DTD or DOCTYPE | validator | yes | — |
| attribute grammar (source policy) | quick-xml's checked iterator accepts a missing separator | **no** | `an_attribute_with_no_separator_…` |
| `xml:space` is `default` or `preserve` | not read | **no** | `an_invalid_space_value_…` |
| 250,000 attributes | no aggregate count | **no** | `an_attribute_budget_exceeded_only_by_the_replaced_span_…` |
| 4 MiB token, 16 MiB text, 32 MiB part | no equal bounds | **no** | (not constructed: multi-MiB fixtures) |
| depth 256, 1,000,000 events | parser bounds happen to be equal | yes on the shared traversal | — |

Two of the witnesses are refused **only** by the audit of the original: the
defect sits in the span the rewrite replaces (the edited cell's `<v>`), so the
replacement is valid and republishes when it is itself someone's original.
Eliding the original audit would publish those two inputs. Change 0705 rejected
dropping an audit because readback passed, and that stays rejected: the audit
of the original bytes remains a complete scan.

The existing typed proof route is no substitute either. A replacement that
carries a `SourceXmlPart` skips both audits, relying on `validate_source_xml`,
which checks against `ReadLimits` and does not read `xml:space` or count
aggregate attributes, text or token bytes. Moving the XLSX editor onto it would
change the audit of published bytes, which this task forbids.

What would make a proof for the original available is described under *What
remains* below.

## The replacement audit stays, and the writer is not trusted

`an_insertion_can_push_an_admitted_original_over_the_attribute_budget` builds a
worksheet exactly at the 250,000-attribute budget. A count-neutral edit of it
publishes; inserting one cell adds one `c/@r` and only the replacement audit
refuses the result. The value writer can therefore emit bytes the audit
rejects from an original the audit accepts, so no proof by construction
exists and the replacement's verdict must still be established.

## The duplicate that does exist

The replacement audit re-proves, over identical bytes, what the original audit
has just proved. The captured measured worksheet pair differs in **one** byte of
63,294 (medium) and 462,568 (dense-sparse). The replacement's audit scanned the
other 63,293 / 462,567 bytes a second time, in the same parser state, moments
after the original's audit had accepted them.

`xml_minifier::audit::verify_source_replacement(original, replacement,
limits)` returns exactly the verdict of `verify_source(original)` followed by
`verify_source(replacement)`: the original's error if it fails, otherwise the
replacement's error if it fails, otherwise `Ok`, and an error value identical
to the one `verify_source` returns, offset included. It works in three steps.

1. **Locate.** The common prefix and suffix of the two payloads are found by
   block comparison. The complete audit of the original runs through the same
   token loop as `verify_source`, with a read-only observer that watches
   elements close, innermost first. The first element that is not the document
   element, starts at or before the first differing byte, ends at or after the
   start of the common tail, and whose replacement bytes begin with `<` and end
   with `>`, is the *window*: `original[a..b]` became `replacement[a..b+Δ]`.
   Once it is found the observer stops listening and the rest of the original's
   audit runs as `verify_source` runs; when the differing bytes alone exceed
   half the replacement, no window could be used and the search never starts.
2. **Scan the window.** The replacement bytes in the window are scanned by the
   same token loop, starting from the state the original's audit had at `a`:
   `d` elements open (`d ≥ 1`, the document element among them), and no
   enclosing element closable from inside. The window must end at depth `d`,
   and its last token must be markup ending at its last byte.
3. **Re-total.** Events, attributes and character-data bytes are the
   original's totals less what its element was charged plus what the window was
   charged, and each must be within its limit; the replacement's length must be
   within the byte limit.

If any step fails, or the window is more than half the replacement, the
replacement is audited completely, so every refusal is `verify_source`'s own.
The shortcut only ever decides *passes*.

Why the verdict is the same. Outside the window the replacement **is** the
original's bytes: `replacement[..a] == original[..a]` and
`replacement[a+|W|..] == original[b..]` hold by the prefix and suffix
comparison, which is computed from the bytes, not assumed. quick-xml 0.41's
slice reader starts every markup token at its `<` and never consumes it when
ending text (`read_text` returns `UpToMarkup` with the cursor on `<`), so a
token boundary before `<` or after a markup token's `>` is a boundary in every
payload with the same bytes on both sides of it. The window begins with `<`,
so the token before it ends where it did; it ends with a markup token, so the
token after it starts where it did. At a token boundary the reader's only
state that matters is its stack of open names for end-tag checks, and a
balanced window leaves that stack as it found it; the initial state of a fresh
reader over the window is equivalent for input that starts with `<`. The
source policy's per-token checks depend on the auditor's state only through
the depth (character context, CDATA placement, the depth limit) and the
aggregate counters; the `xml:space` stack and whitespace-run state serve
compactness checks that the source policy does not make. Every token outside
the window is therefore the same token, checked from the same depth, with the
same result, and the counters, which only grow, are within their limits at
every step if and only if the final totals are. UTF-8 composes because the
window starts and ends at an ASCII delimiter. The document-element and
root-count checks see a balanced window inside the root.

`ReplacementProof` reports which path decided (`Identical`, `Window`,
`Complete`), with offsets only. The identical-payload case is itself a
duplicate this change removes: the multi-Part overlay door with external
relationship removals audited a byte-equal payload twice before comparing it.

Every one of the seven paired audit sites in `litchi-opc` now calls one helper,
`validate_source_part_replacement_xml`, which maps either side's error to the
unchanged `OpcError::XmlPublication { part, source }`.

## Review follow-up

An adversarial review of the first two commits (reported by the coordinator:
about 595 million differential cases, no mismatch) asked for one fix and three
record corrections. They are applied in one further commit, without history
rewrite:

* **The invariant is now stated where it can be broken.** The window proof is
  sound only because no source-policy check reads a token's offset or ordinal,
  whether another token was seen, the `xml:space` scope, the root count away
  from depth zero, or other state of the enclosing elements. `scan`'s doc comment
  now lists exactly what a check may read, and the obligation on any new check:
  be replayed by the window, or disable window proofs. `Policy::SOURCE` points
  to it. `State::within` sets every field explicitly, each with its reason, so
  a new field forces a decision; it no longer fills fields from `Self::new()`.
* **Debug builds re-derive every window proof.** `debug_check_window_proof`,
  compiled only under `debug_assertions`, asserts that `verify_source` accepts
  the replacement whenever a window proof is returned. Every debug-build test
  that reaches a window, in `xml-minifier`, `litchi-opc` and `litchi-xlsx`,
  therefore re-checks the equivalence. Release builds do not contain it, so the
  measured binaries' code path is unchanged and no timing is repeated.
* **The reviewer's counterexample is now a test, and it is caught.** A new
  test compares the pair with the complete audit on position-sensitive tokens
  placed inside a window: declarations, an `xml-stylesheet` instruction, a
  comment, a byte-order mark, CDATA, a DOCTYPE and a closing-and-reopening
  root. As a mutation check, the reviewer's scratch rule (refuse a declaration
  whose offset is not the start of the input) was applied temporarily. The test
  then failed in the debug cross-check with `window proof Window { original:
  3..7, replacement: 3..24 } contradicts the complete source audit of the
  replacement: Err(Malformed { offset: 3, … })`. The mutation was reverted;
  the diff hash before and after the check is identical.
* **The differential evidence is restated** by window proofs rather than cases,
  with the generator's shape limits. The generator now also makes two
  separated edits. See *Correctness evidence*.
* **The worst case of the fallback is measured and stated** under *What is not
  claimed*. The docx round-4 figure is corrected to +129.24% / +126.95% /
  +129.19% at p50 / p95 / mean. The auditor's pre-existing gaps are listed
  under *What remains*, not fixed.

## Measured

Host AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`, shared with other agents;
every measured process pinned to CPU 24 with `taskset`, and no build of this
record ran while a timed process ran. Harness `tools/perf-baseline`, unchanged
by this record; corpora are its deterministic generators.

**The before leg was rebuilt.** A first campaign used the coordinator's
prebuilt base binary (`fb535ebb…`) against the first candidate (`52528d5b…`,
commit `1ccfe6b354`). It showed the same headline effect, but also 2.7–3.4%
shifts on control paths this change does not touch
(`xlsx_ordinary_save_lifecycle` −2.70% [−2.91%, −2.19%], eager dense-sparse
−3.32%). Both binaries report the same rustc 1.95.0 and the same Cargo profile
and target fingerprints; they differ in source path and therefore in code
layout. The before leg was then built from a detached base worktree with the
identical command and flags (`0578e443…`) and measured against the final
candidate (`5c63a831…`, commit `6b49ce999a`). With both legs built alike the
controls are flat. The final campaign is the one reported; the preliminary one
is kept in the packet, unedited.

### Timing: final campaign

Four rounds; in each, every case ran A B B A, so eight processes per leg,
paired within the round (A1 with B2, A4 with B3). Per-process p50, p95 and mean
are in the packet; the table shows the median of the process p50s and p95s, the
median paired p50 change, and a percentile bootstrap 95% interval over the
eight paired changes (20,000 resamples, seed 747).

| case | processes × samples | before p50 ms | after p50 ms | paired p50 change | 95% CI | p95 ms, before → after | output |
| --- | --- | ---: | ---: | ---: | --- | --- | --- |
| `xlsx_source_backed_cell_values_one_edit_save` medium | 8+8 × 30 | 4.448 | 4.165 | **−6.63%** | [−7.39%, −5.55%] | 4.549 → 4.217 | identical |
| `xlsx_source_backed_cell_values_one_edit_save` dense-sparse | 8+8 × 20 | 28.783 | 27.005 | **−6.32%** | [−6.83%, −5.84%] | 29.102 → 27.278 | identical |
| `xlsx_source_backed_cell_values_one_percent_edit_save` medium | 8+8 × 20 | 16.560 | 16.528 | −0.66% | [−0.73%, −0.07%] | 16.713 → 16.725 | identical |
| `xlsx_source_backed_cell_values_one_percent_edit_save` dense-sparse | 8+8 × 20 | 32.319 | 32.166 | −0.64% | [−0.91%, +0.47%] | 32.659 → 32.406 | identical |
| `xlsx_eager_cell_values_one_edit_save` medium | 8+8 × 30 | 6.078 | 6.181 | −0.16% | [−7.11%, +10.15%] | 6.199 → 6.276 | identical |
| `xlsx_eager_cell_values_one_edit_save` dense-sparse | 8+8 × 20 | 37.737 | 37.346 | −0.73% | [−4.66%, +0.64%] | 38.378 → 37.674 | identical |
| `xlsx_ordinary_save_lifecycle` | 8+8 × 20 | 6.303 | 6.303 | −0.00% | [−0.82%, +1.40%] | 6.449 → 6.405 | identical |
| `docx_source_backed_one_edit_save` | 8+8 × 40 | 1.823 | 1.787 | −2.57% | [−57.56%, −2.13%] | 1.869 → 1.826 | identical |
| `pptx_source_backed_one_edit_save` | 8+8 × 40 | 6.703 | 6.777 | +0.36% | [+0.07%, +1.84%] | 6.853 → 6.910 | identical |
| follow-up: eager medium | 16+16 × 40 | 6.180 | 6.078 | −1.49% | [−6.58%, −0.06%] | 6.274 → 6.127 | identical |
| follow-up: pptx | 16+16 × 40 | 6.681 | 6.716 | +0.31% | [−0.73%, +0.71%] | 7.010 → 6.933 | identical |

"Identical" means one output SHA-256 per case across all processes of both
legs: the change moves no published byte.

The harness splits the source-backed interval into phases (medians of the
per-process medians, ms):

| case | open | planning | commit | publication |
| --- | --- | --- | --- | --- |
| one-edit medium | 0.098 → 0.097 | 1.556 → 1.576 | 0.818 → 0.813 | **1.984 → 1.677 (−15.5%)** |
| one-edit dense-sparse | 0.108 → 0.111 | 10.376 → 10.601 | 5.253 → 5.234 | **13.042 → 11.025 (−15.5%)** |
| one-percent medium | 0.089 → 0.092 | 6.054 → 6.104 | 3.591 → 3.550 | 6.838 → 6.768 |
| one-percent dense-sparse | 0.104 → 0.099 | 11.395 → 11.637 | 6.756 → 6.734 | 13.886 → 13.653 |

Publication saves 0.307 ms and 2.017 ms per one-cell edit. The unit probe
below predicts 0.264 ms and 1.855 ms for the audit pair in isolation, where
the replacement's bytes are warm in cache. Planning and commit code is
unchanged. Their 1–2% moves, in both directions across cases, are within what
this host shows for unchanged code (see *Regressions*).

The one-percent route edits cells scattered over each worksheet, so the
differing bytes alone span more than half of each part. No window is searched,
and both sides are audited completely, as before. Its publication phase is
unchanged within noise.

### Regressions and every over-5% flag

No case's median paired change is adverse by more than 0.36%, and no interval
that excludes zero is adverse by more than 1.84% (pptx, +0.36% [+0.07%, +1.84%];
the 16-process follow-up gives +0.31% [−0.73%, +0.71%]). The medians of the
process p50s move adversely by at most 1.7% (eager medium, 6.078 → 6.181 ms,
whose paired change is −0.16% and whose follow-up is −1.49%). All 72 paired
comparisons over 5% (p50, p95 or mean) in the final campaign are in
`timing/final/flags.json`:

* 60 are favourable. 46 are the one-edit source-backed rows: all 48 of their
  comparisons except dense-sparse round 1's second pair at p50 (−4.69%) and
  mean (−4.91%), which improve by less than 5%. 14 are on control paths, in
  the same noise as the adverse ones below: six eager medium, six docx pairs
  whose before process ran in the slow mode, and two eager dense-sparse.
* **12 are adverse, all on control paths this change does not reach.** Nine
  are `xlsx_eager_cell_values_one_edit_save` medium, +6.60% to +12.45% at
  p50/p95/mean in three pairs. The same case also shows −6.85% to −9.95% in two
  other pairs. Its per-process p50s span 5.71–6.36 ms before and 5.62–6.45 ms
  after, overlapping, and the 16-process follow-up gives −1.49% [−6.58%,
  −0.06%]. The follow-up's own 35 flags (15 adverse) are this case and pptx, in
  both directions. The other three are `docx_source_backed_one_edit_save`
  round 4, +129.24% at p50, +126.95% at p95 and +129.19% at mean, where the
  after process ran in the slow mode
  this case shows in both legs: three processes near 4.1–4.3 ms (two before,
  one after) against 1.76–1.84 ms for the rest. That bimodality is also why its
  interval is so wide.
* The ~0.1 ms open phase moves between −5.0% and +3.7% across the four
  source-backed rows, code unchanged.

In the preliminary campaign every case's median paired change was favourable
(its layout offset), and its 72 over-5% comparisons (11 adverse) are in
`timing/preliminary/flags.json`. Its phase medians include an open-phase +6.4%
(one-percent medium) and −11.7% (one-edit medium), with the same code.

### Unit probe: the audits on the captured bytes

`gdb` stopped the coordinator's prebuilt base binary at every
`verify_with_policy` call of one
`--samples 1 --warmup 0` run of each shape. Its backtraces identify the four
calls of the measured publication (and the other 43 calls, all in corpus
construction, lifecycle gates and the eager expected-output oracle). It dumped
the exact bytes each audited. The packet keeps the backtraces and each
payload's SHA-256 and length, not the bytes. A probe built twice from the same
source file with the same flags, once against base `xml-minifier` (`a0a9fe5e…`)
and once against the final candidate (`a130a033…`), timed each audit on those
bytes: 41 batches of about 20 ms per measure, eight processes per build, A B B A
over four rounds, CPU 24. Medians of the process p50s:

| worksheet | two `verify_source` calls | one `verify_source_replacement` | change | one `verify_source`, before → after |
| --- | ---: | ---: | ---: | --- |
| medium, 63,294 bytes | 547,969 ns | 283,822 ns | −48.2% | 278,704 → 272,291 ns |
| dense-sparse, 462,568 bytes | 3,833,706 ns | 1,978,763 ns | −48.4% | 1,943,792 → 1,912,584 ns |

Both pair columns are timed inside the candidate probe, the base one having
no pair API, so they share one code layout. The pair now costs one audit plus
the block comparison of the common suffix.
Standalone `verify_source` moves −2.3% / −1.6%. The probe's tokenizer-only
loop, which neither build changes, moves +3.7% to +5.8%, so that is layout, not
a claim. The workbook pair (424 → 527 bytes) differs inside the document
element's end tag, so no window exists. It is audited completely, and costs
about 70 ns more than two calls (the observer and the comparison), under a
microsecond in all.

A first probe build (`5751be27…`, commit `1ccfe6b354`) called the observer on
every token even after its window was found, and cost 6.7% more than one
audit on dense-sparse. Commit `6b49ce999a` gates the observer; both runs are
in the packet.

### Instructions: callgrind isolation pairs

Isolation pairs as in 0654. Each self-built binary ran each case under
callgrind at `--samples 1` and `--samples 3` with no warmup, pinned to CPU 24.
Half the difference is one measured iteration, which cancels corpus
construction, the lifecycle gates and the expected-output oracle. Inclusive
instructions per measured iteration, from the retained summaries:

| function | medium before | medium after | dense-sparse before | dense-sparse after |
| --- | ---: | ---: | ---: | ---: |
| publication (`publish_multi_commit_to_stream`) | 31,235,589 | 25,043,733 (−19.8%) | 158,856,374 | 114,644,023 (−27.8%) |
| XML audits | 13,060,524 (4 calls) | 6,837,322 (2 calls, −47.6%) | 92,098,713 (4 calls) | 48,068,488 (2 calls, −47.8%) |
| preservation writer | 17,699,770 | 17,729,488 | 65,881,400 | 65,699,362 |
| planning (`edit_sheets`) | 27,384,719 | 28,024,422 | 187,912,108 | 187,607,756 |
| commit | 13,938,691 | 13,945,284 | 94,864,278 | 94,867,208 |

"XML audits" is `verify_with_policy` before and `verify_source_replacement`
after. The latter includes the workbook pair's complete replacement audit, one
`verify_with_policy` call of 21,500 / 21,558 instructions. Medium planning's
+2.3% is run-to-run variation of code this record does not change: the
coordinator's prebuilt base measured 27,996,979 for it. The eager control
`xlsx_eager_cell_values_one_edit_save` dense-sparse never calls the pair API. It
executes 1,096,157,133 → 1,096,131,257 instructions per iteration (−0.002%),
and its two slice audits 46,051,909 → 45,853,940 (−0.43%).

### Correctness evidence beyond the tests

The seeded differential test (`the_pair_verdict_equals_two_complete_audits_on_generated_edits`)
requires the pair's verdict, failing side and error value to equal those of
the two `verify_source` calls. It generates small documents (a few hundred
bytes to a few KiB) holding a random element tree at most five levels deep over
eight names, with attributes, text, references, comments, CDATA and
processing instructions. About 1–3% of tokens are malformed or refused, and 5%
of documents start with a byte-order mark. Limits are the defaults half the
time, and otherwise have each budget narrowed to within two of what one side
needs. The generator has no namespaces beyond one prefix, no multi-KiB tokens,
no real-producer layouts and no sizes near the default budgets. Two campaign
generations ran:

| generator | cases | refused on the original | window proofs | two separated edits (their window proofs) |
| --- | ---: | ---: | ---: | ---: |
| one edit per case (the campaigns first reported here: 8 seeds × 5,000,000 on `1ccfe6b354`, 8 × 3,000,000 on `6b49ce999a`) | 64,000,000 | about 36% (the review's replay of the default seed: 35.9%; 3.2% identical) | 13,242,171 | none: this generator never makes them |
| extended in the review follow-up: in a quarter of the cases a second, separated edit, half of them inside the first edit's enclosing element (8 seeds × 3,000,000) | 24,000,000 | 8,489,897 (35.4%) | 4,510,930 | 3,746,101 (214,761) |

The extended campaign also refused 6,266,432 replacements, found 619,092
identical pairs and scanned 4,113,649 replacements completely. **No verdict,
side or error value ever differed**, in either generation or in the 40,000-case
suite run by every `cargo test`. The window path, not the case count, is what
this evidence covers: 17,753,101 window proofs, of which 214,761 cover two
separated edits.

The coordinator also reports an independent adversarial review. It ran about
595 million differential cases: a 25-minute mutation fuzz over 4,144 real OOXML
parts, plus exhaustive 4-token and single-byte-pair campaigns. It found no
mismatch and reproduced the packet's numbers. That evidence is the reviewer's
and is not in this packet.

## What remains

* **The audit of the original is still a complete scan.** A proof for it would
  have to come from planning, which already tokenizes the identical allocation.
  That needs three things this record does not build: `xml_minifier` exposing
  its source-audit state machine to an external token stream with raw spans; the
  raw worksheet parser's shared traversal (and the workbook validator) driving
  it; and an unforgeable token, constructible only by that audit, that binds
  the `PartData` allocation, its length, the policy and the limits, which
  `SourceTopologyPlan` would carry and publication would honour only for the
  same allocation (or equal bytes) and fall back otherwise. Its prize is bounded
  by what fusing would share, the tokenizer: 0.140 ms of the 0.279 ms medium
  audit and 0.981 ms of 1.944 ms dense-sparse (the final probe's base leg),
  about 3% of the one-edit case.
  It restructures the parser record 0744 is changing concurrently, and was not
  attempted.
* **Multi-edit commits get no window.** The one-percent and batch routes edit
  cells scattered over a worksheet, so the differing bytes span more than half
  the part, no window is searched, and the replacement is audited completely,
  as before. That is 0705's larger share: 54.56% / 58.13% of one-percent
  publication instructions are the audits, half of them replacements. A
  writer-declared list of windows, each proved by byte comparison and scanned
  as above, would extend the proof to them; the value writer already knows its
  spans.
* **Pre-existing gaps in the source audit itself, not fixed here.** The
  adversarial review found inputs the base `verify_source` accepts although
  they are not well-formed XML. The scratch probe confirms each is accepted,
  and that the pair agrees with the complete audit on each:
  `<r/>& ;` (character data outside the root, read as a reference), `<r>&;</r>`,
  `<r>&a b;</r>`, `<r>&#xZZ;&#0;</r>`, a declaration inside the root
  (`<r><?xml version="1.0"?></r>`), `<r>]]></r>`, and an undeclared namespace
  prefix (`<p:r/>`). The coordinator is queuing them as a separate correctness
  item. Fixing some of them is exactly the kind of check the window invariant
  forbids: "a declaration must come first" reads a token's position, and the
  reviewer's version of it produced a false window pass. Any such fix must be
  replayed by the window or must disable window proofs, as the obligation at
  the auditor's token loop now states.

## What is not claimed

* No performance claim is registered; `performance_claim: none`. The numbers
  above are evidence for this host, these synthetic corpora and these cases.
* Only an edit confined to one non-root element benefits: the one-cell route.
  The one-percent route scatters its edits over `<sheetData>`, finds no usable
  window, and is unchanged within noise (−0.66%, −0.64%); the batch route has
  the same shape and was not measured. The
  workbook's `<calcPr>` insertion touches the document element's end tag and
  never has a window. Real producer packages were not measured.
* The unit-probe timings are warm-cache timings of the audit alone. The
  publication phase saved slightly more than they predict (0.307 vs 0.264 ms,
  2.017 vs 1.855 ms); no cause is assigned.
* The docx and pptx controls use other routes for their main Parts. The docx
  median −2.57% sits inside a bimodal process distribution and is not claimed;
  the pptx +0.36% (follow-up +0.31% [−0.73%, +0.71%]) is below every review
  trigger and is reported, not explained.
* No allocation-count, peak-memory or RSS measurement was taken. The new state
  is a stack with at most one entry per open element, bounded by the depth
  limit (256), plus a few counters. Its allocation failure disables the search,
  it does not fail the audit.
* One environmental difference is disclosed rather than hidden: a replacement
  whose complete audit would fail only because the allocator could not grow its
  depth stack is not scanned completely on the window path, so that failure,
  which says nothing about the bytes, is not reproduced there.
* The equivalence argument is stated above, and the generated differential
  campaigns (17,753,101 window proofs among 88,000,000 cases) found no
  counterexample. That is evidence, not a machine proof. The
  fall-back-on-any-doubt structure bounds the risk to a false *passes*, which
  is exactly what the differential oracle checks, and which debug builds now
  re-check on every window proof (see *Review follow-up*).
* **The fallback has a worst case, and it is slower than before.** When a
  window is found but a check made only on the window path then fails, the
  replacement is scanned about one and a half times, plus the byte comparison.
  A scratch probe measured constructed ~64 KiB documents whose window is
  just under half the document, four processes on CPU 24. A valid replacement
  whose window ends with character data containing `>` costs 1.290× the two
  audits (median; range 1.248–1.311×). An invalid replacement whose defect sits
  at the end of such a window costs 1.288× (1.247–1.301×). The same document's
  one-element edit costs 0.540× (0.506–0.553×). Only invalid or unusual
  replacements pay this, which is 0652 trade-off 3: the malicious minority
  may pay more. No real input measured here takes that path.
* The first campaign's 2.7–3.4% control shifts are attributed to the two
  builds' code layout, on the evidence that same-flag builds remove them. The
  mechanism was not isolated further.

## Verification

Tests added: `crates/xml-minifier/tests/source_replacement.rs` (13: the one-cell
window, a length-changing window, identical payloads, original-first precedence,
seven in-window defects and an invalid-UTF-8 window with exact offsets,
tampering outside the edit, document-element and between-children edits,
window widening at text boundaries, sibling insertion, the half-size rule, a
BOM-marked original, every budget re-totalled at and one past its limit, and
the seeded differential test); `crates/litchi-opc/tests/source_replacement_audit.rs`
(4, each through the topology, single-overlay and multi-overlay doors: a local
edit republishes every other member byte-exactly, in-window defects and a
distant substituted defect are refused with the complete audit's error and an
empty sink, and the original's error precedes the replacement's); and the six
witnesses in `crates/litchi-xlsx/src/cell_values/publication_audit_tests.rs`.

Gates, run in the worktree on the committed candidate `6b49ce999a` with the
record and packet in place (commands, test totals and exit codes in
[`gates.txt`](results/change-0747/gates.txt)); every one exits 0:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo check -p xml-minifier -p litchi-opc -p litchi-xlsx -p litchi-docx -p litchi-pptx -p litchi-xlsb -p litchi-ppt -p litchi-ooxml-common -p litchi-spreadsheet-drawing -p litchi --all-targets --locked` | clean |
| `cargo check -p litchi-odf-common -p litchi-odt -p litchi-odp -p litchi-imgconv --all-targets --locked` (compile-only: these also use the auditor; ODF is not otherwise touched) | clean |
| `cargo clippy -p xml-minifier -p litchi-opc -p litchi-xlsx --lib --no-deps --locked -- -D warnings`, and `--all-targets` | clean, both |
| `cargo test -p xml-minifier -p litchi-opc -p litchi-xlsx --locked` | 2,191 passed, 0 failed, 2 ignored |
| `cargo test -p litchi-docx -p litchi-pptx -p litchi-xlsb -p litchi-ooxml-common -p litchi-spreadsheet-drawing --locked` | 3,536 passed, 0 failed, 46 ignored |
| `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked` | 382 passed, 0 failed, 7 ignored |
| `RUSTDOCFLAGS="-D warnings" cargo doc -p xml-minifier -p litchi-opc -p litchi-xlsx --no-deps --locked` | clean |
| `python3 tools/check_crate_boundaries.py` | pass |
| `python3 tools/non_iwork_gate.py verify` | pass |
| `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` | 10 claims validated |

`cargo check --all-targets` reports three dead-code warnings in the facade's
`unexpected_format` test target under the default features. That file is not
touched by this record (the facade has no diff), and the warnings are
pre-existing.

The review follow-up was gated again with a fresh `CARGO_TARGET_DIR`
(`targets/0747`, deleted afterwards), with the results in
[`gates-review.txt`](results/change-0747/gates-review.txt). `cargo fmt --all --check`, `cargo clippy -p xml-minifier
-p litchi-opc -p litchi-xlsx --lib --no-deps --locked -- -D warnings` and the
same with `--all-targets`, `cargo test -p xml-minifier -p litchi-opc -p
litchi-xlsx --locked` (2,192 passed, 0 failed, 2 ignored, with the debug
cross-check active) and `RUSTDOCFLAGS="-D warnings" cargo doc -p xml-minifier -p
litchi-opc -p litchi-xlsx --no-deps --locked` all exit 0. The harness is unchanged, so its own test suite and the coverage
validator were not required. No shared log, registry or coverage index was edited; the
ready-to-paste sections are in
[`log-sections.md`](results/change-0747/log-sections.md).

## Cleanup

After the evidence was copied into the packet and the gates passed, the
candidate target directory (`targets/0747`, 76.3 GB), the before leg's target
directory (`targets/0747-before`, 1.05 GB), the detached base worktree
(`0747-before-src`, 10.2 GB; `git worktree remove --force`, then `prune`) and
the scratch directory (224 MB) were deleted. Each is verified absent in
[`cleanup.json`](results/change-0747/cleanup.json). Binary identities are in
[`binaries.txt`](results/change-0747/binaries.txt). Raw callgrind profiles, the
captured payload bytes and binary copies are not kept; their summaries and
hashes are. The worktree and branch are kept.

For the review follow-up a fresh `targets/0747` and scratch directory were
created and used for the gates, the extended campaign and the worst-case probe.
Both were deleted again afterwards (17.4 GB and 252 KB), and that second
cleanup is recorded in the same file.
