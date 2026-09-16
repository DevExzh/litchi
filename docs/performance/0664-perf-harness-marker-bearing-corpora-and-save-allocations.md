# 0664: a harness selector now reproduces change 0649's 77× — the real deck's edit is 77.21× the generated one, and 20.23× a byte-identical control that differs only in the codec branch

Status: retained, harness and corpus work. **No file under `crates/` was
modified.** `performance_claim: none` — the counts and paired medians below are
reported as evidence, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Change [0649](0649-pptx-opened-transaction-real-deck-edit.md) found two harness
gaps and left both open. First: every prior PPTX record in this program measured
on generated corpora whose members **never mention the markup-compatibility
namespace**, so `litchi_ooxml_common::mce`'s rewriting branch — which produces
16.25× its input and costs 93.9% of a real deck's opened-transaction edit — was
never priced on a harness selector; change [0601](0601-perf-harness-real-producer-shape.md)
added producer-shaped corpora, but only its XLSX family carries markers across
the package. Second: `litchi-perf-baseline-alloc` emitted **no** allocation
metrics for change [0638](0638-facade-and-ordinary-save-selectors.md)'s
`*_ordinary_save_*` family, because that family was registered without ever
opening an allocation region. This change closes both, and adds the DOCX
text-sink selector change [0643](0643-docx-paragraph-count-and-sink-text.md)
recorded as missing.

## What was changed

Five files under `tools/` (one of them new), two documents and one packet.
Nothing in any production crate.

### A new generator: `tools/perf-baseline/src/marker_shape.rs`

Two families, each in two variants, all generated in memory as a pure function
of (family, variant) — no clock, no PRNG, no ambient state, no file read:

| Corpus | Members | Uncompressed | Marker-bearing members | Marker-bearing bytes | `mc:AlternateContent` | Archive |
|---|---:|---:|---:|---:|---:|---:|
| `pptx-marker-deck-marker` | 63 | 351,108 | 27 | 314,824 (89.66%) | 13 | 49,161 |
| `pptx-marker-deck-control` | 63 | 351,108 | 0 | 0 (0.00%) | 13 | 49,589 |
| `docx-marker-medium-marker` | 12 | 72,956 | 6 | 61,292 (84.01%) | 0 | 10,898 |
| `docx-marker-medium-control` | 12 | 72,956 | 0 | 0 (0.00%) | 0 | 11,019 |

**The control is the same archive with the other branch taken.** It replaces
every occurrence of the markup-compatibility namespace URI with an inert URI of
*exactly the same length* — `http://litchi.invalid/perf-baseline/marker-stripped/0664/xx`,
59 bytes each, `.invalid` being reserved by RFC 2606. Member names, per-member
uncompressed lengths, start-tag counts, attribute counts and the projected text
are identical to the marker variant and are **proved** before any sample runs by
`prove_control_is_byte_comparable`; only the compressed archive differs, because
deflate sees different bytes. Change [0588](0588-mce-codec-namespace-emission.md)
introduced the technique and 0649 used it on the real deck. A marker/control
pair is therefore a measurement of the codec branch and not of two differently
shaped packages.

**The shape is derived, not invented.**
[`results/change-0664/scripts/derive_marker_shape.py`](results/change-0664/scripts/derive_marker_shape.py)
censuses every member of three fixtures that are ordinary tracked files in this
repository (`git ls-files` lists all three; there is no submodule and no LFS
filter) and emits the declaration sets the generator authors. Its `verify` mode
re-derives them and fails if the generator drifts; it is run as a gate.

| Fixture | Members | Uncompressed | Marker-bearing | Share | Root declarations |
|---|---:|---:|---:|---:|---:|
| `sd/qa/unit/data/pptx/slide-section-test.pptx` (0649's deck) | 103 | 796,725 | 43 | 93.03% | 6 |
| `sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx` (real Word) | 21 | 463,565 | 10 | 95.01% | 32 / 15 / 10 |
| `sd/qa/unit/data/pptx/tdf89064.pptx` (the notes parts) | 42 | 68,699 | 17 | 57.66% | 6 |

Three facts the census settles that this program had not recorded:

* the deck's 43 marker-bearing members are **13 slides, 18 layouts, 11 masters
  and `ppt/presentation.xml`** — the 11 themes are *not* marker-bearing and take
  the borrowed path;
* the deck carries **no `mc:Ignorable` anywhere**. The bare `xmlns:mc`
  declaration is what puts a part on the rewriting branch, because both gates
  that matter — the codec's presence scan and `source_stream_eligible` — key on
  the URI occurring anywhere in the bytes;
* every marker-bearing PPTX root declares exactly **six** bindings
  (`a`, `p`, `r`, `p14`, `p15`, `mc`), which is 0649's "6.00 namespace bindings
  per emitted element" read off the fixture rather than off a profile. The
  third fixture supplies the notes parts, because 0649's deck has **no notes
  slides at all**.

**Sizing.** The PPTX shape is the production writer's own package at 13 slides
of 48 text boxes, chosen so its **slide** byte total lands near the fixture's,
because that is what the cost is a function of: 0649 found that all three
per-slide sites of one capture read every slide in full and that the capture
reads no layout and no master at all. Thirteen slides give **247,741** slide
bytes against the fixture's 269,178 (92.0%), where the generated corpus every
earlier PPTX record used has 40,788 (15.2%). Matching the fixture's *package*
total instead would need 130 text boxes a slide; that was built and measured at
**384.56 ms** against this shape's 150.37 ms while both were the same multiple
of their own control, so it buys no signal — this harness's writer emits eleven
small layouts and one master where the fixture has eighteen large layouts and
eleven masters, and no slide/layout ratio reproduces both. The DOCX shape is
`semantic_docx_bytes(Medium)`, the same shape the existing `docx_ordinary_save_*`
selectors measure.

### Twenty-six selectors, 501 → 527

Sixteen extend 0638's ordinary-save family with two new origins,
`marker-bearing-producer-shape` and `marker-stripped-control`, over the same four
phases and for DOCX and PPTX only (an XLSX marker origin is refused with a typed
error, because 0601's XLSX producer family already exists):
`{docx,pptx}_marker[_control]_ordinary_save_{lifecycle,edit,atomic_publish,counting_publish}`.
`pptx_marker_ordinary_save_edit` is exactly 0649's phase.

Eight read the complete projected text through both facades:
`{pptx,docx}_marker[_control]_{eager,source}_full_text`.
`SourceBackedPresentation` has no whole-presentation text entry point, so the
PPTX source-backed scenario is every slide's `text()` in order and its oracle is
derived the same way at construction.

Two close 0643's gap in the shape the RTF, ODT, ODS and ODP
`*_semantic_text_to_sink` selectors already use, over the ordinary generated
corpus: `docx_semantic_text_to_sink` (eager `Document::write_text_to`) and
`docx_source_text_to_sink` (source-backed `Package::write_text_to`).

None is in `Case::DEFAULT`; the checked default catalog SHA-256 and
`content_set_sha256` do not move, exactly as for 0601, 0627 and 0638.
`--marker-evidence PATH` writes the per-member census under the schema
`litchi.perf-baseline.marker-shape-evidence.v1`.

### Allocation metrics for the ordinary-save family

All forty selectors in the family now open an allocation region, and it is
exactly the interval the phase reports: it opens immediately before the clock
and closes immediately after it, in each of the four phase arms.

## Authority

This is queue row 17 of change [0651](0651-queue-refresh-after-the-second-wave.md)
("harness gaps"), harness-only work that change 0652 does not need to authorize
because it moves no contract: the third wave's three standing trade-offs are
about production code, and **no file under `crates/` changed**. The work is the
prerequisite 0649 named in as many words, and the corpus it builds is what
0652's row-1 decision — *"the MCE codec's namespace re-declaration may change
its consumers' public API"* — will be measured against: change 0653 is
implementing that codec change in this wave and now has a selector to land
against.

## Breaking changes

None. Nothing under `crates/` changed, so no public item of any published crate
moved. The harness is a separate Cargo project with no published API. Two
harness-internal shapes grew: `ordinary_save::Origin` gains `MarkerShape` and
`MarkerControl` (its `ALL` goes from 2 to 4), and `Case` gains 26 variants.
Neither is public API.

## Why it is sound

* **No production code changed**, so no invariant, error identity, limit,
  defence or output byte could move. `git diff --stat 70d7768cc` touches only
  `tools/`, `docs/` and this packet.
* **Every corpus is deterministic.** Two independent builds of each of the four
  corpora produce the same archive SHA-256, asserted in the module's own tests
  and re-derived by every run that writes a census.
* **The control is not asserted to be comparable — it is proved.** Member names,
  per-member uncompressed lengths, start-tag counts, attribute counts and the
  projected text must match the marker variant before any sample runs, and the
  control must mention the namespace zero times while the marker variant
  mentions it at least once. A corpus that fails either check refuses to build.
* **Every corpus's oracle comes from the library, not from the generator.** The
  eager full text, the source-backed full text, the object count and the sink
  projection are all read back out of the built package before any sample runs,
  and every retained sample must reproduce them.
* **A refusal is an outcome, not a dropped scenario** (0627's and 0638's rule).
  The marker DOCX's sink projection is refused; the refusal string is frozen in
  the corpus census as `sink_refusal` rather than the census omitting it.
* **The census states its own baseline.** Change 0032 recorded that the
  generated XLSX worksheets are marker free and later records generalized it.
  The DOCX writer is **not** marker free: it already declares `xmlns:mc` on
  `word/settings.xml`, `word/numbering.xml` and `word/fontTable.xml`, which is
  3 members and 12,300 bytes of the DOCX skeleton. Every corpus reports
  `skeleton_marked_member_count` and `skeleton_marked_bytes` beside what the
  generator added. The PPTX skeleton's are both zero.
* **The allocation region is the timed interval.** The owner is still alive when
  the region closes (`drop(owner)` has always been outside the timer in every
  phase), so `live_bytes_after − live_bytes_before` is *retained* memory, not a
  leak; it is reported as `retained_live_bytes`, as the ODP and XLSX
  source-backed families already report their endpoints.
* ADR 0001's layering and ADR 0005's no-leakage rules are untouched: the module
  is in `tools/`, holds no archive type, no raw lock and no executor, and the
  harness library keeps `#![forbid(unsafe_code)]`.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
Base `70d7768cc`. Every measured process pinned with `taskset -c 19` while seven
other agents built and measured on the other cores. Binaries staged outside
every Cargo target directory (0627's lesson) and hashed, because `lto = true`
makes them non-reproducible byte for byte (0635):
`litchi-perf-baseline` `095c21f94de740bc827c6080cde46a54b8c32382ec7e8b535bb0b5bcda05ca62`,
`litchi-perf-baseline-alloc` `6a65a09d706ae0063baebf61943a6f91971309bac26b68b42fb8dd9f31dae574`.

20 warm-ups and 50 retained samples per run; three repeats (R1, R2, R3) of every
group and a dedicated A/A pair (A1, A2) for the two heaviest groups. The table
reports the median of the repeat p50s, the widest p50 spread across every repeat
of that selector, and the dedicated A/A floor where one was taken.

### The result this record exists to produce

| selector | p50 (ms) | repeat spread | A/A floor |
| --- | ---: | ---: | ---: |
| `pptx_marker_ordinary_save_edit` | **150.368** | 3.39% | **0.73%** |
| `pptx_marker_control_ordinary_save_edit` | **7.432** | 87.38% † | 86.56% † |
| `pptx_ordinary_save_edit` (generated, change 0638's) | **1.948** | 159.72% ‡ | — |

| ratio | this record | change 0649 |
| --- | ---: | ---: |
| marker / generated | **77.21×** | 77.3× (capture), 67.7× (edit) |
| marker / byte-identical control | **20.23×** | 13.9× (per byte, real deck) |
| control / generated | **3.82×** | 5.57× (bytes, real deck) |

The product decomposes the same way 0649's does — 3.82 × 20.23 = 77.3 against
0649's 5.57 × 13.9 = 77 — with a different split, because this corpus carries
more of its bytes in slides and fewer in layouts than the fixture does, and the
capture only reads slides. **0649's 77× is now a harness number.**

† and ‡ are honest caveats, not rounding, and the † one has a diagnosis. The
control edit's five p50s are 7.400, 7.409, 7.432, 7.677 and 13.866 ms — four
within 3.7% and one at 1.87× — and the outlier is the A1 run of the A/A pair,
in which **all four control selectors are 1.8× to 3.4× their other p50s while
all four marker selectors of the same process are within 0.7%**. The control
selectors execute in the second half of each process, so A1's second half is a
contaminated window on a host eight agents share, not a property of the corpus.
‡ is the same kind of event: the generated edit's three p50s are 1.928, 1.948
and 5.008 ms. The medians above exclude both outliers, both selectors'
*within-run* distributions are tight (`p95/p50` of 1.09 and 1.007), and the
ratio on R1 alone — a single uncontaminated process — is **78.0×**.

### Marker against control, on every scenario that has a pair

Median of the repeat p50s, in milliseconds. The two packages differ **only** in
which branch of `process_markup_compatibility` each part takes.

| scenario | marker | control | ratio | marker A/A | control A/A |
| --- | ---: | ---: | ---: | ---: | ---: |
| PPTX opened-transaction edit | 150.368 | 7.432 | **20.23×** | 0.73% | 86.56% † |
| PPTX eager full text | 58.160 | 4.517 | **12.88×** | 0.34% | 3.07% |
| PPTX source-backed full text | 58.164 | 4.407 | **13.20×** | 1.68% | 3.29% |
| DOCX eager full text | 6.251 | 0.245 | **25.55×** | 0.71% | 0.83% |
| DOCX source-backed full text | 6.228 | 0.204 | **30.59×** | 0.10% | 0.10% |
| DOCX ordinary-save edit | 7.594 | 0.595 | **12.75×** | — | — |
| PPTX open+edit+save | 163.460 | 15.032 | 10.87× | 0.37% | 119.08% † |
| DOCX open+edit+save | 13.861 | 6.743 | 2.06× | — | — |
| PPTX save-to-path | 7.346 | 6.675 | 1.10× ‡ | 82.64% | 240.56% |
| PPTX serialize-to-sink | 0.212 | 0.171 | 1.23× ‡ | 13.13% | 60.70% |
| DOCX save-to-path | 6.150 | 6.978 | 0.88× ‡ | — | — |
| DOCX serialize-to-sink | 0.587 | 0.562 | 1.04× ‡ | — | — |

‡ The four publication rows are **not relied on**. Their floors run from 13.13%
to 240.56% because the interval is one `litchi_opc::atomic::replace_with`,
whose two `fsync`s are on a device eight agents share; the counts below are the
evidence for those phases. This is the program's own rule applied: a floor above
5% at p50 means rely on counts. Every row is reported anyway, including the one
where the marker corpus is *faster* than its control.

The read rows are the clean ones. The marker/control ratio is 12.88× and 13.20×
on PPTX — within 8% of 0649's 13.9× per-byte factor measured natively on the
real deck by a completely different instrument — and 25.55× and 30.59× on DOCX,
where the roots carry 33 declarations instead of 6.

### Deterministic counts: the allocation metrics the family never had

From `litchi-perf-baseline-alloc`, 5 warm-ups and 20 retained samples. **Every
work counter is identical across all 20 retained samples of every selector**, so
these are counts, not estimates. The allocator observer perturbs scheduling, so
no timing is taken from these runs.

| selector | allocation calls | allocated bytes | reallocations | retained live bytes | region peak |
| --- | ---: | ---: | ---: | ---: | ---: |
| `pptx_marker_ordinary_save_edit` | **418,273** | **146,233,844** | 27,576 | 17,059 | 2,074,589 |
| `pptx_marker_control_ordinary_save_edit` | 99,367 | 6,986,277 | 1,244 | 17,059 | 1,026,706 |
| `pptx_ordinary_save_edit` | 26,794 | 1,614,311 | 1,029 | 1,557 | 776,921 |
| `docx_marker_ordinary_save_edit` | 11,848 | 8,886,818 | 877 | 407,879 | 1,875,978 |
| `docx_marker_control_ordinary_save_edit` | 8,292 | 5,454,751 | 852 | 407,879 | 773,719 |
| `docx_ordinary_save_edit` | 8,277 | 1,768,329 | 837 | 388,061 | 544,685 |

The PPTX edit's allocated bytes are **20.93×** its byte-identical control's and
**90.6×** the generated corpus's; the timing ratios are 20.23× and 77.21×. The
allocator and the clock agree about the marker/control factor to 3.4%. One edit
of a 49 KB archive allocates **146 MB — 2,975× the archive** — against 0649's
91 MB and 841× on the real deck; this corpus is 1.16× the real deck's allocation
calls and 1.61× its allocated bytes, so it sits in the same regime rather than
exaggerating it.

Retained live bytes are identical between each marker corpus and its control
(17,059 and 407,879), which is what a corpus pair that differs only in a codec
branch should show: the branch costs transient allocation, not retained state.

### The two DOCX text-sink selectors change 0643 asked for

| selector | p50 (ms) | repeat spread |
| --- | ---: | ---: |
| `docx_semantic_text_to_sink` (eager) | 0.1971 | 2.95% |
| `docx_source_text_to_sink` (source-backed) | 0.2815 | 3.88% |

The source-backed facade is 1.43× the eager one on the 200-paragraph generated
corpus, with both floors under 4%. 0643 measured this pair with a throwaway
probe and recorded that no harness case covered the path; two now do.

### A refusal the pair found

`Document::write_text_to` **refuses the marker DOCX corpus** with the typed
`invalid DOCX format: semantic DOCX XML exceeds 4096 namespace bindings`, while
its byte-identical marker-free control is admitted and projects 10,199 bytes
over 200 objects. `MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS`
(`crates/litchi-docx/src/paragraph/codec/text.rs:413`) is a **cumulative** count
of `xmlns:` attributes over the whole parse, and the MCE codec re-declares every
in-scope binding on every emitted start tag, so a 33-declaration root exhausts
the budget after about 124 elements. The eager `Document::text()` path on the
same corpus succeeds, so this is specific to the sink parser.

This is **not** a claim that real Word files are refused. Both real DOCX
fixtures censused here are admitted (`layout-in-cell-2.docx` projects 6 bytes
over 4 objects, `NumberedList.docx` 82 over 12) because their `document.xml`
carries very little body text. What is established is that the refusal is caused
by the markers and not by the shape: the two corpora have the same members, the
same lengths, the same element counts and the same attribute counts, and one is
refused. It is the reason there is no marker text-sink selector, and it is
frozen in each corpus's census rather than dropped.

## Correctness evidence

No production code changed, so there is no differential to run. What is checked
instead:

* **Five new unit tests** in `marker_shape.rs`: the control namespace is the
  same length as the real one and is not an OpenXML URI; every corpus builds
  byte-identically twice; every corpus marks the part kinds the fixture marks
  (slides, layouts, masters, notes master and `presentation.xml`, and **not**
  themes) with exactly six root declarations each and one
  `mc:AlternateContent` per slide; the control is byte-comparable and hashes
  differently; and both marker corpora carry at least 80% of their bytes in
  parts the codec would rewrite.
* **Four new source-policy tests** in `tools/test_perf_baseline_source_policy.py`:
  the module owns no `unsafe`, no global allocator, no `std::env`, no `Command`
  and no file read; it names its three derivation fixtures, its derivation
  script, its control proof and its census fields; the ordinary-save family
  opens exactly four allocation regions and folds them into
  `InProcessObservation`; and all 26 new selectors are present and absent from
  the default matrix.
* **The derivation gate.** `derive_marker_shape.py verify` re-censuses the three
  fixtures and fails if any declaration URI, `mc:Ignorable` string,
  `mc:AlternateContent` fragment or fixture path in `marker_shape.rs` no longer
  matches them. It passes.
* **Two existing harness tests were updated rather than relaxed**:
  `formats_origins_and_phases_describe_themselves` now asserts the four origins
  and that only the marker ones carry a variant, and
  `every_ordinary_save_case_is_opt_in_and_round_trips` now walks 40 triples,
  asserts that the three XLSX × marker triples have **no** selector, and asserts
  that `build_corpus` refuses an XLSX marker corpus with the typed message.
* **Every retained sample of every new selector reproduced its corpus's frozen
  oracle**; a mismatch fails the case, and none failed across 6,700 retained
  timing samples and 680 retained allocator samples.
* **0638's own determinism invariants still hold on the new corpora**: each
  builds its workspace and proves, untimed, that two fresh open/edit/save cycles
  publish the same digest and that saving one edited owner twice publishes the
  same digest, before any sample runs.

### Gates

`cargo fmt --all --check`, `cargo clippy --release --locked --all-targets` (with
and without `--features allocator-metrics`), `cargo doc --no-deps` and
`cargo test --release --locked` in `tools/perf-baseline`;
`python3 tools/non_iwork_gate.py harness-tests`;
`python3 tools/validate_crud_coverage_index.py`;
`python3 -m unittest tools.test_corpus_manifest_v2 tools.test_crud_coverage_index
tools.test_perf_baseline_source_policy tools.test_perf_corpus_binding`;
`python3 tools/non_iwork_gate.py verify`; and the derivation gate. Every one
exits 0. The harness's own suite is **527 passed, 0 failed, 1 ignored** in both
the release run (80.8 s) and the debug run `non_iwork_gate.py harness-tests`
drives (588.8 s); the ignored one is the opt-in
`security_corpus::real_producer_security_corpus_is_bounded_and_deterministic`,
and the two allocator tests in `docx_bounded_tail_append_compare` the briefing
lists as flaky passed in both windows. Tails in
[`results/change-0664/gates.txt`](results/change-0664/gates.txt).

The derivation gate failed on its first run and its **check** was corrected, not
its subject: it had demanded that every declaration of a PPTX fixture root be a
literal in `marker_shape.rs`, including the three the production writer already
emits and the generator therefore does not re-declare, and it had compared the
`mc:AlternateContent` template as one unescaped string against a Rust `concat!`
of escaped literals. It now requires the URIs the generator adds, the fixture's
root binding count as a named constant, and the template's fragments. No harness
source changed between the two runs; both are retained in `gates.txt`.

## Validation preserved

Nothing under `crates/` changed, so every limit, defence and typed refusal is
the base commit's. The new corpora are *inputs* to those defences, and two of
them fire: the DOCX sink's cumulative namespace-binding limit refuses the marker
corpus (recorded above), and `source_stream_eligible` is defeated on every
marker-bearing part, which is the point of the corpus. Neither is weakened,
relocated or bypassed; both are recorded as outcomes.

## Limitations

* **No timing, allocation, physical-I/O, cold-cache or speedup claim**, and no
  claim-registry entry. These are descriptive baselines.
* **The four publication rows are not usable for latency** at floors of 13.13%
  to 240.56%; their counts are.
* **The corpus is not the fixture.** It reproduces the fixture's slide byte
  total to 92.0%, its root declaration count exactly, its marker byte share to
  within 3.4 points and its `mc:AlternateContent` count per slide exactly. It
  does not reproduce the fixture's layout and master bytes, because this
  harness's writer emits eleven small layouts and one master; the capture reads
  neither, which is why that was the term chosen to drop.
* **The marker/control ratio is not the codec's cost in isolation.** It is the
  cost of the whole scenario on a package that trips the rewriting branch versus
  the same package that does not, which also includes what the rewritten output
  costs the XML parser downstream.
* **The DOCX sink refusal is a corpus fact, not a claim about Word files.** The
  two real fixtures censused are admitted; no survey of real files was run.
* **No native cycles, instructions or per-symbol profile** were taken. 0649 has
  them for the real deck; this record's job was to make the scenario selectable,
  and it reports what the harness reports.
* **The XLSX marker origin does not exist.** 0601's XLSX producer family already
  covers that ground; `build_corpus` refuses one with a typed error rather than
  silently building the wrong thing.
* **`Case::DEFAULT` is unchanged** and the checked catalog hashes do not move,
  so no default-matrix baseline is affected or re-based by this change.

## Retained evidence

[`results/change-0664/README.md`](results/change-0664/README.md).
