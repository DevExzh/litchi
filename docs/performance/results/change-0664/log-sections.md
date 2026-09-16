# Log sections for change 0664

The coordinator merges each block into the newest position of the named file.

## For `HOTSPOTS.md`

## 0664 — the codec's rewriting branch is on a selector at last: 77.21× the generated corpus, 20.23× a byte-identical control

Record: [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md).
Change 0649 found that every prior PPTX record measured on corpora whose members
never mention the markup-compatibility namespace, so the branch that costs 93.9%
of a real deck's edit was never on a harness selector. Four new corpora close
that: a marker-bearing PPTX deck and DOCX document derived from three real
tracked fixtures by a retained script, each with a **marker-stripped control**
whose members, member lengths, start-tag counts, attribute counts and projected
text are proved identical and whose only difference is which branch of
`process_markup_compatibility` each part takes. `pptx_marker_ordinary_save_edit`
— 0649's phase — is **150.368 ms at a 0.73% A/A floor, 77.21× the generated
corpus's 1.948 ms and 20.23× its own control's 7.432 ms**, decomposing 3.82 ×
20.23 = 77.3 where 0649's real deck decomposed 5.57 × 13.9 = 77. The read
scenarios give the cleanest per-byte figures the program has: PPTX eager and
source-backed full text at **12.88× and 13.20×** their controls (0649 measured
13.9× natively on the real deck with a different instrument), DOCX at **25.55×
and 30.59×**, where the roots carry 33 declarations instead of 6. The allocator
agrees: the marker edit allocates **146,233,844 bytes over 418,273 calls**
against the control's 6,986,277 over 99,367 — a 20.93× byte ratio against the
clock's 20.23× — and 2,975× its own 49 KB archive. Change 0653 now has a
selector to land its codec change against. Two facts the census settles: the
deck's 43 marker-bearing members are its slides, layouts, masters and
`presentation.xml` but **not** its themes, and the deck carries **no
`mc:Ignorable` at all** — the bare `xmlns:mc` declaration is the whole trigger.

## For `GOAL_AUDIT.md`

## 0664 — two named harness gaps closed, and one `docs/GOAL.md` measurement rule finally satisfied for PPTX

Record: [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md).
`docs/GOAL.md`'s scoping rule — every claim is scoped to scenario, corpus,
machine, build and metric — has been formally satisfied by every PPTX record in
this program and substantively defeated by one of those five: the corpus. Change
0649 measured that the generated PPTX corpus takes a path **13.9× cheaper per
byte** than a real deck because none of its members mentions the
markup-compatibility namespace, so a PPTX number measured on it was scoped
correctly and generalized wrongly. This change removes the excuse: 26 opt-in
selectors, four generated corpora, and a derivation script that fails if the
corpora stop matching the real fixtures they were derived from. Change 0649's
second gap is closed in the same batch: `litchi-perf-baseline-alloc` emitted no
allocation metrics for the 24 `*_ordinary_save_*` selectors because change 0638
registered them without opening an allocation region; all **forty** (24 existing
plus 16 new) now open one, bounded to exactly the interval each phase reports,
and every work counter is identical across all 20 retained samples of every
selector. Change 0643's third gap — *"the harness has no DOCX text-sink
selector"* — is closed by `docx_semantic_text_to_sink` (0.1971 ms) and
`docx_source_text_to_sink` (0.2815 ms), both under a 4% floor. The registry goes
from 501 to 527 selectors; `Case::DEFAULT` is unchanged at 41 and neither
catalog hash moves, so no existing baseline is re-based.

## For `REPORT.md`

## 0664 — the marker-bearing corpora, their controls, and what the pair costs

Record: [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md).
Harness-only; no file under `crates/` changed. Four corpora:
`pptx-marker-deck` (63 members, 351,108 uncompressed bytes, 27 marker-bearing
members holding 89.66% of them, 13 `mc:AlternateContent` blocks, six root
bindings on every marker-bearing part) and `docx-marker-medium` (12 members,
72,956 bytes, 6 marker-bearing holding 84.01%, 33 root declarations and Word's
own `mc:Ignorable` on `document.xml`), each with a marker-stripped control of
identical member names, lengths, start-tag counts, attribute counts and
projected text. The shape is derived by
`results/change-0664/scripts/derive_marker_shape.py` from
`slide-section-test.pptx` (change 0649's deck: 103 members, 796,725 bytes, 43
marker-bearing holding 93.03%), `layout-in-cell-2.docx` (real Word output: 10 of
21 members, 95.01%) and `tdf89064.pptx` (the only fixture whose marker coverage
reaches the notes parts); its `verify` mode is a gate. Twelve marker/control
pairs were measured at 20 warm-ups, 50 samples and three repeats plus a
dedicated A/A pair, pinned to CPU 19. The six clean rows are 12.75× to 30.59×;
the four publication rows run at floors of 13.13% to 240.56% — two `fsync`s on a
device eight agents share — and are reported and **not relied on**, one of them
showing the marker corpus *faster* than its control. A limit fired that no
record had seen: `Document::write_text_to` refuses the marker DOCX with
`semantic DOCX XML exceeds 4096 namespace bindings` while its byte-identical
control is admitted, because the limit counts `xmlns:` attributes cumulatively
and the codec re-declares every in-scope binding on every emitted start tag.
That is a corpus fact, not a claim about Word files: both real DOCX fixtures
censused are admitted, because their `document.xml` carries little body text.
One correction to the record set: change 0032's "the generated corpora are
marker free" holds for the XLSX worksheets it measured and **not** for DOCX —
`litchi_docx::Package::new()` already declares `xmlns:mc` on `settings.xml`,
`numbering.xml` and `fontTable.xml`, 3 members and 12,300 bytes of the skeleton
— so every corpus now reports its skeleton's own census beside the generator's.

## For `ADR_COMPLIANCE.md`

## 0664 — no boundary moved, and two defences recorded firing on new input

Record: [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md).
Nothing under `crates/` changed, so no ADR boundary could move: no public item,
no error identity, no limit, no defence and no output byte differs from the base
commit. ADR 0001's layering is untouched — the new module lives in `tools/` and
depends on the published facades only. ADR 0005's no-leakage rules hold: the
module holds no archive type, no raw lock and no executor, and the harness
library keeps `#![forbid(unsafe_code)]`, with the single `unsafe` boundary still
confined to the allocator binary's `GlobalAlloc` wrapper. ADR 0003's
measurement discipline is what this change is for, and it is applied to itself:
every corpus proves its own determinism and its control's byte-comparability
before any sample runs, every oracle is read back out of the built package
rather than asserted by the generator, and every floor is published beside its
ratio including the four that disqualify their own rows. Two production defences
are recorded **firing**, neither weakened nor relocated: the DOCX sink parser's
cumulative `MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS` refuses the marker corpus with
its typed `InvalidFormat` while admitting the byte-identical control, and
`source_stream_eligible` is defeated on every marker-bearing part. Both are
frozen as outcomes in the corpus census (`sink_refusal`,
`markup_compatibility_namespace`) rather than omitted, which is change 0627's
and 0638's rule applied to a corpus instead of a selector. The four
harness-internal shapes that grew — `ordinary_save::Origin` from two variants to
four, `Case` by 26 — are not public API and are covered by the harness's own
registry tests.
