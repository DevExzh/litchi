# 0587: what remains, surveyed across the whole OLE2 and OOXML data path

Status: retained, survey only. `performance_claim: none`. **No production code
changed.** This record carries a ranked inventory of the optimization
opportunities that remain at `2fc5fc6572ac322f4aae17f88f90033bcb119f0f`, each
screened against the 586 records before it. Nothing here is implemented, and
nothing here is authorized by this record; every item names the measurement or
design record it needs first.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded and was not surveyed.

## Why this record exists

`docs/GOAL.md` requires `HOTSPOTS.md` to carry "ranked opportunities by expected
total CRUD impact, risk, and ADR compatibility". The ranked work queue it carried
until this change was headed "provisional until baseline measurements are
recorded", cited nothing later than change 0190, and ranked programmes that have
since been implemented, rejected or superseded. Every batch since then has
ranked *its own* area — change 0574 the XLS open, 0581 the OPC retention, 0577
the OOXML open, 0584 the OLE2 read path — but no record ranks across areas, so
each next batch has been chosen from the most recent record's tail rather than
from the whole path.

This record replaces that queue. It is the survey `docs/GOAL.md`'s REPOSITORY
EXPLORATION section asks for, taken after 586 changes rather than before the
first, so that the next batches can be chosen in Amdahl order across the whole
OLE2 and OOXML path.

## Method

Eleven surveys ran in parallel, one per area of the data path: the CFB
substrate; XLS; DOC and PPT; the ZIP and OPC read path; the OPC save path; XLSX;
DOCX and PPTX; the shared XML substrate; XLSB; `litchi-core`, the facade and the
parallel sessions; and the program's own evidence and harness. Each read the
current tree, built a path map for its scenarios, and screened every candidate
against the record set by search before proposing it, stating for each which
records cover, reject, defer or partially implement it. Each was allowed to take
one fresh measurement where the retained packets could not answer a question —
a count, a callgrind isolation pair, an allocator run or a `strace` — under a
ten-minute cap, on the shared warm build tree, with no git worktree and no edit
to a tracked file. The eleven reports are retained verbatim in
`results/change-0587/survey/` with every output they cite; this record is the
synthesis and the ranking, and the coordinator checked eight of the cited
mechanisms against the source before ranking them (the codec's per-element
namespace emission, the three identity fingerprints per DOC snapshot open, the
PPTX fingerprint sites, the per-part `.rels` audit, the eager error construction
in the XLS frame loop, the facade's `.doc` ingress, the per-read ZIP session, the
unconditional XLSB reparse).

The first wave of eleven agents was terminated by a session usage cap before any
of them had read a file, and was relaunched in two waves under a token-economy
instruction. No evidence was lost, because none had been produced; the
instruction capped how much each survey could read, and the limitations section
says where that shows.

### Evidence tiers

Every size in this record carries one of four labels, and the label matters
more than the number:

| tier | meaning |
| --- | --- |
| **measured** | counted or profiled in this batch, on this host, from the cited output retained in `results/change-0587/` |
| **retained** | read from an earlier record's retained packet, cited by change number; not re-measured here |
| **modelled** | arithmetic over retained or measured figures, or over file structure; a prediction, not a measurement |
| **unknown** | the mechanism is established from the source but no size exists; the first step is to measure |

Instruction counts rank work, not latency: change 0579 removed 1.24% of
instructions and 6.19% of cycles on one fixture, and `GOAL_AUDIT.md` carries the
standing instruction to price pointer-chase work in cycles. Callgrind's per-byte
accounting of `rep movsb`/`rep stosb` makes every bulk-copy share an upper
bound, and callgrind runs SHA-256 in software because valgrind masks the SHA
CPUID bit, so every hashing share is about five times its native cycle share and
is paired with a native figure where one was taken. These caveats apply to every
figure below and are not repeated per row.

## The four findings that set the ranking

**1. The measured path is not the path real files take.** Every planning,
selected-cell and commit record from 0362 to 0553 measured generated corpora,
and the generator writes worksheets with no `mc` namespace, no `x14ac`, no
`dyDescent`, no `<cols>`, no shared-string part and no worksheet relationships;
the eager PPTX corpora are built through `Package::from_bytes` and so carry no
physical provenance; every XLS and OOXML selector wraps bytes in memory, so
`FileSource` observation costs never appear. Real Office files have all of these:
41 of 60 XLSX fixtures declare `mc` in their first sheet, 27 of 60 contain
`<cols>`, 77 of 95 have a shared-string part and 57 of 95 have worksheet
relationships. On those files the largest single cost in the whole OOXML read
path is one this program has never profiled: the MCE codec re-declares every
in-scope namespace on every element it emits, expanding a real Excel worksheet
16.9× before the parser sees it, so that an eager open plus one cell costs
**12.9×** the marker-free control (XML-1). The same gap disables three landed
optimizations on real files: 0546's fused traversal admits 19 of 60 fixtures
(XML-3), 0525's reduced readback is off whenever any cell is a shared string
(XLSX-2), and the selected-cell stream runs to end of sheet after `<cols>` has
already made it ineligible (XML-2). The first thing the next batches need is a
real-producer shape in the harness; until then instruction differentials, not
timings, are the usable metric on the small real fixtures that exist.

**2. Whole-artifact hashing is the dominant term in three paths, and callgrind
hid it.** A source-backed DOC or PPT snapshot open hashes the entire artifact
with SHA-256 between two and six times — a PPT text-edit snapshot open costs
13.4× the *complete eager presentation parse* of the same file in native cycles
(DOC-1). Change 0589 later corrected a reading this record first made here:
0586's zero was structural, no chain link removed on any fixture, and hashing
did not mask it; what DOC-1 explains is why no read hint on that path could have
registered in latency. A PPTX
opened-document transaction hashes the complete package four times per
lifecycle, 34.9% of the profiled instructions (PPTX-1), and a cross-package slide
copy spends 85% of its instructions serializing and hashing the source,
destination and candidate at least five times over (PPTX-2). No record found
these because the profiles that could have were callgrind profiles, where the
software SHA-256 was either discounted as harness cost (0584 attributes 53% of a
raw profile to the harness's per-sample identity check) or never taken on these
paths.

**3. The OLE2 substrate is at its floor for this corpus; what remains on OLE2
is in the format crates and the writer.** A CFB open's remaining symbols are ADR
0005 mandatory validations or work rejected three times over, a file-backed open
is 34 `statx` and 53 `pread64`, and no fixture exceeds 1.6 MB or carries a DIFAT
sector. The OLE2 opportunities with measured size are the XLS worksheet frame
loop, where 46% of a one-cell query is framing overhead and 0584's candidate 2
was never implemented (XLS-1); the eager DOC open's three terms, text decoding,
paragraph-property resolution and stream copies, which explain 0584's
size-independent cost (DOC-2 to 4); and the whole-stream zero-fill on the slurp
paths (CFB-1). The OLE2 writer has no copy-through for length-changing saves at
all (CFB-2), and no selector measures one.

**4. Nothing is authorized, and the prerequisites are known.** Of the 36 ranked
items, ten can proceed on a paired measurement alone, twenty-one need a frozen
design record of the kind changes 0565, 0566 and 0575 wrote, and five need a
proposed ADR or an ADR clarification before any code. Six correctness or
compliance findings and a handful of smaller oddities were found on the way and
are reported, not fixed.

## The ranked queue

Rank is a judgment over expected CRUD reach times size, divided by risk, with
the prerequisite stated; it will move as the top items are priced natively.
`step` is the `docs/GOAL.md` optimization-order step. Identifiers refer to the
area sections below.

| # | id | opportunity | step | size and tier | reach | risk | needs first |
| ---: | --- | --- | --- | --- | --- | --- | --- |
| 1 | XML-1 | MCE codec re-declares every namespace on every element: a real Excel sheet grows 16.9×, an eager open plus one cell costs 12.9× the control | 1 | 1,037.8 M vs 80.5 M Ir on one fixture, **measured**; 41 of 60 sheets affected, counted | every XLSX planning and edit route, PPTX text, DOCX source reads, on producer files | medium | frozen design; corpus-wide Ir differential |
| 2 | DOC-1 | source-backed DOC and PPT snapshot opens hash the whole artifact 2-6 times | 1 | PPT text-edit open 13.4× the eager parse in cycles; DOC 2.5-6.5×, **measured** | every source-backed DOC/PPT read and commit | medium | frozen design of the ADR 0006 fence; 0582 harness re-run |
| 3 | PPTX-1 | four complete-package revisions per opened transaction | 1 | 34.9% of whole child, **measured**; ~86% of each commit, modelled | every PPTX opened-document CRUD | low (a,b); medium (c,d) | (a,b) paired measurement; (c) frozen design |
| 4 | DOCX-1 | a one-paragraph edit scans the main part four times and recompacts it whole | 1 | ~64% + 14% of timed Ir, **measured** | every ordinary DOCX edit and save | low (a); medium (b,c) | (a) measurement; (c) frozen record and preservation tests |
| 5 | PPTX-2 | cross-package slide copy re-serializes and re-hashes source, destination and candidate five or more times | 1-2 | 85% of whole child, **measured** | cross-document copy and assembly | medium | frozen design on proof reuse |
| 6 | XML-2 | selected-cell stream runs to EOF after the sheet is known ineligible, then the store re-parses | 1 | 63.6% of a source-backed read, **measured**; 27 of 60 sheets, counted | XLSX selected-cell reads on producer files | low-medium | frozen design with an error-order guard |
| 7 | SAVE-1 | publication plan reserializes and audits unchanged `.rels` and content types | 1 | 32-36% of every save, **measured** | every OOXML save | low | A/B on the 132-member fixture |
| 8 | ZIP-1 | a fresh Deflate decoder per structural member and per cold part | 2 | 60% of open allocation bytes; ≤18% Ir, **measured** | every OOXML open and cold read | low | ABBA open timing |
| 9 | XLS-1 | lean worksheet frame loop; 0584's candidate 2 never landed | 1, 3 | 46.1% of the `54016` one-cell query, **retained** | XLS one cell, all cells, text | low | cycles A/B |
| 10 | DOCX-2 | build the paragraph index lazily | 1 | 75% of a full-text read, **measured** | DOCX text reads | low | measurement |
| 11 | DOC-2/3/4 | eager DOC open: four-pass text decode, triple PAPX resolution, FIB stream copy | 1-2 | 43% / 22-38% / 14% of opens, **measured** | every DOC open, both validations of every DOC edit | low to low-medium | measurement |
| 12 | XLSX-2 | admit shared-string and relationship-bearing worksheets to 0525's readback and the value editor | 1 | 0525's 96%/30% is zero on 77 and 57 of 95 real files, counted | XLSX edit and save on producer files | medium | frozen design; SST harness shape |
| 13 | XML-3 | admit marker-bearing worksheets to 0546's fused traversal | 1 | 0546's −22% planning is zero on 41 of 60 real files, counted; modelled | XLSX planning on producer files | medium | frozen design (0541 guards) |
| 14 | ZIP-2 | one bounded positional read per member first-read | 2 | open 87 → 45 requests; 0572's 354 −40-50%, modelled | range-source OOXML reads | medium | frozen design; 0582 harness |
| 15 | ZIP-5 | the accessor 0577's coalesced structural prefetch is blocked on | 2 | 87 → 10 open requests, modelled | range-source OOXML opens | medium | 0577's record, with error precedence stated |
| 16 | SAVE-2 | buffer the atomic tempfile | 2 | 531 → ~11 write syscalls, **measured** count; 0.8-1.6 ms per save, modelled | every `save(path)` | low | `strace -c` on the target store |
| 17 | XLS-2b | cheaper eager SST walk: single-segment fast path, no per-string `Vec` | 3 | up to half of 49.5% of the `54016` open, modelled | XLS open with strings | low | 0576 differential |
| 18 | XLS-3 | retained validated sheet index and an all-cells iterator | 4 | 70.5% of a query per repeat, **retained** | XLS repeated queries, all cells | medium | frozen design (bounded cache) |
| 19 | XLSB-1 | commit reparses the workbook for a proven no-op, twice for an edit | 1 | +84% no-op, +178% one edit, **measured** single leg | XLSB edit and no-op | low | a large synthetic XLSB fixture |
| 20 | PPT-1 | record tree copies bytes once per nesting level and retains the stream | 2-3 | memcpy 30% of a PPT open; copy factor 1.3-2×, **measured** | every PPT open and edit | medium | design (public field) |
| 21 | SAVE-5 | C2′: lazy decode behind the fallible package accessors | 2 | `load_parts_eager` 49% of an edit operation, **measured**; 3.58× retention, **retained** (0581) | every eager OOXML open, edit, save | medium | proposed ADR (0581 gates) |
| 22 | CFB-1 | append-style reads so the whole-stream zero-fill never exists | 2 | 9-32% of DOC/PPT open Ir, **retained** upper bound; in-memory sources only | eager DOC, PPT and XLS opens | medium | frozen design; cycles A/B, may fall inside the floor |
| 23 | XLS-2 | lazy SST indexing | 1 | 49.5% of the `54016` open, **retained**; moves when a malformed SST is refused | XLS open and list with strings | high | frozen design (0576) |
| 24 | XLS-6 | skip never-interpreted globals payloads, frame once | 2 | ~30% of the flagship open Ir, **retained**; +46 requests, modelled | XLS open | medium-high | frozen design with the density gate |
| 25 | XLSX-1 | compact per-cell source facts from planning into commit | 1-2 | ≤25% of a one-percent edit interval, modelled from 0550/0520 | XLSX source-backed edit | high | 0551's design plus its differential oracle |
| 26 | XLSX-3 | reuse the publication audit of the unchanged original | 2 | ≤28% of publication, modelled from 0528 | XLSX publication | medium | memo: design; proof door: proposed ADR |
| 27 | SAVE-3 | the PPTX eager save regenerates every slide | 1 | unknown | PPTX ordinary edit and save | medium-high | measurement, then frozen design |
| 28 | ZIP-3 | five observations per cold part read, and a sticky monitor flag | 2 | 5 → 2 per read; 54 per stream, **measured** counts | file-backed OOXML reads | low-medium | 0563-style per-site analysis |
| 29 | ZIP-4 | allocation-free member-name lookup, cheaper `PackURI` admission | 2 | ~19% of open-plus-part Ir, **measured** | every OOXML open | low | cycles A/B |
| 30 | CORE-1 | the facade slurps every `.doc` and parses it eagerly | 1-2 | unknown; blocked by 8-of-57 admission | facade DOC opens | high | DOC-1 first; frozen design |
| 31 | XLS-9 | XLS edit and save readback is a complete eager open | 1 | unknown; no attribution exists | XLS edit and save | medium-high | attribution, then proposed ADR or record |
| 32 | CFB-2 | copy-through writer for length-changing OLE2 saves | 1-2 | unknown; no selector | XLS, DOC, PPT edit and save | high | frozen design and an ADR clarification |
| 33 | XLS-4 | full-text per-string and per-row fences, rectangle walk | 1-2 | ≥48,165 `fstat` per extraction on a file, modelled | XLS text on file sources | low-medium | design note |
| 34 | CORE-2 | 25 fences per file-backed XLS open, one per boundary | 2 | 1-1.5% locally; a count claim, modelled | file-backed XLS opens | medium | design note on fence placement |
| 35 | CORE-3 | complete `ExecutionContext`: I/O concurrency, CPU budget, executor injection | 5 | prerequisite | every parallel path | low | proposed ADR (0005 amendment) |
| 36 | CORE-4 / SAVE-6 | parallel deflate of changed members, gated | 5 | unknown; gate two or more members of ≥256 KiB | multi-part OOXML saves | medium | frozen design |

Smaller items — XLS-5, 8 and 10, XLSX-4 to 8, XML-4 and 6, SAVE-4, PPT-2, ZIP-6
to 8, CFB-3 to 6, CORE-5 and 6 — are recorded in their area sections and are
not ranked; each is either inside the noise floor on this corpus, measurable
only with a fixture that does not exist, or bundled with a ranked item.

## Area findings

### The XML substrate (`litchi-ooxml-common`, `xml-minifier`, the OOXML readers and writers)

**Path.** One reader everywhere: quick-xml 0.41 with its `memchr` feature; no
hand-written tokenizer. Between the reader and every semantic parser sit two
Markup Compatibility (MCE) implementations. The *codec*
(`crates/litchi-ooxml-common/src/mce/codec.rs:529` onward) runs a naive
`windows()` presence scan for the MCE namespace string and, if it occurs
anywhere in the part, re-tokenizes the whole part and re-emits every byte into a
`BoundedOutput`; it is called from 47 `litchi-xlsx`, 19 `litchi-docx`, 34
`litchi-pptx` and 6 `litchi-drawingml` files, on the XLSX planning and eager
paths (`raw/worksheet/mod.rs:198`, `:209`), on PPTX semantic text
(`parts/slide.rs:812`) and on DOCX source-backed reads (`source_backed.rs:748`).
The *semantic stream* (`mce/stream.rs`) drives the XLSX selected-cell scan and
builds an owned `ElementData` per element — qualified name, expanded name, and
per attribute a qualified name, raw value, decoded value and expanded name —
then clones the expanded name again for the event (`stream.rs:2440-2497`,
`:2557-2570`). The fused validate-and-parse traversal of change 0546 is admitted
only by `source_stream_eligible` (`raw/worksheet/mod.rs:60-68`): no `mc`
namespace, no `AlternateContent`, no `x14ac`, no `dyDescent`, at most 8 MiB,
UTF-8. Writers build `Vec<u8>` by `extend_from_slice`, retain numbers lexically,
and `escape_xml` always allocates a `String`; `xml-minifier`'s runtime role is
the compactness auditor (`audit.rs:1300-1440`), its macros compile-time only.

**All fresh numbers here are synthetic single-run callgrind instruction counts
on one real fixture** — `Excel_file_with_trash_item.xlsx`, `sheet1.xml` 209,931
bytes, 11,578 elements — against a byte-identical control with only the `mc`
and `x14ac` markers stripped (`results/change-0587/xml-substrate/EVIDENCE.md`,
`cg-*.txt`). They rank work, not latency.

**XML-1. The MCE codec re-declares every in-scope namespace on every element,
expanding real-producer parts 14-17×** (steps 1 and 2). `write_start`
(`mce/codec.rs:1145-1194`) calls `ctx.ns.for_each_effective` (`:361-372`) and
emits *all* effective bindings — the MCE namespace included — on every start
tag it writes; attribute values and URIs are re-escaped character by character
through `esc` (`:1196-1212`) into a `BoundedOutput` whose `reserve` grows by
`try_reserve_exact(additional)` (`:486-497`) and whose attribute vector grows one
element at a time (`:657-662`); `expand` allocates a `Name` per attribute
(`:1178`). **Measured:** the codec alone turns the 209,931-byte sheet into
3,540,261 bytes (16.9×; 46,312 `xmlns` declarations in the output against 4 in
the input) at 841.8 M Ir, 4,008 Ir per input byte, of which `esc` is 81.7% and
2,954,106 `__rust_realloc` calls 38.8%; the control's codec pass is 4.66 M Ir
and borrows. The public eager open plus one cell is **1,037.8 M Ir on the real
file against 80.5 M on the control, 12.9×**, the codec 81.9% of it, and the
parser then parses the expanded stream (152.6 M against 59.4 M). The
source-backed read of the same cell is 1,075.5 M against 221.1 M. A Word 2013
`document.xml` with 31 root namespaces expands 3,108 → 50,987 bytes (16.4×) and
a PPTX slide 27,582 → 385,997 (14.0×). **Breadth, counted:** 41 of 60 fixtures'
`sheet1.xml` declare `mc`, because Excel writes `mc:Ignorable="x14ac …"` by
default. Records: none — the only MCE work is the presence scan (0530 sized it
at 5.18% of planning, 0531 rejected `memmem` for it), and 0032 states explicitly
that the generated benchmark worksheets contain neither `dyDescent` nor MCE
markup; the rewrite path has no retained profile in 586 records. Fix: emit only
the element's own declarations (hoisting an unwrapped `mc:Choice`/
`AlternateContent` wrapper's declarations onto the first emitted child), copy
unchanged start-tag byte spans verbatim, amortize output growth, borrow names.
Scenarios: every XLSX planning and edit route through `raw::worksheet::parse`,
PPTX semantic text and selected-slide reads, DOCX source-backed reads, and the
checklist's "Unknown data is bounded" row (see the defect below). Constraints:
a read-side transform whose output is not published (to be confirmed in the
design record; if any writer publishes it, the "Compact XML output" row is
violated today); namespace scoping when unwrapping must stay exact, proven by
differential tests against today's output modulo redundant declarations. Frozen
design record; no ADR change. Risk medium. Falsified if a corpus of real Office
parts takes the rewrite path on under 5% of opens — the 41-of-60 count makes
that unlikely — or a native ABBA on such a corpus moves under 2%.

**XML-2. The selected-cell stream keeps streaming after the worksheet is
known ineligible, then the eager store re-parses everything** (step 1).
`mark()` only records the reason (`raw/worksheet/selected.rs:1328`); the scan
runs to EOF, ineligibility is published at EOF (`:474-476`), and the caller
falls back to `store()` (`workbook/source.rs:672-678`). A `<cols>` element,
which precedes `sheetData`, marks `NotEligible(Styles)` (`:650-651`).
**Measured on the control fixture:** the selected stream is 140.7 M Ir (63.6% of
the read) before `store()` adds 76.5 M (34.4%); on the real fixture the fallback
also pays XML-1 and the x14ac pass. **Breadth:** 27 of 60 fixtures' `sheet1.xml`
contain `<cols>`, 8 have row styles. Fix: a `memmem` pre-gate before opening the
stream, or stopping at the first mark; mandatory validation is still performed
in full by the fallback parser, and what changes is which typed error is
reported for inputs malformed after the mark, so a 0541-style error-identity
guard is needed. Records: 0362/0363/0365 designed EOF-only publication and the
mandatory fallback; 0574 opportunity 6 rejected an exit that would have skipped
validation, which this does not. Risk low-medium; frozen design record.
Falsified if the post-mark share on ineligible real worksheets is under 10%.

**XML-3. The fused planning traversal of 0546 excludes every real Excel
worksheet** (step 1). Only 19 of 60 fixtures pass all of `source_stream_eligible`'s
gates; the rest take the four-pass fallback (validator, x14ac collector, codec
rewrite, parser, plus `from_utf8` of the expanded buffer). Admitting `x14ac`-only
and namespace-only `mc` declarations (the x14ac observer of 0361 already exists)
lets 0546's measured −21.8 to −22.4% planning instructions apply to producer
files, and with XML-1 fixed the codec pass becomes a scan. Modelled from 0540
(validator and parser passes are equal) and XML-1's numbers. Records: 0546,
0032, 0540-0545 (fallback and "historical x14ac retry" preserved). Risk medium
(error precedence, 0541 guards); frozen design record.

**XML-4. Owned expanded names in the MCE semantic stream** (step 2). Per
element and attribute the stream clones names into `String`s twice although
consumers only compare them (`selected.rs:1377-1394`); `mce::model::Name` is
public, so borrowed or interned names are an API-shaped change. **Retained:**
`clone_bounded_name_part` 16.2% → 10.9% of selected leaf weight after 0410;
**measured:** 7.4% of the whole control read, and the stream costs 12,150 Ir per
element against 5,160 for the raw `NsReader` parser on the same bytes. Records:
0409, 0410 (partial), and the 0539/0527 warning that allocation-count wins can
fail native gates. Risk medium; rank behind XML-2, which removes the stream from
most real reads anyway.

**XML-5 / XLSX-3.** Publication audits proportional to the package (0528 owner;
0529 rejected the attribute probe); a source-version-keyed compactness proof for
the untouched prefix and suffix plus a fragment audit needs an ADR proposal
because 0528 requires both audits to stay. Risk high. **XML-6.** The DOCX sink
text path copies every event twice (`paragraph/codec/text.rs:472-476`,
`into_owned()` on a byte-slice source) where the sibling borrowing path (0229)
does not; unmeasured, low risk.

**SIMD, Workstream G: nothing to propose.** No retained profile shows escaping,
UTF-8 validation or delimiter scanning hot: quick-xml already dispatches
`memchr2`/`memchr3` (under 5% of planning in 0540's rows), `from_utf8` is 1.7%
of the eager control, `escape_xml` and `esc` are hot only because of XML-1's
expansion, number lexing is not on OOXML read paths (numbers are retained
lexically), and `crc32fast` and the inflate loop are settled on the ZIP side
(ZIP-9). UTF-16 conversion belongs to the OLE2 paths, where 0576 has already
removed the largest instance.

**Not opportunities.** `memmem` for the presence scan alone (0531; subsumed by
XML-1 and XML-3, do not re-propose standalone); borrowing attribute decoding in
the raw parser (0537/0539); the row arena, attribute probe, event preflight and
exact-bound scanner (0527, 0529, 0544/0545); lazy `.rels` or early exit from
mandatory validation (0577, 0574 opportunity 6); a redundant PPTX `unescape`
that is `Cow`-borrowed on the common path; `xml-minifier` macros (compile-time).

**Blockers and a limits defect.** The harness corpus cannot exercise the
substrate's real-producer path: generated worksheets contain no `mc`, `x14ac`,
`dyDescent` or `<cols>`, the "vendor-extension" shape adds an unknown package
part rather than worksheet markup, and no option feeds a real file; every
planning record from 0530 to 0546 and every selected-cell record measured the
marker-free fast paths. Real marker-bearing fixtures are all small (the largest
`sheet1.xml` with `dyDescent` is 210 KB), so instruction counts are the usable
metric until a large real-producer fixture exists. **Defect (reported, not
fixed):** output limits are enforced on the codec's self-expanded stream —
`Limits::default().max_output_bytes` (`mce/model.rs:133`), `litchi-pptx`'s
`MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES` (`parts/slide.rs:812-815`),
`litchi-docx`'s `mce_limit` (`source_backed.rs:748-752`) and the XLSX
`within_mce_limits` gate (`raw/worksheet/mod.rs:70-73`) — so with 14-17×
expansion a legitimately sized part (a 30 MB worksheet, or a slide or document
with dozens of root namespaces) can be refused as oversized output although
nothing in it is oversized; and the eligibility gates key on substring presence
anywhere in the part, so a comment containing "dyDescent" disables the fused
path (documented as conservative in 0032/0544).

### OOXML read path below the format crates (`soapberry-zip`, `litchi-opc`)

**Measured on an in-memory `ReadAt` with every request, observation and
allocation logged** (`results/change-0587/zip-opc-read/counts-*.txt`): a
source-backed open of the 132-member workbook is 87 positional requests and
22,546 bytes — 3 for the locator and central directory plus 2 per structural
member — with 5,686 allocations and 5.59 MB, and 4 `version()` observations;
`shapes.pptx` is 45 requests, 2,114 allocations, 2.20 MB; `comment.docx` 12
requests (3 of them descriptor reads), 425 allocations, 448 KB. A cold part read
is 2 requests, 27 allocations and 86 KB, of which one `DeflateDecoder::new` is
80,320 bytes, and **5** `version()` observations; a warm read is 0 requests and 2
observations (0563). The verified/stream path reproduces 0580 exactly: 21
requests and 54 observations for one 4,381-byte part, because
`monitor_publication()` sets a sticky flag (`source_backed.rs:2888`) after which
every positional read pays `version()` for the package's lifetime. The 0580
reverse-order regression never reaches `PartView::data()` consumers (forward and
reverse are both 208 requests); only stream/verified readers see it. No `mmap`
exists under `crates/`.

**ZIP-1. One Deflate decoder per open and per package, not per member** (step
2). `read_structural_member` builds a fresh `IndexedReadSession` and a fresh
`DeflateDecoder` for every structural member (`crates/soapberry-zip/src/office.rs:4217-4222`,
`1310-1318`), and cold part loads do the same although
`read_part_with_observer_and_capture` already accepts a session and every caller
passes `None` (`crates/litchi-opc/src/source_backed.rs:10113-10119`).
**Measured:** 42 × 80,320 = 3.37 MB of the workbook open's 5.59 MB (60%; 77% on
the PPTX, 54% on the DOCX); `memset` is 17.98% of open-plus-one-part
instructions, an upper bound. Fix: thread one session (which resets its decoder)
through `source_catalog`, `walk_relationship_graph` and `load_rels_lazy`, and
pool a session for cold loads, 0402's pattern. Records: 0402 (overlay
validation only), 0567; none for the open's structural reads. Risk low; no
design record needed. Falsified if an ABBA open timing sits inside the 4% floor
on all three fixtures.

**ZIP-2. One bounded positional read per member first-read** (step 2). A
first read costs 2-3 requests: the 30-byte local header (`archive.rs:2038-2042`),
the payload (`reader_at.rs:408-440`) and the descriptor (`archive.rs:2703`). The
central record already gives the offset, compressed size, name and extra
lengths and the descriptor flag, so one read of `30 + name + extra + csize +
24·desc` bytes, clamped by a named fallible ceiling and falling back to today's
path when the local lengths differ, makes a first read one request at the same
bytes. **Modelled from the measured shape:** open 87 → 45 (workbook), 45 → 24,
12 → 6; a part 2-3 → 1; change 0572's 354-request single-cell scenario loses
40-50% of its requests with none of the byte cost that made forward read-ahead
windows a trade (0493: +37% bytes; 0572: windows cannot collapse the scattered
open). Records: 0561 named this span model as target 1 and it was never
implemented; 0562/0573 removed the duplicates; 0577 is the complementary run
coalescing; no record rejects it. Constraints: ADR 0005 limits and per-member
error identity unchanged; the read grammar moves, so the 0582 differential
harness is the gate and a frozen design record is required. Risk medium.
Falsified if on the 1 ms/request transport the request drop does not appear as
wall clock; on a local file no change is expected and none should be claimed.

**ZIP-3. Two source observations per cold part read, not five; and a sticky
monitor flag** (step 2). The five sites are `source_backed.rs:5904`, `10001`,
`10092`, `10105`, `10128`; on `FileSource` each is a mutex plus `fstat`
(`crates/litchi-core/src/source/file.rs:146-152`). 0563's strict-subset
argument applies at least to the pair between cache admission and the cold-load
closure head, keeping the 0317 brackets that fix `SourceChanged` precedence. The
sticky `monitor_reads` flag converts every later positional read into an
observation (54 for one 21-request stream). Records: 0563, 0558/0560, 0317. Risk
low-medium. Falsified by the per-site failure-ordering analysis 0563 performed.

**ZIP-4. Allocation-free member-name lookup and cheaper `PackURI` admission**
(step 2). About 19% of open-plus-one-part instructions are name normalization,
hashing and validation (`callgrind.open.top45.txt`: `normalize_str_fallibly`
2.46%, `rfind`/`split` 3.60%, SipHash 2.85%, three `Chars::try_fold::any` scans
4.38%, `PackURI::new` 1.35%, `PartNameIndex::insert` 1.26%). `lookup_member_name`
allocates and re-normalizes on every `metadata`, `read` and `contains`
(`office.rs:4713-4750`), at least twice per structural member; already-canonical
input can be looked up as `&str` with identical results, and `enqueue_target`
copies each validated `PackURI` again as the visited key
(`pkgreader.rs:1349-1388`). Records: 0395 (case-fold index only), 0557. Risk
low. Falsified if cycles do not follow instructions here (0579's caveat).

**ZIP-5. The accessor change 0577's structural prefetch is blocked on** (step
2). It is a per-entry central-directory fact triple — local header offset,
compressed size, descriptor flag — plus physical order, all already held
(`archive.rs:3754` is `pub(crate)`; `office.rs:3606`) and none of it ZIP framing,
so exposing it on `Metadata` or as `IndexedArchive::local_span_hint(EntryId)`
keeps ADR 0011's line: `litchi-opc` prefills its own `ArchiveReadAhead` window
(`read_ahead.rs:274-423`) per run and parses nothing. **Modelled:** 87 → 10 on
the workbook, or 45 → about 10 after ZIP-2. Risk medium; 0577's frozen record
must state the intra-run error-precedence change. Falsified if, after ZIP-2, the
delayed-transport open saving is under 20 ms on the workbook.

**ZIP-6 to 9, small or nothing.** A one-read tail-window locate saves 2
requests per open (ZIP-6; below noise alone). Validate-on-probe with the
gap-sized window (54 bytes here, not 640) would make reverse-order streaming
cost n reads instead of 2n−1 (ZIP-7; no format scenario exhibits the order). The
64 KiB central-directory scratch is zero-filled although the directory size is
known (9,724 bytes here), and flate2's 32 KiB input buffer splits payload reads
(118 payload requests for 90 parts) (ZIP-8). **SIMD: nothing to propose.** No
retained OOXML read profile shows a CRC or inflate loop; on the open profile
inflate is about 13%, dominated by Huffman table construction because members
are tiny, and `crc32fast` already dispatches to PCLMULQDQ (ZIP-9).

**Not opportunities.** Lazy `.rels` (0577); wider read-ahead (0572/0493);
removing or widening the strict proof (0575/0580/0583); memoising descriptor
framing (0561's correction: zero reads removed); retaining probed windows
(0580); passthrough (0578); C1/C2/C3 (0581); a shared pool (ADR 0005); `mmap`
(ADR 0005 names mmap-like owners as `ReadAt` adapters in `litchi-core`, not the
ZIP layer); skipping CRC (ADR 0006; 0349 keeps it even on the borrowed stored
path); stored-borrowed access on source-backed archives (positional by
construction).

**Blockers and a validation asymmetry.** No file-backed cold-part-read
selector exists, so the observation counts are measured on memory and priced by
model on files. `IndexedArchive::read_entry` accepts a member on the 30-byte
header, sizes, CRC and directory bound only (`archive.rs:2038-2107`: no
local-versus-central name, flag or method comparison, no neighbour
disjointness) while `stream_to` runs the full strict proof, so the same member
can be accepted by `data()` and refused by `stream_to()`; 0583's
`residual-descriptor-flag-disagreement.zip` is a ready witness. Reported, not
proposed.

### OOXML save and publication (`litchi-opc` writer, `soapberry-zip` writer, format saves)

**Path.** Every ordinary documented save (`docs/OFFICE_API_GUIDE.md`) ends in
`atomic::replace_with_impl` (`crates/litchi-opc/src/atomic.rs:48-93`): a
tempfile in the destination directory, a bare `&mut File` with no buffering
anywhere on the path, `sync_all`, rename, directory `sync_all`. The semantic
step differs by format: `litchi-xlsx`'s `Edit::commit` rewrites only touched
parts (**measured:** 2 of 44 members for a three-cell edit, 1 of 132 for a tab
hide); `litchi-docx`'s `write_plain` regenerates `document.xml` and every
model-owned part whenever the single whole-model `is_modified` flag is set
(`crates/litchi-docx/src/writer/doc/model.rs:737`); `litchi-pptx`'s
`materialize_presentation` (`crates/litchi-pptx/src/package/codec.rs:395-470`)
deletes every slide and notes part and regenerates all of them although a
per-slide `is_modified` exists (`writer/slide.rs:382`). `PublicationPlan`
(`crates/litchi-opc/src/pkgwriter.rs:110-168`) then rebuilds and audits
`[Content_Types].xml` from all parts, reserializes and audits `_rels/.rels`, and
for **every** part with relationships reserializes its `.rels` and audits it,
changed or not; `try_write_preserved` (`:189-706`) re-parses the resident
source, rebuilds the `PreservationIndex` (deferred from open by 0313),
byte-compares each blob against its provenance `Arc` without a `ptr_eq`
shortcut, byte-compares each `.rels` against the canonical form captured at open,
and raw-copies unchanged members through a 64 KiB stack buffer. Regenerated
members are deflated fully into a `Vec` through a one-entry writer and a
**fresh** level-6 `DeflateEncoder` each (`crates/soapberry-zip/src/preserve.rs:2167`),
then the mini-archive is re-parsed to locate framing. Central records are one
small write each.

**Measured** (`results/change-0587/opc-save/cg-*.txt`, `strace-cfs-*.txt`): a
tab hide on the 132-member workbook (654,688 bytes) raw-copies 131 members and
regenerates `xl/workbook.xml`; the save is 5,325,511 Ir, 14.0% of the
37,998,187 Ir open-edit-save (eager `load_parts_eager` 49.3%, `Edit::commit`
16.4%). Inside the save: the physical writer 31%, `verify_authored` 18.6% in
**42 calls**, `try_to_xml_bytes` 9.3% in **41 calls**, `memcmp` 357,671 Ir in
631 calls. The real `save(path)` route issues **531 `write(2)` calls** for
654,681 bytes (399 of them 64 bytes or less) and 2 fsyncs. A three-cell edit on a
44-member workbook spends about 1.6 M Ir on the deflate side for 3.8 KB of XML,
fixed per-encoder cost dominating. On the 6.8 MB-XML `no_drawing_patriarch.xlsx`
a tab rename's save is 0.49% of the operation and `Edit::commit` 94.4% (1.07 G
Ir): the save is not the eager XLSX edit bottleneck, `Edit::commit` is
(0520/0521's owner). The bounded path has 45 `publish_*_to_stream` entry points
(DOCX 24, XLSX 16, PPTX 5) and 64-part ceilings; `.blob()` has 1,049 call sites
under `src/` (OPC 70, common 47, DOCX 128, XLSX 278, PPTX 390, XLSB 136),
`iter_parts()` 289.

**SAVE-1. Pristine-member proof reuse in the publication plan** (step 1). For
every part whose relationships are unchanged the plan reserializes the `.rels`
XML, audits it, then discards it, because `try_write_preserved` byte-compares it
with the open-time canonical form and copies the source member; the same happens
to `[Content_Types].xml` when nothing changed, and the blob compare `memcmp`s an
`Arc` against itself. **Measured:** 40 of the 42 audits and all 41
reserializations on the hide save are for unchanged members, about 1.7-1.9 M Ir,
**32-36% of the save** and 4.5-5% of the whole operation; proportionally larger
on PPTX (one `.rels` per slide, layout and master). Fix: a revision counter on
`Relationships` (`rel.rs:218-224`, six `&mut self` methods) recorded at open,
skip the reserialization and audit when unchanged, compute
`content_types_changed` from provenance before building the XML, `Arc::ptr_eq`
before `memcmp`. Records: 0519 exempts unchanged XML *parts* only; 0489/0517 are
the splice route; none for unchanged relationships. ADR 0006 unaffected, since
the audited bytes are never published. Risk low; no ADR. Falsified if an A/B on
the workbook shows under 2% save-instruction reduction.

**SAVE-2. Buffer the temporary file inside `atomic::replace`** (step 2). A 64
KiB `BufWriter` between the counting sink and the tempfile (`atomic.rs:75`)
takes the 531 syscalls to about 11, and to 1-2 for a typical DOCX, without
touching caller-supplied sinks' `IncompleteOutput.written` accounting.
**Modelled:** 0.8-1.6 ms per 132-member save at 1.5-3 µs per page-cache write,
against 1-2 ms of save CPU; the two fsyncs (79-82% of the tiny route on a real
store, 0490) are unchanged. Records: none (0490 counted 3,205 writes per child
and attributed the medians to `fdatasync`). Risk low. Falsified if write-syscall
time is under 5% of save wall time after fsync.

**SAVE-3. The PPTX eager save regenerates every slide** (step 1). Retaining
unmodified slides' source parts as copies requires keeping their names and
relationship ids, a change in the mutable model rather than the writer. Size
**unmeasured** — no example opens a real deck and edits through the model;
modelled as N slides times serialize, audit, encoder init and deflate versus 1.
Records: none (0475/0478 cover the streaming writer; 0158 the source-backed
route). Risk medium-high; frozen design; joint with the DOCX/PPTX area.

**SAVE-4. Fresh deflate state per regenerated member** (step 2). 0476's reuse
covers the owned writer only; the preservation writer constructs a new encoder
per member. **Measured:** about 1.6 M Ir for a 3.8 KB member, init and flush
dominating; modelled 0.3-1 M Ir plus one ~300 KB allocation per extra member —
DOCX multi-part saves and SAVE-3 pay it. Risk low; output bytes identical.

**SAVE-5. C2′: lazy decode behind the fallible accessors** (step 2). 0581's C2
costs 1,049 `.blob()` sites and C3 needs a general save that does not exist. A
narrower shape decodes on first access inside `get_part`/`get_part_mut`/
`main_document_part` (already `Result`), adds `try_iter_parts`, and lets the
publication plan treat never-decoded parts as pristine copies without decoding;
the migration surface is the 289 `iter_parts()` sites. Retention stays 0581's C2
column; the save-side gain is the 131 blob compares and, at open,
`load_parts_eager` at 49% of the hide operation. Gate 1 (a proposed ADR) and
gate 2 (the `PartBytes`/`TotalPartBytes` refusals at open; declared central
sizes can be charged at open with a decode-time recheck, a reviewed relocation)
apply. Recommendation: C2′ first, C3 as the end state. Risk medium.

**SAVE-6.** Compression is fixed at level 6 everywhere and no record has
measured level; per-member parallel deflate under an explicit context keeps
bytes identical but pays only at two or more regenerated members of 256 KiB or
more (0498/0499: four 1 MiB parts 576 → 82 µs at width 4; sixty-four 16 KiB
parts regress), so one-cell and one-paragraph saves gain nothing. Low priority.

**Not opportunities.** Skipping or deferring fsync (0490/0497); caching the
`PreservationIndex` across saves (0313 deferred it deliberately; 4.7% of one
save); precompressed passthrough (0578); wider copy chunks (memcpy-bound from a
resident slice); skipping audits of *changed* XML (ADR 0006, 0528); C1 (0581);
intra-member parallel deflate (changes bytes, fails 0581 gate 3); dropping
`source_archive` retention (ADR 0005 exact-source rule); partial output on a
preservation refusal (rule 3).

**Blockers and one defect.** **Defect (reported, not fixed):** the ordinary
eager XLSX save refuses two ordinary edits on an Excel-produced file. Renaming
or activating a sheet of `ConditionalFormattingSamples.xlsx` through the `tabs`
example fails with `XmlPublication { part: "/docProps/app.xml", source:
NotCompact(Violation { kind: FormattingWhitespace, offset: 55 }) }` (and the same
on `/xl/worksheets/sheet10.xml`): offset 55 is the `\r\n` Excel writes after the
XML declaration, which the rewriters preserve and `validate_authored_xml`
(`pkgwriter.rs:988`) then refuses for the whole save; hide and cell edits
succeed. It is recorded nowhere and blocks measuring sheet-view edits on real
producer files. There is no harness scenario for the ordinary save at all; the
only vehicles are examples. And the ordinary path performs **no candidate
readback** (`litchi-docx` `codec.rs:1193`, `litchi-pptx` `codec.rs:296`,
`litchi-xlsx` `writer.rs:33`) where ADR 0003 requires typed readback for staged
CRUD and the bounded path pays it (0489): a compliance observation, and a cost
C3 would add that the ordinary path does not carry today.

### XLSX (`litchi-xlsx`)

**Measured sizes:** `Cell` 72 bytes, `Value` 24, `Formula` 72, `Text` 16
(`Arc<str>`), `Number` 16 (`Box<str>`), `Address` 8; the crate-private `Stored`
is modelled at 184 bytes per stored cell plus one heap numeral per numeric cell
(`results/change-0587/xlsx/sizes-output.txt`). On the eager door the first
`cell()` parses and retains the whole worksheet: on the dense-wide 2×256×256
corpus, open is 1.41 ms and the first cell 34.5 ms, a 65,536-cell parse at about
0.50 µs per cell, 24× the open (`xlsx/alloc-read.json`). On the source-backed
door a **one-cell edit costs about 2.5 µs per cell of the touched sheet**, which
is roughly six whole-sheet passes — planning validate-and-parse, the commit's
layout scan, output validation, the reduced readback parse and merge, two
publication audits and deflate: on the dense-sparse corpus, planning 116,915
allocations and 13.65 MB (7.1 allocations per cell of the sheet), commit 151,522
and 12.0 MB (9.2 per cell), publication 33,342 and 4.18 MB, p50 41.1 ms; the
eager door's one-cell commit and save on dense-wide is 180 ms, or 2.7 µs per cell
(`xlsx/alloc-commit.json`, `alloc-eager-commit.json`). The phase allocation
counters reproduce change 0538's packet within 0.2%.

**XLSX-1. Carry compact per-cell source facts from planning into commit so the
commit's second whole-sheet layout scan disappears** (steps 1 and 2). Planning
already visits every row and cell with positions; commit re-scans the same bytes
into a `Layout` (`crates/litchi-xlsx/src/raw/worksheet/edit/package.rs:163-260`,
`scan.rs:830-854`). **Retained:** 53.24-54.94% of commit instructions (0550),
commit about 45% of the edit interval (0520); **modelled** at most about 25% of a
one-percent edit interval. Records: 0550 names it next, 0551 selected compact
cell facts over layout caching and narrowed the prerequisites, 0553 rejected a
commit-local variant, 0514/0516 rejected fusion; nothing is implemented. Risk
high; the frozen design exists and needs its differential oracle first. Falsified
if the proof builder's retained state costs more than the scan on sparse or
refusal rows, the 0553 outcome.

**XLSX-2. Extend the 0525 omission readback and the source-backed value editor to
real-producer worksheets** (step 1). `stored_entry_is_supported`
(`cell_values/snapshot.rs:365-371`) disables the 96% reduced-parse win whenever
*any* cell of the source or output is a shared string, and planning refuses any
worksheet that carries a relationship (`:441-443`). **Measured over the 95 real
`.xlsx` fixtures:** 77 have `sharedStrings.xml`, 57 have worksheet relationships,
63 have `x14ac`/`dyDescent`/`AlternateContent` in their first sheet (which also
disables change 0546's shared traversal); every harness XLSX corpus has
`shared_strings: null`, so the retained commit evidence never exercises these
fallbacks. Records: 0525 leaves shared-string boundaries authoritative;
0542-0545 rejected widening the traversal; none on relationship-bearing
worksheets. Size: 0525's 96% parser-instruction and 30% commit-p50 reduction,
currently zero on any Excel-produced file; absolute unknown. Risk medium; needs a
frozen design and an SST-bearing harness shape. Falsified if the complete
candidate parse is under 10% of commit instructions on an SST-bearing input.
**Corrected by change 0602:** the mechanism above is wrong. The value editor
admits none of the 95 real fixtures — 93 refuse at the package-root relationship
allow-list, two gates before `stored_entry_is_supported`, whose four clauses are
all refused upstream and cannot fire — and the publication audit of the
*original* part bytes (defect 1 below) refuses 94 of 95 real packages outright.
XLSX-2 is therefore an admission-surface widening gated on a `litchi-opc`
compactness-contract prerequisite, not a readback widening; 0602 measured the
complete candidate parse at 32.9-85.3% of the commits that pay it, so the
falsification condition is not met.

**XLSX-3. Reuse the publication audit of the unchanged original** (step 2).
`litchi-opc` audits both the original and the replacement bytes of every
replaced XML part (`crates/litchi-opc/src/source_backed.rs:7696-7697`,
`8121-8122`), and `litchi-xlsx` never supplies a `SourceXmlPart` proof, so the
0519 skip (`:7687-7691`) never fires for XLSX. **Retained:** `validate_overlay_xml`
is 56.2% of publication instructions (0528); **modelled** at most about 28% of
publication, which is about 11% of a one-cell edit's allocations. Two new
mechanisms: memoize the original-audit verdict per (source lineage, part name)
inside `SourceBackedPackage`, or a proof-carrying replacement door, which needs a
proposed ADR. 0528 requires both audits to stay; this reuses, it does not drop.
Risk medium.

**XLSX-4. Overlay the candidate `Store` instead of clone-and-resort** (steps 2
and 4). `merge_omitted_cells` (`raw/cell.rs:802-850`) clones every omitted
`Stored`, one `Box<str>` per numeric cell, then re-sorts and rebuilds indexes.
**Retained:** provenance merge 4.64-6.03% of commit instructions (0550), Store
merge 5.97% (0525); **modelled** about 16,000 of the 151,522 one-edit commit
allocations. Commits are capped at 256 actions (`source.rs:24`), so an
`Arc<Store>` plus a sorted overlay is bounded. Risk medium; design record needed.

**XLSX-5 to 8, smaller.** `capture_auxiliary_source` parses the whole stylesheet
at every planning for a count (`snapshot.rs:1974-1980`; unknown size, invisible on
the harness's minimal styles); every value commit rewrites and re-validates
`workbook.xml` for recalculation even when the workbook has no formulas
(`snapshot.rs:761-770`; an ADR 0006 discussion, since non-Excel producers omit
`calcChain`); `visit_cells` (`workbook/source.rs:777-800`) materializes the whole
`Vec<SourceCell>` before visiting, modelled at over 5.2 MB for 65,536 cells; and
inlining short numerals in `Number` would remove one of about seven planning and
nine commit allocations per cell, with 0539 as the likely outcome.

**Not opportunities.** Edited-row-only validation (ADR 0005; 0540/0541/0546, and
0551 shows edited rows cover 16,769 of 17,792 cells); fusing candidate validation
with the writer (0514/0516/0521); the row arena, combined tag scan, transient
attribute borrowing and commit-local proof (0526/0527, 0522, 0539, 0553); widening
the shared traversal (0542-0545); dropping a publication audit (0528);
recompressing unchanged members (already passthrough, 0578); the eager door's
retention (0581) and any eager lazy-SST work it subsumes; widening the
validated-store handoff (0025/0466); streaming-writer memory (bounded, 0432).

**Blockers.** The harness XLSX corpora carry no shared-string part and
integer-only payloads; the source-backed value editor refuses 57 of 95 real
fixtures (worksheet relationships), plus unknown cells and cell metadata; the
1,311-update workload exists only on the eager door. One item to verify rather
than a confirmed defect: publication applies the authored-compactness audit to the
*original* part bytes, so a real-producer worksheet with indentation whitespace
could be refused at publication after planning accepted it.

### DOCX and PPTX facades (`litchi-docx`, `litchi-pptx`)

Fresh callgrind of three harness selectors at one sample each
(`results/change-0587/docx-pptx/incl-*.txt`), with native `perf record`
self-symbol shares quoted beside them because callgrind runs software SHA-256
and overstates hashing about fivefold against SHA-NI. The harness corpora are
synthetic and the PPTX eager corpora are built with `Package::from_bytes`, which
carries no physical provenance, so every eager save in those profiles
re-deflates all media where production `open`/`from_vec` copies unchanged
members; the shares below are stated for the regions that finding does not
touch.

**DOCX path.** `Package::open` parses nothing of `document.xml`. Every
`document()` call runs the MCE visibility pass and then builds a
`ParagraphIndex` by a full `scan_word_element_ranges` pass
(`crates/litchi-docx/src/parts/document_part.rs:576-584`, `:50`) before any
query, although `text()`, `write_text_to`, `tables()` and `blocks()` never use
it. **Measured:** `DocumentPart::from_part` 165.8 M Ir over 6 calls (index 114.5
M, MCE 51.2 M) against `Document::text` 151.9 M over the same 6 calls — the
eager index alone costs about 75% of a full text extraction. The ordinary
one-paragraph edit and save (`edit_document` → `commit` →
`publish_document_edit`; `document/transaction.rs:1834-1910`, `4410-4460`,
`7241-7275`) scans the whole main part **four times** — the snapshot, the
rewrite's rescan, the rescan after `compact_changed_document_xml` over the whole
part, and a second `document_snapshot` built only to feed `Patch::apply`'s
`same_source` byte compare — and compacts the whole document. **Measured**
(`docx_semantic_one_edit_save`, 24/200/10,000 paragraphs summed):
`Snapshot::from_xml` 280.4 M Ir in 12 calls, about **64%** of the roughly 440 M
timed instructions; compaction 62.6 M (14%); the one-paragraph rewrite and its
readback about 1 M.

**PPTX path.** The eager open parses the presentation root only; but
`Presentation::slide(i)` re-parses `presentation.xml` on **every** call
(`presentation/package.rs:158-167`; 2.95 M Ir per call on a 200-slide deck,
proportional to N), `slide_count()` resolves and content-type-checks every
slide part, and `find_slide(Key::Name)` parses every slide's XML. The opened
transaction (`opened/model.rs:308-440`, `opened/transaction.rs:1195-1220`,
`opened/patch.rs:753-870`) hashes the complete package — every part blob,
media included — **four times per lifecycle**: at capture, in `commit` to decide
`unsign()`, in `commit`'s recapture, and again in
`apply_opened_presentation_commit`, which discards `commit.snapshot` and calls
`apply` rather than `apply_exact_revision`; the notes and slide-name scans run
three times. **Measured** (`pptx_eager_batch_edit_save`, 200 slides with 8
text boxes each and 8 × 2 MiB PNGs, 39.6 G Ir whole child): `package_fingerprint`
13.84 G, **34.9%**, in 16 calls of about 0.865 G each; `capture_with_provenance`
13.19 G (33.3%) in 13 calls; the two fingerprints are about 86% of each
`Transaction::commit`; native `x86_sha::compress` is 10.8% of whole-child
cycles. The cross-package slide copy (`opened/cross_copy_plan.rs:880-920` and
the lines cited in the retained report) serializes the whole source, destination
and candidate through hashing or `Vec` sinks at least five times and hashes all
blobs at least six times: **measured** (`pptx_cross_copy_media_rich`, 166 G Ir
whole child) SHA-256 79.5%, `package_fingerprint` 77.2 G (46.5%),
`physical_package_fingerprint` 43.0 G (25.9%, 37 full `write_to_stream`
serializations), `bounded_package_bytes` 21.7 G (13.1%, 8 re-deflations of the
candidate), while the closure inventory itself is under 0.5%; native whole-child
SHA-NI 28.0%.

**PPTX-1. One complete-package revision per opened transaction, not four**
(step 1). (a) `commit` reuses the fingerprint it just computed when `unsign()`
changed nothing; (b) `apply_opened_presentation_commit` keeps `commit.snapshot`
and uses `apply_exact_revision` plus `validate_after`; (c) memoize per-part
digests keyed by blob `Arc` identity and length inside `Snapshot`, so a
recapture hashes only parts whose `Arc` changed (the revision becomes a hash of
sorted per-part digests, with the `litchi-pptx-opened-v1` domain string
bumped); (d) build `SlideNameIndex` lazily and cache the notes index per
(presentation blob, slide blob) identity. Records: none for
`package_fingerprint` or `capture_with_provenance`; 0501 left 27.95% of
whole-child SHA-256 unattributed after removing the media input from
`digest_touched`; 0449 assigned about half of SHA period to untimed harness
output hashing. **Modelled** from the measured calls: about 3.46 G of the 4.0 G
Ir in open-transaction, commit and apply per lifecycle on the media-rich
corpus; on media-free decks the notes and name scans (about 130 M Ir per
capture on 200 slides) dominate instead. Scenarios: every PPTX opened-document
CRUD selector (`pptx_eager_batch_edit_save`,
`pptx_eager_multi_slide_batch_edit_save`, `pptx_slide_*_boundary_save`,
`pptx_cross_copy_*`). Constraints: ADR 0003 and 0006 require the revision to
bind the complete package; (a) and (b) reuse a value computed on the identical
`OpcPackage` instance, (c) changes the proof format and needs a frozen design
record, (d) must keep ADR 0013's notes-topology checks before any mutation.
Risk low for (a)/(b), medium for (c)/(d). Falsified if native `perf stat` of the
timed lifecycle on both corpora shows the four passes below the 4% floor after
(a) and (b), or a signature-policy test requires the post-unsign rehash (a)
skips.

**PPTX-2. Cache physical and semantic revisions in the cross-package slide
copy** (steps 1 and 2). Cache `physical_package_fingerprint` per immutable
`Snapshot`, compute the candidate's physical revision once with one hashing
sink instead of `bounded_package_bytes` plus a second serialization, and skip
re-proving source and destination at apply when their revisions already match
the plan. Records: 0454 introduced the physical proof unpriced; 0501 scoped only
the touched digest. **Measured:** 141.9 G of 166 G Ir (85%) whole child; verify
whether `build_candidate` loses provenance, since the candidate serializations
appear to re-deflate the 16 MiB of media. Risk medium; frozen design record on
proof-reuse rules. Falsified if, with plan and apply timers, caching leaves the
media-rich p50 within noise.

**DOCX-1. Stop rescanning and recompacting the whole document on a
one-paragraph edit** (step 1). (a) Compare `main.blob()` against
`patch.before.xml_bytes()` by pointer or `memcmp` and reuse `patch.after`
instead of building a fourth `Snapshot`; (b) after a same-structure paragraph
rewrite, update the layout incrementally (shift ranges after the edit, grow the
enclosing table and control ranges) instead of `from_xml`; (c) compact only the
replacement fragment before splicing, which also removes the post-compaction
rescan. Records: 0500 batched the *managed* scalar route; 0518/0519 reuse proofs
on the source-backed publication path; none covers the eager route's four scans
or the whole-document compaction. **Measured:** about 64% plus 14% of the timed
instructions above; 20.8 ms cold for 10,000 paragraphs. Constraints: readback
and `same_source` stay (ADR 0005); (c) changes output bytes for documents whose
untouched paragraphs contain compactable whitespace — closer to ADR 0006's
preservation-by-default, but a behaviour change needing a frozen record and
preservation tests. Risk low for (a), medium for (b) and (c). Falsified if the
large one-edit p50 improves under 4% after (a) and (b).

**DOCX-2. Build the paragraph index lazily** (step 1). A `OnceCell` filled on
first `paragraph(i)` or `paragraph_count()`; the source-backed `document()`
(`source_backed.rs:625-645`) has the same shape; the MCE pass must stay.
**Measured:** 114.5 M Ir per 6 calls, about 19 M per call. Records: 0283
introduced the index for selected-paragraph queries; none considers text-only
reads. Risk low; no ADR. Falsified if the full-text harness timer starts after
`document()`, in which case the gain is real for callers but invisible to the
selector.

**PPTX-3. Memoize the eager slide catalog** (steps 1 and 4). Cache
`slide_references` in `PresentationPart` and let `slide_count` reuse the
validated catalog. **Measured** per call as above; whole-scenario impact unknown
because the harness calls `slide(i)` once per sample. Records: 0120, 0375
(source-backed); none on the eager view. Risk low.

**PPTX-4 / DOCX-3.** Text extraction runs three passes per slide
(`parts/slide.rs:808-850`: raw budget scan, MCE, parse) and DOCX `write_text_to`
runs `preflight_semantic_xml` before its parse; fusing budget checks into the
parse is the pattern 0514/0516 rejected for XLSX on measurement, so it is listed
unmeasured.

**Not opportunities.** Skipping `.rels` reads (0577); C2/C3 (0581); dropping
readback or `same_source` (ADR 0005); removing `fdatasync` (0490); byte-equal
media dedup on slide copy (0454: byte equality does not prove equivalent
ownership); parallel small local part reads (0498/0499); a precomputed
incoming-relationship index — 94 reverse `iter_parts()` scans exist in
production code but none appears in any profiled top symbol, so there is no
evidence yet that a scenario is bound by them; `OpcPackage` clones (about 1.1 M
Ir on 200 slides, `Arc` blobs).

**Blockers and observations.** No selector times PPTX capture, commit, apply
and publication separately, and no eager-DOCX phase diagnostics exist (0496's
overlay is the managed route), so the shares are callgrind-modelled and the
native figures are whole-child. An `open`/`from_vec` variant of
`pptx_eager_batch_edit_save` is needed before any eager PPTX save timing is
believed. Two observations, not fixed: the ordinary DOCX `commit` compacts
whitespace across untouched paragraphs (`transaction.rs:4422-4429`), to be
checked against ADR 0006's intent; and `apply_opened_presentation_commit`
re-derives the snapshot with `apply` rather than `apply_exact_revision`
(`package/model.rs:413-418`).

### XLSB (`litchi-xlsb`)

The first XLSB performance evidence in the program: all sixteen XLSB records
(0304-0376) carry `performance_claim: none` and are correctness records. Every
cell-CRUD entry point goes through the eager `OpcPackage` door and a fully eager
`Workbook` (`crates/litchi-xlsb/src/workbook/model.rs:20-40`) that parses shared
strings, styles, formula context, pivot caches, tables, chart sheets and
connections at open. Both worksheet representations (`cell_values::worksheet::
read_shared`, `cell_values/worksheet.rs:1563-1651`; `sheet::Worksheet`,
`workbook/codec/codec.rs:207-217`) materialize every cell before returning one,
and nothing is cached across `cell_values()` calls. The source-backed path
(`workbook/source.rs`, change 0304) exists and works, but only feeds a sequential
text writer, reachable through the facade's `.text()` alone. **Measured** on
`testVarious.xlsb` (22,715 bytes, 17 parts, 48 cells; single leg, `xlsb_crud`):
source-backed `full_text` 137 µs against 1.50-2.84 ms for every eager scenario.

**XLSB-1. A commit clones and fully reparses the workbook even for a proven
no-op, and twice for a real edit** (step 1). `Workbook::apply_cell_values`
(`crates/litchi-xlsb/src/workbook/package.rs:137-153`) clones the package before
knowing whether anything changes and reparses the whole `Workbook` afterwards even
when `apply_with_external_link_limits` (`cell_values/workbook.rs:60-95`) took its
byte-identical no-op return (`:66-70`); a real edit clones again (`:73`), reparses
into a throwaway `Workbook` for `validate_dependencies` (`:76-78`, `:92`), discards
it, and the caller reparses the same bytes a second time. **Measured:** read-only
p50 1.506 ms; no-op commit and save 2.766 ms (+84% for zero changed bytes);
one-cell edit and save 4.189 ms. Records: none; nearest precedent 0525. Risk low
for the no-op branch, low-medium for reusing the validated parse. Falsified if
`validate_dependencies` rather than the reparse dominates the delta.

**XLSB-2.** The eager door is 0581's question and is not re-priced here; the new
fact is that XLSB is the one OOXML format with a *shipped* deferred reader, so a
`cell_values`-shaped materialization on that path is a design question inside one
crate rather than a new ADR. **XLSB-3.** A bounded single-cell read has no early
exit and the `unique_index` duplicate-coordinate contract to preserve; it is the
XLSB twin of 0574 opportunity 6 and is recorded as ADR 0005-blocked, not proposed.

**Blocker.** The corpus ceiling: the largest of the fifteen `.xlsb` fixtures is
22,715 bytes with one sheet and 48 cells, so every scaling question is buried
under a 1.5 ms open floor (one cell and a full scan are 0.3% apart). The
`xlsb_crud` harness is wall-clock only, with no instruction or allocation
counters. A synthetic large XLSB shape is the prerequisite for any XLSB batch.

### CFB substrate (`litchi-cfb`)

For this corpus the substrate's open is at its floor. Its remaining open-time
symbols are the ADR 0005 mandatory validations (`validate_stream_allocations`
1.31%, `validate_physical_sector_layout` 1.38%, `collect_exact` 3.94% of the
flagship open, all rejected as closure-proportional candidates by 0574
opportunity 6 and by 0524/0548/0549/0533) and thrice-rejected micro-work. The
per-read version fence as a syscall storm on file sources is **refuted by
measurement**: a flagship open in `file-source` mode is 34 `statx` and 53
`pread64` (`results/change-0587/cfb/strace-summary.txt`); 0558/0560 already
single-fence and the read-ahead layer absorbs the roughly 2,099 logical range
reads. Every CFB-attributable cost above noise is on the *slurp* paths the eager
readers use, and every one of those is a whole-stream zero-fill. A census over
212 parsed fixtures finds none above 1.6 MB, none with a DIFAT sector, none with
4,096-byte sectors, and at most 67 directory entries (`cfb/cfb_open_model.txt`).

**CFB-1. Append-style stream reads, so the whole-stream zero-fill never exists**
(step 2). `read_stream_from_fat` (`crates/litchi-cfb/src/file.rs:2133`, the DOC
site) and `read_fat_stream` (`crates/litchi-cfb/src/shared.rs:2427`, the PPT
site, which also feeds `load_ministream`) allocate `size` zeroed bytes that the
run reads then overwrite in full except a truncated final sector. **Measured:**
`memset` instructions per open equal the bytes of the streams slurped whole —
1,595,476 Ir against 1,595,422 bytes on the 1.6 MB DOC (98.5% of the file),
315,656 against 315,620 on the PPT; the two callers are 32.15% / 9.13% / 0.98% of
the three DOC opens and 21.6% / 26.7% of the two PPT opens (retained 0584
caller attribution). The safe shape without `unsafe`: a provided `ReadAt`
method that appends into a caller's `Vec<u8>` (default keeps today's zero-fill;
`OwnedSource`/`SliceSource`/`OwnedArcSource` override with `extend_from_slice`),
and an appending `read_chain_into` that pushes each physical run and
zero-extends only the missing tail, keeping 0570's typed refusal for a run at or
past EOF. `FileSource` cannot avoid it on stable Rust (`read_at` needs an
initialized slice), so on real files only the tail trim (CFB-4) applies.
Records: 0574 opportunity 4 names only the XLS site, 0584 candidate 5 names the
CFB sites and the rule 10 block; no record designs the safe shape. Constraints:
GOAL rules 10, 11 (a `litchi-core` trait addition), 12 (`try_reserve` before
append); the three truncated-final-sector zero-fill tests are the parity gate.
Risk medium; frozen design record, no ADR. Falsified if a cycles A/B of the DOC
and PPT opens on an owned source sits inside the 4% p50 floor — plausible, since
a 1.6 MB `memset` is 60-80 µs on this host against an open of 1-2 ms; the
callgrind share is a `rep stosb` upper bound.

**CFB-2. A copy-through writer for length-changing OLE2 saves** (steps 1 and 2;
GOAL's "LEGACY CFB-SPECIFIC WORK"). `OleWriter` (`writer/core.rs:250-1222`) is a
from-scratch builder: every OLE2 edit-save in `litchi-xls`, `litchi-doc` and
`litchi-ppt` decodes every unchanged stream (zero-fill plus copy), retains all of
them until `write_to`, and re-serializes header, FAT, DIFAT, MiniFAT, directory
and every stream; the overlay path (`overlay.rs:651-870`) covers same-length
edits only. A copy-through writer would stream untouched streams' sectors
verbatim and rebuild only the allocation tables, directory and changed streams.
Invariants it must prove: every physical sector claimed once and every unclaimed
sector `FREESECT`; FAT/DIFAT counts and markers; directory red-black order and
SID ownership (ADR 0026); v3 size masking; mini-stream cutoff migration;
byte-identical untouched streams (rule 2); deterministic output; reopen through
the normal parser before a sink observes a byte (the 0175 policy). Nothing in
ADR 0005/0006 forbids it; whether `FREESECT` reuse versus append is
preservation-neutral is an open policy question. Records: none for CFB
length-changing copy-through (0387/0420 are OPC; 0103/0142/0143/0172/0175 are
same-length; 0003/0036 are the builder). Size **unknown**: no OLE2
length-changing save selector exists; modelled, peak memory O(file) to
O(bounded buffer) and one fewer decode per unchanged stream, bytes written
unchanged. Risk high; frozen design plus an ADR clarification of physical-layout
policy. Falsified if a phase attribution of one XLS/DOC/PPT save shows the
crates' own record re-encoding, not the container copy, dominates.

**CFB-3 to 6, small on this corpus.** The eager stream readers' `collect_sector_chain`
(`file.rs:2774-2812`) allocates a FAT-sized zeroed bitset and a full-chain `Vec`
per `open_stream` and walks to end-of-chain regardless of declared size, where
the shared `read_chain_into` walks exactly `required` sectors with no allocation:
below noise here (at most 2,738 FAT entries), 256 KiB plus 8 MiB per `open_stream`
on a 1 GiB v3 file, measurable only with a synthetic large fixture (CFB-3).
`read_sector_into` and `read_sector_run_into` fill whole buffers then overwrite
`[..present]`: 0.3-0.5% of DOC/PPT open instructions, to bundle with CFB-1 and
never land alone (CFB-4). `load_ministream` copies the ministream twice more
(`Vec` into `Arc<[u8]>`, then per mini-sector into a zeroed buffer,
`shared.rs:2480-2536`): zero on this corpus, whose ministreams are 128 B to
4,928 B and stay unmaterialized under the 0146-0152 direct-read policy (CFB-5). A
resolved-entry handle for repeated reads would remove `directory_name_data` plus
`find_entry` at 0.90% of a flagship one-cell query (CFB-6).

**Not opportunities.** Coalescing reads across non-contiguous runs (0574
opportunity 7, ADR 0006 ownership); a single directory parse or open-time name
handoff (0554, every primary p50 regressed; the second parse also backs the
public `DirectoryEntry` API and ADR 0026); lazy CLSID formatting (already
empty for all-zero CLSIDs); fewer hash passes on the generic-source overlay
(0143 coalesced, 0175 deliberately withheld generic sources because a
`SourceVersion` fence is weaker than the content guarantee: an ADR 0006 policy
amendment, not an optimization); parallel bulk reads of DOC's three open-time
streams (small tasks regress; the bulk session has no production caller); the
DIFAT loop (0 of 212 fixtures).

**Blockers.** No registered selector opens a DOC or PPT or performs any OLE2
length-changing save (0584's numbers came from a throwaway driver). Callgrind
runs `sha2` on its software backend because valgrind masks the SHA CPUID bit:
the retained flagship profile's 53% `sha256::soft` is the harness's per-sample
identity check, and any instruction pricing of the overlay fingerprints is
wrong — use `perf stat` cycles (see DOC-1). `xls_source_attribution --operation
one-cell` defaults to worksheet index 1 and fails on single-sheet fixtures such
as `54016.xls` unless `--worksheet-index 0` is passed.

### XLS (`litchi-xls`)

**Path.** The open (`crates/litchi-xls/src/workbook/source.rs:1013`) frames every
globals record once into `GlobalsBuffer`, retains all globals bytes, frames them
a second time (`BiffRecords::with_limits`, `:1830`) and interprets twelve record
kinds; the SST is measure-walked (0576) into 24-byte segments and 16-byte
entries. A one-cell query (`query_cell`, `:2735`) builds a cursor with a cold
chain walk (599 links on `54016.xls`, up to 2,477 on the flagship's last sheet),
then loops `next_frame` **to end of stream**, parsing every cell and resolving
only the target; nothing is retained, so "all cells" is N full re-scans and no
whole-sheet iterator exists. Full text runs the same scan per sheet, decoding
every cell into a `HashMap<(u16,u16), CellValue>` with three hash operations and
two eagerly built error values per insert (`:664-716`), resolving each shared
string through a linear segment scan, a `Vec<Vec<u8>>` of zero-filled chunks and
**three source observations** (`:3018`, `:3070`, `:3109`), and then walks the
declared `DIMENSIONS` rectangle with one hash lookup per position and one
observation per row. On a `from_path` workbook every observation is an `fstat`;
the harness never sees this because every XLS selector wraps bytes. **Measured**
type sizes (`results/change-0587/xls/type-sizes-summary.txt`): `SourceBackedError`
48 bytes, `Result<WorksheetFrame, SourceBackedError>` 48, `WorksheetFrame` 16,
`CellRecord` 88. The retained 0584 figures: on `54016.xls` the open is 7,239,535
Ir of which the five SST symbols are 3,581,589 (49.5%, **454 Ir per string**) and
the one-cell query 24,525,331 Ir framing 37,929 records for one value.

**XLS-1. A lean worksheet frame loop, including 0584's candidate 2, which never
landed** (steps 1 and 3). Two `ok_or({ SourceBackedError::ResourceLimit{..} })`
calls in `next_frame` (`:2307`, `:2321`) build a 48-byte error on **every** frame
and drop it on the `Some` path — 0584's 75,858 drops per query are exactly
2 × 37,929 frames — and each frame also pays two non-inlined `ensure` calls, one
`consume_payload` and two `check_execution`, each returning a 48-byte `Result`
through memory. **Retained:** `drop_in_place<SourceBackedError>` is 4.20% of the
`54016` one-cell query, and framing overhead in total (`next_frame` 17.32%,
`ensure` 7.17%, `read_payload` 6.07%, `query_cell` self 10.67%, the drops,
`skip_payload`) is **46.1%** of the query against 18.0% for cell semantics. Fix:
`ok_or_else`, an inlined already-resident fast path, cold-path error
construction; same checks, same order, no refusal moves. Records: 0584
candidate 2 was proposed and, by `git log` on `source.rs`, never implemented
(57b25a820 is the 0585 hint). Risk low; no design record; price in cycles.
Falsified if paired callgrind and `perf stat` on the `54016` one-cell query show
under 2% of instructions and no cycle change.

**XLS-2. Lazy SST indexing, 0584 candidate 3, and its cheap sibling** (step 1).
`Continue` boundaries are already recorded at open for free — `segments` is
built from framed records (`records.rs:1170-1183`) before any string is walked
— so only `entries` is deferrable, as a prefix index under a lock in the snapshot
that ADR 0005's lazy-payload clause permits and `max_sst_entries` still bounds;
the header checks (`:1190-1205`) can stay at open, so only per-string refusals
move to first resolve past the defect, and open-and-list callers never see them
(0576's error-identity trap). **Retained:** 49.5% of the `54016` open, 6.6% of
the flagship; a saving only when few strings are resolved, since a full text
resolves all 16,055 and pays the walk anyway. Requires its own frozen design
record (0576, 0584). The cheaper sibling **XLS-2b** keeps the eager walk and cuts
its 454 Ir per string: a fast path when header and characters lie within one
segment (direct indexing instead of `read_exact` into a 2-byte `copy_from_slice`,
`records.rs:694-718`; 7,900 such calls per open) and a measure instantiation of
`read_formatting_runs` without its per-string `Vec` (`:824-830`). Modelled at up
to half the 49.5%; 0576's differential harness proves error identity; no design
record. Falsified if segment-boundary logic dominates.

**XLS-3. A snapshot-scoped retained sheet index and a whole-sheet iterator**
(step 4). The first validated scan records per cell its stream offset,
`(row, col)`, kind and XF (16-24 bytes per cell; at most 0.7 MB on `54016`) plus a
validated-to-EOF mark, so later queries seek within the validated range; even
without the index, a public `SourceBackedWorksheet::cells()` turns N scans into
one and is the missing all-cells selector. ADR 0005: an index keyed by
`expected_version` on an immutable snapshot is a clean-value cache, but it must
be weighted, bounded by `max_worksheet_scan_records` and evictable, and
`litchi-xls` has no such cache today — a frozen design record. **Retained per
query:** the scan is 17,285,796 Ir (70.5%) on `54016` and 237,761 (9.1%) on the
flagship, N-fold across repeated queries. This is not 0574 opportunity 6: the
first scan still validates to EOF; 0584 called the retained scan "the tractable
form". Risk medium. Falsified if a second query is not at least 5× cheaper than a
re-scan, or the retained weight breaks the bounded-memory contract.

**XLS-4. Full text: per-string and per-row freshness fences, and the rectangle
walk** (steps 2 and 1). Three observations per string and one per row are, on a
file source, at least 48,165 `fstat` calls for one `54016` text extraction
(**modelled**; zero in the harness, where `version()` is a field read). ADR 0005
requires that mutation during a read return `SourceChanged`; fences at scan
start and end detect the same mutations but move *when* the refusal surfaces, so
a short design note is needed. Emitting retained cells in sorted order instead
of walking the declared rectangle produces identical bytes. Records: none.
Risk low-medium. Falsified if `strace -c` on a `from_path` extraction puts
fences under 5% of wall time.

**XLS-5. Measure-only validation of non-target cells in `query_cell`** (step
2). `CellRecord::parse` materializes `Label { value: String }` and `Formula {
formula: Vec<u8> }` (`records.rs:2124`, `:2150`) for every record and
`process_cell` drops all but the target; `drop_in_place<CellRecord>` is 2.05% of
the `54016` query. A `MeasuredText`-style instantiation (0576's pattern)
validates without allocating. Size modelled; small where the mix is `LabelSst`
and RK, larger on formula-heavy sheets, and the corpus mix is unknown because
the 0584 census counted `LabelSst` only. Risk low-medium.

**XLS-6. Skip never-interpreted globals payloads and frame once** (0574
opportunities 2 and 5; step 2). Still open; the consumed kinds are confirmed in
the current tree. Unlike the worksheet gate (untestable, 0568), the globals
density gate can fire on a real fixture: the flagship's `MsoDrawingGroup` chain
is 524,839 bytes over about 261 records, about 2 KiB per record, against a 1 KiB
threshold (modelled from 0574 and the census; verify). **Retained:** about 30%
of the flagship open as an instruction upper bound, 5.1% of aggregate open time,
at a cost of about 46 more requests, a range-source regression. Risk
medium-high; frozen design record. Falsified if the 103 dense-globals fixtures
regress in cycles more than the sparse ones gain.

**XLS-7. Zero-fill of the fill buffers** (0574 opportunity 4): 23.0% of the
flagship open (`GlobalsBuffer::ensure`'s `memset`, upper bound), 3.7% of the
`54016` query. No append-style read exists anywhere in `litchi-cfb` or
`litchi-core`. The safe form helps only in-memory sources, which is the harness
and not `from_path`, so it must not land on harness evidence alone; see CFB-1
and defer behind XLS-6.

**XLS-8. The per-sheet cursor construction hint** (0585 limitation; step 4):
28,143 links on the 16-sheet flagship, 91% resumable, 0 on single-sheet
fixtures; needs a second `StreamChainHint` with a lifetime disjoint from the
resolver's. Multi-sheet text only; measure on `HyperlinksOnManySheets.xls` or
`WithCustomViews.xls`, since the flagship's text extraction is refused.

**XLS-9. Edit and save: the candidate readback is a complete eager open**
(step 1). `Snapshot::from_bytes` (`cell_values/mod.rs:516`) runs a complete
eager `Workbook::new`; the generic commit copies the workbook stream and the
file into a second `PackageEditor` and rebuilds the snapshot; the plan-only
fixed-width path (0137/0138/0168/0172) still runs `Workbook::new` over the
complete composed target (`:5230`). HOTSPOTS rank 13's "second complete target
artifact" is gone only for plan-only Number and RK families, not for string,
formula, style or structural edits. A source-backed candidate readback (globals
plus a scan of the edited sheets) would replace it if the three coverage,
protection and macro requirements (`:5231-5233`) are proven equivalent on the
source-backed owner. Size **unknown**: no attribution of the XLS edit path
exists (0138 has only p50s: Number 105.3 ms plan-only against 145.4 ms
source-backed). ADR 0006 keeps the readback mandatory; changing its owner needs a
proposed ADR or frozen record. Risk medium-high. Falsified if a callgrind of
`xls_numeric_plan_only_*` shows `Workbook::new` under 15% of the commit.

**XLS-10.** `resolve_shared_string`'s per-resolve allocations (chunk `Vec`s, a
`slices` `Vec`, a linear segment scan) belong in the 0585 `SharedStringResolver`
as a reused scratch buffer and a binary search; 16,055 resolves per `54016`
text; unmeasured. Low risk.

**Not opportunities.** Early exit or `INDEX`/`DBCELL` seeking in `query_cell`
(0574 opportunity 6; XLS-3 is the sanctioned form); `directory_name_data`
handoff (0554); collector variants (0524/0548/0549/0533/0536); pending-role
accounting (0555); the DOC-style resolve hint (0586); one chain hint shared by
resolver and cursor (thrashes, 0585); windows wider than 64 KiB (0568, 0572);
zero-fill via `MaybeUninit` (rule 10); a decoded `Vec<String>` SST cache (0300
retains offsets deliberately); parallel per-sheet scans.

**Blockers.** No source-backed XLS full-text or all-cells selector exists
(`xls_source_attribution` supports `open`, `list` and `one-cell` only), so
XLS-3, 4, 8 and 10 cannot be sized today; every XLS harness source is in-memory,
so `FileSource::version` costs are invisible; the flagship's text extraction is
refused (`Invalid record 0x0006: shared Formula metadata requires a leading PtgExp
token`) and whether the eager path refuses the same file is unrecorded; no
attribution of the XLS edit/save path exists; the corpus Formula/Label record
mix is unknown.

### DOC and PPT (`litchi-doc`, `litchi-ppt`)

Change 0584's open question — why a 65 KB DOC costs 9% more instructions to
open than a 1.6 MB one — is answered. The eager `Document` open
(`crates/litchi-doc/src/document/package.rs:68-330`, which is what the harness's
`doc_semantic_open` and both validation passes of every `doc_semantic_*_edit_save`
run) is the sum of three terms with different drivers: text decoding, proportional
to text units; PAPX/CHPX resolution, proportional to formatting entries; and stream
zero-fill plus copy, proportional to bytes. The 65 KB fixture has three times the
paragraphs and eighteen times the text of the 1.6 MB one, whose bytes are mostly
an 893 KB table stream and pictures (`results/change-0587/doc-ppt/survey-doc.tsv`).
Measured per open, fresh callgrind isolation pairs: `saved-by-table.doc`
5,582,903 Ir (`TextExtractor::new` 42.9%, `PapBinTable::parse` 37.7%);
`FloatingPictures.doc` 3,467,311 Ir; the 1.6 MB form 4,971,503 Ir (`open_stream`
zero-fill 35.4%, PAP 22.4%, `FileInformationBlock::parse` 14.05%).

**DOC-1. Source-backed DOC and PPT opens are dominated by whole-artifact SHA-256,
computed two to six times per open** (step 1). `overlay::fingerprints`
(`crates/litchi-cfb/src/overlay.rs:1320-1344`) streams the whole artifact through
*two* hashers even when the splice list is empty and source and target are the
same bytes; `SourceSnapshot::open_with_options`
(`crates/litchi-doc/src/body_text/source.rs:499-538`) computes the identity three
times (`:526`, `:591`, `:599`) and parses the CFB index twice (`:510`, `:532`);
`ensure_current` hashes twice more (`:766`, `:774`) and `write_validated` twice
again (`overlay.rs:913-940`). The PPT `text_edit::SourceSnapshot::open`
(`crates/litchi-ppt/src/text_edit.rs:399-449`) does no semantic work at all: it is
one `plan_splices(Vec::new())`, two SHA-256 passes. **Measured:** the DOC snapshot
open of a 64 KB fixture is 41,107,039 Ir, 99.6% in `identity_fingerprint`; of
`picture.doc` (1.4 MB) 912,551,293 Ir, 99.2% SHA-256; the PPT text-edit snapshot
open of a 385 KB file is 81,168,004 Ir against 1,181,516 for the *full eager
presentation parse* of the same file. Callgrind runs the software SHA-256 while
the host has SHA-NI, so the record prices this in cycles: native `perf stat` puts
the PPT text-edit open at 3.31 M cycles per operation against 0.25 M for the eager
open (13.4×), and on the two DOC fixtures both readers accept the snapshot open
costs 2.5× and 6.5× the eager open (`doc-ppt/perf-stat-native.txt`). Change 0586's
zero was structural (it counted chain links and removed none), so hashing did
not mask it, as change 0589 corrects; what this term does explain is that no
read hint on the snapshot path could have registered in latency against 41 M
instructions of hashing. Records: 0100/0105/0119 designed the complete-artifact
fingerprint, 0143 coalesced its reads, 0165 made the owned editor's fingerprint
lazy; none prices the CPU cost. Constraints: this is a redesign of the ADR 0006
source-identity fence, so it needs a frozen design record and a re-run of change
0582's differential harness; a one-hasher empty-splice plan halves every identity
pass without touching the fence's semantics. Risk medium. Falsified if native
cycles per `SourceSnapshot::open` do not fall by the removed passes' share, or the
fence fuzz finds a stable-token mutation the reduced fence misses.

**DOC-2. DOC text extraction makes four passes and one `memcpy` per character**
(step 2). `TextExtractor::new` (`crates/litchi-doc/src/parts/text.rs:45-56`,
`:359-420`) decodes every piece into an unsized `Vec<u16>` with one
`extend_from_slice` per ANSI character, then `from_utf16_lossy`, then
`encode_utf16().count()`, then an 8-bytes-per-code-unit `cp_to_byte` table
(`:117-125`). **Measured:** 2,396,879 Ir per open on `saved-by-table.doc` (42.9%),
10,722 `memcpy` calls. Records: none. Risk low; text must stay byte-identical
(the lossy surrogate handling at `text.rs:98-105`). Falsified if that fixture's
open falls by under 10%.

**DOC-3. Each PAPX entry is resolved about three times** (step 1).
`PapBinTable::parse` (`crates/litchi-doc/src/parts/pap_bin_table.rs:57-140`)
concatenates sprms per entry (`:218`), allocates a `HashSet` per entry (`:196`),
parses the sprms once for the huge-PAPX check, again in `apply_direct_sprms`
(`parts/pap/parser.rs:175`), and re-runs the whole cascade in
`from_sprm_with_stylesheet` (`:186-195`) only to copy five table fields; the
single-entry adjacent-style cache of change 0051 misses whenever consecutive
paragraphs alternate styles. **Measured:** 37.7%, 28.7% and 22.4% of the three
opens; 1,931 `parse_sprms` calls for about 626 entries, 103 baseline
re-resolutions for 17 FKP pages, 548 `ParagraphProperties::clone`. Records: 0051
(cache), 0056 (containment index); nothing on the double cascade. Risk
low-medium. Falsified if `parse_sprms` calls per entry do not fall to one.

**DOC-4. The FIB owns a copy of the whole `WordDocument` stream** (step 2).
`FileInformationBlock::parse_at` (`crates/litchi-doc/src/parts/fib.rs:107`) does
`data.to_vec()` over the stream suffix although its doc comment (`:71-75`) says it
owns only the FIB; the glossary FIB copies again. **Measured:** 698,295 Ir per
open of the 1.6 MB fixture, one copy of its 697,827-byte stream (14.05%). Records:
none. Risk low.

**PPT-1. The PPT record tree copies every byte once per nesting level and retains
the stream as well** (steps 2 and 3). `Record::parse_impl`
(`crates/litchi-ppt/src/records/record.rs:175-200`) copies each record payload
into an owned `Vec<u8>` and then parses children from the same slice; the retained
`Presentation` also keeps the stream (`presentation/model.rs:133-148`).
**Measured:** owned payload bytes are 1.32× to 2.01× the stream across three
fixtures (`doc-ppt/ppt-tree.txt`), so retention is roughly 2.3× to 3× the stream;
on `45543.ppt` (385 KB, 1,181,516 Ir per open) `memcpy` is 30.4%, the record
parser 47.2%, and 877 `memcpy` calls. `Record.data` is a public field, so this is
an API change inside `litchi-ppt`. Records: none for PPT; 0199/0200 did the
analogue for ODS events; 0301 leaves the complete document parser out of scope.
Risk medium.

**PPT-2. The presentation parser extracts the full text of every slide at every
open** (step 1). `parse_document_with_limits` ends with
`extract_slide_text_from_document` (`crates/litchi-ppt/src/parsers/parser.rs:83`,
`:108-118`), 259 `extract_text` calls per open, stored in `slide_atoms_sets` and
consumed only by `RecordParser::slides()`/`slide_count()`, which nothing outside
`litchi-ppt` calls. **Retained:** about 2.7% of change 0584's `pptmid` open. Risk
low; `RecordParser` is public, so a deprecation note is needed.

**Deferred, already priced.** Stream zero-fill on open (0574 opportunity 4,
0584 candidate 5) is 32.8% of the 1.6 MB DOC open and 27.6% of the PPT open, the
largest payers in the program, and stays blocked on `docs/GOAL.md` rule 10 until a
reader API appends into the buffer (see CFB-1).

**Not opportunities.** Lazy parsing of the ~45 DOC tables or of the FKPs on a
one-paragraph read: the eager `Document` is both the documented general reader
and the mandatory public-reader gate (`validate_ole_file`, 0160), so it needs the
DOC analogue of change 0576's frozen design, not a shortcut. Borrowing the owned
editor's second validation: 0161 rejected. Hinting `resolve_paragraph`: 0586,
explained by DOC-1. Routing the facade to `SourceSnapshot`: it refuses 49 of 57
fixtures by design and is slower than the eager open today. Parallelism inside a
1-6 M-instruction open: a loss under rule 9.

### Core, facade, execution and parallelism (`litchi-core`, `litchi`, the three parallel sessions)

**Measured costs of the source layer** (`results/change-0587/core-facade-parallel/microbench-results.txt`,
a scratch project on `litchi-core` alone, pinned, best of five): `FileSource::version()`
167 ns (mutex plus `statx` plus fingerprint compare), `len()` 157 ns, a 512-byte
`read_at` 144 ns and a 4 KiB one 161 ns; `OwnedSource::read_at` of 4 KiB 16.8 ns;
`Budget::consume` 6.0 ns at depth 1 and 15.8 ns at depth 3, `reserve` plus drop
44 ns and 99 ns. One source observation therefore costs the same as one 4 KiB
`pread`, and budget charging is not on any OLE2 per-read path at all (two sites
in `litchi-cfb`, both in the bulk session). `ReadAt` is copy-out by contract
(`crates/litchi-core/src/source.rs:84`); there is no borrowing accessor, no
`read_vectored`, and no `mmap` anywhere in the workspace. Hashing is SipHash
everywhere in scope, and no hashing symbol appears in any retained top-25.

**Whole-process syscalls per open** (`core-facade-parallel/strace-*.txt`, 1
against 11 samples): `54016.xls` through the facade on a path costs 42 `pread64`
and 36 `statx`; the flagship 55 and 40. The harness's own attribution of the
file-source open records 25 `version()` calls and 1 `len()` for 317,171 bytes
read of a 984,576-byte file: three fences in `SharedOleFile::open_with_limits`
(`crates/litchi-cfb/src/shared.rs:599`, `:601`, `:623`) and one
`check_source_version` per `read_stream_range_hinted` (`:1047`, twelve call
sites), plus six `ensure_path_source_current` in the facade's XLS route
(`crates/litchi/src/detection_smart/detected.rs:835-1000`) and eight in the PPT
route (`:3069-3150`).

**CORE-1. The facade slurps every `.doc` and parses it eagerly** (steps 1 and
2; GOAL hypothesis 1's last remnant on a priority format). Path opens dispatch
XLSX, XLSB, XLS, DOCX, PPTX and native PPT to source-backed readers over a
`FileSource`, but `DocumentSourcePathDetection` has only `Odt`, `Docx` and
`Bytes` variants (`detected.rs:2156-2171`), so a `.doc` path is read whole into a
`Vec<u8>` (`read_path_source_bytes`, `:1198-1225`, capped at 2 GiB), wrapped in
`OleFile<Cursor<Vec<u8>>>` (`:42`) and parsed eagerly by `doc::Package::from_ole_file`
plus `package.document()` (`document/doc.rs:928-941`); the facade never references
`litchi-doc`'s `SourceSnapshot::open(Arc<dyn ReadAt>)`. Records: 0584 and 0586
cover the readers, HOTSPOTS records the 49-of-57 refusal; none covers the facade
route. Size: the eager open is 3.5-5.4 M Ir per fixture (0584), 26-58% of it
`memset` and `memcpy`, and the whole-file read is a further full copy retained
for the document's lifetime (modelled); the facade delta is unknown because no
harness selector opens a `.doc` through the facade. Constraints: ADR 0005's
positional source and ADR 0006's typed refusals — the source-backed reader must
refuse exactly what the eager reader refuses or fall back per fixture, and
`SourceSnapshot` admits eight of 57 fixtures today, so admission coverage is the
gate (0586) and DOC-1's hashing cost must fall first or the route is slower.
Risk high; frozen design record. Falsified if on the eight admitted fixtures the
source-backed open is not lower in both instructions and peak RSS, or widening
admission requires weakening a validation the eager path performs.

**CORE-2. Collapse the 25 per-open `version()` fences on a file-backed XLS open
to one per operation boundary** (step 2). ADR 0005 requires observing before
work and comparing after; it does not require a `statx` per stream read.
Records: 0558 (two fences per read to one), 0560, 0563 ("24 removed `statx` at
roughly 0.2 µs each explains at most a third of the observed movement"); the
residual 25 is attributed by no record. **Modelled:** 25-36 × 167 ns is 4-6 µs
against an open of about 1.19 M cycles, 1-1.5%, below the p50 floor on a warm
local file — so this is a **count** claim, and its latency relevance is on
network filesystems where `statx` is a round trip. Constraints: the per-read
fence is what makes a mid-operation truncation surface as `SourceChanged`
rather than a garbled read, so fence placement must keep every typed refusal at
the same boundary; a frozen design note. Risk medium. Falsified if moving fences
changes any `SourceChanged` outcome in the change-under-read tests, or the count
cannot drop below about 10 per open.

**CORE-3. `ExecutionContext` completeness, a prerequisite rather than a
speedup.** Against GOAL Workstream F's list, `ExecutionLimits`
(`crates/litchi-core/src/execution.rs:38-44`) has max workers, in-flight task and
byte caps (the decompression-in-flight budget), `min_parallel_bytes`,
cancellation, affinity and the hierarchical `Budget`; it lacks an I/O
concurrency limit (nothing bounds concurrent `read_at`s issued from workers), a
CPU task budget distinct from byte-denominated `Work`, and a caller-provided
executor or scoped-worker facility — each of the three sessions builds its own
pool (`office.rs:276`, `shared_bulk.rs:217`, `batch.rs:517` uses
`std::thread::scope`), so three pools can coexist in one process with no shared
cap. Records: 0009, 0088, 0240, 0498, 0499 use the current shape; none proposes
the missing fields. ADR 0005 names "scheduling, affinity, cancellation, thread
and memory budgets", so adding I/O concurrency and executor injection is an ADR
0005 amendment: proposed ADR required. Risk low.

**CORE-4. Parallel compression of changed members on multi-part saves** (step
5). Every writer is serial (zero `rayon` or `scope` hits under the OPC, ZIP and
CFB writers) and unchanged members are bounded passthrough (0578), so deflating
changed members is the only CPU-heavy save work. The retained scaling evidence
(0009: OPC 4.52× at 12 workers with a serial fraction of about 15%; CFB 5.93×,
about 9.3%; many-small tasks 0.73× and 0.52×) sets the gate at two or more
changed members each at or above `min_parallel_bytes`, which single-part edits
never meet (SAVE-6 is the same item from the save side). Size unknown; frozen
design record; constraints are rule 9, preserved output order, and `OutputBytes`
and cancellation covering in-flight buffers. Falsified if the median multi-part
edit changes fewer than two members above the threshold, or deflate of the
changed set is under 20% of save wall time.

**CORE-5. Invert the effect index in `JoinedSubEdits::join`** (step 4).
`join_failure` iterates every accepted sub-edit per incoming edit with three
`BTreeSet::intersection`s each (`crates/litchi-core/src/patch.rs:1754-1790`,
`:1889-1920`), O(n²·k) bounded only by `CompositionLimits`; in-tree users set
limits of 4-8 sub-edits, so today it is microseconds. Records: none. The CRUD
checklist's "hot paths are indexed" row names accidental quadratic overlap
detection. Risk low; a hygiene item until a composition exceeds about 16
sub-edits.

**CORE-6.** Per-`PartData` `Objects` and `Memory` reservations in OPC batches
(`source_backed/batch.rs:317-330`) are at most 200 ns per part, about 26 µs on
a 132-member workbook; closed by the micro-measurement unless batches reach
thousands of parts. **CORE-7. ILP in serial loops: nothing rankable.** The
non-copy leaders on the flagship one-cell query are `next_chain_sector` 4.06%
and `collect_exact` 3.58%, both true pointer chases where 0579 measured that
extra per-sector state cut instructions 7.7% and raised cycles 3.7% (IPC 3.629
to 3.230); BIFF `Records::next` (1.61%) is length-chained framing. Price any ILP
candidate in cycles with `perf stat` before writing code.

**Not opportunities.** Swapping SipHash for a faster hasher (no hot map in any
profile; ADR 0006 favours DoS-safe defaults); `mmap` (copy-out `ReadAt` means an
mmap source still copies into every consumer, and the XLS open's 42 `pread`s are
about 7 µs); `read_vectored` (no scatter/gather pattern exists; the mandatory
rels reads are 0577's coalescing problem); removing the `version()` mutex (10 ns
of 167); skipping `Budget` on hot paths (not there); parallelizing the open with
the existing sessions (0009, 0498/0499, 0577); caching `FileSource::len()` (its
`statx` is the freshness observation).

**Blockers and notes.** No harness selector opens `.doc` or `.ppt` through the
facade, so CORE-1 cannot be A/B-measured today; lock wait and thread counts are
unavailable by design in the harness envelope (0240) and 0405's seam adds timer
overhead, so contention can only be counted; no real range source exists.
`refine_workbook_format` (`crates/litchi/src/sheet/workbook_types.rs:133-160`)
is a public, dead, whole-input slurp of any `Read` — a latent hypothesis-1
ingress if it is ever wired; `FileSource::from_file` calls `metadata()` twice
(`source/file.rs:93`, `:102`).

### Program evidence and harness gaps

Ranked by how much each gap blocks `docs/GOAL.md`'s DEFINITION OF DONE for
OLE2 and OOXML (`results/change-0587/survey/evidence-gaps.md` has the full
table with file and line citations).

| # | gap | what exists | what is missing |
| ---: | --- | --- | --- |
| 1 | the CI smoke check cannot detect a regression | `.github/workflows/perf-baseline.yml`'s `smoke` job builds `baseline` as a byte-copy of `current` with only the revision label changed, then asserts the comparator passes against itself | any comparison against prior history on push or pull request; the `reference-regression` job, comparator and policy already exist and need a fetched artifact instead of an operator-supplied run id |
| 2 | the harness corpora do not take the real-producer path | generated worksheets carry no MCE markers, `<cols>`, shared strings or relationships; eager PPTX corpora carry no provenance; every source is in memory | a real-file option or a generator shape with markers, an SST and relationships; a `from_vec` PPTX variant; `from_path` selectors |
| 3 | most real DOC fixtures and a share of real XLS, XLSX, DOCX and PPTX fixtures cannot reach the measured path | 8 of 57 `.doc`, 10 of 93 XLSX cell, 1 zero-byte DOCX full text, 1 PPTX genuine middle slide | a first-class refusal census (reason histogram per format) instead of each record re-deriving the counts |
| 4 | no OLE2 range-source coverage | three OOXML mechanisms (`SimulatedRangeSource` over OPC and XLSX, a PPTX module, provider pacing) | any CFB, XLS, DOC or PPT range-source case; the CFB readers already take `ReadAt`, so this is harness plumbing |
| 5 | no selectors for the paths this record ranks highest | XLS `open`, `list`, `one-cell`; source-backed DOCX and PPTX lifecycles | DOC and PPT facade opens, OLE2 length-changing saves, the ordinary OOXML save, XLS full text and all cells, file-backed cold part reads, PPTX capture/commit/apply timers, eager DOCX phases |
| 6 | fuzz targets exist and never execute | 13 targets under `fuzz/` | `cargo-fuzz` and a nightly toolchain on any CI runner or this host (0582) |
| 7 | corpus ceilings | largest OLE2 fixture 1.6 MB, no DIFAT, no 4,096-byte sectors; largest XLSB 22.7 KB; largest marker-bearing worksheet 210 KB | synthetic large CFB past the ~7 MB DIFAT threshold; a large synthetic XLSB; a large real-producer XLSX |
| 8 | hardware counters | instructions, cycles, branches validated; L2 request counters validated locally | L1 aliases return zeroes and LLC events are absent in this guest (0408/0409); not fixable in the repository |
| 9 | contention evidence is scoped to one producer | the OPC managed cache's opt-in mutex observation (0405) | any equivalent for the CFB bulk session; lock wait remains unavailable by design (0240) |
| 10 | the regression-gated corpus is almost entirely synthetic | real-producer breadth exists ad hoc per record | any real-producer fixture in the schema-2 default catalog beyond the RTF watermark |
| 11 | the coverage index is thin | 33 rows over 15 categories; creation-from-scratch has no OOXML row; the security row is XLS and XLSX only | XLSB in the index at all: none of `xlsb_crud`'s 8 selectors is mapped |
| 12 | verification breadth | a 3-OS correctness matrix in `rust-ci.yml` | perf numbers on macOS or Windows; any Miri, sanitizer or loom job |

The claim registry's stall at change 0467 was checked and is **not** a gap:
every record from 0468 to 0586 carries `performance_claim: none` by design. The
cheapest high-value additions, in order: wrap the CFB `ReadAt` sources in the
existing `SimulatedRangeSource` (closes gap 4); fetch the last successful full
run's artifact as the smoke baseline (closes gap 1); a synthetic CFB shape past
the DIFAT threshold; the eight `xlsb_crud` selectors mapped into the coverage
index; a refusal census case per format; and a real-producer worksheet shape in
the generator, which every item in the top half of the queue needs.

## Correctness defects and compliance observations found on the way

Reported, not fixed; none is performance work.

1. **The ordinary eager XLSX save refuses ordinary sheet-view edits on an
   Excel-produced file.** Renaming or activating a sheet of
   `ConditionalFormattingSamples.xlsx` fails with `XmlPublication { part:
   "/docProps/app.xml", source: NotCompact(Violation { kind: FormattingWhitespace,
   offset: 55 }) }`: offset 55 is the `\r\n` Excel writes after the XML
   declaration, which the rewriters preserve and `validate_authored_xml`
   (`crates/litchi-opc/src/pkgwriter.rs:988`) refuses for the whole save. The
   related XLSX source-backed publication audits the *original* part bytes for
   compactness (`crates/litchi-opc/src/source_backed.rs:7696`, `8121`), so a
   worksheet accepted at planning could be refused at publication; to verify.
2. **MCE output limits are enforced on the codec's self-expanded stream.** With
   14-17× expansion, `Limits::default().max_output_bytes`,
   `MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES`, `mce_limit` and `within_mce_limits`
   can refuse a legitimately sized part as oversized output (XML-1's section
   cites the four sites). The eligibility gates also key on substring presence
   anywhere in the part.
3. **A validation asymmetry between `data()` and `stream_to()`.**
   `IndexedArchive::read_entry` accepts a member on the 30-byte header, sizes,
   CRC and directory bound only (`crates/soapberry-zip/src/archive.rs:2038-2107`)
   while the stream and verified readers run the full strict proof, so the same
   member can be accepted by one and refused by the other; 0583's
   `residual-descriptor-flag-disagreement.zip` is a ready witness.
4. **DOC facade refusals on ordinary Word files.** Two field-table consistency
   refusals (`watermark.doc`: `grffldEnd.fNested`; POI `test.doc`: `fHasSep`)
   merit a check under ADR 0006, and five fixtures are refused for duplicate
   style names although `StyleSheet::parse_with_leniency` and
   `OpenOptions.leniency` exist (`crates/litchi-doc/src/document/package.rs:228-230`;
   whether they admit these files is unverified). The 49-of-57 `SourceSnapshot`
   refusals, by contrast, are typed refusals under 0105's contract (Drawing 34,
   Field 4, AmbiguousTopology 4, Encrypted 3, Revision 2, Macro 1, one CFB
   error on a deliberate failure fixture) and are correct.
5. **The ordinary DOCX commit compacts whitespace across untouched paragraphs**
   (`crates/litchi-docx/src/document/transaction.rs:4422-4429`); whether this is
   the intended preservation behaviour under ADR 0006 is a question for the
   owner, and DOCX-1(c) depends on the answer.
6. **The ordinary OOXML save performs no candidate readback** (`litchi-docx`
   `package/codec.rs:1193`, `litchi-pptx` `package/codec.rs:296`, `litchi-xlsx`
   `writer.rs:33`) where ADR 0003 requires typed readback for staged CRUD and
   the bounded path pays it (0489). A compliance observation for human review,
   and a cost 0581's C3 would add that the ordinary path does not carry today.
7. Smaller: `monitor_reads` never clears once set
   (`source_backed.rs:2888`); `refine_workbook_format` is a public, dead,
   whole-input slurp of any `Read` (`crates/litchi/src/sheet/workbook_types.rs:133-160`);
   `FileSource::from_file` calls `metadata()` twice; `fib.rs:71-75`'s doc comment
   contradicts `:107`; `presentation/model.rs:137` carries a wrong `dead_code`
   annotation and `RecordParser::slide_count()` is semantically wrong
   (`parsers/parser.rs:112`); `xls_source_attribution --operation one-cell`
   defaults to worksheet index 1 and fails on single-sheet fixtures without
   `--worksheet-index 0`.

## Looks like an opportunity but is not

Each area section carries its own list; the items most likely to be
re-proposed by a reader of the code are collected here with the record that
settles them. Early exit from mandatory validation, in `query_cell` or any CFB
whole-container check (0574 opportunity 6, ADR 0005, rule 12). Lazy or skipped
`.rels` reads at open (0577, ADR 0005/0006). An infallible lazy `Part::blob`
(0581, C1 refused). Dropping either publication audit (0528). Skipping or
deferring `fsync` (0490/0497). Wider source read-ahead windows (0572, 0493).
Coalescing CFB reads across non-contiguous runs (0574 opportunity 7, ADR 0006).
`collect_exact`, `directory_name_data` handoff, pending-role accounting and the
DOC resolve hint (0524/0548/0549/0533/0554/0555/0586). The XLSX row arena,
combined tag scan, transient attribute borrowing, event preflight, exact-bound
scanner and commit-local proof (0522/0526/0527/0529/0539/0544/0545/0553).
`memmem` for the MCE presence scan alone (0531; subsumed by XML-1). A shared
Rayon pool or parallel small local reads (ADR 0005, 0009, 0498/0499). `mmap`
(copy-out `ReadAt`; ADR 0001 and `docs/GOAL.md` restrict it to an audited
low-level owner after proof; no borrowing consumer exists). Swapping SipHash
(no hot map in any profile; ADR 0006). SIMD anywhere (no retained profile shows
a hot scanning, escaping, UTF-8 or CRC loop; the hot `esc` is XML-1's expansion).

## Measurement blockers

Consolidated from the area sections. **Corpora:** the harness corpora miss the
real-producer paths (finding 1); the largest OLE2 fixture is 1.6 MB with no
DIFAT or 4,096-byte sector; the largest XLSB is 22.7 KB with one sheet and 48
cells; the largest marker-bearing worksheet is 210 KB. **Selectors:** none for
DOC or PPT facade opens, OLE2 length-changing saves, the ordinary OOXML save,
XLS full text or all cells, file-backed cold part reads, PPTX phase timers or
eager DOCX phases; every XLS and OOXML source in the harness is in memory.
**Refusals:** 49 of 57 `.doc` by the source snapshot (typed and correct) and 15
of 57 by the facade; 57 of 95 XLSX fixtures by the source-backed value editor
(worksheet relationships) and 10 of 93 by the cell scenario; the flagship XLS
fixture's text extraction (`Invalid record 0x0006`); 4 of 30 PPT (encrypted).
**Instruments:** callgrind prices SHA-256 in software and bulk copies per byte;
L1 and LLC events are unavailable in this guest; no `cargo-fuzz`; the ineligible
selected-cell case and the 2n−1 reverse-order regression are unobservable
through any format-crate scenario.

## ADR compliance

Nothing was implemented, so no boundary moved. The table records, for each
ranked item, which accepted ADR governs it and what it needs before code; it is
the compliance matrix `docs/GOAL.md` deliverable 4 asks for, filled in advance.

| items | governing ADR or rule | verdict |
| --- | --- | --- |
| XML-1, XML-2, XML-3 | ADR 0005 mandatory validation (preserved by the fallback parser); ADR 0006 (read-side transform, nothing published); 0541 error-order guards | aligned; frozen design record each; XML-1 must confirm no writer publishes the processed stream |
| DOC-1 | ADR 0006 source-identity fence; ADR 0005 `SourceChanged` | a fence redesign, not a weakening: frozen design record, 0582 harness re-run |
| PPTX-1, PPTX-2 | ADR 0003 revision binding; ADR 0006; ADR 0013 notes topology | (a)/(b) reuse a value computed on the identical package: aligned; (c) changes the proof format: frozen design; PPTX-2: frozen design on proof reuse |
| DOCX-1, DOCX-2 | ADR 0005 readback and `same_source` stay; ADR 0006 for (c) | (a)/(b) and DOCX-2 aligned; (c) a behaviour change needing a frozen record and the owner's answer to defect 5 |
| SAVE-1, SAVE-2, SAVE-4, ZIP-1, ZIP-4, XLS-1, XLS-2b, DOC-2/3/4, XLSB-1 | ADR 0006 preservation (bytes unchanged); ADR 0003 atomicity for XLSB-1 | aligned; paired measurement only |
| ZIP-2, ZIP-3, ZIP-5 | ADR 0005 limits and error identity; ADR 0011 ownership (facts, not framing, cross the boundary); 0317 precedence | frozen design records; ZIP-5 is 0577's record |
| XLS-2, XLS-3, XLS-6, XLS-4, CORE-2 | ADR 0005 lazy payloads, weighted evictable caches, `SourceChanged` semantics | aligned in principle, each moves *when* a refusal or observation happens: frozen design records |
| XLSX-1, XLSX-2, XLSX-3, XML-5 | 0541 guards; ADR 0006 relationship preservation; 0528 both audits stay | XLSX-1: 0551's design; XLSX-2: frozen design; XLSX-3 memo: design, proof door: proposed ADR |
| CFB-1 | rules 10, 11, 12 (no unsafe; a `litchi-core` trait addition; `try_reserve`) | frozen design record; no ADR |
| CFB-2 | ADR 0026, ADR 0006, ADR 0005 reopen-before-sink (0175) | frozen design plus an ADR clarification of physical-layout policy |
| SAVE-5 | ADR 0005 lazy payloads (favours); 0581 gates 1 and 2 | proposed ADR |
| CORE-1, XLS-9 | ADR 0005 positional source; ADR 0006 typed refusals and mandatory readback | proposed ADR or frozen record; CORE-1 after DOC-1 |
| CORE-3, CORE-4 | ADR 0005 execution budgets; rule 9 | CORE-3 proposed ADR (0005 amendment); CORE-4 frozen design |
| SAVE-3, PPT-1 | ADR 0006 preservation; public type changes inside one crate | measurement, then frozen design |

## Limitations

Every fresh number is a single leg: warm page cache, one host, no A/A control,
no host-quiescence log, taken while other survey agents were building and
measuring on the same machine. Instruction counts rank work, not latency;
callgrind's per-byte copy accounting and software SHA-256 are stated once above
and apply throughout; the DOCX/PPTX figures are whole-child shares because no
selector times those phases separately. XML-1 is measured on **one** real
fixture and its breadth is a count, not a size: the 12.9× is the size of the
mechanism on that file, and the corpus-wide figure is the first measurement the
design record must take. No timing, cold-cache, physical-device, real
range-source, peak-RSS, concurrency or cross-platform result is claimed anywhere
in this record.

The survey is not a proof of absence. Each area was read under a token-economy
cap after the first wave was terminated by a session usage limit, so the areas
are covered at the depth their reports show and no deeper; encrypted, signed
and macro-bearing paths (`litchi-crypto`, `litchi-sign`, `litchi-vba`),
conversion and export sinks, and the DOCX and PPTX streaming writers beyond
0473-0478 were not surveyed. The coordinator verified the eight mechanisms named
in the method section against the source and did not re-run any agent's
measurement; the raw outputs are retained so that anyone can.

The ranking is a judgment. Reach is estimated from which scenarios and which
share of the real fixtures a mechanism touches, size from the tiered figures,
and risk from the ADR reading above; a natively priced top item can move down
and a currently unknown one up. What the ranking is not is a plan: the
prerequisite column says what has to exist before any of it becomes one.

## Retained evidence

[`results/change-0587/`](results/change-0587/README.md): the eleven survey
reports as returned, under `survey/`; each area's cited outputs under its own
directory; `decision.json`; and `cleanup.json` recording what was removed.
