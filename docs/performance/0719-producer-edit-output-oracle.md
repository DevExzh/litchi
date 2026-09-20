# 0719 — verify producer-shaped XLSX edit/save output

`performance_claim: none`.

The existing producer edit/save benchmark checked source immutability and
repeatable output hashes. Those checks can accept a consistently wrong output:
a stable archive need not contain the requested edit. This batch strengthens
that benchmark's untimed output oracle before it is used to gate further
optimizations. Production code, selector names, generated corpora and the
sheet-0 numeric edit target remain unchanged.

The timer covers worksheet planning, staging the scalar edit, commit and
sequential publication. Editor opening and counting-sink reservation occur
before it. Publication consumes the editor and drops its returned temporary
snapshot inside the interval; remaining commit and sink destruction, source
hashing and output checks occur afterwards. This is not an open-through-drop
lifecycle benchmark. The shared-string worksheet is an untouched control,
not the edited worksheet. These generated packages model producer features;
they are not independent Office-produced files or Office round-trip evidence.

The semantic oracle reopens source and output through the source-backed XLSX
reader and compares all stored cells over `Rect::ALL`, along with worksheet
count and names. Only the selected target may differ, and it must contain the
expected numeric replacement. This includes cells beyond the generated grid,
so an unexpected extra stored cell cannot evade the check.

For the edited worksheet, a separate generated-corpus oracle constructs the
expected XML by replacing exactly one known source numeric-cell fragment.
The whole decoded worksheet must equal that splice. Dimension, root, row,
namespace and all non-target lexical bytes therefore remain checked, rather
than being inferred from equal cell values. A focused corruption test inserts
semantically harmless whitespace outside the target and requires this exact
check to refuse it while semantic readback still passes.

The package oracle requires an identical member set. Only `xl/workbook.xml`
and `xl/worksheets/sheet1.xml` may change. Every other member must retain its
raw local record and central record, except the relocated local-record offset.
This includes shared strings, unselected sheets, relationships, content types
and auxiliary parts present in these corpora. Within the workbook, stripping
the direct `calcPr` element must leave identical bytes; its typed properties
must equal the required calculation invalidation. The check does not compare
archive member ordering or certify arbitrary ZIP64 layouts.

All extra work is outside the measured interval, but it can affect later allocator state. Therefore unchanged timer boundaries
do not make old and new harness latency captures interchangeable. A future
candidate requires a fresh matched baseline and independent process repeats.

## Verification scope

Four successful release children cover medium/dense in forward and reverse
order, with one warmup and three retained samples per child. These are
correctness captures, not sufficient samples for latency claims. The medium
source corpus and published digest match the retained 0705 control exactly.
Dense output identity also agrees across both fresh children. No before/after
performance comparison is made.

The first oracle build was superseded after adding the exact worksheet check.
The first executable correctness attempt then exposed an overly strict oracle
assumption: the generated workbook already has the required invalidation
flags, so `workbook.xml` need not change. The final oracle permits that exact
unchanged case while retaining the typed flag and outside-`calcPr` byte checks.
The selected worksheet must change. Both earlier source/build records and the
failed capture remain archived; they do not contribute to the four successful
children.

The final standalone harness passes 535 default-library tests with zero
failures and one existing ignored test, including four focused oracle tests.
Formatting, all-target Clippy with the allocator/process feature combination
and warnings denied, and warning-denied library rustdoc pass. An earlier
quality attempt passed the same tests but failed Clippy on an unused `mut`
in the test helper; that attempt and the final correction remain archived.
All six repository evidence gates also pass: crate boundaries, strict and
structural performance claims, report classification, CRUD coverage and the
non-iWork boundary. The packet audit binds the final source, commands, logs,
corpus/output identities, archived attempts and cleanup witness.
These checks cover the benchmark change, not a fresh whole-workspace or native
Office certification. No fuzz, cross-platform or physical-I/O claim is added.

## Current work selection

The ranking in `HOTSPOTS.md` stopped its XLS evidence at 0687. Source review
and later records show that 0688–0690 already implement checked-link inlining,
SST chain checkpoints and borrowed ASCII directory lookups. The 0690 retained
first-cell profile assigns 12.05% self samples to chain walking; the obsolete
51–59% figures must not drive another first-cell redesign. A late 54016 target
still traverses 1,044 worksheet links in 0689's diagnostic. Any new checkpoint
needs measured benefit, explicit weight and unchanged replay/refusal fences.

The next concrete DOCX work-elimination experiment is a shared structural walk
for altChunk metadata and active paragraph/table ranges. The current two MCE
selection calls and their ordering must remain: unioning their inputs is not
proven equivalent at marked-byte/count limits or on malformed input. This is
separate from the rejected 0711 event-borrowing and 0715 section-collection
pilots. [The qualification plan](results/change-0719/next-experiment.md)
records the exact ownership and parity obligations. No fusion is implemented
or measured in this batch.

XLSX planning and publication remain substantial measured work in 0705/0707;
the rejected 0706 attribute-cardinality and 0708 name-storage pilots stay
rejected. The 0718 allocation result does not justify compressor pooling or an
allocator-policy change. The broader non-iWork goal remains active.

[Evidence and reproduction](results/change-0719/README.md).
