# Non-iWork `docs/GOAL.md` audit

## 0519: reuse the immutable OPC publication XML proof

[0519](changes/0519-opc-publication-xml-proof-reuse.md) reuses complete XML
validation when destination and proof `ReadLimits` are exactly equal. The
publication owner retains Work, source/context, PartBytes, destination identity,
security, topology, transfer, and DOCX candidate readback checks. Different
limits retain full XML validation; matching limits no longer reserve unused
parser workspace.

Across 48 matched synthetic DOCX comparisons, lifecycle p50 improves
1.43–26.41% and publication p50 improves 36.31–80.85%. Scoped publication
instructions fall 47.33–47.37% for p128 and 77.47–77.53% for p512. The separate
allocator probe records 109→96 calls and 75,001→74,218 allocated bytes;
incremental region peak remains 72,535 bytes. Work, source I/O, exact output,
and release counters are unchanged. RSS has no >5% adverse flag.

All 107 phase flags remain explicit, including two second-campaign lifecycle
p99 regressions dominated by edit outliers. The separate eight-child tail
guard improves lifecycle p99 by 14.47–18.07% in all four comparisons, while
retaining three short commit-phase flags and all original flags. All nine
quality gates pass, including 5,012 executed tests. No general tail-latency,
cold/range, native Office producer, or scaling improvement is claimed. The
remaining local publication profile is mostly ZIP preservation and copying;
its instruction count alone does not prove further removable work. OLE2/OOXML
remains active, ODF deferred, and iWork excluded.

## 0518: reuse the retained DOCX source snapshot at publication

[0518](changes/0518-docx-source-snapshot-reuse.md) reuses the immutable
patch source after OPC checks the current Part, source identity, exact bytes,
limits, and security state. Candidate reparse/readback and destination
publication validation remain intact. Across 48 same-API comparisons,
lifecycle p50 improves 1.42–34.59% and publication p50 improves 43.75–67.42%.
Method instruction counts fall about 54–66%. In the separate instrumented
probe, publication allocation calls fall 96.65–99.13%, while incremental
region peaks improve only 2.70–6.55%.
All lifecycle/publication tails improve; RSS stays within the 5% review
threshold. The 69 adverse open/commit/drop flags remain explicit, with causes
unproven. These results cover the named synthetic DOCX matrix only.

The repeated current-snapshot owner is now negligible in these profiles;
topology publication dominates the remainder and is the next proof-review
candidate. Historical 0499/0500 flags, broader CRUD/corpus coverage, cold/range
and scaling requirements remain open. OLE2/OOXML work continues, ODF is
deferred until that goal is complete, and iWork is excluded.

## 0517: remove a duplicate shared OPC source-XML validation pass

[0517](changes/0517-opc-source-xml-validation.md) retains the initial complete
source/destination/XML proof and replaces the later duplicate scan with a
source freshness and cancellation check. All 48 same-route DOCX comparisons
improve lifecycle p50 by 1.87–14.09%; scoped publication instruction counts fall
about 17.9–21.1%. No adverse lifecycle, publication or whole-child RSS threshold
is triggered; short open/commit/drop flags remain explicit in the report.
This measures synthetic DOCX cases, not every OPC caller. The source-snapshot
construction path remains roughly half to two-thirds of publication CPU and
is the next measured investigation target. Historical 0499/0500 flags and
broader coverage/scaling requirements remain open. OLE2/OOXML remain the
priority; ODF is deferred until that optimization goal is complete.

## 0516: reject XLSX emitted-output parser fusion

[0516](changes/0516-xlsx-output-fusion-rejection.md) passes all 1,284 candidate
all-features tests but fails performance admission: changed-edit pilot p50
rises 4.46–8.30% across the 12 main rows and 5.82–8.64% across the six warm
changed-edit guard rows. Dense allocation calls also increase, and scoped
commit instruction references rise 3.95%. No formal ABBA campaign or speedup
claim follows. The exact candidate and all adverse evidence are archived;
production is restored. Two independent no-op fixture fixes remain, with
1,263 tests passing on unchanged production. OLE2/OOXML work remains active;
ODF is deferred and iWork excluded.

## Current priority after 0519: OLE2 and OOXML

Per the user's instruction, prioritize OLE2 and OOXML performance
until their full optimization goal is complete. Further ODF optimization is
deferred until then. This overrides the ordering of older entries below.
CFB profiling and FAT batching are complete in 0511. The 0514 source-pass
fusion and 0516 emitted-output fusion are rejected. The 0517–0519 DOCX publication
changes remove repeated XML and current-snapshot construction work. Further
publication investigation must distinguish required byte copying from
removable work and compare its end-to-end value with broader OLE2/OOXML
coverage. Required candidate validation and readback remain intact.
The [next-priority review](results/change-0519/next-priority-review.md)
selects current source-backed XLSX one-percent edit/save phase and resource
attribution before another DOCX publication change. Retain the 0499/0500 review flags. This investigation queue is not a completion checklist. The broader requirements and
outstanding coverage remain open; iWork remains outside this workstream.
See the [priority review](results/change-0510/ole2-ooxml-priority-review.md).

## 0515: isolate XLSX changed-output work

[0515](changes/0515-xlsx-output-attribution.md) separates changed-output parsing
from source Store parsing with caller-context profiles and captures compaction
alone. The unchanged-source commit profiles differ by 0.008% in simulated
instructions. Output parsing remains about 25.8% of commit work and compaction
about 20.4%, but sharing the reader directly targets only about 6.9% of commit
instructions. Semantic cell processing, materialization and namespace resolution
remain required; the full validation share is not a removable-work estimate.
The 720 normal samples show p50 repeat drift from −1.01% to +1.10%, with no
frozen latency or RSS review flag. Whole-child RSS is 138,480/135,732 KiB;
this is repeatability context, not an optimization or document-memory result.

The next implementation should investigate feeding emitted-equivalent events
from compaction into the existing parser only for effective changed outputs.
Reuse needs a proof for actual normalized bytes, exact-output x14ac/MCE/UTF-8
checks, deferred error ordering and an authoritative fallback. Style checks,
web validation, change readback, package reopen and the bounded Store handoff
remain intact. See the [semantic review](results/change-0515/output-semantics-review.md)
and [scope review](results/change-0515/scope-review.md). This evidence-only batch
makes no speedup or new phase-memory/hardware/cache/scaling claim. The full
OLE2/OOXML goal remains active; ODF is deferred and iWork excluded.

## 0514: reject speculative XLSX parser/layout fusion

[0514](changes/0514-xlsx-fusion-rejection.md) evaluates a shared-event source
parser and lossless rewrite layout. The candidate passes all 966 XLSX owner
unit tests, including 17 new differential/concurrency tests, but fails the
predeclared admission gate: cold same-value pilot p50 increases 64.43–84.64%
across all six size/update rows. Repeated dense no-op allocation observations
show incremental peak live demand rising 27.04% for one-cell edits and 18.64%
for one-percent edits. A 6.30% reduction in scoped save Callgrind instructions
does not justify that extra work. Changed-edit pilot results are mixed; no
full candidate native comparison or accepted speedup is claimed.

Production and candidate tests are restored exactly to the control revision.
The verified candidate patch, source-bound evidence and reusable public no-op
and cold-read guard remain available. Next work should investigate the larger
changed-output XLSX validation/compaction path or DOCX publication CPU cost;
0511's completed CFB profiling and FAT batching should not be repeated as an
unstarted item. See the [conditional follow-up](results/change-0514/follow-up-options.md).
The full OLE2/OOXML goal remains active, ODF stays deferred, and iWork remains
excluded.

## 0513: XLSX operation allocation baseline

[0513](changes/0513-xlsx-operation-allocation.md) adds allocation observations
to four existing XLSX commit/save cases and a private save profiling boundary.
The 4,800 matched native samples show p50 changes of −0.48% to +2.32%, with
no latency/throughput/RSS or repeat-drift flag. The separate 240 allocator
samples establish candidate-only baselines: dense one-percent commit allocates
273,128,176 bytes in 2,400,578 allocation calls; commit/save allocates
286,872,324 bytes in 2,532,326 calls. Both have 58,496,820 bytes of derived
incremental live demand above region entry. Absolute region peaks differ
because commit-only retains the prior result; neither is document peak or RSS.
This is a measured harness enabler, with no format speedup or memory-reduction
claim. The next production candidate must measure parser/snapshot reuse while
preserving error order, original-source spans and bounded temporary overlap.
The full OLE2/OOXML goal remains active; ODF stays deferred.

## 0512: current XLSX commit attribution

[0512](changes/0512-xlsx-commit-attribution.md) isolates three dense one-percent
commit bodies per profile, resetting after fixture generation. Two captures
differ by 0.012% in simulated instruction references. Direct source Store and
validation parsing consume about 52%, rewrite 27% and compaction 20%; nested
snapshot scanning is 25.40%, while shared-formula resolution is only 0.05%.
These are scoped instruction diagnostics, not phase clocks or a speedup.
Twelve unchanged native cases retain 720 durations across two repeats with no
same-build drift flags. Whole-child hardware/RSS remain separate; operation
allocation and exact save-phase boundaries are still missing. Next work should
add those observations before attempting parser/snapshot fusion with preserved
error order and bounded memory. No production or default-matrix change occurs.

## 0511: CFB FAT helper work reduced; broader goal remains open

[0511](changes/0511-cfb-fat-entry-reservation.md) removes repeated FAT-entry
helper work using the proven exact reservation. Eager XLS and CFB native
opening improve; plain OwnedSource results are mixed. Required chain,
sector-ownership and physical-layout checks retain their work and dominate
the remaining source-constructor profile. FileSource, cold/provider behavior,
operation-local hardware scope and broader scaling coverage remain open.
Continue the OLE2/OOXML investigation queue with dense XLSX commit/save and
DOCX provider/publication costs, revisiting CFB only with fresh attribution
and an exact safety proof. This batch does not complete the full optimization
goal or unlock deferred ODF work. Default coverage remains 41 cases, 213 rows,
43 corpora and 18 mapped correctness-only selectors; iWork remains excluded.

## 0509: profiled ODT export allocation reduction

[0509](changes/0509-odt-sink-buffer-reuse.md) follows fresh allocation evidence
with one bounded operation-local paragraph buffer. Stack allocation calls
fall 99.99% in the large synthetic export, with exact output/sink identities
and 1,491 passing Rust tests/doctests. Native captures retain 14,000 samples,
including a longer tail follow-up after an initial +9.14% large p99 flag.
The follow-up stays below 5%, but no tail-latency or peak-memory improvement
is claimed. This is scoped implementation progress; the remaining 18
correctness-only selectors, provider overhead, native-producer coverage,
historical comparisons and full non-iWork requirements remain open. The
program goal remains active.

## 0508: concrete default conversion coverage expansion

[0508](changes/0508-default-semantic-text-export.md) preserves all 201 prior
default identities and adds twelve measured text-export rows. The default
now has 41 cases/213 rows/43 corpora. Two fresh full matrices validate 6,390
samples; 15 mapped measured selectors bind 60 case/corpus rows, with 18
selectors still correctness-only. The XML-compaction checklist mismatch is
corrected. This advances representative conversion/export coverage without
claiming native-producer, cold-provider or complete checklist certification.
The [ODG audit](results/change-0508/odg-priority-review.md) distinguishes
matched 0504–0507 improvements from the historical old-parser comparison.
Remaining priorities include the 0499 local Part-batch overhead, 0500 K1
latency/RSS flags, the remaining coverage gaps and real-producer evidence.
The program goal remains active.

## Current audit: 0507 reduces repeated ODG value work

[0507](changes/0507-odg-attribute-value-batches.md) records this batch.
The next attributed bottleneck is reduced by two fixed-size request groups,
retaining scalar error precedence, semantic validation order and exact source
preservation. All 124 tests pass; the capture improves median time 17–34%
and plain-large RSS about 9%. This is scoped progress with no >5% adverse
paired flag, not closure of all older-parser comparisons or the broader goal.
The ten-entry strict registry and outstanding CRUD/provider requirements
remain unchanged.

## Current audit: 0506 batches ODG source-span discovery

[0506](changes/0506-odg-shape-attribute-span-batch.md) records this batch.
The measured source-span bottleneck is reduced while exact field ordering,
namespace resolution, checked attributes and fallback errors remain covered.
The 108-test suite includes all 16 field mappings and private error parity.
The plain-large ~10% RSS regression is accepted for the measured latency/work
reduction, with unresolved allocator/working-set attribution explicitly retained.
This does not complete the broader goal, close older-parser comparisons or
change the ten-entry strict registry.

## Current audit: 0505 reduces ODG attribute lookup work

[0505](changes/0505-odg-attribute-name-prefilter.md) records this batch.
A local-name filter retains checked iteration and exact source preservation,
while reducing p50 8.43–9.88% in the matched richer-parser capture. Four new
tests cover namespace selection and trailing invalid attributes. Full CRUD,
provider/scaling coverage and older-parser regression attribution remain open.
The strict claim registry remains at ten entries; this is scoped progress.

## Current audit: 0504 reduces a measured ODG open bottleneck

[0504](changes/0504-odg-direct-transition-reuse.md) implements local reuse of
successfully parsed direct page-transition values. The 3,200-sample matched
richer-parser comparison reduces metadata-large open/traversal p50 by about
34% and metadata-small by about 10%, with no paired >5% adverse latency or
whole-child RSS flag. Per-page inheritance validation remains in force.
This is production and scoped measurement progress, not completion of the
ODG regression review: `parse_content` remains the dominant profiled subtree,
unique-style-heavy input is unmeasured, and the older 0502 baseline exposed
less metadata. Full CRUD, provider, memory-attribution and scaling requirements
remain open; the ten-entry strict registry is unchanged.

## Current audit: 0502 adds bounded ODG metadata but retains large open regressions; the full goal remains open

[0502](changes/0502-odg-metadata-open.md) adds typed ODG transitions and
auxiliary/enhanced shape metadata, with presence-gated validation, parsed-style
reuse, and boxed cold metadata. The committed probe evidence is four
deterministic corpora, two serial repeats, 25 warmups, and 200 samples per
child: 1,600 samples per phase and 3,200 matched samples in
[the retained summary](results/change-0502/summary.json). The timing comparison
is clean HEAD `f3f9221` versus the final working tree; source manifests and all
raw reports are retained under `results/change-0502/`. The summary was
recomputed from those reports and matches the retained JSON.

The final p50 changes are +5.903/+6.386% for plain-small, +6.062/+5.949% for
plain-large, +53.402/+55.215% for metadata-small, and +110.090/+111.136% for
metadata-large across repeats. The p95/p99 rows show the same direction,
including +110% metadata-large tails. Boxing reduces `Shape` from the 936-byte
pre-boxing candidate to 800 bytes, versus 792 bytes at clean HEAD. A separate
whole-child heaptrack comparison reports peak heap 13.69M to 12.57M with the
same 174,853,447 allocation calls; its instrumented p50 and RSS are not
operation-local measurements. Hardware counters were unavailable, so no CPU,
cache, branch, or IPC conclusion follows. The ODG work is therefore a bounded
correctness/layout enabler with an explicit regression queue, not an accepted
end-to-end performance improvement.

| Goal area | 0502 evidence | Audit status and boundary |
| --- | --- | --- |
| ODG open and metadata traversal | 3,200 matched samples over plain and metadata corpora | Measured regression; follow-up parser/representation work is required |
| Layout and allocation shape | `Shape` 792 → 800 bytes; pre-boxing 936; heaptrack peak display −8.2% with unchanged call count | Whole-child/layout evidence only; no allocation or RSS claim for the timed operation |
| Native producer and semantic correctness | Retained ODG tests include LibreOffice fixture checks and bounded/refusal cases | Correctness evidence only; the probe uses generated packages and does not establish native-producer performance |
| Full non-iWork goal | No CRUD selector, edit/save, cold/range, concurrent, or scaling timing is added | Open |

## Current audit: 0501 replay evidence remains retained, but its current verifier cannot be replayed

The matched [0501](changes/0501-pptx-exact-payload-comparisons.md) reports,
catalog checks, and 38 static coverage tests remain present and pass their
current validators. The committed `verification-final.json` is the historical
formal custody receipt for the eight before and eight after reports. A fresh
run of `verify-final.py` is currently unavailable because the frozen before
executable under `/tmp/litchi-goal-0501` is missing; the replay reports
`before: frozen binary is missing`. This is a replay/custody limitation, not a
reason to discard the retained report rows or to describe the historical receipt
as a new run. The strict claim-registry check currently validates ten claims.

The formal comparison and its 52 favorable flags remain scoped to the named
PPTX source-backed media/plain workflows; native producer closure, cold and
physical-I/O behavior, and exhaustive CRUD promotion remain open. The prior
0501 timing, profile, cleanup, and exact-payload boundaries below continue to
apply.

## Current audit: 0501 scopes PPTX payload hashing; the full goal remains open

[0501](changes/0501-pptx-exact-payload-comparisons.md) removes redundant
private image/chart payload hashing while retaining exact payload comparisons,
candidate reread, graph and source proofs, cancellation, budgets, and partial
publication behavior. The fresh before phase has eight reports and 240 measured
samples, and the matched after phase has the same eight reports and 240
measured samples. The supplementary whole-child
SHA-256 profile reports 31.01% owned and 30.08% warm-file, with incomplete
caller recovery, so it does not establish touched-digest attribution or a
timer-local causal speedup. Same-data export recovery and two subsequent owned
profiling workload reruns are recorded separately and do not alter the formal
matrix. The comparison retains all 208 rows, passes every lifecycle oracle, and
retains 52 favorable timing/throughput flags with no adverse change above five
percent. The separate default CRUD refresh
passes two serial 201-row lanes, totaling 6,030 samples across 37 cases and 31
corpora; both generated report/catalog validators and 38 static coverage tests
pass. It closes the timing-report baseline gate only. The scoped production
gates also pass: default all-targets 848, all-features library 552, doctests 6
with 2 ignored, focused 58, private guards 5, Clippy, fmt, rustdoc, downstream,
boundaries, and 38 CRUD static tests. Independent strict verification passes.
Final cleanup removed 2,154,708,992 unique-inode allocated bytes and
retained eight replay files in local tmpfs: two binaries and six raw perf files,
including two failed attempts. Historical 15.5–16.1%
touched-digest attribution is from 0449. Native producer notes/chart closure,
cold and physical-I/O behavior, exhaustive CRUD promotion, and other goal
areas remain open.

## Current audit: 0500 targets managed paragraph batching; the full goal remains open

[0500](changes/0500-managed-paragraph-batches.md) targets the managed refusal
at the existing `Edit::replace_body_paragraph_texts` seam. The repeated
baseline has 12 children, 720 measured samples, and 72 warmups; at 32 selected
paragraphs, edit accounts for about 93% of the timed lifecycle. The candidate
builds one source-checked final projection and preserves finite ownership,
atomic failures, inverse proofs, and monotonic accounting. The 24-child after
phase completes a 36-child, 2,160-measurement comparison. Against repeated
scalar in the same final executable, K=32 lifecycle p50 improves 6.947–7.138x
and edit p50 12.640–13.456x; K=8 lifecycle improves 2.324–2.346x. K=1 p128
owned retains a 7.34% lifecycle flag, and p512 K=8 warm-file retains a 5.77%
RSS flag. These are scoped API-choice results; CRUD-index promotion, native
producers, controlled-cold behavior, one-percent coverage, and other goal
areas remain open.

## Current audit: 0499 targets bounded Part worker lifetime; the full goal remains open

[0499](changes/0499-operation-local-part-workers.md) reuses a bounded,
operation-local scoped worker set when an ordered Part batch spans waves, with
the existing single-wave path retained. The implementation preserves source
and cancellation fences, worker/task/byte limits, owner-retained results,
monotonic `Work` and `InputBytes`, typed error ordering, and complete joins.
An unwind guard also closes a provider-panic cache flight; recovery is limited
to unwind builds because release uses `panic = abort`. The matched capture has
60 children, 3,600 measured samples, and 360 warmups. Many-small owned p50
improves 3.59x/2.92x/2.08x and warm-file 2.64x/2.26x/1.81x at widths 2/4/8,
but both remain slower than their ordinary serial controls. Five aggregate and
ten per-repeat flags remain retained, so this is scoped workload evidence
rather than a program-wide speed result. CRUD, native-producer, controlled-cold,
history/composition, and other goal areas remain open.

## Current audit: 0498 adds explicit bounded Part batches; the full goal remains open

[0498](changes/0498-bounded-source-backed-part-batch.md) adds an owning ordered
OPC Part batch with caller-controlled worker, task, byte, cancellation, and
memory limits. The 30-child same-executable matrix has 1,800 measured samples
and 180 warmups with byte/budget/cleanup checks. Delayed-source scaling is
useful, while many-small owned/file reads regress materially. Apparent
superlinear few-large observations remain unexplained; no scheduler-only or
program-wide speedup claim follows. Broad CRUD, native producers, controlled
cold behavior, allocation attribution, and history requirements remain open.

## Current audit: 0497 adds an atomic DOCX publication capability; formal analysis is verified

[0497](changes/0497-docx-atomic-publication.md) adds consuming
`ParagraphStreamPlan::write_to_path` and
`ParagraphStreamCommit::write_to_path` methods for the existing bounded DOCX
logical-tail append route. The methods use OPC's sibling-temporary atomic
replacement helper and retain the existing source, replay, budget,
cancellation, candidate, and inverse-proof boundaries. A callback failure before
replacement preserves the destination and removes the private temporary;
post-replacement parent-directory sync failure returns typed
`OpcError::Committed` and must not be blindly retried.

The default hashing route remains the only before/after comparison and keeps
its existing report shape. Counting is a bounded non-retaining after-only
capability; atomic publication is also after-only because the before revision
has no atomic destination method. The frozen formal inventory contains 288
children and 8,640 samples, with a separate 72-child/216-sample pilot. Formal
verification passes all 288 child terminals, including the 97-child resume
after the original `ENOSPC` interruption. The descriptive analysis retains 864
matched default-hashing comparison cells and 144 after-only capability rows.
It retains all 100 >5% adverse flags: 72 allocator live-byte endpoint, 16
latency, and 12 whole-child RSS. The original ordinal-191 raw observation
remains archived without fabricated evidence. Cleanup overlapped two formal
whole-child intervals, so no isolated-host or individual-outlier attribution is
made. Final cleanup verification passes; the evidence inventory is recorded
by `results/change-0497/seal.py`.

| Goal area | 0497 candidate evidence | Audit status and boundary |
| --- | --- | --- |
| Atomic filesystem publication | Consuming DOCX plan/commit path methods delegate to the existing OPC atomic sibling replacement | Capability scope only; no general atomic-save or durability claim |
| Default publication comparison | Historical hashing sink remains unchanged and is the only before/after route | 864 descriptive comparison cells; no broad speedup or causal claim |
| Counting and atomic routes | Non-retaining counting and filesystem atomic routes are balanced after-only capabilities | 144 descriptive capability rows; no atomic before/after comparison or synthetic write-call/digest values |
| Broader non-iWork goal | No new representative CRUD selector; borrowed input, producers, cold behavior, scaling, history, and broad CRUD remain open | Open |

## Current audit: 0496 verifies descriptive phase attribution; the full goal remains open

[0496](changes/0496-docx-edit-phase-attribution.md) is a harness-only,
opt-in diagnostic follow-up to the unresolved 0495 whole-child RSS and latency
flags. Its verified formal run has 32 children and 960 samples: 12
before-unmanaged, 12 after-unmanaged, and 8 after-managed, with two reversed
repeats, three warmups, and 30 measured samples per child. The default report
schema and existing correctness oracles remain unchanged. The claims scope is
descriptive phase latency, whole-child RSS, and full-lifecycle allocation;
managed-after rows remain capability observations and no CPU, optimization, or
causal claim is authorized.

The paired unmanaged review contains 294 comparison cells, with 74 threshold
flags retained and classified by evidence scope.

Normal unmanaged lifecycle p50 latency is milliseconds, paired by repeat:

| Arm | Before R1 | After R1 | Before R2 | After R2 |
| --- | ---: | ---: | ---: | ---: |
| owned | 2.221 | 2.153 | 2.126 | 2.176 |
| file-warm | 4.794 | 2.302 | 2.279 | 2.326 |
| short | 5.359 | 5.340 | 5.238 | 5.419 |

Publication is the largest named normal phase in all 16 normal children. The
after-managed file-warm lifecycle/publication p50s are 3.563/2.465 ms and
3.553/2.453 ms; short is 6.509/5.406 ms and 4.092/2.993 ms, showing repeat
instability. Phase clocks are wall time, not CPU time; allocator and RSS remain
full-lifecycle/separate scopes.

The 74 retained threshold flags are 60 phase rows (18 commit-drop, 18
published-snapshot-drop, 10 diagnostics/XML identity, 7 open, 6 publication,
1 phase-sum), 9 allocator reallocation rows, 4 whole-child RSS rows, and 1
full-lifecycle latency row. The allocator p50s are 22,859 calls/5,721,334
bytes/+606,959 peak increment before-unmanaged, 9,696/1,623,696/+609,903
after-unmanaged, and 20,535/4,270,196/+622,454 after-managed. These rows are
descriptive flags, not independent regressions; `historical_flags_resolved` is
false and causality remains unresolved.

Formal verification passes all 32 children and 960 samples. The final helper
gate passes 8/8 seal-helper tests; the evidence seal remains a separate custody
gate. Cleanup passes source-manifest and protected-file checks after removing
10.709 GiB of disposable custody data and retaining four replay binaries. The
full non-iWork goal remains open: borrowed input, atomic save, independent
producers, cold intersections, scaling, durable history/composition, and
broader CRUD/security evidence still require separate work.

| Goal area | 0496 evidence | Audit status and boundary |
| --- | --- | --- |
| Opened-document edit/save | Verified descriptive phase/lifecycle capture over the 0495 path | No optimization or causal claim |
| CPU/RSS attribution | Wall-clock phase vectors and whole-child RSS | Wall time is not CPU attribution; RSS is not phase-local |
| Allocation attribution | Full-lifecycle calls, bytes, reallocations, and peak increments | No nested phase allocation region; managed-after remains descriptive |
| Existing 0495 flags | 0496 retains 74 flagged comparison rows alongside unresolved historical flags | No flag is dismissed or deleted |
| Complete non-iWork goal | No new CRUD selector or production capability | Open |

## Current audit: 0495 enables and measures ordinary managed DOCX edit/save; the full goal remains open

[0495](changes/0495-docx-managed-document-edits.md) closes the finite-owner
and source-authority seam that 0494 identified as a prerequisite for an
ordinary managed opened-document edit. Its verified formal run retains 72
processes and 2,160 samples; pilot2 retains 36 processes and 108 samples. The
matrix covers six providers (`owned`, `instrumented`, `file-warm`, 4 KiB
`short`, delayed 64 KiB `delayed`, and zero-fixed-delay `range-zero`) in normal
and allocator roles over a 200-paragraph, 20-member DOCX with eight 2 MiB media
members. Every formal row emits the expected 16,793,048-byte artifact and
passes source, sink, semantic, and untouched-media checks. Managed rows also
pass their finite-budget checks, while allocator-role rows pass
allocator-conservation checks; normal and unmanaged rows do not expose those
allocator fields.

The normal managed p50 observations in milliseconds are 5.592/5.556 (owned),
3.277/3.298 (instrumented), 5.682/5.853 (file-warm), 4.066/4.060 (short),
572.748/576.147 (delayed), and 182.327/181.990 (range-zero) for repeats one
and two. These are managed capability and budget observations. The protocol has
no managed-before baseline, so they do not authorize a managed speed, RSS,
allocation, or provider-ranking claim. The only performance comparison is the
paired unmanaged before/after review, which retains eight whole-child RSS flags
and three latency flags above the five-percent threshold. Their causality is
unresolved, and repeat instability remains visible rather than being dismissed
as host noise.

All 720 managed formal rows report zero reservation failures and release
resource memory, objects, and depth to baseline. The lifecycle labels are
separate: `cache_before` is post-open, `cache_live` is post-edit/pre-publication,
resource `live` is post-publication, and `after_drop` follows package
consumption plus returned-snapshot/commit release. No post-publication cache
gauge exists. The profile review is whole-child evidence that includes setup,
output verification, and serialization; it does not attribute CPU or RSS to the
timed operation. The next measured follow-up is phase-local CPU and whole-child
RSS attribution, with repeats of the unstable file-warm and short/allocator
arms, plus a distinct post-publication cache observation if retention is being
claimed.

| Goal area | Current evidence | Audit status and boundary |
| --- | --- | --- |
| Opened-document edit/save | 0495 verified managed capability/budget baseline plus paired unmanaged before/after review | Descriptive evidence only; no broad speedup or managed-before comparison; atomic filesystem save remains open |
| Managed ownership and source authority | Owner-retained snapshots/views, finite admission, source-checked publication and complete-artifact inverse tests | Necessary ordinary-edit enabler; durable history, composition, broad mutators, and dependency-bearing edits remain open |
| Provider and cache boundaries | Six explicit provider arms, two roles, two repeats, and distinct cache/resource lifecycle gauges | Cache diagnostic is pre-publication; release gauge is post-publication/drop; no cache-after-drop or provider ranking claim |
| CPU/RSS attribution | Core counters, syscall traces, and owned stack captures in the profile bundle | Whole-child diagnostics only; phase-local operation attribution remains the follow-up |
| Borrowed and independent-producer coverage | No genuine borrowed source or native producer round trip | Open |
| Concurrency and scaling | One worker and serial lifecycle execution | Open; bounded 1/2/4/8-worker evidence and Amdahl analysis remain required |
| Complete non-iWork CRUD checklist | 0495 adds no representative selector or default row | Open; structural/deletion, cross-document, merge/split, patch, repair, dynamic-content, security, and broader format rows remain |

The final validation bundle records 478 harness tests (one ignored), 937 DOCX
unit tests, 119 DOCX integration tests, 388 OPC unit tests, 79 doctests (31
ignored), and 58 Python helper tests. These are scoped correctness, custody,
and measurement gates; they do not turn 0495 into a broad performance claim.
The full non-iWork goal remains open, and iWork is untouched.

## Current audit: 0494 adds opened-edit provider baselines; the full goal remains open

0494 supplies the missing descriptive baseline for one opened DOCX paragraph
replacement, commit, and sequential publication across six explicit warm
provider arms and a verified-cold `FileSource` lane. The warm capture retains
24 formal processes and 720 samples (plus 36 pilot samples); the cold capture
retains 120 formal samples and six pilot samples. The deterministic corpus has
200 paragraphs, 20 archive members, and eight 2 MiB media members. Warm and
cold rows use the same logical content; the cold lane uses a page-aligned copy
with a padded ZIP tail and therefore a distinct physical archive hash. Output,
semantic edit, untouched-part/media, and source-version checks pass for the
retained rows. Patch replay, inverse, and stale-source checks are untimed
preflight gates.

The result is baseline evidence with `claim_authorized: false` and no
before/after optimization claim. Warm normal p50 ranges from 2.298 ms for the
first instrumented repeat to 569.581 ms for the delayed-range arm; the two
repeat values vary substantially for several providers. The verified-cold
normal p50 is 148.065 / 238.430 ms and allocator p50 is 98.139 / 93.802 ms.
The cold lane has its own source-open and residency/I/O boundary, so these
values are not a warm-versus-cold performance comparison. The 4 KiB short-read
rows produce 4,217 source calls and 3,840 short reads; other traced rows
produce 377 source calls. Warm allocator rows retain a 606,986-byte peak
increment and 22,859 calls; cold allocator rows retain 607,702 bytes and
22,864 calls. The independent 63-row cold allocator audit passes.

| Goal area | Current evidence | Audit status and boundary |
| --- | --- | --- |
| Opened-document edit/save | One existing paragraph replacement through commit and sequential publication across six warm providers and verified-cold file input | Baseline coverage is present; 0495 subsequently enables the managed path, while this 0494 capture remains descriptive and general atomic filesystem save remains open |
| Read-only managed provider lifecycle | 0493 synthetic opt-in open/load/text evidence | Accepted as a separate read-only slice; it must not be conflated with 0494 edit/save |
| Provider and cache boundaries | Owned, instrumented, warm file, short-read, delayed-range, range-control, and verified-cold file cells | Descriptive provider baselines; no provider ranking or warm/cold delta is authorized |
| Borrowed and independent-producer coverage | No genuine borrowed source and no native producer round trip | Open |
| Concurrency and scaling | 0493/0494 lifecycle captures are serial; 0498/0499 add bounded low-level Part batches and 0500 adds managed batch comparison | Open; full lifecycle 1/2/4/8 scaling, lock-wait evidence, and Amdahl analysis remain required |
| Complete non-iWork CRUD checklist | The 0494 path adds no selector or representative-index row | Open; broader format and CRUD rows remain to be evidenced |

The ordinary managed edit boundary was intentional at the 0494 revision. The
0495 follow-up now provides the owner-retained managed capability; 0494 remains
its descriptive provider/cold baseline and does not itself establish a managed
before/after speed claim. iWork remains outside this audit while its separate
work is in progress.

## Current audit: 0493 production read-ahead is accepted; the full goal remains open

0493 completes the production OPC integration of the bounded forward-start
read-ahead policy selected by 0492. The policy is explicit and opt-in, the
default source-backed constructors remain exact, and publication or exact
preservation paths permanently close forward admission before their first
exact archive read. The DOCX facade forwards the policy without exposing
archive implementation types. The source and focused integration tests cover
budget charging, retained-window release, source-version changes, re-entry,
panic/poison recovery, queued cancellation, short reads, and untouched-member
preservation.

The accepted 0493 evidence is a policy comparison using the same retained
executable within each normal or allocator role for a fresh managed DOCX open,
main-document load, text extraction, and drop over a pinned in-memory provider.
It contains 24 pilot and 480 formal samples. The modeled
delayed provider reduces managed p50 by 81.82–84.87%; physical calls fall
19→3 while accepted bytes rise 3,966→5,445 (+37.29%). Zero-delay changes stay
within −2.01% to +2.28%, allocator cost is +3 calls and +4,384 bytes, and no
same-repeat latency or whole-child RSS regression exceeds five percent. These
figures establish a bounded provider/read-only result, not a default-local,
real-network, cold-filesystem, or broad format claim.

| Goal area | Current evidence | Audit status and boundary |
| --- | --- | --- |
| Explicit caller-supplied range reads | OPC-owned bounded managed window, physical-fill accounting, exact fallback, and DOCX forwarding | Partial completion for the opt-in source-backed path; default behavior remains exact |
| Read-only content extraction | Fresh synthetic managed DOCX open/load/text lifecycle with strict source, text, cache, physical-trace, and budget oracles | Accepted but narrowly scoped to the pinned in-memory provider and one worker |
| Opened-document edit/save | 0494/0495 add descriptive provider/cold coverage and owner-retained managed edit/save; 0497 adds an atomic logical-tail capability | Partial capability and baseline evidence; broad before/after coverage and general atomic save remain open |
| Provider, cold, and producer breadth | 0491 and 0494 add explicit provider and verified-cold baselines | Borrowed lifetimes, native producers, and cold intersections across the CRUD matrix remain open |
| Bounded concurrency and scaling | 0498/0499 provide explicit ordered Part batches; 0500 compares managed batch versus repeated scalar edits | Partial low-level evidence; full lifecycle 1/2/4/8 curves, lock wait, and Amdahl analysis remain open |
| Complete non-iWork CRUD checklist | No 0493 selector or representative-index promotion | Open; conversion, structural/deletion, cross-document, merge/split, patch, repair, dynamic-content, security, malformed, and broader format rows remain to be evidenced |

The source validation for this batch passed 628 OPC tests, 1,390 DOCX tests,
456 harness tests, and 27 selected Python helper tests, alongside warning-
denied lint/documentation and boundary gates. Those are scoped correctness and
integration gates; they do not turn the read-only synthetic slice into
opened-edit, native-producer, cold-cache, borrowed-lifetime, scaling, or full
CRUD evidence. iWork remains outside this audit while its separate work is in
progress.

## Current audit: 0492 measures a range read-ahead enabler; 0493 is the bounded production integration

[0492](changes/0492-docx-bounded-range-read-ahead.md) is a benchmark-only
4 KiB forward-window experiment over 24 pilot and 480 formal samples. On its
simulated 1 ms range source, p50 falls from about 20.34 ms to 3.46 ms while
physical calls fall 19→3 and accepted bytes rise 3,966→5,445 (+37.29%). The
zero-delay normal p50 changes +4.88%/+0.50%, with a repeat-one p99 increase of
5.20%; the window allocation is outside the operation allocator region. The
private mutex and one-worker route are correctness details, not scaling
evidence. [0493](changes/0493-managed-opc-source-read-ahead.md) carries the
bounded policy into OPC with explicit InputBytes and Memory accounting; neither
batch establishes real-network, cold-filesystem, borrowed-lifetime, or broad
CRUD performance.

## Current audit: 0491 establishes DOCX provider and verified-cold baselines

[0491](changes/0491-docx-provider-and-cold-baseline.md) adds descriptive
source-provider and filesystem observations: 600 provider samples, 360 fresh
child filesystem samples, and four prepared-query controls that are explicitly
cold-ineligible. The simulated 1 ms range arm has roughly 20.3 ms p50 and 19
calls; local provider medians are roughly 0.27–0.30 ms. Verified-cold p50 is
2.753/4.361 ms normal across repeats, with 2.724/2.797 ms allocator, and each
row proves one main-part materialization plus positive process I/O. Repeat
tails vary materially. Whole-child profiles include setup and report work, so
they do not attribute operation CPU or RSS. This is a baseline, not a speedup;
0492/0493 provide the subsequent bounded read-ahead experiment and integration.

## Current audit: 0490 leaves file-store tails unresolved and bounds sync conclusions

[0490](changes/0490-file-store-variance-and-sync-attribution.md) alternates
retained before/after binaries through six blocks, preserving 72 processes and
4,320 samples with exact source/output identities. Normal file-store p50 has a
mean 6.73% reduction with a block-bootstrap interval of −9.71% to −4.41%, but
normal and allocator p95/p99 intervals span both directions and adverse blocks
remain. Four sync-only diagnostics attribute 78.99–81.61% of traced median
operation time to the unchanged `fdatasync`; selected whole-child syscall
counts match. The derived 1.23–1.27× bound is a scoped Amdahl model, not a
durability optimization or hardware limit. No production change follows;
0491 onward supplies the provider/cold work that this follow-up identified.

## Current audit: 0489 reuses candidate XML audit work with retained adverse rows

[0489](changes/0489-opc-candidate-xml-audit-reuse.md) reuses one successful
candidate XML audit inside an immutable prepared splice plan while retaining
source/candidate byte and EOF authentication, freshness, cancellation, Work,
limits, and final reopen. Its 144-process, 4,320-sample comparison reports
roughly 20–21% lower source-heavy medians, 14–15% lower authored-heavy file
medians, and about 19.89% lower deterministic operation heap. Candidate bytes
are unchanged. A file-store tail regression and four whole-child RSS increases
above 5% remain in the review; the result does not authorize dropping freshness
or independent candidate proofs.

## Current audit: 0488 records evidence-preserving cleanup only

[0488 cleanup](results/change-0488-disk-cleanup/README.md) changed no
production code or performance behavior. It removed 503,605,403,648 allocated
bytes of rebuildable Cargo output and 12.25 GiB of inactive worktree output,
481.26 GiB in total, after process-use checks; protected and dirty worktrees
were excluded and reachable commits were retained. The post-cleanup 0487 seal
still verifies 144 formal processes, 4,320 samples, 12 diagnostic children,
and two fuzz campaigns. This is custody evidence; rebuilding is required
before new workspace tests, and the cleanup does not strengthen any performance
claim.

## Current audit: 0487 reduces replay sink fences but retains small-workload regressions

[0487](changes/0487-opc-replay-consumed-prefix-retention.md) keeps consumed
replay bytes in the existing adapter allocation across short reads. The matched
144-process/4,320-sample comparison lowers authored-heavy file p50 by
35.61–35.99% and leaves source-heavy medians within 1%; diagnostic `statx`
falls 59.90% on authored-heavy file input while `pread64` remains unchanged.
Four latency rows and three whole-child RSS pairs above 5% remain, all in the
small-workload controls, and operation-heap rows do not increase. The parallel
allocator harness assertion failed because its process-global counter is not
isolated; the serial retry passes and the failure remains retained. Exact
source/output and preservation oracles pass. The broader goal and candidate
audit reuse remained open at this change.

## Current audit: 0486 attributes DOCX replay metadata cost without a production change

[0486](changes/0486-docx-replay-metadata-callers.md) retains four tiny
whole-child caller profiles over the 0485 executable. The authored-heavy file
profile places about 27.38% of sampled period weight in stacks containing
`statx`; this is inclusive sampled attribution, not a syscall count, CPU
measurement of the operation, or speedup. Its review identified the consumed
prefix retention candidate and an unresolved partial-output contract boundary,
which 0487 implemented under additional tests. Source-heavy 0485 strace data
showed 25,219 `statx` calls, but the 0486 sampled absence of a `statx` frame in
other profiles does not imply zero calls. No production code changed and the
full provider, cold, scaling, native-producer, and CRUD requirements remain
open.

## 0485: bounded consumed-window batching; full goal remains open

[0485](changes/0485-opc-splice-consumed-window-batching.md) applies a private
OPC adapter optimization to the replayable DOCX logical-append route. It
batches hashing and sink output for bytes already consumed inside the existing
bounded window while retaining the separate source XML audit, replay
authentication, source freshness policy, per-fragment Work charges, and
source-bound preservation checks.

The matched evidence contains 144 formal processes and 4,320 samples. At 64
existing and 16,384 authored paragraphs, file-input p50 falls from
443.916/440.383 ms to 239.499/238.865 ms across the two repeats. At 131,072
existing and 64 authored paragraphs, owned-input p50 falls from
482.325/476.215 ms to 385.058/385.006 ms. Operation heap peaks are effectively
unchanged. Three latency quantiles and nine small-workload whole-child RSS
observations exceed the five-percent review threshold and remain part of the
accepted scope.

The profile bundle records authored-heavy file `statx` falling from 3,735,939
to 1,475,055 while `pread64` remains 114, and source-heavy file `statx` rising
from 15,415 to 25,219 while `pread64` remains 264. These are one-sample,
one-warmup whole-child diagnostics that include setup and oracle work. They
justify caller-level attribution before any further change; they do not justify
relaxing freshness checks or deleting the independent source proof.

The route passes its scoped all-feature/no-default tests, formatting, warning-
denied Clippy and rustdoc, boundary, helper, benchmark, and sanitizer gates.
It does not close cold-cache, concurrent, atomic-save, all provider/input
intersections, or broad native Office validation. It also does not add a
representative CRUD selector. The machine-readable index therefore remains 15
categories and 34 rows: 11 measured, 22 correctness-only, and one explicitly
unsupported dynamic-content row (33 selector-backed mappings).

## 0482: shared primitives complete; end-to-end append remains open

[0482](changes/0482-bounded-xml-opc-splice.md) implements bounded XML auditing and decoded OPC insertion publication.
Focused tests cover source/candidate proof refusal, opaque ZIP preservation,
no-ops, inverse restoration, partial output, cancellation and budgets. Primitive
measurements cannot prove the program's bounded append requirement: public
DOCX append still uses the materialized lifecycle, and generated fragments,
package metadata and caller storage have distinct ownership. The
[next work](results/change-0482/next-work.md) retains DOCX scanner integration,
replayable paragraph generation, durable patches, provider variants and
end-to-end scaling evidence. Broader CRUD and parallel scaling obligations
remain in scope. The full goal is not achieved.

## 0481: scanner allocation progress; full goal still open

[0481](changes/0481-docx-borrowed-scanner-names.md) removes measured unnecessary
name allocation work and records lower normal lifecycle means in both repeats.
It leaves peak operation heap and document-sized XML/index ownership unchanged.
The [explicit-window contract](results/change-0481/window-contract.md) separates
the first decoded-splice milestone from the required multi-paragraph producer,
replay, durable inverse and large-stream scaling work. The broader non-iWork
requirements remain open; this batch does not redefine the definition of done.

## DOCX publication duplicate payload removed; append bound open (0480)

[0480](changes/0480-docx-shared-publication.md) removes one complete target XML copy at the existing OPC handoff.
The large operation peak is 35,371,102 bytes instead of 41,793,870 bytes;
retention still grows with document size. This is measured ownership progress,
not a bounded-window append completion claim. Repeated scanner/layout work,
the separate tail-only capability and the broader non-iWork scenario/source/
native/parallel requirements remain open.

## DOCX logical append baseline captured; window requirement open (0479)

[0479](changes/0479-docx-tail-append-baseline.md) records 720 samples of the existing one-paragraph tail-copy lifecycle.
The 41,793,870-byte incremental heap peak at 131,072 source paragraphs confirms
that this materialized transaction is not a bounded-window append solution.
Its current one-operation and section-property restrictions also leave the
proposed 64/256 repeated-append workloads uncovered. Production changes follow
measured ownership/CPU attribution. Broader CRUD, native/source variants,
parallel scaling and the full non-iWork objective remain open.

## Explicit public PPTX metadata window implemented (0478)

[0478](changes/0478-pptx-generated-metadata-spool.md) integrates a bounded
checked name plan and caller-supplied central-directory scratch through the
public fresh PPTX writer. The final evidence bundle records the matched policy
comparison and the operation allocator window gate. Caller scratch storage is
separate and grows with serialized directory bytes; this is not a constant-RSS
or all-PPTX claim. Ordinary constructors retain their metadata indexes.
The full non-iWork objective remains open, including logical append, broader
CRUD and native/source coverage, arbitrary repackaging and measured parallel
scaling. The [next batch](results/change-0478/next-work.md) starts with measurement
of the existing DOCX append path before proposing a bounded implementation.

The final 24-process, 720-sample matrix passes byte and reopen verification.
All 180 spool allocator samples have a 432,436-byte operation peak, versus
8,875,252 bytes for the 8,192-slide control. Scratch extents are 4,074, 40,842
and 1,237,692 bytes at 8, 256 and 8,192 slides. Small-deck normal mean latency
increases 3.11% and 3.31% in the two repeats; no registered 5% review threshold
is crossed. This is an operation-heap result with separate caller storage.

## Central-directory scratch implemented; name indexes remain (0477)

[0477](changes/0477-zip-central-directory-spool.md) removes growing central
header/name retention in an explicit low-level spool mode. Its measured heap
peak is flat across three tested member counts, while caller scratch and
per-member allocations/I/O remain real costs. ZIP Office and OPC validation
indexes still grow; the public PPTX route has not met its total-memory
requirement. A checked generated-name plan and semantic scratch integration are
next. The full non-iWork goal, including broader CRUD, source/native variants
and scaling, remains open.

## Repeated compressor allocation removed; total-memory gap remains (0476)

[0476](changes/0476-zip-deflate-state-reuse.md) implements the allocation owner
selected by 0475. All 720 main samples preserve output and the large operation
reduces requested bytes by 99.538473%. Peak heap remains approximately
8.88 MB for 8,192 slides, with retained name/directory metadata. This advances
implementation without closing the explicit total-memory requirement, logical
append, native breadth, source variants, repackaging or scaling.

## PPTX allocation ownership investigated (0475)

[0475](changes/0475-pptx-streaming-attribution.md) adds repeated attribution
to the 0474 memory-growth finding. Exact writer context is separated from
materialized preflight, raw traces and the failed demangled Heaptrack filter
are retained, and a separately declared mangled-symbol export corrects that
filter. This advances diagnosis; it does not close bounded total memory,
logical append, native breadth, arbitrary repackaging or scaling. The full
non-iWork goal remains open.

## PPTX fresh streaming memory requirement remains open (0474)

[0474](changes/0474-pptx-streaming-operation-memory.md) adds a missing public
fresh-creation baseline and rejects constant total memory for the tested path:
operation peak rises from 435,541 to 8,875,092 bytes over 8/256/8,192 slides.
All 360 samples and authored-slide oracles pass, with zero allocator live exit
delta. This advances measurement, not closure of the explicit-window requirement.
ZIP/OPC metadata and repeated allocation work need profiling and implementation.
Fresh creation remains distinct from logical append, Part addition and edits
followed by repackaging. Native breadth, source variants, scaling and the full
non-iWork goal remain open.

## DOCX fresh-creation evidence added (0473)

[0473](changes/0473-docx-streaming-operation-memory.md) advances the explicit-window
streaming requirement with 360 samples and full paragraph/run/package oracles.
The observed operation allocation peak is constant at 414,732 bytes over the
three tested sizes, separately from the 64-byte scratch reservation. Production
and the default matrix are unchanged. This adds fresh plain DOCX coverage;
logical append, Part addition, arbitrary repackaging, PPTX streaming, native
breadth, source variants and parallel scaling remain open. The full non-iWork
goal is not complete.

## Current scoped progress (0472)

[0472](changes/0472-xlsx-plain-cell-tags.md) retains plain-cell tag elision after
a 14.330% whole-process allocation reduction and 1,263 passing XLSX tests.
The full guard retains 86 latency flags; there is no registered latency or
peak-memory claim. Independent audit identifies public DOCX fresh streaming
creation as a missing allocator/scaling measurement, distinct from buffered
creation and the other three append meanings. Native, source-variant, CRUD,
streaming and scaling requirements remain open; the full goal is not complete.

## Latest measured rejection (0471)

[0471](changes/0471-xlsx-rewrite-buffer-lifetime.md) rejects a safe explicit
release of obsolete worksheet rewrite bytes because matched measurements do
not demonstrate the required peak-memory benefit. Seven-row diagnostic ABBA,
full-guard results and Heaptrack totals remain reproducible. This closes one
hypothesis, not a goal requirement or the broader non-iWork program. Larger
snapshot allocation and duplicate parsing work, CRUD coverage, source variants,
bounded streaming and measured scaling remain open.

## Current audit: 0470 bounded XLSX pass reuse (2026-09-08)

The [0470 record](changes/0470-xlsx-empty-web-proof.md) advances the measured
ordinary XLSX commit bottleneck: compaction can establish empty web bindings
without another worksheet traversal, while every unproven input retains the
original reader and error phase. All 1,257 XLSX tests and six scoped gates
pass. The protocol retains six-row ABBA, full-default guard and whole-process
allocation evidence, with its exact results and limitations recorded separately.

This changes neither the default scenario matrix nor the taxonomy's remaining
correctness-only mappings. Eager parsing and lossless snapshot scans remain
substantial leads. Full native-producer, physical-cold/range, bounded-streaming,
parallel-scaling and comprehensive CRUD evidence is not established by this
batch. Existing caches and publication checks retain their bounds. The full
non-iWork goal remains open; no completion or general speedup claim follows.

## Current audit: 0466 dense XLSX investigation (2026-09-08)

[0466](changes/0466-xlsx-dense-commit-profile.md) advances the required
profile-before-optimization work for the default dense one-percent XLSX
commit/save lead. It retains repeated normal timings, whole-process counters,
allocation stacks, initial incomplete CPU callchains and a same-source diagnostic
frame-pointer build. Supplemental normal p50 is 399.470314/398.785007 ms.
Wrapper recovery and postprocessing overlap remain explicit. Production,
37-case/201-row default coverage, 11 measured/22 correctness-only mappings,
and all accepted ADR contracts are unchanged. This is actionable profiling
evidence, not a production speedup or closure of the program's largest-bottleneck,
native, cold/remote, bounded-streaming or scaling requirements. The goal remains open.

All 18 focused Python tests, exact cross-process summary replay, fresh-copy
portable verification and postcleanup verification pass. The 103-artifact seal
is complete; both temporary executable copies and generated bytecode are removed.

## Current audit: 0465 checked default coverage (2026-09-07)

0465 completes the measured capture for the existing materialized
`odp_existing_append_lifecycle` case. Preflight preserves all prior 198 row
identities and the checked catalog now binds 37 default cases, 201 rows and 31
corpora. The mixed append-incremental category has exactly one measured ODP
row and two correctness-only fresh-streaming rows, enforced by the validator;
the taxonomy is 15 categories, 33 mappings, 11 measured and 22
correctness-only.

The formal lanes retain 6,030 normal samples and 90 ODP-only allocator samples.
ODP normal p50 is 1.691887/1.687518 ms (tiny), 67.786424/67.745553 ms
(medium) and 136.334131/136.843521 ms (large) in R1/R2. All four lanes,
353 harness tests (one ignored), warning-denied Clippy, rustdoc, scoped format,
167 latest Python tests, both full-report CRUD validators and boundaries pass.
Initial stale Python hash/count-pin failures remain retained. The sealed
precleanup verifier, five resealed negative probes, finalize precleanup,
fresh-copy flagless portable verification and owned cleanup pass; the portable
seal is unchanged and temporary directories are absent. This evidence makes no
regression, speedup, independent native-producer, bounded-memory streaming or
scaling claim; the broader non-iWork goal remains open.

## Current audit: 0464 evidence (2026-09-07)

0464 adds descriptive harness evidence for a generic PPTX source/destination
pair and changes no production implementation. The pair is a same-source-derived
positive control: its destination is a Litchi self-copy publication of the
source archive, not an independently authored native package. The frozen matrix
contains eight reports and 240 samples across serialized R1 forward/R2 reverse
normal and operation-scoped allocator bytes/range lanes. API-sum p50 values in
R1/R2 are 1.9536/1.9380 ms (normal bytes), 1.9872/1.9865 ms (normal range),
2.0768/2.0705 ms (allocator bytes), and 2.1182/2.1213 ms (allocator range).

The logical-range adapter returns at most 256 bytes with zero fixed delay. Each
formal output is three slides and 55,891 bytes and passes the independent
package oracle. The API sum covers four public calls and excludes input/setup,
sink/adapter construction, artifact/oracle work and teardown; it is not a
contiguous end-to-end measure. Normal allocator totals are unavailable by
design. No native Office acceptance, independent-producer, network, physical-I/O,
cold-cache, scaling or copied-byte/compression claim follows.

Across all 120 allocator rows, summing the four allocator regions gives exactly
8,623,012 allocated bytes, 9,616 allocation calls and 1,455 reallocation
calls. This does not infer a lifecycle peak by summing phase peaks. Bound GNU
time resource logs report maximum RSS in kB for R1 normal-bytes/normal-range/
allocator-bytes/allocator-range as 16,912/16,908/17,068/17,152 and for R2 as
16,900/16,684/16,612/16,584. The overall 16,584–17,152 KiB (~16.2–16.75 MiB)
range is process-level and includes setup, oracle work and teardown. The
summary does not derive this GNU-time resource field; raw
`captures/R1/*/resource.log` and `captures/R2/*/resource.log` files are bound
in the bundle.

The final source review and harness Clippy receipt are clean at custody epoch
`a35f4507a74e7678f29f91dd6857d07a87579ba872785a1eaf19e296540debbe`; harness
validation records 391 passing tests with one ignored, and all 240 formal
samples plus focused smoke (four positive and seven negative cases) pass.
LibreOffice 26.2.5.2 saves the formal three-slide output. Native-r2 passes the
scoped application-save plus source-backed/eager validation of all three slide
counts, slide sizes and ordered text; its saved output is 39,830 bytes with
SHA-256 `d2ca3109448f5a0797f00561beff4eafa84616f7dbf897ff2456667fbf555915`.
All three post-save source-backed image inventories are unavailable because
`UnsafeEdit` / `source-backed picture inventory` refuses markup compatibility.
The failed native-roundtrip-r1 receipt and raw saved output remain preserved.
This does not establish full Office acceptance; image equivalence and rendering
compatibility remain unproven. A supplemental inventory-only source epoch
changes only that diagnostic binary; its build and warning-denied Clippy receipt
pass. Source compatibility records 7,031
unchanged files, while the measured copy Rust/binary/captures remain immutable
and no retiming occurs. The documentation
receipt and owned performance-crate scoped format check pass. A broad `cargo
fmt --all --check` receipt remains failed only at pre-existing formatting in the
out-of-scope iWork `crates/litchi-keynote/src/document.rs`; no full-workspace
format pass is claimed. Boundary checks pass. The corrected `profiling-r1`
receipt supersedes the initial UID-0-unavailable diagnostic. It passes a
whole-process user-counter and DWARF profile over three warmups and 30 retained
samples (33 checked iterations); separate counter and sampling runs add 60
diagnostic samples outside the 240 formal samples. Totals are instructions
1,758,215,198, cycles 631,274,012, branches 327,372,796, branch misses
3,514,541, cache misses 2,295,690, and page faults 1,709. PMU counters ran at
83% because of multiplexing, and the raw zero cache-reference value is
uninterpreted. The 20-sample cycle symbol view places
`sha2::sha256::x86_sha::compress` at 19.96% and
`zlib_rs::inflate::inflate_fast_help_avx2` at 14.25%; the whole-process scope
includes binary hashing, setup and oracle work, so this is diagnostic only and
does not establish operation hotspots or stable ranking. Sealed precleanup, fresh-copy portable verification and resealed-summary
tamper rejection pass. Owned cleanup removed 12 files totaling 252,885,687
bytes plus two archived audit temporaries. The full non-iWork goal remains open, counts remain 439
selectors / 36 defaults,
0463 remains retained, and iWork is untouched. See the [0464 summary](results/change-0464/summary.json),
[0464 protocol](results/change-0464/protocol.json),
[pair](results/change-0464/pair.json), [fixture provenance](results/change-0464/fixture-provenance.md),
[source review](results/change-0464/source-review.md), [native-r1 receipt](results/change-0464/native-roundtrip/receipt.json),
[native-r2 receipt](results/change-0464/native-roundtrip-r2/receipt.json),
[profile summary](results/change-0464/profiling-r1/profile-summary.json),
[top symbols](results/change-0464/profiling-r1/samples/top-symbols.txt), and
[source compatibility](results/change-0464/source-compatibility.json).

## Current audit: 0463 evidence (2026-09-07)

0463 retains a private writer-origin proof for the ordinary ODP publication
path after the frozen 3% gate passes. The private ODP serializer records proof
accounting around the existing common `PackageWriter` audit path, which is
unchanged. The A1/B1/B2/A2 evidence contains 24
reports and 720 samples. Normal p50 candidate-minus-baseline deltas are
-10.0491% / -8.9678% for tiny, -6.9642% / -7.3050% for medium, and
-7.3657% / -6.8199% for large in R1/R2. All normal and allocator bootstrap
upper bounds are below zero; no adverse >5% elapsed or RSS flag or allocation
increase review flag is present.

Allocator bytes fall by 2,087,682 / 11,072,106 / 20,202,090 for
tiny/medium/large in both repeats. Every lane reduces allocation calls by
1,039, reallocations by 93 and deallocations by 946; peak above entry changes
are 0 / -49,674 / -15,438 bytes and retained-live deltas remain zero. The proof
is eligible only for the exact writer/source owners and conservative audit
bounds; candidate reopen/readback, source precheck, media/domain checks,
no-op and patch behavior remain. The source review reports 381 ODP tests and
warning-denied Clippy passing.

Supplementary phase clocks exclude setup, warmups and checks; whole-process
counters include them. The proof changes the commit validation path, whose p50
phase delta is -12.6863% / -13.0797%; other phase movement is diagnostic. The
full non-iWork goal remains open, counts remain 439 selectors / 36 defaults,
0460 remains accepted, and iWork is untouched. The harness and final gates are
complete: 387 harness tests pass with one ignored, for 768 passed ODP/harness
tests in total. Warning-denied rustdoc, scoped formatting, boundaries,
precleanup and source replay, fresh-copy portable replay, resealed +1ns tamper
rejection, and owned cleanup of four executables totaling 233,058,712 bytes
pass. See the [0463
comparison summary](results/change-0463/summary.json), [phase summary](results/change-0463/phase-summary.json),
[source review](results/change-0463/source-review.md), and [proof design](results/change-0463/proof-design.md).

## Current audit: 0462 evidence (2026-09-07)

0462 records partial normal-lifecycle improvements from a private seventeen-key
shape-attribute index, but rejects the candidate after the predeclared 3%
medium/large gate. The frozen A1/B1/B2/A2 evidence contains 24 reports and 720
samples. Normal p50 candidate-minus-baseline deltas are -2.3623% / -3.0049%
for tiny, -3.4649% / -3.1917% for medium, and -2.6602% / -2.6279% for large
in R1/R2. Both large rows miss the threshold while every normal p50 bootstrap
interval remains below zero. Allocation bytes, calls, reallocations,
deallocations, regional peak and retained-live metrics are exactly unchanged;
the R1 tiny allocator interval crosses zero, and no adverse >5% elapsed or RSS
flag is present. The index adds 280 bytes per element, so no retained speed or
memory benefit claim follows.

Manual assembly review finds `shape_builder` changing from 17 generic getter
static call sites and a 1,400-byte frame to 17 indexed typed getter call sites and a 1,688-byte
frame. Generic `ElementAttrs::get` remains 328 bytes. Indexed known-key hits
bypass the cached loop, but the defensive fallback remains; the automatic
elimination field is a flawed direct-call heuristic and is excluded from the
evidence. Supplementary phase clocks cover public API calls and exclude setup,
warmups and checks; whole-process counters include them and remain diagnostic.

Candidate validation records 379 ODP tests and warning-denied owner Clippy as
passing; the initial compilation failure remains preserved. All 387 harness
tests pass (one ignored). Both Rust files are restored byte-exact to baseline
revision `dbd2f8ece`; final Clippy/docs/format/boundaries, portable replay, tamper
rejection and owned temporary cleanup pass.
No selector or corpus coverage is added: the registry remains 439 selectors /
36 defaults, 0460 remains accepted, the full non-iWork goal remains open, and
iWork is untouched. See the [0462 comparison summary](results/change-0462/summary.json),
[phase summary](results/change-0462/phase-summary.json), [source review](results/change-0462/source-review.md),
and [assembly review](results/change-0462/assembly-review.md).

## Current audit: 0461 evidence (2026-09-07)

0461 records a partial normal-lifecycle improvement from splitting ODP
attribute matching from value decoding, but rejects the candidate after the
predeclared 3% practical gate. The frozen A1/B1/B2/A2 evidence contains 24
reports and 720 samples. Normal p50 candidate-minus-baseline deltas are
-2.0578% / -2.2940% for tiny, -2.0604% / -3.8754% for medium, and
-2.1547% / -3.1281% for large in R1/R2. R1 medium and large fail the gated
threshold; all normal p50 bootstrap upper bounds remain below zero. Allocation bytes,
allocation calls, reallocations, deallocations, regional peak and retained-live
metrics are exactly unchanged, and no adverse >5% elapsed or process-RSS flag
is present. No retained speedup claim follows.

The assembly receipt confirms that the out-of-line `ElementAttrs::lookup` body
is removed, namespace/local-name checks are inlined into `get`, and the `get`
stack frame moves from `0x148` to `0x128`. Supplementary phase clocks measure
public API calls separately from the primary matrix. Whole-process counters
include setup, warmups and checks and are not operation-only totals. Neither
diagnostic establishes causal attribution.
The candidate records 372 ODP tests, warning-denied all-target Clippy and
scoped formatting as passing, plus 387 harness tests with one ignored. Source
restoration to `05f432d48`, final Clippy/docs/formatting/boundaries, portable
verification, tamper rejection and owned temporary cleanup pass.

The experiment adds no selector or corpus coverage; the registry remains 439
selectors / 36 defaults. 0460's fused staging optimization remains accepted,
the full non-iWork goal remains open, and iWork is untouched. See the
[0461 comparison summary](results/change-0461/summary.json), [phase summary](results/change-0461/phase-summary.json),
[source review](results/change-0461/source-review.md), and [assembly receipts](results/change-0461/candidate-assembly.json).

## Current audit: 0460 evidence (2026-09-07)

0460 retains the private ODP fused staging/source-scanning optimization. The
authoritative A1/B1/B2/A2 lifecycle matrix contains 24 reports and 720 samples.
Normal p50 candidate-minus-baseline deltas are -3.5430% / -3.8257% for tiny,
-6.0849% / -5.0679% for medium, and -5.4732% / -5.8268% for large in R1/R2.
All four medium/large rows pass the predeclared 3% improvement gate with
negative independent bootstrap upper bounds; no adverse >5% elapsed or RSS
flags are present.

Allocator p50 deltas are -4.7840% / -4.1309% for tiny, -7.3100% / -6.7785%
for medium, and -7.3197% / -6.9282% for large. Every allocator lane reduces
allocated bytes by 4,642, allocation calls by 16, reallocations by 12 and
deallocations by 4, while regional peak and retained-live deltas are unchanged.
Supplementary phase clocks show transaction p50 reductions of -24.2681% /
-24.0697%; whole-process counters show instructions -6.1959%, cycles -4.9100%
and branch misses +0.5503%. The phase clocks cover public API calls without
setup, warmups or checks; the whole-process counters include those activities.
Neither scope establishes causal API attribution.

The source review finds the fused namespace-aware traversal preserves state
machines, error precedence, BOM-relative spans, limits, source ownership and
readback/patch/no-op contracts. The corrected owner retry passes 371 tests and
all-target warning-denied Clippy passes. The optimization is retained under
this scoped evidence; no selector or corpus coverage is added, counts remain
439 selectors / 36 defaults, the full non-iWork goal remains open, and iWork is
untouched. All builds, 387 harness tests (one ignored), documentation, boundaries,
portable verification and owned temporary cleanup pass. See the [0460 comparison
summary](results/change-0460/summary.json), [phase summary](results/change-0460/phase-summary.json),
and [source review](results/change-0460/source-review.md).

## Current audit: 0459 evidence (2026-09-07)

0459 makes measurement progress: it repairs sampled phase ancestry and rejects
a local-name-first attribute lookup experiment after 720 matched operations.
No tested production change is retained. Candidate readback and repeated staging
scans now have stronger profile attribution, while caching family reopen is
rejected as low impact. Every regression flag remains visible.

The full non-iWork objective remains open. This batch adds neither a CRUD
selector nor source/output/native/scaling coverage; counts remain 439/36.
See the [0459 evidence](results/change-0459/README.md). Historical audits below
retain their original scoped conclusions.

## Current audit: 0458 evidence (2026-09-07)

0458 completes a supplementary ordinary ODP append phase diagnostic: 24 lanes,
720 samples, separate profiles, exact allocation-volume conservation, and
recomputable sealed evidence. Commit is the largest individual large-input
phase; opening plus transaction setup consumes about half the phase time.
Sampled stacks do not resolve phase-marker ancestry, so internal cost ranking
needs another diagnostic. This batch adds no production optimization or new
CRUD/source/output coverage. The registry remains 439 selectors / 36 defaults.

The full non-iWork goal remains open: wider selective CRUD, source/input/output
matrices, native application roundtrips, cold/range behavior and measured
bounded-worker scaling still need completion. The historical 0457 audit below
retains its original scope. See the [0458 evidence](results/change-0458/README.md).

## Current audit: 0457 evidence (2026-09-07)

The full non-iWork goal remains open, but the bounded existing-ODP evidence
now has a formal current-revision ordinary control. The
`odp_existing_append_lifecycle` baseline retains two repeats, normal and
allocator lanes, 64/4,096/8,192 source slides, three warmups and 30 samples
per lane (360 samples across 12 reports); the final candidate summary retains
the same 360-sample, 12-report matrix. The paired
`odp_source_tail_append_lifecycle` path validates bounded source XML and
replays the changed ZIP member through a sequential sink. It is a specialized
publication plan with a different retained-result contract and does not
integrate the ordinary `edit::Snapshot`/`edit::Transaction`/`edit::Patch`/
`edit::Commit` lifecycle. Those owned ordinary APIs are implemented; source-tail
integration with them is the remaining gap.

Normal p50 candidate-minus-control deltas are -33.352% / -33.185% for the
64-slide shape, +3.262% / -2.461% for 4,096 slides, and -1.940% / -4.221% for
8,192 slides in R1/R2. The allocator evidence reports control versus
source-tail regional peak above entry of 781,342 vs 620,381 bytes (tiny),
18,027,568 vs 620,385 (medium), and 35,958,388 vs 620,385 (large). Allocated
bytes are 10,613,383 vs 2,607,922, 110,226,105 vs 36,593,937, and 211,442,207
vs 71,164,177. These are operation-scoped allocator observations; source and
fixture ownership at entry are excluded, and candidate allocation volume
continues to grow with document size.

The comparison receipt explicitly withholds ordinary Commit/Patch speedup or
regression, general CRUD or retained-result equivalence, causal, bounded-memory,
scaling, cancellation and physical-I/O claims. R1 normal tiny is +8.301% and
R2 normal tiny is +6.517%; allocator tiny is +10.184% in R1, while allocator
large is -5.113% / -5.740% in R1/R2. The remaining process-peak rows are
within the five-percent threshold. The final identity-bound set has ten native
records and three synthetic source/output pairs passing the independent ZIP/XML
oracle, with no Office application launch. ZIP, ODF-common and ODP suites pass
1,348 tests with three ignored (ODF-common 499, ZIP 481, and ODP 368); the
harness passes 383 with one ignored, OPC passes 497 with one ignored, both
ZIP/XML sanitizer fuzz lanes retain 1,000 runs, and candidate derivation and
comparison receipts pass. Initial pre-fix large-normal profiles retain 1,000 sampled
`cycles:u` observations: candidate SHA-256 compression is 11.00%, XML
`validate_name` 9.08%, `memcmp` 6.01% and `validate_start_element` 5.42%;
control `memcmp` is 9.67%, Quick-XML attribute iteration 6.00%, `memmove`
5.76% and namespace-prefix resolution 5.24%. These are whole-process sampled
stacks, not API-attributed CPU percentages or causal hotspots, and no hard
counter totals were collected. These whole-process samples remain initial
pre-fix diagnostics, not final candidate attribution; no sampling profile was
captured for the final candidate epoch. Candidate profile checking passes for
this initial capture; the unchanged control report/raw profile passes the
existing amended oracle after the original frame-pointer expectation failed.

The source mapping verifies 439 selectable selectors and 36 defaults. The
coverage index keeps 15 categories, 33 representative mappings, 10 measured
mappings and 23 correctness-only mappings; its ODP append row remains
correctness-only because generated-per-run 0457 corpora do not satisfy the
index's checked-catalog/default measured-status contract. The ordinary control
is nevertheless recorded as the formal baseline in the 0457 evidence and
reports. The [final source read review](results/change-0457/final-code-review.md)
records the retained-fragment lease fix as resolved and found no additional
verified blocker. The sealed batch passes precleanup, portable-copy replay,
and three altered-copy rejection checks. Owned temporary artifacts are removed
with a retained inventory; see the [bundle](results/change-0457/README.md). See the [final candidate summary](results/change-0457/candidate-final/summary.json)
and [final comparison](results/change-0457/comparison-final.json).

## Current audit: 0456 evidence (2026-09-07)

The full non-iWork goal remains open. [0456](changes/0456-zip-shared-payload-framing.md)
retains verified/shared ZIP payload storage during publication, removing a
payload-sized preparation allocation and copy. Media publication allocates
4,243,083 bytes (-79.85%) and peaks 593,892 bytes above entry (-96.59%); output,
read/write work and managed budgets remain unchanged. Plain publication metadata
costs 3,360 more allocated bytes. The complete document lifecycle is not bounded
by this change, and managed admissions still include the old writer allowance.

The 720 formal observations retain all 15 timing flags. Ordinary bytes/media API
medians regress 15.411%/16.347% with higher page faults. Separate 240-sample fixed
allocator-policy diagnostics improve API medians 2.626–10.266%. Acceptance is for
the memory reduction with allocator-sensitive latency disclosed, not a general
speedup. Release checks pass 2,654 tests (7 ignored), strict lint, documentation,
formatting, workspace/boundaries and 1,000 ASAN fuzz iterations. Native self-pair
output remains exact; six distinct source/destination probes refuse shared-graph
incompatibility. An evidence SHA transcription failure and corrected retry are
retained alongside the initial compile correction.

Coverage remains 438 selectors, 36 defaults, 15 categories, 33 representative
mappings, 10 measured mappings and 23 correctness-only mappings. No coverage row
is promoted. Broader CRUD, native application roundtrips, distinct-package pairs,
physical cold I/O, bounded existing-document append/repackaging and representative
worker scaling remain incomplete. [Next-work notes](results/change-0456/next-work.md)
identify the streaming XML/ODP append dependency. See the
[0456 bundle](results/change-0456/README.md) for replay and owned cleanup.
User-owned `docs/GOAL.md` remains unchanged.

## Prior audit: 0455 evidence (2026-09-07)

The full non-iWork goal remains open. [0455](changes/0455-zip-preservation-transfer-chunks.md)
retains a 64 KiB ZIP preservation buffer: 256 fewer media publication reads and
sink writes, identical byte counts/output, and simulated range API median
improvements 4.775%/4.449%. It adds 32 KiB fixed stack per active publication;
operation allocation counts and bytes do not change.

The evidence contains 720 formal samples, 240 separate ordinary confirmation
samples and 240 fixed allocator-policy diagnostic samples. Original adverse
whole-process CPU counters and all 19 formal timing flags remain disclosed.
Slow/high-page-fault processes appear in both builds; controlled allocator
policy supports paging variability without reconstructing historical mapping
decisions. No CPU or large bytes-only gain is claimed. Release checks pass
2,188 tests (6 ignored), strict lint, documentation, formatting, minimal workspace,
boundaries and 1,000 sanitizer fuzz runs. The native self-pair produces its
previous exact output via bytes/range; this adds no native-application roundtrip
or distinct-package evidence.

Coverage remains 438 selectors, 36 defaults, 15 categories, 33 representative
mappings, 10 measured mappings and 23 correctness-only mappings. This transfer
optimization promotes no coverage rows. Broader semantic CRUD, native application
roundtrips, distinct-package pairs, physical cold I/O, bounded existing-document
append/repackaging and representative worker scaling remain incomplete.
See the [0455 bundle](results/change-0455/README.md) for capture custody,
verification and owned cleanup. User-owned `docs/GOAL.md` remains unchanged.

## Prior audit: 0454 evidence (2026-09-07)

The full non-iWork goal remains open. The 0454 production source epoch
implements three recorded boundary fixes: the baseline optional
producer-visible-name refusal, the name-only historical intermediate
candidate's noncanonical relationship XML refusal, and the ordinary
authored-XML compactness refusal through OPC source XML proof. The opaque
`source_xml`, `checked_range`, `AuthoredXmlFragment`, and splice publication
path preserves source whitespace and literal namespace grammar while retaining
source identity, limits, resource accounting, cancellation, and exact no-op
copying. Ordinary authored XML remains strict. This capability evidence makes
no baseline-refusal speedup claim.

The final native inventory covers 588 files and 189 image-bearing direct-picture
rows: four self-pair cases publish and 185 outcomes remain unchanged. Each
case reopens the same unmodified archive independently as source and
destination. The pinned LibreOffice QA fixture additionally passes the exact
ZIP/XML preservation checks in its external bytes/range captures. This is a
self-pair proof, not an independent package pair or a native-application
roundtrip ([inventory](results/change-0454/final-native-inventory.json),
[outcome comparison](results/change-0454/final-outcome-comparison.json)).

The formal capture passes all 18 lanes, 16 provider plus 2 external, with 30
samples each: 540 retained samples
([measurements](results/change-0454/measurements.md),
[machine-readable rows](results/change-0454/measurements.json)). The matched
whole-API p50 movement stays within about 1.3%, no RSS comparison exceeds 5%,
and all 15 absolute review flags remain retained, including the three positive
range/media-rich R2 open p99 flags (+5.84%, +9.93%, +7.87%). The
range/media-rich R1 open-source p99 flag is -8.97% and does not repeat.
Logical read/work counters remain consistent; scheduler or sleep variability is
a plausible inference for the tail behavior, not a proved cause.

The core release epoch passes build, format, strict, harness, OPC, PPTX,
documentation, workspace, boundary, oracle, and native checks: 497 OPC tests,
854 PPTX tests, and 381 aggregate harness tests pass. The standalone amended
ASAN fuzzer also passes its 1,000-run seed-454 smoke. The original fuzzer
build failure is historical; `fuzz-source-amendment.json` records the only
post-measurement source change, an isolated `parse_opc` harness bridge. The
amended fuzzer source manifest is separate from the measured production
manifest, which remains unchanged
([amendment](results/change-0454/fuzz-source-amendment.json),
[validation history](results/change-0454/validation-notes.md)).

Custody corrections are disclosed. The Markdown-renderer correction temporarily
rewrote `protocol_sha256` in all 18 formal receipts; the renderer-r1 restoration
reconstructs those receipt bindings and retains the intermediate hashes, so
custody is reconstructed rather than uninterrupted
([restoration](results/change-0454/custody-corrections/renderer-r1/restoration.json)).
Raw timing reports, sampling, statistics, and timing inputs remain unchanged.
The separate derivation amendment records the corrected `derive-final.py`
display lookup apart from the original `derive.py`; it does not create a new
measurement ([derivation amendment](results/change-0454/derivation-amendment.json)).

The coverage taxonomy remains 438 selectors, 36 defaults, 15 categories, 33
representative mappings, 10 measured mappings, and 23 correctness-only
mappings ([coverage index](crud-coverage-index-v1.json)). The 540 formal samples
are one matched workload control and do not promote those coverage rows. Native
application roundtrip, distinct-package, cold-cache/physical-I/O, allocator,
scaling, and the broader semantic CRUD rows remain unmeasured or incomplete.

Final precleanup and post-cleanup verification pass. Owned temporary binaries,
fuzz build files, and all seven generated PPTX outputs have been removed; the
retained evidence records their identities. No future measured hotspot is asserted.
The next measurement scope is the 64 KiB destination pass-through hypothesis:
compare a bounded 32-to-64 KiB or adaptive candidate with exact request
histograms, timing, resource and sink gates. It remains a measurement scope,
not an established bottleneck or performance claim. The full non-iWork goal
remains open.

The sections below retain historical audit records. Their older “current” labels
refer to their original audit dates and must not override this section.

## Historical evidence through 0453

[0453](changes/0453-pptx-shared-decoded-payload.md) removes the duplicate staged decoded payload after successful PPTX image/chart
capture. Both allocator repeats save 16,777,408 planning allocated/retained bytes;
bytes/media API medians improve about 3%. Full fallback allowance remains and
media destination staging grows 128 bytes. All 1,702 final tests and 1,000 existing
OPC fuzz runs pass. Primary plain p99 increases remain visible; a separate fixed
240-sample ABBA investigation does not reproduce them.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Native application
breadth, cold I/O, bounded existing append, repackaging and scaling remain
required. This is progress; the full non-iWork goal remains active and uncompleted.

## Earlier evidence through 0452

[0452](changes/0452-pptx-retained-capture.md) integrates retained OPC captures into PPTX image/chart plans with
independent publication reservations. Complete simulated-range API p50 improves
31.324%/31.048%; separate balanced bytes/media confirmation supports about 8%
improvement with about 6% extra planning time. Publication source data reads
fall to zero. Plan-held source reservations increase 16,815,144 bytes until drop;
existing decoded staging remains. All 1,700 final tests and 1,000 fuzz runs pass.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Native application
breadth, cold I/O, bounded existing append, repackaging and scaling remain
required. Shared staged decoded ownership is the next local memory opportunity.
This is progress; the full non-iWork goal remains active and uncompleted.

## Earlier evidence through 0451

[0451](changes/0451-opc-combined-capture.md) implements combined OPC read/authorization under cache, source identity,
work and memory/object budgets. Eight deterministic I/O cases preserve whole
publication output with a compressed source pass removed. Tests cover ordinary
and combined loader coordination, rollback, short reads, budget boundaries and
native-input compressed transfer after package/data drop. The API is opt-in;
semantic PPTX plans have not adopted it and no complete latency gain is claimed.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Reusable PPTX publication
reservations, matched timing, native breadth, cold I/O, bounded existing append,
repackaging and scaling remain required. The full non-iWork goal remains active.

## Earlier evidence through 0450

[0450](changes/0450-zip-combined-capture-decode.md) adds the low-level combined ZIP capture/decode primitive needed for OPC
first-read transfer authorization. Deterministic I/O evidence removes a compressed
source pass while preserving decoded bytes and writer token output. All 455 ZIP,
436 OPC and 59 PPTX targeted tests pass, as do strict lint and the instrumented
1,000-run fuzz smoke. This is a measured enabler; OPC/PPTX has not adopted it yet.

No timed or native coverage is promoted. Registry/default counts remain 438/36;
representative coverage remains 15 categories/33 mappings/10 measured/23
correctness-only. Combined OPC reservations, token/cache lifetime, matched timing,
native breadth, cold I/O, bounded existing append, repackaging and scaling remain
required. The full user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0449

[0449](changes/0449-pptx-caller-source-attribution.md) corrects the attribution used to rank the next PPTX optimization: untimed
harness hashing is about half of lifecycle SHA period, and publication read totals
combine source compressed capture with destination passthrough. Replaying all
240 existing samples confirms source-cache hits rather than cold rereads during
publication. This is diagnostic evidence; no new workload or native test ran.

Registry/default counts remain 438/36 and representative coverage remains
15 categories/33 mappings/10 measured/23 correctness-only. Native breadth, cold
I/O, bounded existing append, repackaging and scaling remain incomplete. The full
user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0448

[0448](changes/0448-pptx-minimum-service-pacing.md) completes the scoped minimum-service timer calibration identified in 0447.
Eight reports/240 samples and four profiles support retention of the opt-in
policy; both plain median gates and all service floors pass. All 381 harness
tests pass and strict lint adds zero diagnostics. No repeat trigger exceeds 5%.
This is a measured tooling enabler, with no production or physical-network claim.

Registry/default counts remain 438/36; the representative index remains
15 categories/33 mappings/10 measured/23 correctness-only. Its default full-run
contract is not promoted by opt-in captures. Native breadth, cold I/O, shared-link
concurrency, bounded existing append, repackaging and full scaling evidence are
incomplete. The full user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0447

[0447](changes/0447-pptx-range-transfer-pacing.md) fills a range-simulation tooling gap with optional configured transfer rate,
checked requested-delay counters and matched managed PPTX lifecycles. Eight
reports/240 samples and four profiles retain exact source-work/output equality.
The 378 harness tests pass; strict lint adds zero diagnostics. Twelve repeat
tail flags and sleep-granularity limitations remain explicit. This is a measured
enabler, with no production speedup, actual network-bandwidth or allocator claim.

The representative index intentionally reserves measured status for its default
full-run contract; later opt-in timing bundles do not automatically promote
those rows. Counts remain 438 selectors/36 defaults and 15 categories/33 mappings,
10 measured/23 correctness-only. Shared-link concurrency, cold I/O, native
breadth, bounded existing append, repackaging and full scaling evidence remain
incomplete. The full user-owned non-iWork goal stays active and uncompleted.

## Earlier evidence through 0446

[0446](changes/0446-opc-owned-content-type-name.md) retains a one-line ownership optimization in content-type parsing.
The 720-sample matrix passes the allocation gate: 5.963%/5.990% fewer calls at
medium/large in both repeats. The separate normal latency gate fails; peak
memory is unchanged and no paired latency/RSS or repeat trigger crosses 5%.
All 458 OPC/373 harness tests pass, owner strict lint is clean and harness lint
adds no diagnostics. Independent oracles and corruption probes remain retained.

This improves the existing synthetic package-addition slice without promoting
semantic or native coverage. There remain 438 selectors/36 defaults and 15
categories/33 representative mappings, 10 measured and 23 correctness-only.
Native breadth, repackaging, semantic dependency closures, bounded existing
append, cold/range and scaling are incomplete. The full user-owned non-iWork
goal remains active and uncompleted.

## Earlier evidence through 0445

[0445](changes/0445-opc-part-add-plain-source.md) adds a matched plain-source Part-addition lifecycle and a single-build
24-report/720-sample calibration. Plain large normal p50 is about 18.9 ms versus
62.1 ms observed; this is observer overhead, not a production speedup. Allocation
calls/requested bytes/above-entry peaks match. One instrumented p99 repeat flag
is disclosed. The 373 harness tests, 37 topology tests and 53 corruption probes
pass; strict harness Clippy retains inherited debt with zero new diagnostics.

Plain stack evidence points to content-type map construction. Source inspection
finds an avoidable owned-string clone as the next small measured candidate;
required validation, freshness and memory accounting must remain. The standalone
selector count is 438, with unchanged defaults and semantic representative
coverage. Native breadth, repackaging, semantic dependency closures, bounded
existing append, cold/range and scaling remain open. The full user-owned
non-iWork goal is active and uncompleted.

## Earlier evidence through 0444

[0444](changes/0444-opc-part-add-baseline.md) adds a low-level source-backed OPC Part-addition baseline:
12 reports/360 samples, three sizes, normal and allocator repeats, actual
independently verified ZIP fixtures, 29 corruption probes, and complete source
and sink observations. The 372 harness and 458 OPC tests pass; strict harness
Clippy retains inherited debt with zero new diagnostics. No production code
changes or comparative performance claim are made.

The observed source reader accounts for 55.24% of whole-process sampled self
time and scans all ordinary ranges per read. A matched plain-source baseline
must precede production attribution for this path. This partially fills package
Part-addition evidence only: semantic owner creation, broader dependency closures,
repackaging, bounded existing append, native breadth, cold/range and scaling
remain open. The user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0443

[0443](changes/0443-odp-compact-fragment-frames.md) replaces copied namespace/local-name frame data
in ODP source-fragment scanning with exact element kinds. The 720-sample matrix
passes the frozen allocation-call gate at 14.886%/14.943% medium/large reductions.
The normal latency gate fails; timing costs, two instrumented tail flags and
eight repeat flags remain disclosed. Peak and retained bytes are unchanged.
The original scanner remains a byte-identical differential reference.

This removes one source of temporary allocation work. Repeated validation,
one-shot attribute lookups, Part addition, repackaging, bounded existing append,
native breadth, cold/range and scaling remain open. The full user-owned
non-iWork goal is unchanged and uncompleted.

## Earlier evidence through 0442

[0442](changes/0442-odp-shared-staging-traversal.md) combines three ODP staging XML traversals while
preserving complete-pass error priority. The 720-sample matrix passes the frozen
normal p50 gate: medium/large improve 9.830–13.307% across both repeats.
All five repeat flags remain disclosed. Peak and retained bytes are unchanged;
no general tail, RSS or bounded-append benefit is claimed.

The three auxiliary scans are consolidated. Source-fragment scanning, one-shot
cache costs, Part addition, repackaging, bounded existing append, native breadth,
cold/range and scaling remain open. This batch does not complete or narrow the
user-owned non-iWork goal.

## Earlier evidence through 0441

[0441](changes/0441-odp-shared-preservation-projection.md) shares the immutable
ODP preservation comparison projection and keeps the mutable draft detached.
The 720-sample comparison shows 9.834%/9.857% lower medium/large operation peak,
with unchanged retained live bytes. Its original practical gate failed;
retention is explicitly based on a post-hoc peak-memory review. All raw timing
costs and the baseline tiny p99 repeat flag remain visible. No normal latency,
RSS or bounded existing-append claim is made.

Repeated XML staging scans remain a larger CPU target. Part addition,
repackaging, bounded existing append, native breadth, cold/range and scaling
remain open. This measured ownership improvement does not complete or narrow
the user-owned non-iWork goal.

## Earlier evidence through 0440

[0440](changes/0440-odp-borrowed-attribute-namespaces.md) reduces temporary
namespace allocation work in owned ODP append: about 20% fewer medium/large
calls and 6.3–6.6% fewer requested bytes across both repeats. Exact archive,
semantic, preservation and patch gates pass. The main 720 samples, additional
120-sample confirmation, four selected profiles, 352 ODP tests and 368 harness
tests are retained. Peak and retained live bytes are unchanged; no normal
latency or RSS benefit is claimed. Original adverse tails and two excluded
overlapping attempts remain documented.

This closes one private ownership experiment. One-shot cache cost and repeated
staging/validation are separate next hypotheses. Bounded existing append,
Part addition, repackaging, native/producer breadth, cold/range input, scaling
and the remaining non-iWork goal stay open. The user-owned goal is unchanged.

## Earlier evidence through 0439

[0439](changes/0439-odp-existing-append-lifecycle.md) establishes the full
owned existing-ODP append interval that the older one-edit timer omitted.
Twelve reports/360 samples and two accepted profiles cover 64/4,096/8,192
source slides plus an opaque package member. The 368-test harness suite,
independent fixture/resource gates, and no-new-Clippy-debt comparison pass.
Large normal p50 is 170.742/175.416 ms, with 236,704,188 requested bytes and
a 39,890,548-byte region peak above entry. This is a baseline addition, with
no production optimization or bounded-commit-memory claim.

The next evidence question is operation-specific attribution of owned open
and commit validation, including repeated attribute/namespace work. Part
addition, arbitrary repackaging, native/producer breadth, cold/range input,
scaling and the remaining non-iWork goal stay open. The user-owned goal is
unchanged; this batch does not narrow or complete it.

## Earlier evidence through 0438

[0438](changes/0438-odp-markup-batching-negative.md) tested the proposed ODP
fixed-markup accounting change and rejected it under the predeclared 5%
practical-gain gate. The 24-report/720-sample comparison found only
0.703–1.982% normal p50 improvements. Production was restored; candidate
source, measurements, four profiles, and validation remain reproducible.

The next coverage target is logical append to an existing ODP through the full
owned open/transaction/commit/sequential-output lifecycle. The existing
`odp_semantic_one_edit_save` timer excludes opening. This is distinct from
fresh streaming creation, adding a package Part, and arbitrary repackaging.
The broader non-iWork goal remains open, including rich/native producer
coverage, cold/range I/O, and scaling. `docs/GOAL.md` remains user-owned.

## Earlier evidence through 0437

[0437](changes/0437-odp-bounded-plain-slides.md) implements and measures fresh
plain titled-slide ODP streaming with explicit source/sink/resource ownership.
The common fixed-prelude constructor is opt-in and preserves the old strict
contract. Full release suites, resource/refusal/sink tests, independent
semantic/package gates, and the 36-report/1,080-sample/six-profile matrix pass.
The API is kept for its measured 420,352-byte operation peak across three
sizes, 98.247% below large Builder, with the 1.574/1.588 large p50 ratios and
all other regression flags disclosed. Native visual compatibility is unproven.

This closes the current plain fresh ODP creation slice, not the program goal.
Fixed-markup execution-accounting cost was the next measured-hypothesis target;
existing append, richer creation, Part addition/repackaging, real-producer
breadth, cold/range/scaling, and other taxonomy gaps remain. The user-owned
`docs/GOAL.md` stays unchanged and excluded from commits. The batch retains
reproducible raw evidence and portable cleanup proof.

## Earlier evidence through 0436

[0436](changes/0436-odt-bounded-text-spans.md) retains measured private ODT
ordinary-text batching: 24 formal reports, 720 samples, four profiles and
1,002 passing ODT tests. Normal p50 is 19.382–30.156% lower across three sizes
and two repeats, with exact archive/sink identities and identical aligned
allocator vectors. Region peak above entry remains 420,091 bytes; RSS is
unchanged in practice. No matched comparison crosses 5%; candidate tiny p99
repeat drift of −5.489% remains visible. Formal whole-process consume-self
share falls 44.27% → 19.24%; the scope includes setup and oracle work.

The immediate ODT accounting hypothesis is now measured. Next establish ODP
fresh creation evidence. Rich authoring, logical append, Part addition,
repackaging, native compatibility breadth, cold/range I/O, scalable parallelism
and the other original completion criteria remain open. The original user goal
is unmodified, and this batch neither completes nor narrows it.

## Earlier evidence through 0435

[0435](changes/0435-odt-bounded-plain-paragraphs.md) implements bounded fresh
plaintext ODT publication and retains 36 formal reports, 1,080 samples, six
profiles, 1,445 passing ODT/common tests, and 315 passing harness tests.
The new API lowers large-case operation allocator peak from 22.45 MB to
0.42 MB and requested bytes by 86.159%, with roughly 3.35–3.37 times candidate
Builder latency. Whole-process RSS is essentially unchanged. Every regression
and repeat flag remains visible; this is a measured sequential-output enabler.
The next hypothesis targets ODT per-text Work charging, sampled at 45.74% self
in the whole-process profile, before ODP fresh creation. Logical append, Part
addition, repackaging, rich authoring, native compatibility breadth, cold/range
I/O, scaling, and the other original completion criteria remain open.
`docs/GOAL.md` is unmodified. This batch does not complete or narrow that goal.

## Earlier evidence through 0434

[0434](changes/0434-ods-bounded-text-spans.md) retains a matched ODS
streaming comparison with 24 formal reports, 720 samples, and four passing
whole-process profiles. The same selector and deterministic scalar rows are
used before and after the private ordinary-text span batching. Descriptive
normal p50 deltas (after versus before) are tiny **−12.499% / −16.211%**,
medium **−13.731% / −13.121%**, and large **−13.562% / −13.492%** for R1/R2.
No matched comparison exceeds the 5% regression trigger, including RSS; the
only repeat flag is baseline tiny p99 at **+9.844%**.

Allocator calls, requested bytes, and regional peak-above-entry vectors are identical
before and after for every shape/repeat, with a 419,347-byte regional peak
above entry. The four profiles are whole-process samples that include setup
and the untimed oracle; `ExecutionContext::consume` self samples move from
25.01% to 10.57%. Ten before and eleven after addr2line warnings remain in the
retained diagnostics, zero L1 readings are not a zero-miss proof, and LLC was
not captured. The derived `claims[]` is empty: no broad 10x, RSS, physical-copy,
total-memory, or scaling claim follows. Copied-bundle replay and six mutation probes pass before and after task cleanup.

The ODS release receipt reports 454 tests passed; scoped Clippy retains the
known common `ArchiveReaderKind` large-enum debt. This is descriptive evidence
for one fresh ODS scalar-creation path. The next implementation slice is
bounded ODT paragraph creation, then bounded ODP creation; logical append,
package-Part addition, arbitrary repackaging, native breadth, and the overall
non-iWork goal remain open.

## Prior evidence through 0433

[0433](changes/0433-ods-bounded-fresh-scalar-creation.md) records 36 formal
reports and 1,080 samples for fresh one-sheet ODS scalar creation. The new
sequential writer's normal p50 is 59.017–62.137% below the before-buffered
role across the retained shapes and repeats; the after-buffered control retains
a +7.037% tiny R1 p99 regression flag. In allocator mode, the large
operation-region peak is observed at 71,050,076 bytes for the buffered role and
419,347 bytes for streaming. These are descriptive harness observations, not
an accepted speedup, total-memory, RSS, or allocator-internal claim. The
[result bundle](results/change-0433/README.md) retains the protocol, summary,
source/build custody, profiles, and checks. Copied-bundle replay and mutation
checks pass before and after removal of the task temporaries. Fresh creation
is the only covered semantic;
logical append, package-Part addition, arbitrary repackaging, and broader
native/I/O/scaling coverage remain open.

## Prior evidence through 0432

[0432](changes/0432-xlsx-streaming-operation-memory.md) closes the missing
operation-allocation observation for the existing XLSX scalar streaming writer.
Across 64/8,192/131,072 rows, all 180 allocator samples retain the same
420,110-byte incremental peak and zero exit live-byte change. This supports
the tested creation path; it does not close richer authoring, logical append,
all-feature memory bounds, cold/remote I/O or explicit scaling. Medium normal
latency repeat drift remains flagged. The [next-work audit](results/change-0432/next-work.md)
identifies ODS bounded scalar-row creation; the [native audit](results/change-0432/native-gap.md)
records the PPTX identity/fixture gap. The non-iWork goal remains active.

## Prior evidence through 0431

[0431](changes/0431-verified-compressed-source-transfer.md) implements and
measures verified compressed source-part transfer in the source-backed PPTX/OPC
workflow. Matched synthetic media API medians improve 86.5–89.0% for bytes,
warm files and short ranges, and 19.4–20.1% for simulated delayed range. The
initial delayed-range regression and its bounded-read refinement remain
reviewable. Refined validation includes 1,755 tests, strict checks and ASAN
smoke. This closes one measured publication hotspot; representative CRUD,
broader native/size coverage, true cold/remote I/O, semantic streaming/append,
allocator/physical-copy attribution, concurrent scaling and remaining strict
gates stay open. See [batch scope](results/change-0431/goal-scope.md). The
broader non-iWork goal remains active.

## Prior evidence through 0430

[0430](changes/0430-pptx-publication-cpu-attribution.md) closes the missing
publication CPU attribution in the synthetic media-rich provider experiment.
Frame-pointer capture on the unchanged binary resolves the Deflate callers;
its measured iteration share is 83.22% / 83.38% for bytes/warm files. The
[source and transfer audit](results/change-0430/transfer-design.md) identifies
compressed-media reuse as the next optimization, subject to preserved
validation, source authority, budgets and archive ownership. This batch makes
no production change or performance-improvement claim. Normal matched
captures, transfer implementation and adversarial validation remain required;
the broader non-iWork goal remains active.

## Prior evidence through 0429

[0429](changes/0429-pptx-provider-native-baselines.md) completes the 32-process/960-sample provider and native selected-
image baseline. It adds a bounded ZIP short-read correctness fix, not a measured
speedup. All final managed gauges are zero, and non-RSS phase observations match
exactly across samples and repeats. All 27 repeat flags are RSS points already
different at entry; comparative RSS conclusions remain withheld. Files are warm,
ranges are explicit simulations, and native selected-image ownership does not
establish native cross-copy. The 100-sample CPU-profile validator correction is
explicitly amended with original validators and unchanged captures retained.
Broader producer/size and CRUD coverage, true cold I/O, allocator/physical-copy
attribution, semantic streaming and explicit scaling remain open. The global
goal is still active; see [batch scope](results/change-0429/goal-scope.md).

## Prior evidence through 0428

[0428](changes/0428-managed-pptx-cache-lifetimes.md) completes the fixed synthetic
managed-cache/budget matrix: 16 processes, 480 samples, exact admission and
refusal, pinning/eviction, oversized bypass, and three publications at an exact
cumulative output ceiling. Every final releasable caller budget is zero. It
adds no production optimization and does not close the global goal. All 27
repeat flags are RSS points with differing entry values; comparative RSS
claims require a new setup/warmup attribution protocol. Native-producer and
cold/range-source lifecycles, the full CRUD taxonomy, streaming, and explicit
scaling remain open. See the [next-work record](results/change-0428/next-work.md).

## Prior evidence through 0427

[0427](changes/0427-pptx-allocator-drop-checkpoints.md) adds explicit allocator
phase/drop observations for the current plain/media-rich PPTX APIs. Eight
fresh processes retain 240 samples, all returning to their entry callback
live-byte value after the final sink drop, with identical repeat phase changes.
This closes the narrow missing caller-drop callback observation. It does not
close the full retention/cache/managed-budget/near-limit gate or establish RSS
release, object-owned bytes or a leak finding. The
[next-work record](results/change-0427/next-work.md) identifies additive
fallible PPTX cache-diagnostic forwarding using existing managed constructors.
Composite harness coverage is 288 passes and one ignored after a baseline-
reproduced stale selector assertion is corrected; the full failed command is
retained. Existing harness strict debt remains the same 29 findings.

An earlier retained allocation optimization is
[0424](changes/0424-staged-pptx-payload-reuse.md), based on production revision
`d18bf7db4` and evidence revision `340cc91ae`. Its matched source-backed media
lifecycle records 14.266% fewer requested allocation bytes and a 7.317% lower
mean V3 operation-region peak. Logical reads and endpoint retention remain
essentially unchanged; normal timing is diagnostic. The standalone bundle
replays 16 reports, four traces and 104 mutation probes after original
worktree/binary cleanup. It establishes no release latency, physical-copy,
managed-budget or post-drop claim.

The intervening [0419](changes/0419-pptx-bounded-archive-growth.md),
[0420](changes/0420-opc-owned-payload-reuse.md),
[0421](changes/0421-allocator-peak-counter.md),
[0422](changes/0422-operation-region-allocator-peak.md), and
[0423](changes/0423-matched-source-backed-pptx-lifecycles.md) records distinguish
allocation requests, retained endpoints, corrected lifetime peaks, V3 region
peaks and aligned lifecycle boundaries. Historical high-water values affected
by the 0421 correction remain unsuitable for current comparisons. There are
nine strict registered claim replays and 15 index categories with 32
representative selectors in the retained 0424 verification; coverage remains
representative rather than complete.

The 0425 audit finds 64 workspace members: 45 non-iWork leaf/shared packages,
the separately configured `litchi` facade, 17 iWork owners, and the Python
binding that unconditionally enables iWork. Shared ZIP and XML owners remain
in scope. The current strict layout finding in ODF's `ArchiveReaderKind`
requires a measured ownership decision; its borrowed reader is retained inline.
Mechanical constant-chunk and test setup findings are closed in the 45 leaf/shared
packages, with strict Clippy passing for 44 and the ODF layout gate open.
[0425](changes/0425-non-iwork-verification-maintenance.md) retains composite
coverage of 15,778 passing tests and 98 ignored tests. The separate facade had
six baseline-reproduced test failures, now closed by
[0426](changes/0426-xls-formula-ancillary-preservation.md): five fixture/assertion
corrections and bounded standard BIFF8 Formula ancillary preservation. The
final facade suite passes 461 tests with 11 ignored; XLS passes 1,341 tests
with one ignored doctest and its scoped strict gate passes. This is correctness
evidence, with static metadata layout cost documented separately and no
performance claim. The facade strict rerun retains the same 18 findings from
the prior audit; no allowance is added.
The standalone harness has 29 lint findings outside the modified code, and
native-resave requires a lockfile refresh before its locked gate can run.
These strict-gate debts remain open; no blanket strict pass is claimed.

Resolve the remaining scoped strict-gate debt and return to the performance
evidence program with facade correctness restored.
Explicit caller-drop snapshots, near-limit memory/cache evidence, broad native
producer matrices, cold/range sources, bounded semantic streaming and append,
scaling/CPU counters, full CRUD coverage and final strict gates remain open.
The full non-iWork goal is not achieved. Sections below retain earlier
revision-specific audits and their historical counts; this section supersedes
only the current facts explicitly listed above.

Change [0418](changes/0418-pptx-cross-copy-candidate-reuse.md) retains a scoped
owned PPTX candidate-reuse optimization. Media lifecycle p50 improves
38.27–38.69%, with an explicitly reviewed approximately 8% whole-process RSS
increase. Two opt-in lifecycle selectors bring the registry to 427; the default
36-case/198-row matrix and representative index statuses remain unchanged.
The batch retains 6,400 normal observations, 480 separate allocator observations,
paired profiles and source/output checks. Retained-memory, broader corpus,
source-backed, native/cold/remote and scaling work remains open.

Change [0417](changes/0417-representative-crud-baseline.md) broadens current
descriptive evidence to all 30 representative selectors across 14 executable
categories at clean revision `b10d6c25a13242ca260a8c897946f4d80ae06c61`.
It retains 30,000 normal observations, 1,800 separate allocator observations,
30 preflight checks, explicit timer boundaries, repeat uncertainty and portable
verification. Five selectors have >5% repeat drift on at least one quantile;
28 lack operation allocation attribution. This is progress on baseline coverage,
with no speedup or completion claim.

The index's default/full-run status contract is unchanged, and several cases
time only an already-open query or commit. Aligned end-to-end workflows,
operation allocation, broad corpus/producer matrices, cold/remote, native,
failure and scaling evidence remain open. The media-rich PPTX path provides
a concrete investigation target with an existing source-backed counterpart
capability; matched timing boundaries and preservation checks are required.

Change [0415](changes/0415-zip64-streaming-deflate.md) adds explicit streamed
ZIP64 Deflate framing, automatic Office transport selection from raised limits,
and strict support for descriptor-backed zero local-size placeholders. Its
transport tests and independent large Python OPC corpus narrow the ZIP64 gap;
semantic streaming, broader append/failure matrices and the performance program
remain open.

Change [0416](changes/0416-local-zip64-read-preservation.md) adds local-only
forced-ZIP64 read/preserve compatibility: 13 deterministic Python fixtures, 10
focused ZIP integration cases, and 7 OPC no-op/edit/failure cases, with strict
path fuzz coverage. The fixtures cover signed and unsigned descriptors,
descriptor-signature CRC collision, seekable no-descriptor output, empty and
many-small members, and central/local or ZIP32-tail combinations. This records
capability coverage only; the source identity, performance guards, and final
gate status belong to the 0416 change record. The overall non-iWork goal
remains open.

The prepared ODF catalog's existing `has_zip64_metadata` observation is based
on central-directory/tail metadata. It does not classify a seekable local-header
ZIP64 sentinel when the central and tail metadata remain ordinary. ODF catalog
parity for that input form, together with semantic, cold/remote, native,
concurrency/scaling, and broader failure-matrix evidence, remains open.

Change [0414](changes/0414-zip64-output-promotion.md) implements preservation
output promotion at ZIP32 local-offset and member-count boundaries, including
OPC provenance for reopened packages above 65,535 members. Generated ZIP64
Deflate through that revision's streaming local-header path remained a typed
refusal; 0415 replaces that refusal with explicit framing. These capability
changes do not close the performance program.

Change [0413](changes/0413-cfb-chain-scratch-reservation.md) retains a scoped
production CFB optimization with nine XLS ABBA comparisons, exact allocation
and logical-I/O checks, a reviewed ~3% CFB guard cost, paired CPU/PMU evidence
and eight strict registered claims. This is verified progress, not completion
of the non-iWork program. Broad scenario, native/cold/remote and scaling gaps
below remain open.

**Audit date:** 2026-09-05
**Audit basis:** the 0418 candidate `f8f9e6667` against control `79dfee502`,
the 0417 representative baseline at `b10d6c25a` and retained evidence,
with the 0416 local-framing candidate `d18cd04a2` and retained evidence,
with the 0415 streaming source batch and its retained evidence,
with the 0414 source batch and its retained verification evidence,
with the 0413 committed control `6b632726b` and the 0412 captured candidate at
`63c95bc22d5883c8ecab0872030757e5584254f7`, with the verified 0411 baseline at
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d` and the earlier
`9c6742c5212dd0e7ff2367da585abe357aae8975` ZIP64 control retained for the
historical comparison. iWork is outside this audit by the user's instruction.

**Disposition: OPEN.** The repository has substantial correctness and scoped
performance work, but the definition of done in `docs/GOAL.md` is not met. The
coverage index explicitly describes itself as representative, and the current
records do not provide a complete, independently reproducible baseline and
optimization result for the non-iWork CRUD matrix.

## 0412 current implementation

Change [0412](changes/0412-xls-observer-isolation.md) at committed revision
`b8f61970d` corrects the XLS observer's category range catalog by coalescing
only exactly adjacent spans, preserving overlap multiplicity, and disabling
the unused generic repeated-read union. It adds three explicit opt-in plain
`OwnedSource` lifecycle selectors: `xls_owned_source_open`,
`xls_owned_source_open_list_worksheets`, and `xls_owned_source_open_one_cell`.
Their timed reports expose operation/allocation metrics with `source: None`;
separate instrumented observations retain the logical locality evidence.

Eleven focused XLS observer/owned/allocator tests plus the registry test pass,
along with scoped formatting, crate boundaries, coverage-index validation, the
seven-claim strict checker, and report classification. The clean candidate
`63c95bc22` now has 18 schema/corpus-verified timing, allocation, profile and PMU
reports. Plain one-cell p50 is 0.166–0.169 ms; separate allocator captures record
126 calls / 223,774 bytes. Instrumented locality remains unchanged across the
observer correction. All 91 comparator tests pass, including mismatched observer
rejection. This is a measurement enabler and baseline, with no production
speedup claim. The registry is now 425 selectors; the default remains 36 cases / 198 rows and the
index remains 15 categories / 30 representative selectors. The non-iWork goal
remains open.

## 0411 current update

Change [0411](changes/0411-xls-read-allocation-baseline.md) now retains a
verified, descriptive baseline for six explicit opt-in XLS/CFB lifecycle
selectors: eager and source-backed open, open plus worksheet listing, and open
plus one-cell selection. The fixed generated corpus is
`xls-comments-opaque-heavy` (`litchi-xls-comments-opaque-heavy-v1`): two sheets
(`Comments`, `Untouched`), selected `Untouched!E21 = 42.0`, 257 logical entries,
10 archive members, 16,995,840 archive bytes, an 80,946-byte `Workbook` stream,
and eight 2 MiB opaque streams. The eager and source-backed paths use the same
two-sheet semantic oracle; this does not assert parity with the separate real
producer corpus discussed in older records.

The [0411 evidence bundle](results/change-0411/) contains four fresh CPU-2,
one-worker normal processes with 20 warmups and 500 samples per selector, plus
two allocator processes with 3 warmups and 30 samples per selector. Samples
share each process; the protocol is warm in-memory corpus evidence, not cold,
remote, concurrent, or per-sample process-isolated evidence. Source-backed
reports retain classified `ReadAt` counters and prove zero worksheet/opaque
payload reads for open/list and selected-only worksheet reads for one-cell.
Allocator reports retain operation-scoped allocation vectors; their peak values
are process-lifetime snapshots, and neither allocation nor timing is an A/B
speedup claim.

The capture and corpus verifier passed on clean revision
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d`. All-feature, all-target
`litchi-xlsx` Clippy passed; the full XLSX test run passed 1,238 tests with no
failures, and the combined harness/allocator test selection passed 9 tests with
four test threads. This evidence expands the current descriptive baseline but
does not promote the selectors into the default matrix or claim completion of
the non-iWork goal. The index remains 15 categories and 30 representative
selectors; the default remains 36 cases / 198 rows and the selector registry
remains 422.

## 0410 current update

Change [0410](changes/0410-mce-attribute-name-reuse.md) is a bounded MCE
expanded-attribute ownership experiment on the XLSX selected-cell path. It
uses candidate `e4d477466718a8fad38cd55b9babe0b826e7f3a7` and control
`972dc25be0dbd6690c74429839a48288d637e2d5`, the fixed four-sheet 9,216-cell,
17-member corpus with 4 MiB of media, and Rust/Cargo/Rustdoc 1.98.1 release
builds on CPU 2. The primary ABBA p50 candidate-minus-control changes are
`-3.9447%` and `-4.1373%`; operation-local allocator calls move from 81,918
to 77,212 and allocated bytes from 10,690,444 to 10,309,094. The seventh
strict-registry claim entry has been added, and the strict checker passes all
seven claims.

The initial eager edit/save guard is adverse at +0.53% / +5.66% p50 elapsed
time. A diagnostic repeat is lower by 1.106% / 0.505%, while same-role p50
drift is roughly 5–6% for both roles. The edit claim is therefore withheld and
the initial adverse result retained. Review found that the eager fixture does
not execute the changed MCE stream; this limits causal interpretation and does
not prove an eager no-regression result. The residual selected-path profile
attributes 10.91% of leaf weight to `clone_bounded_name_part` and reports
`parse_element` at 5.89% self / 23.42% inclusive. These are sampled profile
figures, not paired CPU evidence.

The final all-feature run passes 1,918 tests across the three crates. The
OPC/common warning-denied Clippy passes after two preexisting test-lint fixes;
XLSX all-target Clippy remains blocked by 9 library diagnostics and 28
library-test diagnostics, including 19 additional test diagnostics. This
update records scoped progress only; the non-iWork goal remains open.
The [0410 evidence bundle](results/change-0410/) retains the build identities,
ABBA reports, allocator captures, profile attribution, and gate logs. Rustdoc,
crate-boundary, and scoped-format checks pass.

## Authorities and reading boundary

This audit uses the current goal, the accepted ADRs, and the current
non-iWork performance records:

- [`docs/GOAL.md`](../GOAL.md), especially the mission, measurement contract,
  CRUD matrix, deliverables, and definition of done.
- ADR 0001 (public layers and typed refusals), ADR 0002 (downward crate
  ownership), ADR 0003 (immutable snapshots and atomic edits), ADR 0005 (the
  `ReadAt`/budget/output/measurement contract), ADR 0006 (preserve-by-default
  validation and security), ADR 0008 (migration gates), ADR 0010/0011 (facade
  and OPC physical ownership), and ADR 0024 (current topology).
- [`REPORT.md`](REPORT.md), [`CRUD_COVERAGE.md`](CRUD_COVERAGE.md),
  [`crud-coverage-index-v1.json`](crud-coverage-index-v1.json), and the
  claim-registry policy.

The decisive constraints for the open ZIP work are: preserve is the default;
validation must not mutate; an owned changed OPC source may publish only
through a proven preservation plan; unsupported framing must return a typed
refusal before output; and physical ZIP ownership stays in `soapberry-zip`.
Those constraints come from ADR 0005's exact-source amendment, ADR 0006, and
ADRs 0010/0011.

## Current evidence through 0502

| Goal requirement | Current evidence | Assessment |
| --- | --- | --- |
| Reproducible CRUD baseline | The current refresh covers 15 categories, 33 selector-backed mappings, 37 cases and 31 corpora with two serial 201-row lanes and 6,030 samples; generated report/catalog validators and 38 static coverage tests pass. | Timing-report baseline gate is closed for the 11 measured mappings (48 rows); 22 mappings remain correctness-only. The complete checklist, missing metrics, native producers, and correctness-only rows remain open. |
| Scoped claims | The strict registry currently validates 10 claims. 0501 retains 208 comparison rows and 52 favorable flags; 0502 retains a matched timing summary and an explicit regression review. | Claims are scoped to named workflows and evidence; they do not establish program completion. |
| Correctness and boundaries | Recent bundles retain source/build manifests, exact output/source oracles, budget/cancellation tests, preservation checks, and ADR matrices; 0495–0497 cover managed edit and an atomic logical-tail capability. | Strong scoped DOCX/OPC and ODG correctness/custody evidence; general atomic save, all formats, and full CRUD remain unproven. |
| Provider and cache states | 0491 and 0494 cover owned, file, instrumented, short, simulated delayed/range, and verified-cold DOCX lanes; 0492/0493 measure and integrate bounded read-ahead. | Descriptive and opt-in evidence only. Genuine borrowed lifetimes, physical-device cold behavior, native producers, and cross-format intersections remain open. |
| Hardware/resource profiling | 0490 sync traces, 0496 phase clocks, 0501 whole-child profiles, 0502 heaptrack/layout observations, and allocator/RSS reports are retained with scope limits. Hardware counters are unavailable for 0502. | Attribution is partial: whole-child/setup-inclusive evidence is not operation-local CPU, allocation, or RSS proof; lock-wait and full scaling evidence remain open. |
| Parallelism and batching | 0498/0499 provide explicit bounded ordered Part batches, and 0500 measures managed paragraph batching; adverse local rows and serial controls remain visible. | Partial capability evidence. Full lifecycle 1/2/4/8 scaling, serial-fraction/Amdahl analysis, and stable lock-contention evidence remain open. |

The retained 0486–0502 evidence establishes several bounded capabilities and
descriptive baselines, including managed edit ownership, atomic logical-tail
publication, source read-ahead, Part batching, and PPTX/ODG scoped experiments.
It still does not establish the `docs/GOAL.md` requirements for complete
p50/p95/p99 coverage, throughput, allocations, peak RSS,
copied/decompressed/recompressed bytes, physical I/O, lock wait, borrowed
inputs, native producers, all cold/warm intersections, or Amdahl scaling. The
0502 open regressions, 0499/0500 local adverse rows, and 0501 missing replay
binary remain explicit rather than being absorbed into a completion claim.

## ZIP64 passthrough audit

The low-level foundation is now materially further along than the last
committed control. `soapberry-zip` retains ZIP64 field origins and tail bytes;
its preservation plan preflights the complete output, patches only movable
offset fields, retains central/local metadata and ZIP64 extensible data, and
has malformed, multi-disk, truncation, descriptor, limit, offset, sink, and
no-output tests.

The earlier integration changed public preservation construction to
`AllowZip64` and admitted already-ZIP64 sources. Change 0414 also promotes
generated/copied central offsets and synthesizes ZIP64 tails for ZIP32 sources.
OPC size/count capability guards and the provenance count ceiling are removed;
arithmetic, limits and structural refusals remain. Earlier synthetic tests
continue to cover:

- `zip64_source_targeted_save_preserves_untouched_records_and_tail` changes
  one Part in a ZIP64 package, retains a different Part's raw local and central
  records, retains ZIP64 tail extensible bytes and the comment, reopens the
  output, and compares `to_bytes` with `write_to_stream`.
- `projected_zip64_descriptor_targeted_save_preserves_untouched_member`
  changes one Part while a different member uses ZIP64 data-descriptor fields,
  then checks raw preservation, reopen, and stream output.

Focused source-backed coverage now also includes
`zip64_one_part_overlay_preserves_raw_unknown_members_and_cold_work` and
`topology_zip64_add_remove_preserves_untouched_records_and_tail`. Both pass in the final integrated test run. The
package-writer suite also has a ZIP64 partial-sink failure case.

These tests are meaningful end-to-end package-writer correctness evidence, but
they are synthetic in-memory fixtures. They leave the following gates open:

1. Extend the source-backed selected-Part and topology matrix to signed/encrypted inputs,
   both descriptor forms, non-seek sinks, cancellation, configured metadata
   limits, and atomic filesystem finalization. Each refusal must leave the
   sink/destination untouched where the API promises that property.
2. Exercise ZIP64 add/remove/topology plans with more than the current narrow
   synthetic graph, including unknown physical members and dependency closures.
   The existing focused add/remove test is valuable correctness evidence but
   does not certify every topology operation.
3. Complete generated-size and streaming creation coverage. Change 0414
   implements offset/count promotion and known-size Store/precompressed ZIP64
   headers. Public sparse tests cover generated offsets immediately below, at
   and above `u32::MAX`; OPC tests cover count promotion and repeated owned
   publication beyond 65,535 members. Copied-offset promotion has focused
   layout/metadata tests. Change 0415 adds explicit one-pass streamed Deflate
   framing and raised-limit Office transport admission, including generated
   metadata charging and descriptor-backed local placeholders. Preservation
   regeneration still buffers payloads; transport-level bounded-memory evidence
   does not certify semantic row/paragraph/slide creation or append.
4. Add at least one real-producer or independently generated large ZIP64 OPC
   corpus and validate all semantic Part bytes, raw member identity, archive
   layout, and reopen behavior. Change 0415 supplies an independently generated
   Python OPC source with a real 4 GiB logical member and a small XML overlay
   preservation check. The 0416 batch adds deterministic Python forced-ZIP64
   local-framing fixtures and focused ZIP/OPC no-op, edit, and refusal cases;
   these remain synthetic compatibility evidence, and the final gate status is
   deferred to its change record. Native Office producer coverage, ODF
   prepared-catalog classification of local-only sentinels, and the broader
   dependency/topology matrix remain open.
5. Expand the scoped gates in
   [`changes/0404-zip64-preservation-integration.md`](changes/0404-zip64-preservation-integration.md)
   to the full non-iWork feature and native-producer matrix. The scoped final
   commands, source bindings, and pass/fail logs are now retained.

The control behavior is also important: at the control revision, a changed
owned ZIP64 source was refused rather than normalized. That refusal remains the
safe fallback for unsupported framing, suffixes, opaque topology, and
unrepresentable generated output. A successful targeted test does not authorize
a normalizing fallback for a different unsupported source.

## Prioritized remaining work

| Priority | Requirement from `docs/GOAL.md` | Next reviewable evidence |
| --- | --- | --- |
| P0 | Resolve the 0502 ODG open regressions | Profile parser and per-shape metadata work with operation-local attribution, then capture a same-protocol before/after pair for plain and metadata corpora. Keep the boxed-layout/heaptrack result separate from latency and allocation-count claims. |
| P0 | Restore replayable 0501 custody or recapture it | The retained 0501 reports and historical receipt are usable, but `/tmp/litchi-goal-0501` no longer contains the frozen before executable. Rebuild from the recorded source/build manifest or retain a new verified replay bundle before claiming current independent replay. |
| P0 | Complete the Phase-1 metric and CRUD baseline | The 6,030-sample default refresh closes the timing-report gate for 11 measured mappings (48 rows) within the 33-selector index, but the 15-category checklist still has correctness-only/unsupported rows and lacks complete throughput, copy/decompression, lock-wait, cold/warm, and scaling coverage. Promote only rows with validated timing evidence. |
| P0 | Measure provider and publication intersections | Extend the 0491–0494 DOCX provider/cold baselines and 0497 atomic logical-tail capability to genuine borrowed input, sequential non-seek sinks, filesystem atomic save, source-version/cancellation failures, and physical cold behavior across representative formats. |
| P1 | Address the retained local batch and lifecycle adverse rows | Revisit 0499/0500 serial-versus-batch choices and their flagged latency/RSS rows with operation-local CPU/allocation evidence. Publish full explicit 1/2/4/8-worker curves, lock-wait proxy, and serial-fraction/Amdahl analysis before widening concurrency. |
| P1 | Finish source-backed CRUD adoption across formats | Extend selective open/read/edit/save and dependency-closure publication through DOCX, XLSX, PPTX, XLSB, DOC, XLS, PPT, and ODF owners; measure physical I/O, decompression/recompression, copies, allocations, RSS, and semantic phase boundaries. |
| P1 | Cover the high-impact CRUD categories and real producers | Add or explicitly classify conversion, creation, append variants, structural/deletion/sanitization, cross-document copy, merge/split, patch/inverse/three-way merge, repair/normalize, dynamic content, security, malformed, signed/encrypted/macro-enabled, and independently produced Office corpora. |
| P1 | Close ZIP64/CFB and output-source preservation intersections | Exercise ZIP64 and CFB topology/stream cases with non-seek sinks, cancellation, configured limits, physical cold/high-latency sources, unchanged-member passthrough, and atomic finalization while retaining typed refusals and exact untouched bytes. |
| P2 | Apply layout, cache, or SIMD tuning only from measured hot loops | Require operation-local profiles, scalar fallbacks, differential malformed-input tests, and material end-to-end benefit before any low-level change. |

No row in this audit should be read as a completion claim. The current evidence
supports targeted capability progress, including the 0491–0494 provider/cold
baselines, 0495–0497 managed publication capabilities, 0498–0500 bounded batch
work, the 0501 PPTX comparison, and the 0502 ODG audit. The retained
regressions, missing replay executable, incomplete CRUD metrics and provider
intersections, and unmeasured scaling keep the full non-iWork performance goal
open.
