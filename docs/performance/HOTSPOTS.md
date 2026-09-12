# Performance hotspot inventory

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

## 0509: ODT sink allocation churn reduced

[0509](changes/0509-odt-sink-buffer-reuse.md) profiles and removes repeated
small paragraph-buffer allocation: 200,000→20 calls over twenty large exports.
The parser keeps one cleared spare with actual capacity at most 4 KiB.
Callgrind parser-inclusive references fall 1.83%, including the checking sink;
`normalized_xml10_decoded_len` remains 28.35M exclusive references (9.49% of
the candidate profile). Its validation/normalization boundary needs a separate
proof before optimization. SHA compression occupies 44.07% under Callgrind's
CPU dispatch, which is not a native timing share. ODS repeated row preflight
and local provider wave/channel overhead remain separate attribution tasks.

Initial large p99 +9.14% is retained. An 8,000-sample follow-up reports
+2.18%/+2.78%, below the review threshold, with modest median reductions.
No tail-latency or peak-memory improvement is claimed. The broader priority
queue below remains open.

## 0508: export baselines and current priority reconciliation

[0508](changes/0508-default-semantic-text-export.md) adds twelve descriptive
text-export rows while preserving the old default identities. ODT-large emits
19,999 sink writes and ODS-large 65,535, with 49/54-byte largest writes. Those
output boundaries are candidates for future attribution, not proof that
coalescing is the dominant cost. Matched profiling is required before changing
traversal, error order or sink behavior. Eighteen index selectors remain
correctness-only. The [ODG comparison audit](results/change-0508/odg-priority-review.md)
keeps the accepted recent improvements separate from an unmatched historical
endpoint comparison; further ODG work should use the current residual costs.

## 0507: contiguous ODG value walks batched

[0507](changes/0507-odg-attribute-value-batches.md) records this batch.
Two groups of eight requests remove 14 repeated checked value walks per
shape. Fresh scalar-helper attribution is 355.3M inclusive references, 241.3M
directly from the main scanner. Whole-child references fall 17.15%; plain
p50 improves 32.74–33.51%, metadata 17.06–19.64%. Plain-large RSS falls
9.08%, with no paired >5% adverse flag. Remaining scalar lookups account
for 153.2M references and parse_content for 349.4M. Future priorities must
consider these reduced costs alongside outstanding CRUD/provider coverage;
the older 0502 comparison remains a separate evidence question.

## 0506: repeated ODG shape-span walks batched

[0506](changes/0506-odg-shape-attribute-span-batch.md) records this batch.
One checked pass replaces 16 shape-span lookup walks. The fresh baseline
attributes 18.51% of whole-child instruction references to source-span lookup;
the candidate reduces whole-child references 15.89%. Plain p50 improves
26.81–28.29%, metadata p50 17.00–18.06%. Plain-large RSS rises about 10%
(~2 MiB), an explicitly accepted tradeoff. Repeated semantic value lookups
remain roughly 355M inclusive references and are the next measured ODG
work-elimination opportunity; no validation removal is authorized.

## 0505: avoid unrelated ODG namespace lookups

[0505](changes/0505-odg-attribute-name-prefilter.md) records this batch.
Filtering local names before namespace lookup reduces whole-child Callgrind
instruction references by 12.38%; exclusive namespace resolution falls
149.9M to 36.8M references. Checked attribute iteration and duplicate checking
remain unchanged and prominent. Final p50 improves 8.43–9.88% across all
eight pairs, without a >5% adverse measured flag.

## 0504: repeated ODG transition owner scans reduced

[0504](changes/0504-odg-direct-transition-reuse.md) reuses successful direct
style values within one resolver invocation, retaining per-page inheritance
checks. The fresh Callgrind owner-scan cost falls 557.4M → 34.8M inclusive
instruction references. Metadata-large open/traversal p50 falls 34.43%/34.32%
in reversed repeats; metadata-small falls 9.77%/10.15%. No measured paired
latency/throughput/RSS regression exceeds 5%. These are synthetic owned-byte
scenarios. `parse_content` remains roughly 850M inclusive references in the
matched whole-child profile; repeated attribute walks and unique-style scans
remain candidates for attribution. The older 0502 regression and broader CRUD
requirements are not declared closed.

## Review queue reconciled through 0509

The retained evidence changes the next investigations. These priorities concern
measured scenarios; the full taxonomy and provider matrix remain required.

| Priority | Observed cost | Next evidence needed |
| --- | --- | --- |
| P1 | 0502 recorded historical richer-parser regressions; matched 0504–0507 captures establish subsequent reductions. | Use current residual parse/attribute profiles for optimization. A fresh, fully source-bound old-to-current comparison is needed for a causal historical-closure claim; retain richer semantics. |
| P1 | 0499 many-small local Part batches remain slower than serial despite worker reuse; delayed batch-8 p99 rises 30.79%. | Isolate queue/wave overhead and delayed-provider tails with matched operation-local observations before changing admission policy. |
| P1 | 0500 K=1 p128 owned batch lifecycle p50 rises 7.34%; p512 K=8 warm-file RSS rises 5.77%. | Recheck the publication-phase association and memory in matched repeated runs; larger batches do not erase these flags. |
| P1 | 0508 full runs validate 60 mapped case/corpus rows; 18 selectors remain correctness-only. | Extend measured CRUD/provider and native-producer coverage from the checked taxonomy, retaining source, output, and preservation oracles. |

The 2026-09-11 evidence audit reproduces the 0502 summary, including its
bootstrap intervals, from all 16 raw timing reports. It also validates both
0501 full-run report/catalog pairs. These are retained-data checks, not new
benchmark runs or current-worktree performance measurements.

## 0502: ODG metadata open; regressions and heap reduction retained

[0502](changes/0502-odg-metadata-open.md) records four deterministic ODG
corpora, two serial repeats, 25 warmups, and 200 samples per child (1,600
measured samples per phase). Optional-pass gating and parsed-style reuse remove
duplicate work. The final candidate remains 5.9–6.4% slower on plain p50 and
53.4–111.1% slower on metadata p50 than the old parser, which exposed less
metadata. These remain follow-up hotspots, with within-run bootstrap intervals
and whole-child RSS retained. Boxing optional metadata reduces `Shape` from
936 to 800 bytes (old baseline: 792); a separate heaptrack comparison reduces
plain-large peak heap from 13.69M to 12.57M with unchanged allocation counts.
Hardware counters were unavailable. These scoped measurements do not establish
a CRUD speedup or complete the broader performance goal.

## 0501: redundant PPTX payload hashing is a scoped preparation hotspot

[0501](changes/0501-pptx-exact-payload-comparisons.md) removes only the image
and chart decoded-byte contribution to private `digest_touched`, retaining
exact payload equality, candidate reread, graph/metadata proof, and 64 KiB
cancellation checks. The fresh before media-rich API p50 is 24.297/24.337 ms
owned and 28.588/28.910 ms warm-file. The before supplementary profile reports
SHA-256 compression at 31.01% for owned and 30.08% for warm-file whole-child
samples; setup, untimed work, and incomplete callgraph coverage prevent
touched-digest attribution. The first owned export reused the same recorded
data after local symbol recovery, followed by two owned profiling workload
reruns; this supplementary history does not change the formal 240 samples.
Historical targeted caller attribution is about 15.5–16.1%. Each media-rich
fixture has eight 2 MiB image payloads, so the scoped change removes 32 MiB of
redundant payload hash input across the two preparation passes. The matched after
phase cuts media-rich API p50 by 61.481%/61.467% owned and 51.452%/51.870%
warm-file across R1/R2; all 52 retained >5% flags are favorable and no adverse
change exceeds five percent. Plain warm-file R2 API p50 rises 1.347% and p99
1.632%, below the review threshold. The separate default CRUD refresh passes
both 201-row lanes (6,030 measured samples; 38 static coverage tests), with
47.42/46.57 seconds wall time and 151,028/160,568 KiB maximum RSS. It supplies
the current timing-report baseline only. The scoped gates also pass: default
all-targets 848, all-features library 552, doctests 6 with 2 ignored, focused
58, private guards 5, plus Clippy, fmt, rustdoc, downstream, and boundaries.
The retained final verification receipt passed at capture time. On 2026-09-11,
rerunning that verifier stops because the frozen control binary is missing;
this is distinct from the independently revalidated retained report/catalog
pairs. The historical cleanup receipt records 2,154,708,992 unique-inode
allocated bytes removed and eight replay files retained in local tmpfs at that
time, not a durable binary archive. The full goal remains open.

## 0500: repeated managed reconstruction is the measured hotspot

[0500](changes/0500-managed-paragraph-batches.md) targets the existing managed
scalar paragraph path, which reconstructs and validates a source-backed
candidate for each selected paragraph. The repeated baseline spends about 93%
of lifecycle time in edit at 32 selected paragraphs. The same-final-executable
batch route reduces K=32 lifecycle p50 by 6.947–7.138x and edit p50 by
12.640–13.456x; K=8 lifecycle improves 2.324–2.346x. The p128 K=1 owned
lifecycle rises 7.34%, with publication p50 up 15.23% while edit p50 falls
0.35%; this phase association is descriptive. Whole-child RSS and perf
counters include setup and verification, and the warm synthetic corpus does
not establish cold-source or native-producer behavior.

## 0499: worker creation is reduced, but local batch overhead remains

[0499](changes/0499-operation-local-part-workers.md) reuses operation-local
scoped workers across ordered Part waves with bounded per-worker command and
reply queues. The old single-wave route remains in place. This targets the
repeated worker creation observed in the 0498 scheduler; the matched capture
has 3,600 measured samples and 360 warmups. Many-small owned p50 improves
3.59x/2.92x/2.08x and warm-file 2.64x/2.26x/1.81x at widths 2/4/8, yet both
remain slower than ordinary serial local reads. The capture retains five
aggregate and ten per-repeat latency/throughput flags, including a sparse
few-large width-eight tail cluster. Traced many-small width-four creation
falls 64→4 successful `clone3` calls; traced counts are diagnostic evidence,
not timing evidence, and synthetic delay remains a provider model rather than
a real remote-service result.

## 0498: Part-read concurrency helps delayed providers; small local reads remain serial candidates

[0498](changes/0498-bounded-source-backed-part-batch.md) measures explicit
production Part batches. Delayed reads benefit from overlap, but per-Part
worker waves regress on small owned/file reads. Thread creation is a plausible
cost, not isolated attribution. Few-large local observations are superlinear
and do not support a simple Amdahl fit. No global pool or automatic facade
parallelism was added; future work should measure task grouping or reuse before
expanding the scheduler. Warm file and synthetic-delay results do not close the
cold-cache or real remote-provider requirements.

## 0497: atomic publication capture is descriptive; hotspot attribution remains limited

[0497](changes/0497-docx-atomic-publication.md) moves the bounded DOCX
logical-tail append route toward an explicit filesystem destination. The
consuming plan/commit methods use OPC's sibling temporary, data synchronization,
replacement, and parent-directory synchronization path. The atomic route is
after-only, while the unchanged hashing sink is the only before/after control.
Counting is a non-retaining after-only capability.

The measured interval includes the production atomic publication work and
publication drop; destination readback, semantic/raw verification, fixture
inverse checks, replay cleanup, and process/allocation endpoint snapshots are
outside it. The frozen formal matrix contains 288 children and 8,640 samples,
with a separate 72-child/216-sample pilot. Formal verification passes all 288
child terminals, including the 97-child resume after the original `ENOSPC`
interruption. The analysis retains 864 matched default-hashing comparison cells
and 144 after-only capability rows, with all 100 >5% adverse flags retained:
72 allocator live-byte endpoint, 16 latency, and 12 whole-child RSS.

These records support descriptive route and lifecycle observations only. The
default hashing sink is the only before/after comparison; counting and atomic
routes are after-only, so no general speedup, durability, cross-platform, or
atomic-performance claim is authorized. The original ordinal-191 raw
observation remains archived without fabricated evidence. Emergency cleanup
overlapped two formal whole-child intervals, so no isolated-host or
individual-outlier attribution is made. Final cleanup verification passes. The retained failure custody, diagnostic
syscall scope, and fixture-only inverse oracle continue to bound interpretation.

## 0496: publication dominates named normal phases; attribution remains descriptive

[0496](changes/0496-docx-edit-phase-attribution.md) targets the unresolved
0495 eight whole-child RSS and three latency flags with an opt-in diagnostic
overlay. The verified run has 32 processes and 960 samples across before/after
unmanaged owned, warm-file, and short-range providers plus after-managed
warm-file and short-range roles, in normal and allocator binaries.
The paired unmanaged review contains 294 comparison cells, with 74 threshold
flags retained; those flags remain descriptive and do not establish causality.

Publication is the largest named phase in all 16 normal children. After-managed
file-warm lifecycle/publication p50s are 3.563/2.465 ms and 3.553/2.453 ms;
short is 6.509/5.406 ms and 4.092/2.993 ms, so the short arm is unstable.
The 74 flagged rows separate into 60 phase, 9 allocator-reallocation, 4
whole-child RSS, and 1 lifecycle-latency rows. Named phase values are wall-clock
`Instant` intervals, not CPU attribution. Full-lifecycle allocation counters
and whole-child RSS remain separate; the non-reentrant allocator observer
provides no nested phase allocations. Existing flags retain unresolved
causality, and the full non-iWork goal remains open.

## 0495: managed ordinary DOCX edit is enabled; attribution remains open

[0495](changes/0495-docx-managed-document-edits.md) measures the new finite
owner-retained ordinary managed DOCX edit/save path across six provider arms,
normal and allocator roles, and two reversed repeats. The formal run retains
72 processes and 2,160 samples; pilot2 retains 36 processes and 108 samples.
The normal managed p50s in milliseconds are 5.592/5.556 (owned), 3.277/3.298
(instrumented), 5.682/5.853 (file-warm), 4.066/4.060 (short), 572.748/576.147
(delayed), and 182.327/181.990 (range-zero).

These are managed capability and finite-budget baselines, not managed-before
speed or provider-ranking evidence. The paired performance review compares
unmanaged before/after rows only and retains eight whole-child RSS flags plus
three latency flags over five percent. Causality remains unresolved, and
repeat instability remains visible. Every managed formal row passes output,
source, semantic, media, and finite-budget checks, while allocator-role rows
pass allocator conservation. Normal rows have no allocator counters, and
unmanaged rows have no managed budget fields; managed resource
memory/objects/depth return to baseline.

The phase labels matter: `cache_before` is post-open, `cache_live` is
post-edit/pre-publication, resource `live` is post-publication, and
`after_drop` follows package consumption and returned-snapshot/commit release.
There is no cache-after-drop gauge. Profiles cover whole-child setup,
verification, and serialization, so the next hotspot work is phase-local CPU
and whole-child RSS attribution, repeat capture of unstable file-warm and
short/allocator arms, and a separate post-publication cache observation if
retention is claimed. The full non-iWork goal remains open.

## 0494: opened DOCX edit/save exposes range-request cost

The six-provider baseline records 377 nonempty reads for one paragraph edit and
sequential save, or 4,217 calls through a 4 KiB short-read adapter. Both return
16,799,430 logical bytes. The simulated 1 ms/request provider therefore adds
377 ms of fixed service, above 160.212 ms of modeled transfer service. This
makes request coalescing during publication a measured investigation target;
the 0493 selective-read result does not prove an edit/save improvement.

One main Part is materialized and unchanged media survives every output check.
Warm allocator operations add 606,986 peak live bytes over their preallocated
baseline. Provider-local timings and 120 verified-cold observations vary greatly
between repeats, so these results establish neither a stable local ranking nor
a production speedup. Managed ordinary edits, true borrowed lifetimes, atomic
save, native producers, and concurrent scaling remain open. See
[0494](changes/0494-docx-edit-provider-baseline.md) for individual results and
[methods](results/change-0494/methods.md) for timing and memory boundaries.

## 0493: managed source read-ahead reduces delayed DOCX requests

The opt-in production OPC window, forwarded through DOCX, reduces the pinned
managed full-text lifecycle from 19 physical reads to 3. The 480-sample formal
matrix records 82.17–84.87% lower normal-executable medians under the simulated
1 ms plus 100 MiB/s provider. Zero-delay medians change +0.98% and −2.01%, so
this does not establish a local-source improvement.

The cost is 1,479 extra input bytes (+37.29%), including 69 compressed media
bytes, plus three allocations and 4,384 allocator bytes. Window Memory is
reserved and physical InputBytes are charged; package publication releases the
window before exact traversal. Queued managed readers retain cancellation
polling. Delayed tails vary substantially on the shared host; no scaling or
real-network result follows. The next gaps are opened DOCX edit/save provider
coverage, bounded concurrent scaling, independent producers, and genuine
borrowed lifetimes. See [the change record](changes/0493-managed-opc-source-read-ahead.md)
and [raw evidence](results/change-0493/README.md).

## 0492: bounded range reads remove simulated request service

[0492](changes/0492-docx-bounded-range-read-ahead.md) verifies the range-locality
hypothesis with 480 formal before/candidate samples. A private 4 KiB forward
window reduces 19 synthetic transport calls to three, with identical logical
requests and text. Delayed medians are about 83% lower; normal zero-delay
repeat 1 p99 regresses 5.20%. Physical bytes increase 37.29%, including 69
compressed-media bytes. The fixed window is allocated outside the timed region.

Keep this as an opt-in range-source opportunity. The immediate production work
is OPC-owned physical-fill budget charging and retained-window memory accounting,
with cancellation, source-version, and exact-publication tests. The benchmark
wrapper alone does not satisfy those production requirements. See the
[individual review](results/change-0492/results-review.md) and
[implementation plan](results/change-0492/production-next.md).

## 0491: DOCX source-provider and cache-state baseline

[0491](changes/0491-docx-provider-and-cold-baseline.md) adds a formal source-backed
full-text baseline: 600 provider samples and 360 fresh-child filesystem samples,
with four explicit prepared-query cold-ineligible controls. Normal provider
medians are about 0.27–0.29 ms; the simulated 1 ms-per-request arm is about
20.3 ms with 19 paced calls. Its first repeat has wider tails, retained in the
[individual review](results/change-0491/results-review.md). These are baseline
observations, not before/after improvements.

The actual timed text is now authenticated after its explicit versioned clock.
Verified-cold alignment retains raw EOCD-tail compressed-payload overlaps and
proves zero cache loads at open, followed by one main-part materialization.
Heap peak increments, absolute process memory, source counters and whole-child
profiles remain separate. No production optimization is included.

Bounded range coalescing is now a measured opportunity; evaluate it against
these request counts while preserving source identity, resource bounds and
lossless package semantics. Genuine borrowed lifetimes, concurrency/scaling,
native producers, publication and broader CRUD intersections remain open.


## 0490: file-store synchronization dominates the tiny route

[The controlled follow-up](changes/0490-file-store-variance-and-sync-attribution.md)
retains six alternating blocks and all 4,320 samples. Normal file-store p50
improves 6.73% on average, while tail intervals span both directions and
individual adverse blocks remain. Operation heap decreases; procfs delta/high-
water observations and every flagged quantile remain explicit in the review.

Sync-only traces attribute about 79–82% of these traced operation medians to
fdatasync. The sync policy and selected syscall counts are unchanged. Removing
sync or freshness checks is not justified. The next priority is the
[verified-cold/source-provider contract](results/change-0490/next-implementation.md)
for an existing end-to-end DOCX selector, followed by bounded concurrency and
independent native producer evidence. Prepared warm queries and copied slices
must not be relabeled as cold or genuinely borrowed inputs. The goal remains open.

## 0489: candidate audit reuse; file-store tails and metadata remain

[0489](changes/0489-opc-candidate-xml-audit-reuse.md) reuses a successful initial
candidate XML audit inside one private immutable prepared OPC splice plan.
Source-heavy owned/file medians improve about 20–21%, authored-heavy file
medians improve 14–15%, and deterministic operation heap falls about 19.89%.
Source/replay/candidate byte authentication, initial source and candidate audits,
frozen limits, conservative workspace admission and final reopen remain required.

[The review](results/change-0489/results-review.md) retains a small file-store
repeat with normal p95/p99 increases of 79.00/82.60%, allocator tail increases,
and four RSS observations above +5%. Their cause remains open. All candidate
archive bytes match. Whole-child instructions decrease, but file statx and
pread64 counts are unchanged, leaving freshness metadata as a measured cost.
Broader CRUD/provider intersections, cold-cache, concurrency and native Office
validation remain open. Batch-local Cargo output is removed after validation.

## 0487: fewer replay sink fences; candidate audit cost remains

[0487](changes/0487-opc-replay-consumed-prefix-retention.md) retains consumed
replay bytes across short reads within the existing adapter window. Both
authored-heavy file repeats improve p50 about 36%; statx diagnostics fall
59.90%. Source-heavy medians stay within 1%, and operation heap is effectively
unchanged. The [review](results/change-0487/results-review.md) retains small-input
latency regressions and three RSS increases above 5%.

The 0486 caller attribution and this reduction support sink-fence overhead as
one cost. Repeated candidate XML audits remain a possible next experiment;
source-only proof reuse is insufficient. Any retained candidate-audit capability
must preserve frozen limits, complete byte authentication, freshness, Work,
cancellation, accepted-output accounting and final reopen. This batch retains
all XML audits and does not close the broader non-iWork goal.

## 0485: lower splice overhead with remaining metadata and audit cost

[0485](changes/0485-opc-splice-consumed-window-batching.md) coalesces hashing
and sink output for already-consumed parser bytes in the private OPC window.
Authored-heavy file p50 improves by 45.76–46.05%, with effectively unchanged
operation heap and exact candidate archive bytes. The separate source XML
audit, replay authentication, per-fragment Work charges, and FileSource
freshness policy remain intact.

The [profile review](results/change-0485/results-review.md) shows
authored-heavy file `statx` calls dropping from 3,735,939 to 1,475,055.
Substantial metadata overhead remains: owned p50 is about 107 ms while file
p50 remains about 239 ms. Source-heavy file metadata calls increase from
15,415 to 25,219 even as p50 improves by about 19%; source/replay callback
fences are a retained tradeoff. Logical I/O and `pread64` counts are unchanged.
Whole-child instructions drop by 10.57–32.78% in the selected diagnostics.

Further caller-level profiling is needed before reducing remaining freshness
checks or replay/audit work. The batch also retains a normal p99 regression
on injected-latency authored-heavy input and nine small-workload RSS increases;
these open concerns must not be hidden by lower medians. Broader cold-cache,
concurrency, and native Office validation remain open.

## 0484: authored replay cost and file-input metadata investigation

[0484](changes/0484-docx-replayable-tail-stream.md) implements a replayable
multi-paragraph append route and explicit one-shot storage. The deterministic
route opens five cursors across sealing, validation and publication; chunk
count does not multiply the number of opens. The existing XML audit and replay
costs remain present. The [formal measurements](results/change-0484/README.md)
characterize these routes without claiming a historical improvement over 0483.

The 228-process matrix varies source and authored size independently and adds
input, sink-window and Store/Deflate profiles. At 16,384 authored paragraphs,
memory-store p50 is about 90 ms versus deterministic 117–119 ms, with a larger
operation heap reservation. At 131,072 source paragraphs all three routes
remain near 475–479 ms. One-shot routes emit one authored pass with zero cursor
opens, then authenticate four replay-reader passes.

File input on the authored-heavy case takes about 444 ms in repeat one despite
the same 60 logical reads and 7,651 returned bytes as owned input. Metadata
syscalls need explicit attribution before choosing an I/O optimization. The
[audit investigation](results/change-0484/repeated-audit-investigation.md)
also rules out deleting the standalone source XML audit: candidate validity
does not prove source validity in the generic splice API. Profiling must
justify a replacement that retains both proofs and error ordering.

## 0483: bounded DOCX heap with repeated audit/replay CPU cost

[0483](changes/0483-docx-bounded-tail-append.md) removes full source/candidate
XML and the paragraph range vector from a new one-paragraph tail route. Its
609,875-byte incremental operation peak is invariant across the tested source
sizes, but normal latency increases about 98–104%. Logical source calls rise
from 62 to 119 on the largest source. Whole-process RSS stays near 100 MiB.

The [profile review](results/change-0483/profile-review.md) records 1.934×
whole-process instructions. Bounded lifecycle intersections include streaming
XML auditing and replay; those inclusive stack shares overlap. Phase
attribution should establish which repeated work can be eliminated while
preserving semantic validation, freshness and publication proofs. The
[authored-stream design](results/change-0484/design.md) preserves the separate
requirement for many generated paragraphs, bounded replay and durable patches.

## 0482: prerequisites for bounded DOCX publication implemented

[0482](changes/0482-bounded-xml-opc-splice.md) adds the finite XML reader audit and decoded OPC insertion path selected
by the 0481 window contract. It avoids retaining complete source/candidate XML
inside these primitives, with separate fragment, parser, ZIP and source/sink
ownership. The next measured bottleneck remains the public DOCX scanner's
materialized XML and paragraph index. Integrating the new path, a replayable
paragraph producer and durable patches remains necessary before claiming
bounded end-to-end append memory.

The XML comparison confirms bounded operation heap for the tested reader
profile: 65,587 bytes at all three input sizes. Timing remains a tradeoff:
the 8 MiB normal case is 3.6–4.1% slower, while the 128 MiB case is
10.6–10.8% faster than materialization with a growing `Vec`. Public DOCX
profiling must establish whether those observations transfer to its denser XML.

## 0481: per-event DOCX name copies removed

[0481](changes/0481-docx-borrowed-scanner-names.md) removes temporary local-name
vectors from the measured source-backed scanner. Exact allocation removal is
`24*N+28` callbacks and `24*N+108` requested bytes in the named tail-copy
lifecycle; both repeats show lower normal means. Full XML/index retention,
repeated scans and exact-one range growth remain. The separate
[window contract](results/change-0481/window-contract.md) identifies the OPC
decoded-splice replay and multi-paragraph producer work still required.

## One DOCX publication XML copy removed (0480)

[0480](changes/0480-docx-shared-publication.md) passes the immutable target Arc into the existing shared OPC overlay
entry point. It removes one payload allocation and the separate Arc allocation;
the large measured operation peak falls by 6,422,768 bytes. The underlying
validation/writer is unchanged. Complete XML snapshots, repeated scanner work
and paragraph indexes remain significant; this does not close the explicit
window requirement for logical append.

## DOCX logical append retains document-sized state (0479)

[0479](changes/0479-docx-tail-append-baseline.md) measures the current one-copy tail transaction. Its incremental
operation heap grows from 509,974 bytes at 64 paragraphs to 41,793,870 bytes at
131,072 paragraphs. Snapshot, staged candidate and publication owners retain
whole main-story XML and paragraph indexes. The phase and CPU evidence in the
linked bundle guide the next compatible production change. Requested allocator
bytes include complete realloc sizes and must not be interpreted as physical
memory-copy traffic. The explicit-window append requirement remains open.

## Generated-name retention removed from the explicit PPTX route (0478)

[0478](changes/0478-pptx-generated-metadata-spool.md) removes growing ZIP Office
and OPC name indexes from a checked, finite generated-name mode and integrates
central-directory scratch into the public fresh PPTX writer. The existing
serializer emits the same members in the same order. This addresses the name
and directory owners identified in 0475 after compressor reuse in 0476.
Serialized scratch and per-member work still grow; the provider's storage is
separate from the library's replay buffer and active-entry state. Default
arbitrary-name paths retain their ordinary metadata. The next investigation is
[DOCX logical-tail append](results/change-0478/next-work.md), with a baseline and
attribution required before any production change.

The final 24-process, 720-sample matrix passes byte and reopen verification.
All 180 spool allocator samples have a 432,436-byte operation peak, versus
8,875,252 bytes for the 8,192-slide control. Scratch extents are 4,074, 40,842
and 1,237,692 bytes at 8, 256 and 8,192 slides. Small-deck normal mean latency
increases 3.11% and 3.31% in the two repeats; no registered 5% review threshold
is crossed. This is an operation-heap result with separate caller storage.

## Central-directory working heap separated from scratch (0477)

[0477](changes/0477-zip-central-directory-spool.md) gives the ZIP writer
explicit replayable central scratch. At 8,192 members, Store peak heap changes
from 1,196,032 to 16,574 bytes and Deflate from 1,608,992 to 429,534 bytes. The
File provider adds one syscall per member and the temporary records add
allocations; large Store medians rise to about 4.3 ms from about 3.1 ms.
Retained ZIP/OPC name indexes remain the next public streaming memory owner. See
[next work](results/change-0477/next-work.md).

## Repeated owned compressor allocation addressed (0476)

[0476](changes/0476-zip-deflate-state-reuse.md) removes repeated backend and
output-buffer allocation between successfully finalized owned ZIP members.
Large PPTX requested operation allocation work falls 99.538473%, while peak
heap is effectively unchanged and still grows with member count. The next
memory design must account for ZIP directory/name storage and OPC/streaming
name indexes. See [remaining work](results/change-0476/next-work.md).

## PPTX writer and preflight costs separated (0475)

[0475](changes/0475-pptx-streaming-attribution.md) retains repeated CPU and
heap traces for the unchanged 8,192-slide streaming writer. Materialized
preflight dominates whole-process CPU and allocation counts, so exact run
ancestry is required. Within the writer, Deflate processing has roughly 51%
of sampled cycle weight, while initialization has roughly 5%. Compressor
state reuse is an allocation-work candidate, with independent members and
error/budget behavior to preserve. Growing name/directory metadata remains
a separate memory problem.

Repeated exact-stack attribution assigns 99.543891% of run-context requested
bytes to backend initialization plus encoder output buffers. This selects a
measured allocation-work candidate; it does not establish a speedup or remove
the persistent metadata-memory requirement.

## PPTX streaming metadata growth measured (0474)

[0474](changes/0474-pptx-streaming-operation-memory.md) measures operation peak
growth from 435,541 to 8,875,092 bytes across 8 to 8,192 fresh slides. The large
operation also requests 6.81 GB of allocations. The public writer emits XML
directly, but OPC/ZIP name validation and ZIP header/name storage persist until
finalization. Fresh Deflate setup and transient name preparation are additional
allocation-work leads; no per-owner peak attribution is claimed. The next step
is stack attribution before a shared transport change. An explicit directory
spool alone would leave growing name indexes. See [next work](results/change-0474/next-work.md).

## DOCX streaming memory gap measured (0473)

[0473](changes/0473-docx-streaming-operation-memory.md) measures the already-public
StreamingDocumentWriter rather than buffered Package creation. Incremental
operation heap is 414,732 bytes in every allocator sample across a 2,048-fold
paragraph-count range. Requested bytes/calls grow with input, so this does not
remove allocation work or prove general RSS bounds. The next streaming gap is
PPTX, whose retained ZIP directory/part metadata requires explicit measurement
and accounting before a fixed-memory claim. Larger edit bottlenecks remain open.

## Plain snapshot tag ownership reduced (0472)

[0472](changes/0472-xlsx-plain-cell-tags.md) removes ephemeral owned tags for
exact plain cells while preserving full attribute validation and rich fallback.
Whole-process allocation calls fall 14.330%; peak heap does not fall. Residual
eager parsing and snapshot traversal remain leads, but the next concrete
coverage gap is measured DOCX fresh streaming creation through the public
StreamingDocumentWriter; existing buffered creation cases do not exercise it.
See [next work](results/change-0472/next-work.md). Short-guard regressions and
PPT variability remain explicit.

## Pre-compaction lifetime hypothesis rejected (0471)

[0471](changes/0471-xlsx-rewrite-buffer-lifetime.md) tests the old rewrite
vector's overlap with verification. Explicit release changes neither rounded
whole-process peak heap nor allocation count, and normal RSS is not reduced in
either matched pair. The production experiment is reverted. Local lifetime
shortening alone is not a sufficient performance result.

The snapshot scan remains the next measured lead: ordinary cells allocate
owned names, values and attribute slices, and cell addresses are checked in a
separate attribute traversal. Full eager/snapshot fusion also adds no-op work
and temporary overlap unless carefully deferred; MCE-processed bytes cannot
supply original-byte spans. Compare local attribute-pass reuse with a properly
source-bound tag representation before adding retained caches or larger Store
handoffs. The [0471 analysis](results/change-0471/next-work.md) records both
options and their limits.

## Current XLSX pass-reuse result and remaining leads (0470)

[0470](changes/0470-xlsx-empty-web-proof.md) removes the later web-binding
traversal for ordinary worksheets through a bounded proof during compaction.
Unproven inputs keep the original reader and error phase. Whole-process
allocation calls fall 19.753%; the six-row diagnostic latency probe improves,
with no qualified latency or peak-memory claim. Repeated RSS is variable,
and targeted PPT/CFB guard penalties remain explicit.

The larger eager-parser and lossless snapshot traversals remain next leads.
Their fusion must account for transformed MCE input, original-byte span identity,
error priority and simultaneous temporary storage. A separate lifetime review
also identifies the now-unused pre-compaction vector retained around grid
parsing in both versions: measure releasing it promptly before pursuing another
cache or increasing the existing 4,096-cell/1 MiB Store handoff. Keep the
persistent payload-heavy PPT guard flag in subsequent matched comparisons.

## Current XLSX result and next lead (0469)

[0469](changes/0469-xlsx-borrowed-compaction-events.md) retains borrowed events
in the compactor after byte/error differential checks and a measured 10.985%
reduction in whole-process allocation calls. The pass itself, checked attribute
normalization, publication validation and bounded Store handoff remain intact.
Latency improvement is modest and unqualified; peak-memory improvement is not
established. Full-guard flags and the supplemental CFB identity rejection are
retained. The larger remaining opportunities are eager parsing, snapshot scans
and carefully validated pass reuse from the 0468 profile. Avoid spending further
batches qualifying this small latency effect at the expense of those paths.


## Measured compaction lead before 0469

[0468](changes/0468-xlsx-remaining-commit-profile.md) refreshes CPU attribution
after the 0467 optimization. Eager parsing remains the largest sampled commit
context (40.01%), followed by snapshot scanning (24.41%), compaction (13.27%)
and web metadata validation (10.33%); these inclusive contexts overlap.
The web reader traverses every changed worksheet even for ordinary cell edits,
but skipping it would remove existing malformed-input and web-extension checks.

The next bounded experiment is removing `into_owned()` from the compaction
event loop: each borrowed event is consumed before the next read from the
stable input. Retain the complete pass, attribute normalization, `xml:space`,
entity/namespace handling and error ordering. This targets temporary copies
without increasing Store retention. Entire-compaction elimination would bound
sampled commit-context speedup near 1.153x; event borrowing removes only a
fraction, and no end-to-end benefit is yet claimed. Full parse/validation-pass
reuse remains the larger architectural lead; redundant pre-sort checks are
not justified by this profile.

## Current dense XLSX result and next lead (0467)

[0467](changes/0467-xlsx-cell-attributes.md) removes the five repeated checked
attribute scans identified by 0466. The fixed-path 500-sample ABBA measures
8.68% / 7.86% lower dense one-percent commit/save median latency; mean and
tails also qualify. Whole-process Heaptrack allocation calls fall 20.47%,
without an operation-local allocation or reduced peak-memory claim. All 1,242
XLSX tests and applicable scoped correctness gates pass.

The next profile should measure the candidate before ranking remaining parser
and writer work. The ordinary path still performs complete worksheet Store
parses, XML rewriting/compaction and publication validation. The bounded Store
handoff remains at 4,096 cells / 1 MiB; the earlier unrestricted approach failed
its memory gate. Retained full-matrix latency flags and build-path sensitivity
limit broader claims. No new remote/cold, native or scaling coverage follows.

## Current dense XLSX lead (0466)

[0466](changes/0466-xlsx-dense-commit-profile.md) profiles the ordinary owned
one-percent commit/save path on 131,072 cells and 1,311 updates. The source
path performs two original and two rewritten worksheet Store parses, full XML
rewrite/compaction and publication audit. Repeated checked cell-attribute scans
are a concrete allocation lead. The retained CPU profiles distinguish normal
release callchain limitations from a separate frame-pointer build; whole-process
ancestors are not retained-sample-only phase timers. Use this evidence before
changing parsing or retention. The old unrestricted validated-Store handoff
failed its memory gate and is not reinstated. Broader CRUD, remote/cold-input
and scaling gaps remain open.

The 0466 frame-pointer profile attributes 55.13% of whole-process sampled
weight to exact commit ancestors. Within commit, worksheet Parser ancestors
account for 47.66% and the shared attribute lookup for 15.38% (inclusive and
overlapping). 0467 completes the narrow experiment of eliminating repeated cell attribute
scans while preserving validation and refusal behavior.

## Change 0465: checked-default ODP append coverage

0465 measures the existing materialized ODP append lifecycle as the 37th
checked default case. The four lanes pass over 6,030 normal and 90 allocator
samples; ODP normal p50 is 1.691887/1.687518 ms for tiny, 67.786424/67.745553
ms for medium and 136.334131/136.843521 ms for large in R1/R2. The allocator
observations are descriptive per-iteration counts, and region peaks include
absolute process-live baseline bytes. No hotspot attribution or optimization
comparison follows. The sealed precleanup verifier, five resealed negative
probes, fresh-copy portable verification and owned cleanup pass, with no
regression, speedup, native-producer, bounded-memory streaming or scaling
claim.

## Change 0464: descriptive PPTX pair lifecycle baseline

0464 measures a harness-only generic PPTX pair lifecycle. The pair is a
same-source-derived positive control, so it exercises generalized source and
destination handling without providing independent producer or native Office
acceptance evidence. No production PPTX optimization is present.

The eight-report, 240-sample matrix uses R1 forward and R2 reverse order, bytes
and a bounded logical-range provider, normal and operation-scoped allocator
binaries, three warmups and 30 samples per lane on CPU 2 with one worker.
Normal bytes p50 is 1.9536 / 1.9380 ms in R1/R2; normal range is 1.9872 /
1.9865 ms; allocator bytes is 2.0768 / 2.0705 ms; allocator range is 2.1182 /
2.1213 ms. The range adapter returns at most 256 bytes per logical read with
zero fixed delay.

The reported API sum covers open-source, open-destination, plan and publication
calls. It is not a contiguous end-to-end timer; setup, input loading, sink and
adapter construction, oracle/artifact work and teardown are outside the sum.
Allocator observations are available only in the operation-scoped allocator
binary. No compression or copied-byte count follows. All eight outputs are
55,891 bytes with three slides and pass the independent package oracle within
its declared scope.

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

The source review and final harness Clippy receipt are clean at custody epoch
`a35f4507a74e7678f29f91dd6857d07a87579ba872785a1eaf19e296540debbe`. Harness
validation records 391 passing tests with one ignored; all 240 formal samples
and focused smoke (four positive and seven negative cases) pass.

LibreOffice 26.2.5.2 saves the formal three-slide output. Native-r2 passes the
scoped application-save plus source-backed/eager validation of all three slide
counts, slide sizes and ordered text; its saved output is 39,830 bytes with
SHA-256 `d2ca3109448f5a0797f00561beff4eafa84616f7dbf897ff2456667fbf555915`.
All three post-save source-backed image inventories are unavailable because
`UnsafeEdit` / `source-backed picture inventory` refuses markup compatibility.
The failed native-roundtrip-r1 receipt and raw saved output remain preserved.
This is not full Office acceptance; image equivalence and rendering
compatibility remain unproven. The supplemental inventory binary is diagnostic
only: its build and
warning-denied Clippy receipt pass; source compatibility records 7,031
unchanged files, the measured copy Rust/binary/captures remain immutable, and
no retiming is introduced. The documentation receipt and owned performance-crate scoped format
check pass. A broad `cargo fmt --all --check` receipt remains failed only at
pre-existing formatting in the out-of-scope iWork
`crates/litchi-keynote/src/document.rs`; no full-workspace format pass is
claimed. Boundary checks pass. The corrected `profiling-r1` receipt supersedes
the initial UID-0-unavailable diagnostic. It passes a whole-process user-counter
and DWARF profile over three warmups and 30 retained samples (33 checked
iterations); separate counter and sampling runs add 60 diagnostic samples
outside the 240 formal samples. Totals are instructions 1,758,215,198, cycles
631,274,012, branches 327,372,796, branch misses 3,514,541, cache misses
2,295,690, and page faults 1,709. PMU counters ran at 83% because of
multiplexing, and the raw zero cache-reference value is uninterpreted. The
20-sample cycle symbol view places `sha2::sha256::x86_sha::compress` at
19.96% and `zlib_rs::inflate::inflate_fast_help_avx2` at 14.25%; the
whole-process scope includes binary hashing, setup and oracle work, so this is
diagnostic only and does not establish operation hotspots or stable ranking.
Sealed precleanup, fresh-copy portable verification and resealed-summary
tamper rejection pass. Owned cleanup removed 12 files totaling 252,885,687
bytes plus two archived audit temporaries. Coverage remains 439/36, the
full goal remains open, and iWork is
outside scope. See the [summary](results/change-0464/summary.json),
[native-r1 receipt](results/change-0464/native-roundtrip/receipt.json),
[native-r2 receipt](results/change-0464/native-roundtrip-r2/receipt.json),
[profile summary](results/change-0464/profiling-r1/profile-summary.json),
[top symbols](results/change-0464/profiling-r1/samples/top-symbols.txt), and
[source compatibility](results/change-0464/source-compatibility.json).

## Change 0463: writer-origin proof removes repeated commit audit work

0463 reuses a private proof minted by the ODP serializer only after the
existing common `PackageWriter` audit path has accepted each authored XML
payload and the exact candidate bytes have been reopened. `PackageWriter`
itself is unchanged, and exact source members retain their source-origin
treatment. The eligible commit path skips only the final
`validate_compact_xml_parts` pass; a missing or stale proof, identity mismatch,
conservative-bound failure, or later package replacement uses the existing
validator.

Normal p50 deltas across 24 reports and 720 samples are -10.0491% / -8.9678%
for tiny, -6.9642% / -7.3050% for medium, and -7.3657% / -6.8199% for large
in R1/R2. All normal and allocator bootstrap upper bounds are negative, all
four medium/large rows clear the 3% gate, and no adverse >5% elapsed or RSS
flag is present. Allocator p50 elapsed deltas are -9.2947% / -8.8357% for
tiny, -3.6356% / -3.6243% for medium, and -4.0327% / -2.8299% for large.

Allocated bytes fall by 2,087,682 / 11,072,106 / 20,202,090 for tiny/medium/
large in both repeats. Every lane reduces allocation calls by 1,039,
reallocations by 93 and deallocations by 946; peak above entry changes are
0 / -49,674 / -15,438 bytes and retained-live deltas remain zero. This is
operation-scoped allocator evidence, not a claim about unmeasured I/O or
compression work.

The separate public-API phase clocks show commit -12.6863% / -13.0797%,
transaction -3.8135% / -2.2023%, snapshot opening +0.9812% / +0.4743%, add
+3.7947% / +1.2462%, and publication -0.0322% / +0.0312%. Only commit is
changed by the proof; the other phase movement is diagnostic. Whole-process
counters include setup, warmups and checks and show instructions -5.6821%,
cycles -6.4241%, branch misses -5.2313% and branches -5.3523%.

The source review confirms exact owner binding, conservative default-limit
coverage, proof invalidation after package replacement, and unchanged semantic
readback/error order. Candidate validation records 381 ODP tests and
warning-denied Clippy passing. The harness and final gates are complete: 387
harness tests pass with one ignored, for 768 passed ODP/harness tests in total.
Warning-denied rustdoc, scoped formatting, boundaries, precleanup and source
replay, fresh-copy portable replay, resealed +1ns tamper rejection, and owned
cleanup of four executables totaling 233,058,712 bytes pass. 0460 remains
accepted, coverage remains 439/36, the full goal remains open, and iWork is
outside scope.

## Change 0462: shape index bypasses the hot scan but misses the gate

0462 adds a private seventeen-key first-occurrence index only to `ShapeAttrs`.
Valid indexed typed hits bypass the cached linear loop; the defensive fallback
still handles a cached-key predicate that does not validate. The fixed index
adds 280 bytes per element (`ElementAttrs` 144 bytes, `ShapeAttrs` 424 bytes),
and `shape_builder` grows from a 1,400-byte to a 1,688-byte stack frame. The
generic `ElementAttrs::get` frame remains 328 bytes in both builds.

Normal p50 deltas across 24 reports and 720 samples are -2.3623% / -3.0049%
for tiny, -3.4649% / -3.1917% for medium, and -2.6602% / -2.6279% for large
in R1/R2. Both large rows miss the predeclared 3% medium/large gate, so the
candidate is rejected despite all normal confidence intervals remaining below
zero. Allocation bytes, calls, reallocations, deallocations, regional peak and
retained-live values are exactly unchanged; the R1 tiny allocator interval
crosses zero. No retained speed or memory benefit follows.

Supplementary public-API phase p50 deltas are transaction -1.2255% / -3.8972%,
snapshot opening -3.9477% / -5.6825%, commit -3.6631% / -4.1715%, add +0.1066%
/ +0.0797%, and publication -0.1857% / -0.2066% in R1/R2. These phase clocks
exclude setup, warmups and checks; whole-process counter diagnostics include
them and report instructions -3.3494%, cycles -3.0144%, branch misses +1.2247%
and cache misses +1.2504%. Neither scope establishes causal API attribution.

Manual assembly review confirms the indexed hit path and the retained fallback.
The automatic `known_cached_scan.eliminated` field is a flawed direct-call
heuristic and is not evidence that work disappeared. Candidate validation
records 379 ODP tests and 387 harness tests passing (one ignored), with
warning-denied Clippy. Both Rust files are restored byte-exact to `dbd2f8ece`;
final Clippy/docs/format/boundaries, portable replay, tamper rejection and owned
temporary cleanup pass. 0460 remains
accepted, coverage remains 439/36, the full goal remains open, and iWork is
outside scope.

## Change 0461: attribute-match split helps slightly but misses the gate

0461 moves ODP attribute matching into a pure cached-key path and decodes values
only after a namespace/local-name hit. The 24-report, 720-sample matrix records
normal p50 deltas of -2.0578% / -2.2940% for tiny, -2.0604% / -3.8754% for
medium, and -2.1547% / -3.1281% for large in R1/R2. R1 medium and large miss
the predeclared 3% practical gate, so the candidate is rejected despite
negative bootstrap upper bounds. There are no adverse >5% elapsed or RSS flags.

Allocator metrics are exactly unchanged across all six lanes. This preserves
the practical-gate refusal and supplies no retained speedup claim. Supplementary
phase clocks show transaction -3.0962% / -0.4057%, snapshot opening -3.9421% /
-4.2911%, commit -4.1677% / -3.0561%, add -0.1727% / +1.2645%, and publication
-0.4772% / -0.4562% in R1/R2. These are diagnostic public-phase observations,
not causal API attribution. Whole-process counters report instructions -3.2567%,
cycles -3.2314%, branch misses -3.3587% and cache misses +0.2581%; setup,
warmups and checks are included.

The authenticated assembly explains the partial result: `ElementAttrs::lookup`
has no candidate out-of-line body, matching checks are inlined into `get`, and
the `get` stack frame contracts from `0x148` to `0x128`. The source review
preserves lazy decoding, iterator/error order, namespace shadowing and drawing
attribute harvest semantics. Candidate validation reports 372 ODP tests,
all-target warning-denied Clippy and scoped formatting passing. The candidate also passes 387 harness tests (one ignored). Source restoration
to `05f432d48`, final Clippy/docs/formatting/boundaries, portable verification,
tamper rejection and owned temporary cleanup pass; 0460 stays accepted, coverage remains 439/36, the full goal remains open, and iWork is
outside scope.

## Change 0460: fused staging reduces transaction work

0460 retains the private ODP fused staging/source scan after a 24-report,
720-sample A1/B1/B2/A2 lifecycle matrix. Normal p50 candidate-minus-baseline
deltas are -3.5430% / -3.8257% for tiny, -6.0849% / -5.0679% for medium, and
-5.4732% / -5.8268% for large in R1/R2. The four medium/large keep-gate rows
clear the 3% threshold with negative independent bootstrap upper bounds; no
adverse >5% elapsed or process-RSS flag is present.

Allocator p50 deltas are -4.7840% / -4.1309% for tiny, -7.3100% / -6.7785%
for medium, and -7.3197% / -6.9282% for large. Every allocator lane reduces
allocated bytes by 4,642, allocation calls by 16, reallocations by 12 and
deallocations by 4, while regional peak above entry and retained-live deltas
remain unchanged. This is operation-scoped allocator evidence, not a flat
whole-process memory bound.

The supplementary public phase clocks identify transaction as the main changed
phase, with p50 deltas of -24.2681% / -24.0697% in R1/R2. Add is -1.6664% /
+0.8260%, snapshot-open +0.3413% / +0.0119%, commit +1.4977% / +1.3278%, and
publication -0.1138% / -0.3100%. These phase clocks are separate mechanism
evidence and do not turn the lifecycle result into API-attributed causal proof.
Whole-process counters report instructions -6.1959%, cycles -4.9100% and branch
misses +0.5503%; setup, warmups and checks are included.

The source review confirms namespace-aware shared traversal while preserving
validation/error order, BOM-relative spans, limits, source ownership and final
readback. The corrected owner retry passes 371 tests and all-target warning-denied
Clippy passes. No coverage is added (439 selectors / 36 defaults); the full
non-iWork goal remains open and iWork remains outside scope. All builds, 387 harness tests (one ignored), documentation, boundaries,
portable verification and owned temporary cleanup pass.

## Change 0459: source setup and candidate parsing dominate remaining work

Resolved fp/DWARF profiles place candidate slide readback at about 59–60% of
commit samples and serialization at 27–30%. Transaction setup spends about
64% in staging metadata and 32–34% in source-fragment parsing. Family reopening
is at most about 3%, so retaining its large XML owner is not justified.

The local-name-first cached lookup experiment fails its practical latency gate
and is reverted; allocation metrics are unchanged. The next transaction target
is sharing tokenization/namespace maintenance between staging and fragment
scans while preserving every state machine, error precedence and exact byte
span. Initial/candidate slide parsing remains a larger independent target.
See the [source review](results/change-0459/source-review.md) and
[resolved profile](results/change-0459/diagnostic-summary.json). Inclusive
sample shares overlap and do not authorize skipping validation or readback.

## Change 0458: ordinary ODP commit and setup phase costs

Commit is the largest individual ordinary append phase on large sources:
45.64 / 45.55% of normal phase time, and 118,096,476 of 211,442,207 allocated
bytes. Snapshot opening and transaction construction together account for
about 53% of phase time. Append itself is about 1.55%; final sequential output
about 0.03%. The next measurements should resolve internal commit work and
repeated source setup while preserving readback, validation and patch semantics.

The separate sampled recording contains 1,656 samples, all lacking resolved
phase-marker ancestry. It cannot establish internal-stage dominance. Seven
whole-process counters were available, but include setup and checks. The
[source audit](results/change-0458/source-audit.md) supplies hypotheses for
further measurement; it does not authorize skipping validation. See the
[phase evidence and limitations](results/change-0458/README.md).

## Change 0457: source-tail ODP publication work remains specialized

The ordinary `odp_existing_append_lifecycle` control is now a formal
current-revision baseline over 64/4,096/8,192 source slides. Its final 360
retained samples (12 reports across two repeats and normal/allocator lanes),
matched by 360 final candidate samples, are the reference for the specialized
`odp_source_tail_append_lifecycle` candidate.
The candidate scans bounded source XML, validates a finite insertion window,
and replays the changed ZIP member through a sequential sink. It has a
different retained-result contract from ordinary open/append/commit/output,
so the paired rows are numeric endpoint comparisons only.

The ordinary ODP owner has the implemented owned `edit::Snapshot`,
`edit::Transaction`, `edit::Patch` and `edit::Commit` lifecycle. The hotspot
here is the separate source-tail publication path and its integration with
that lifecycle, not an absent ordinary editor.

The allocator lane identifies the working-set change: source-tail regional
peak above entry is 620,381 / 620,385 / 620,385 bytes for 64/4,096/8,192
slides, versus 781,342 / 18,027,568 / 35,958,388 for the ordinary control.
Allocated bytes are 2,607,922 / 36,593,937 / 71,164,177 versus
10,613,383 / 110,226,105 / 211,442,207. The candidate's allocation volume
still grows with document size, and the operation peak excludes source and
fixture ownership already live at entry; these rows do not establish a flat
whole-process memory bound.

Normal p50 deltas are -33.352% / -33.185% for tiny, +3.262% / -2.461% for
medium, and -1.940% / -4.221% for large in R1/R2. Allocator p50 deltas are
-38.056% / -38.004%, -8.911% / -9.868%, and -9.873% / -7.967% in the same
shape order. The final comparison receipt defines candidate minus control and
withholds ordinary Commit/Patch, causal, scaling, physical-I/O and general
CRUD claims. Process peak RSS remains over the five-percent review threshold
for R1 normal tiny (+8.301%), R2 normal tiny (+6.517%), R1 allocator tiny
(+10.184%), and allocator large in both repeats (-5.113% / -5.740%). Process
RSS is separate from the allocator-region evidence.

The source-tail path makes four content passes and adds bounded parser,
replay and fixed Deflate-window work. The initial pre-fix large-normal sampled profile
puts candidate SHA-256 compression at 11.00%, XML `validate_name` at 9.08%,
`memcmp` at 6.01% and `validate_start_element` at 5.42%. The corresponding
control profile puts `memcmp` at 9.67%, Quick-XML attribute iteration at 6.00%,
`memmove` at 5.76% and namespace-prefix resolution at 5.24%. The 1,000
`cycles:u` samples are whole-process observations; they do not attribute CPU
percentages to the timed API, establish causality, or provide hard counter
totals. These whole-process samples remain initial pre-fix diagnostics, not
final candidate attribution; no sampling profile was captured for the final
candidate epoch. Candidate profile validation passes for this initial capture.
The
control's unchanged report
and raw profile pass the existing amended oracle after the original
frame-pointer expectation failed; the correction is bound in
[the amendment](results/change-0457/profiling/control-oracle-amendment.json).
See the [final candidate summary](results/change-0457/candidate-final/summary.json)
and [final comparison](results/change-0457/comparison-final.json). Sealed
precleanup, portable-copy replay and three altered-copy rejection checks pass;
owned staging cleanup is recorded in the [bundle](results/change-0457/README.md).

## Change 0456: eliminate the verified-payload preparation copy

[0456](changes/0456-zip-shared-payload-framing.md) retains shared payload storage
and allocates only framing/central metadata. Media publication drops 16,814,784
allocated bytes and 16,812,496 regional peak bytes above entry. The remaining
publication allocation is 4,243,083 bytes with 593,892 peak bytes above entry;
managed admissions remain conservative. Read/write counts and payload bytes
are unchanged. Ordinary owned Store and generated Deflate remain buffered.

The ordinary bytes/media API regression of 15.411%/16.347% is retained. Separate
fixed-policy diagnostics reverse the comparison, demonstrating allocator-sensitive
latency without proving historical mapping decisions. No universal latency or
API-level CPU gain is claimed. The next bounded-append gap needs a streaming
common XML transform and a format-owned ODP tail append; the current source
presentation and replacement publisher retain whole XML. Six actual distinct
native PPTX pair probes all refuse incompatible shared graphs. Neither native
applications nor broader pair/append coverage is completed by this batch.
See [remaining work](results/change-0456/next-work.md).

## Change 0455: fewer unchanged destination requests

[0455](changes/0455-zip-preservation-transfer-chunks.md) doubles the fixed ZIP
copy buffer to 64 KiB, removing 256 media publication requests and sink writes.
The range API median improves 4.775%/4.449%. Candidate R2 still spends about
51.493 ms opening, 796.157 ms planning and 834.024 ms publishing (1,681.794 ms
API sum). Publication source reads remain zero; destination reads are 577.
Two roughly 16 MiB transfers at the simulated 25 MiB/s already imply about
1.28 seconds of nominal transfer delay. This is not physical network evidence.

Whole-process profiles are dominated by corpus compression and hashing. The
original candidate counter increases remain adverse and cannot be attributed
solely to the timed API. Slow/high-page-fault baseline processes and controlled
glibc mmap-policy runs support substantial allocator variability; no CPU gain
is claimed. Native application, distinct-package, physical cold-I/O, bounded
existing-document append/repackaging and worker-scaling evidence remain open.
See [diagnostic results](results/change-0455/formal/diagnostic-summary.md).

## Change 0454: source-proof publication and measured controls

The immediate 0454 scope was capability refusal rather than measured CPU: the
baseline met the optional producer-visible-name refusal; the name-only
historical intermediate candidate then met the noncanonical relationship XML
lexical refusal; and formatted `ppt/presentation.xml` met the ordinary
authored compactness refusal. The source-owned
`source_xml`/`checked_range`/splice path provides a bounded way through those
boundaries while retaining exact source whitespace and literal namespace
grammar. The final native inventory publishes four self-pair cases and leaves
185 of 189 outcomes unchanged. It is an enabler; replacing refusal with
success does not establish a speedup.

The formal [measurements](results/change-0454/measurements.md) and
[machine-readable rows](results/change-0454/measurements.json) show matched
whole-API p50 movement within about 1.3% and no RSS comparison above 5%.
Fifteen absolute review flags remain visible. Three positive flags are the
range/media-rich R2 `open_source.p99` +5.84%, `open_destination.p99` +9.93%,
and `open.p99` +7.87%; the range/media-rich R1 `open_source.p99` -8.97% does
not repeat. Logical read/work counters are consistent across the matched
lanes. Scheduler or sleep variability is plausible, but this evidence does
not prove that cause or identify a CPU bottleneck.

The next measurement scope is the unmeasured native-application roundtrip and
distinct-package pair, followed by cold-cache/physical-I/O, allocator and
scaling evidence. The external fixture remains candidate-only, so no speedup
against the preserved baseline refusal is available. Core release gates pass
with 497 OPC tests, 854 PPTX tests and 381 aggregate harness tests; the isolated
ASAN fuzzer passes 1,000 runs. Final evidence verification and owned cleanup pass. The full non-iWork goal remains open;
see the [0454 bundle](results/change-0454/README.md).

## Change 0453: duplicate staged decoded media removed from PPTX plans

[0453](changes/0453-pptx-shared-decoded-payload.md) shares managed OPC decoded bytes after successful capture, eliminating
16,777,408 allocated/retained bytes during media planning in both repeats.
Bytes/media API medians improve about 3%; simulated-range latency is essentially
unchanged. Publication absolute live peak falls because its entry storage is
smaller; publication allocation and peak growth are unchanged. Plan allocation
calls fall 3,404→3,388. Exact reread, source work and semantic checks remain.

Late Memory refusals still need a decoded fallback copy, covered by conservative
destination staging. Inline handle storage adds 128 reserved bytes for this
corpus. Remaining captures/decoded buffers, source work, semantic validation,
destination passthrough and fixture construction are not removed. Broader native,
cold-I/O, bounded append, repackaging and scaling evidence remains required.
The primary plain p99 flags and separate confirmation remain in the review.

## Change 0452: compressed capture is retained through PPTX publication

[0452](changes/0452-pptx-retained-capture.md) removes repeated source capture/verification for prepared images/charts
while preserving planner rerun, byte identity, source checks and semantic guards.
Media source work drops 50,366,359→33,589,143 units; publication retains 23 cache
hits/zero cold loads and makes zero source data reads. Complete simulated-range
API time improves about 31%; separate bytes confirmation improves about 8%.

The next ownership opportunity is the staged decoded copy still retained beside
cache and compressed data. Examine shared decoded ownership and tight-budget
admission before claiming memory improvement; current source reservations grow
16.8 MB until plan drop. Native breadth, cold I/O, bounded append, repackaging and
scaling remain open. Whole-process profiles include untimed hashing and corpus
construction and do not quantify the timed publication CPU reduction.

## Change 0451: OPC captures and decodes once on a cold cache load

[0451](changes/0451-opc-combined-capture.md) removes a compressed source pass for first-read transfer authorization.
The elected loader uses the combined ZIP primitive; warm hits and waiters reuse
the decoded allocation and verify a fresh compressed capture. Tokens pin decoded
memory/object reservations after package/data drop. Eight I/O cases preserve
exact output while returning roughly half the source bytes for larger payloads.

Next, integrate with immutable reusable PPTX plans and explicit per-publication
writer reservations. Cloning the existing token would share a reservation sized
for one writer. Complete old/new planning/publication timings, expected-byte
controls, native breadth, cold I/O, bounded append and scaling remain open.

## Change 0450: ZIP primitive for first-read compressed transfer

[0450](changes/0450-zip-combined-capture-decode.md) implements the prerequisite identified in 0449: a checked compressed
capture produces both decoded bytes and a verified token in one read/decode pass.
Large deterministic Store/Deflate cases return roughly half as many source bytes
as cold decoded read followed by expected-byte capture. CRC, exact size, complete
Deflate consumption, cancellation and private token construction remain shared.

The next task is OPC/PPTX adoption under explicit combined reservations, source
lineage/version checks, cache/single-flight ownership and plan/publication lifetime.
Measure complete old/new workflows, including the existing expected-byte path;
no latency regression clearance or end-to-end gain is established here. Native,
cold I/O, bounded append, repackaging and scaling work remains incomplete.

## Change 0449: separate harness hashing and publication source owners

[0449](changes/0449-pptx-caller-source-attribution.md) shows that roughly half the lifecycle SHA period belongs to the untimed
output hash. Planning/publication touched-digest stacks each account for roughly
one quarter of lifecycle SHA period. Keep those contracts; the prior aggregate
SHA percentage does not justify a production SIMD rewrite.

Publication's 33,617,184 returned bytes comprise 16,786,581 from the source and
16,830,603 from the destination. Source cache hits are 23, with no cold loads.
Logical candidate rereads therefore do not imply fresh payload materialization.
Investigate OPC/ZIP authorization during first cold decode under explicit capture
memory/source fences, then bounded destination copy granularity. Current counters
cannot prove exactly removable reads. See [source review](results/change-0449/source-review.md).
Native breadth, cold I/O, bounded append, repackaging and scaling remain open.

## Change 0448: reduce excess waiting in the explicit pacing model

[0448](changes/0448-pptx-minimum-service-pacing.md) credits elapsed source work and fixed-wait overshoot against a combined
fixed-plus-transfer target. Plain API medians fall about 19.3% in both repeats;
all service floors pass. The default separate-sleep model remains available and
unchanged. Neither policy establishes ideal physical bandwidth or a shared link.

Media-rich planning/publication still return 16,794,014/33,617,184 bytes. Their
freshness checks and dependency-closure publication are the next source-work
investigation; these counters alone do not identify redundant reads. SHA-256
accounts for 63.2–63.9% of run-frame self period, but that subset includes untimed
work and excludes blocked sleep. Preserve source identity and validation before
considering reuse. Native/cold I/O, bounded append, repackaging and scaling remain.

## Change 0447: transfer-sensitive PPTX cross-copy evidence

[0447](changes/0447-pptx-range-transfer-pacing.md) adds explicit requested transfer pacing around caller ReadAt. At the frozen
25 MiB/s rate, media-rich planning/publication return 16,794,014/33,617,184 bytes
and request 640.641/1,282.394 ms of transfer sleep. Underlying work matches the
unpaced control; publication accounts for roughly two-thirds of requested
transfer delay. This identifies an I/O-sensitive path to examine while preserving
source freshness, dependency closure and exact publication. It does not prove
that those bytes are redundant or that an optimization is safe.

The plain result exposes per-request sleep granularity: observed latency growth
exceeds the 2.767 ms requested transfer delay. Calibrate a combined deadline model
before interpreting small-payload comparisons as ideal link-rate effects. CPU
profiles omit blocked sleep and include untimed work; SHA-256 dominates the
lifecycle-frame subset. Native/cold I/O and shared-link scaling remain open.

## Change 0446: remove one owned Part-name clone per override

[0446](changes/0446-opc-owned-content-type-name.md) removes exactly `3*N+1` temporary allocations in the measured Part-addition
lifecycle. Medium/large calls fall about 6%, passing the frozen gate; peak memory
is unchanged. The normal latency gate fails despite 2.096–2.846% lower medians.
No paired latency/RSS or repeat flag exceeds 5%.

Candidate content-type parsing remains 29.810% inclusive within the sampled
run-frame subset; catalog opening and publication overlap that cost. Profiles
include untimed work and do not isolate timer-only causality. Further reductions
in repeated attribute/map allocations warrant measurement. Preserve all three
required parses, freshness and managed-memory accounting; a manifest cache or
SIMD rewrite is not justified by this ownership change. Native/CRUD breadth and
cold/range/scaling gaps remain priorities alongside this scoped substrate work.

## Change 0445: observer separated; content-type allocations next

[0445](changes/0445-opc-part-add-plain-source.md) measures the same Part-addition lifecycle with a plain OwnedSource.
Large normal p50 drops from about 62.1 ms observed to 18.9 ms plain, while
allocation calls/requested bytes/above-entry peaks remain identical. The
instrumented reader accounts for 68.989% of observed run-frame self period.
Its removal calibrates the benchmark and is not a production optimization.

Plain inclusive run-frame samples attribute 59.290% to publication, 40.552%
to opening and 29.291% to ContentTypeMap parsing; those rows overlap and include
setup/probes. [Source review](results/change-0445/hotspot-review.md) identifies
an owned Part-name String passed by reference to an Into<String> constructor,
causing a clone during repeated content-type parsing. Measure that ownership
handoff before broad parser/SIMD changes. Keep freshness, generated-candidate
validation and managed-memory accounting; a naive manifest cache is not justified.

## Change 0444: Part-addition observer cost precedes production attribution

[0444](changes/0444-opc-part-add-baseline.md) establishes a low-level OPC one-Part/root-relationship addition
baseline. Large normal p50 is about 61.7–61.9 ms, but the instrumented source
reader accounts for 55.24% of whole-process sampled self time. Code inspection
shows a full ordinary-range scan on each read, so observer cost grows with both
member count and read count. SHA-256 follows at 9.79%. Profiles include untimed
fixture/gate/report work and addr2line limitations. The zero L1 event on this
guest supports no cache-miss claim.

The highest-priority follow-up for this slice is a matched plain-source lifecycle
with identical output and gates, followed by operation-specific attribution.
Do not treat the observed growth as evidence that production topology is
quadratic. No production optimization or Amdahl speedup is accepted in this
batch. The 372 harness/458 OPC tests and 29 independent corruption probes pass;
strict harness lint retains only inherited debt.

## Change 0443: compact preservation-scanner frames

[0443](changes/0443-odp-compact-fragment-frames.md) eliminates temporary namespace/local-name copies
from the source-fragment scanner. The prior profile attributed 7.908% of sampled
periods to this path under transaction, including warmups and incomplete symbols.
Medium/large operation allocation calls now fall 14.886%/14.943%, passing the
frozen gate. Requested bytes fall only 2.557%/2.663%, while peak is unchanged.

Normal latency does not pass its practical gate. Whole-process cycles rise
1.516% and instructions fall 0.739%; no operation-only causal improvement is
claimed. Continue measuring repeated validation and one-shot attribute lookups
before further changes; bounded existing append and wider I/O/scaling remain open.

## Change 0442: consolidate three auxiliary staging XML traversals

[0442](changes/0442-odp-shared-staging-traversal.md) shares borrowed events among settings, declarations
and page metadata. These accounted for 22.097% of the prior profile's weighted
sampled periods under transaction, including warmups and incomplete symbols.
Normal medium/large owned-append p50 now improves 9.830–13.307% in both repeats.
Whole-process cycles/instructions fall 10.092%/8.824%; these include setup and
oracle work and are not operation-only causal attribution. Peak is unchanged.

Source-fragment scanning, ordinary semantic parsing and commit/publication
readback remain separate traversals. Inspect their current measured contribution
and one-shot attribute-cache cost before another change. General tail speedup,
RSS, bounded existing append, cold/range and scaling remain unproven.

## Change 0441: avoid a temporary preservation-model copy

[0441](changes/0441-odp-shared-preservation-projection.md) removes one deep
source-slide copy from ODP staging. Large peak above entry falls by 3,932,160
bytes to 35,958,388 bytes. The original latency/calls/requested-bytes gate fails;
the explicit post-hoc memory review keeps the 9.857% large peak benefit.
Normal medium/large p50 remains slightly slower, and retained bytes are unchanged.

The previous candidate's observed transaction call chains account for 30.56%
of weighted sampled periods, with only 0.266% visibly under clone frames.
These include warmups and incomplete symbolization, not operation-only causal
fractions. Repeated settings, declaration, page-metadata and source-fragment
XML traversals remain the larger CPU investigation. Preserve error order,
namespace and publication contracts before consolidating any traversal.

## Change 0440: borrowed namespace cache reduces allocation work

[0440](changes/0440-odp-borrowed-attribute-namespaces.md) removes per-attribute
namespace URI vectors while preserving reader-scoped resolution. The large
owned append interval drops from 1,462,779 to 1,167,845 allocation calls and
236,704,188 to 221,171,020 requested bytes. Peak above entry remains 39,890,548
bytes. No normal latency improvement is established; the main tail flags and
fixed confirmation remain visible in the [measurements](results/change-0440/measurements.md).

Next isolate one-shot attribute-cache costs and repeated transaction staging
or commit validation. The historical stack attribution includes setup and
oracle work and is not an operation-only causal share. Keep these hypotheses
separate from the measured namespace-ownership result. Bounded existing
append, Part addition, repackaging, native breadth, cold/range and scaling
remain open.

## Change 0439: owned append exposes XML and allocation costs

[0439](changes/0439-odp-existing-append-lifecycle.md) measures existing ODP
opening through one appended slide, commit and output. The 8,192-slide case
takes 170.742/175.416 ms normal p50 and requests 236,704,188 bytes through
1,462,779 allocation calls. Its region peak is 39,890,548 bytes above entry.
Whole-process self samples include namespace-event processing (6.91%),
attribute iteration (5.21%), `ElementAttrs::get` (4.71%), memcmp (8.38%), and
memmove (4.19%). Setup and oracle work are included, so these are candidates
for operation-specific attribution, not proven causal shares of append.

Next isolate open versus commit validation and the owned namespace snapshots
in `ElementAttrs`; preserve namespace resolution, malformed-attribute ordering,
and exact publication checks before testing an optimization. The earlier
shared generated-XML double-parse hypothesis remains separate future work.
Part addition/repackaging, native breadth, cold/range input and scaling remain
open. No production optimization is included in this batch.

## Change 0438: fixed-markup Work batching is insufficient

[0438](changes/0438-odp-markup-batching-negative.md) rejects the preceding
fixed-markup batching hypothesis for this ODP workload. Medium/large p50 gains
were only 0.972–1.982%, below the predeclared 5% gate; production was restored.
Fresh whole-executable consume self samples changed 9.03% → 8.55%, with no
operation-only causal claim. Output and operation allocation identities held,
and no 5% regression/repeat flag appeared. Next investigate a different measured
cost in required XML/publication validation or compression before optimizing;
retain all validation and budget contracts. The 0437 bounded-memory API remains
the baseline, with append, native breadth, cold/range and scaling work open.

## Change 0437: ODP memory retention removed; accounting remains visible

[0437](changes/0437-odp-bounded-plain-slides.md) keeps a bounded fresh plain
ODP publication API. Its operation allocator peak is 420,352 bytes across all
three measured sizes, compared with 23,973,505 bytes for large Builder.
Streaming large p50 is 1.574 / 1.588 times candidate Builder; the memory benefit
has a disclosed CPU cost. The same-API tiny R1 p99 flag (+5.874%) also remains.
RSS stays around 84.5–84.7 MB, and no repeat comparison crosses 5%.

Whole-executable streaming self samples include execution consume 11.00%,
Deflate longest-match 10.78%, Deflate medium 6.88%, XML audit 5.39%, and fragment
validation 3.17%; SHA setup hashing is 12.85%. These scopes exclude the later
Python oracle and do not establish an operation-only Amdahl fraction. The
next hypothesis is bounded batching of repeated fixed-markup Work charges,
with exact fallback at limits. No benefit is yet measured. Preserve XML and
publication checks; broader append, native, cold/range, and scaling coverage
remain separate open work.

## Change 0436: ODT Work batching measured

[0436](changes/0436-odt-bounded-text-spans.md) closes the immediate ODT
ordinary-scalar accounting hypothesis. The bounded borrowed-span encoder
reduces normal p50 by 19.382–30.156% across three sizes and two repeats while
retaining exact output/sink identities and identical aligned allocator vectors.
Operation peak above entry stays 420,091 bytes; no matched regression flag
crosses 5%, and one tiny p99 repeat drift (−5.489%) remains visible.

Fresh whole-process consume self share falls 44.27% → 19.24%; SHA hashing
(13.70%), XML audit (8.28%), memset (8.16%) and fragment validation (6.31%)
remain visible. These scopes include setup/oracle work and cannot establish an
operation-only Amdahl fraction. Preserve validation unless equivalent proof
supports further work removal. Next establish ODP fresh creation evidence;
existing append, Part addition, repackaging, native breadth and cold/range/
scaling remain open. The following 0435 entry records the original hypothesis.

## Change 0435: ODT streaming execution accounting

[0435](changes/0435-odt-bounded-plain-paragraphs.md) removes whole-document
retention from fresh plaintext ODT publication, with a measured 420,091-byte
operation allocator peak across three sizes. Large buffered peak is 22,450,985
bytes. The new path is 3.351 / 3.369 times slower than candidate Builder at
32,768 paragraphs. Its whole-process stack record attributes 45.74% self
samples to `ExecutionContext::consume`; that scope includes setup and oracle.
Next test bounded ordinary-text Work batching in ODT, keeping per-scalar
cancellation and exact fallback at limits. Retain the measured streaming API
as the next baseline. This is a hypothesis, not a causal speedup claim.
ODP fresh creation follows; append, Part addition, repackaging, native breadth,
and cold/range/scaling work remain separate. All regression flags are retained.

## Change 0434: ODS ordinary-text work batching

[0434](changes/0434-ods-bounded-text-spans.md) measures a private ODS
serialization change that batches already-safe borrowed UTF-8 spans up to 256
bytes while retaining per-scalar cancellation checks and exact scalar fallback
at row/Work limits. Matched normal p50 is descriptively lower by 12.499–16.211%
for 64 rows, 13.121–13.731% for 8,192 rows, and 13.492–13.562% for 32,768
rows across the two repeats. Allocation vectors are identical before/after;
the regional peak above entry is 419,347 bytes. The four profiles are whole-
process diagnostics, including setup and the untimed oracle; sampled
`ExecutionContext::consume` self share changes from 25.01% to 10.57%.

No universal hotspot, release speedup, RSS, physical-copy, or scaling claim is
authorized. The only repeat flag is baseline tiny p99 at +9.844%; no matched
comparison or RSS point crosses 5%. L1 zero readings do not prove zero misses,
and LLC was not captured. The next concrete hotspot/coverage work is bounded
ODT paragraph creation, then ODP creation; existing append and native breadth
remain separate.

## Change 0433: bounded ODS fresh creation

[0433](changes/0433-ods-bounded-fresh-scalar-creation.md) supplies the missing
fresh ODS scalar-row creation evidence with a sequential writer and a fixed
4,096-byte authoring window. Its retained normal p50 is 59.017–62.137% below
the before-buffered role, while the large operation-region requested peak is
71,050,076 → 419,347 bytes. The after-buffered tiny R1 p99 control flag is
+7.037%. These observations do not identify a universal hotspot or authorize
a production speedup: allocator vectors, RSS, profiles, compression, XML
audit, and setup have distinct scopes. The bundle's [profile index](results/change-0433/profile-index.json)
and [summary](results/change-0433/summary.json) retain the evidence.

The path covers fresh one-sheet scalar creation only. Logical append, package
Part addition, arbitrary repackaging, native/cold I/O, total-RSS attribution,
and broader scaling remain open work.

## Change 0432: streaming heap evidence and next coverage

[0432](changes/0432-xlsx-streaming-operation-memory.md) separates XLSX's
configured row buffer from actual allocator observations. The incremental
requested-live peak remains 420,110 bytes across a 2,048-fold row-count range,
while cumulative requests grow. No allocation-stack cause or general memory
bound is inferred. The [source audit](results/change-0432/next-work.md)
recommends bounded ODS scalar-row creation as the next missing semantic
workstream. Its present builder/append paths retain complete XML/package
candidates; performance ranking requires a measured baseline. The separate
[native PPTX audit](results/change-0432/native-gap.md) records the name-gate and
provenance blocker without fabricating a positive fixture.

## Change 0431: compressed media transfer measured

[0431](changes/0431-verified-compressed-source-transfer.md) addresses the
publication recompression identified by 0430. ZIP verifies an opaque compressed
payload, OPC binds it to source authority, and PPTX retains semantic validation
before adding a canonical destination member. Synthetic media API medians
improve 86.5–89.0% for bytes/warm-file/short-range and 19.4–20.1% for simulated
delayed range. The retained first attempt regressed delayed reads; batching
capture to 64 KiB reduces publication calls from 1,193 to 425. Extra source
input and verification work remain explicit. Native cross-copy, broader size
coverage, allocator/copy attribution and real remote behavior remain open.

## Change 0430: copied-media recompression

[0430](changes/0430-pptx-publication-cpu-attribution.md) recovers publication
callers with frame-pointer capture on the unchanged release binary. Deflate
with the measured CountingSink occurs in 83.22% / 83.38% of iteration samples
for synthetic media-rich bytes/warm-file workloads. OPC topology additions
currently retain decoded shared payloads and regenerate Deflate members;
untouched destination members already copy raw. The next measured candidate is
source-bound compressed media transfer owned by ZIP/OPC, retaining all logical
validation and publication checks. Its staging memory, source reads, limits,
and output framing require implementation and matched measurements. The
[design audit](results/change-0430/transfer-design.md) records those boundaries.

## Change 0429: provider boundary and ZIP refill correctness

[0429](changes/0429-pptx-provider-native-baselines.md) extends the matched PPTX lifecycle with bytes, warm files and
explicit capped/delayed sources, plus separate native image ownership probes.
Runtime preflight exposed and corrected a shared ZIP central-directory refill
bug; four red regressions and six expanded capped-read tests bind that
correctness result. Across 960 samples, non-RSS numeric phase observations are
exactly repeatable. All 27 review flags are RSS points with different baseline
values, so API cost cannot explain them without further setup/allocator
attribution. API clocks exclude setup, source copies, observers, checks and
drops. Cold I/O, representative native cross-copy, full CRUD coverage,
semantic streaming and explicit scaling remain open.

## Change 0428: managed PPTX cache and budget lifetimes

[0428](changes/0428-managed-pptx-cache-lifetimes.md) adds fallible diagnostics on
existing PPTX source owners and a separate managed-resource journal. Sixteen
fresh processes retain 480 samples and 5,460 phase points. Final caller Memory,
Objects, and Depth are zero in every sample; non-RSS numeric phase observations
match across repeats. Exact/one-under image admission, pinning, oversized
bypass, and cumulative publication limits pass. All 27 repeat flags are RSS
points; image-lane RSS already differs at entry, so comparative RSS claims are
withheld. The [resource review](results/change-0428/resource-review.md) retains
those limitations. There is no workload optimization or general leak claim;
native/range, broader CRUD, bounded streaming, and scaling work remain open.

## Change 0427: explicit PPTX allocator drop checkpoints

[0427](changes/0427-pptx-allocator-drop-checkpoints.md) adds a separate allocator
journal for opening, planning, publication and explicit caller drops. Eight
fresh release processes retain 240 samples; every final sink-drop point equals
its entry callback live-byte value, and phase changes match across repeats.
The [resource review](results/change-0427/resource-review.md) distinguishes
owned snapshots, consumed source-backed editors, remaining caller source Arcs
and the sink. This is descriptive current-API evidence with no optimization,
RSS-release, cache-eviction, managed-budget or leak claim. Portable replay and
mutation probes pass; the full cache/near-limit and non-iWork goals remain open.

## Change 0424: validated staged PPTX payload reuse

[0424](changes/0424-staged-pptx-payload-reuse.md) shares independently staged
image/chart bytes after full publication revalidation. The matched media
lifecycle records −14.266% operation allocation requests and −7.317% mean
region peak live bytes, with unchanged logical reads and effectively unchanged
RSS. Both allocator repeats agree; the before/after stack evidence shows the
publication's second 16 MiB clone absent. Plain requests are unchanged and its
median is +0.805% / +0.755%, an explicitly accepted diagnostic tradeoff. All
normal repeat limits pass and no timing/RSS pair crosses the 5% review trigger.
There is no release latency, physical-copy, managed-budget or post-drop claim.
The [bundle](results/change-0424/README.md) retains 16 fresh processes, 1,040
observations and 86 applicable passing Rust tests; existing strict lint debt
and broader native/range/scaling/CRUD work remain open.

## Change 0423: matched source-backed PPTX lifecycles

[0423](changes/0423-matched-source-backed-pptx-lifecycles.md) adds plain and
eight-image source-backed opening/planning/publication selectors using the
owned baseline's exact inputs and bounded sink ceilings. The [bundle](results/change-0423/README.md)
retains 16 fresh processes and 1,040 observations, with full report replay and
88 rejected mutation probes. Source-backed p50 is 2.165 / 2.165 ms for plain
and 272.762 / 261.779 ms for media. Owned media p50/mean repeat drift exceeds
5%, so those statistics remain descriptive; the other three pairs pass their
within-role limits. This is a current API baseline, with no cross-role speedup
or memory-reduction claim. Differing retained objects, logical-read adapter
overhead and setup-inclusive RSS require explicit scope. Plan/publication/drop
snapshots, allocation profiles, near-limit and native/range/scaling coverage
remain open; the full non-iWork goal is incomplete.

## Change 0422: operation-region allocator peak

[0422](changes/0422-operation-region-allocator-peak.md) adds a serialized
operation-region maximum, independent from the earlier process lifetime peak.
The current media-rich/plain PPTX baseline records mean region peaks of 272,736,303
and 1,360,003 bytes; lifetime peaks in those same reports are 812,687,524 and
3,559,492 bytes. These are distinct scopes within V3, not a memory-reduction
comparison. The metric includes entry live bytes and other process callbacks;
it excludes hidden realloc overlap and RSS. Retention/drop boundaries,
near-limit cases and matched source-backed lifecycles remain open.

## Change 0421: allocator high-water correction

[0421](changes/0421-allocator-peak-counter.md) fixes a benchmark counter that
used pre-allocation live bytes when updating the process peak. Historical
`peak_live_bytes_*` numbers and derived differences require corrected captures;
raw reports remain unchanged. Live-byte totals, allocation request counts and
bytes, normal timing, RSS and independent Heaptrack metrics are unaffected by
this specific defect. New allocator reports carry `post_update_peak_v2`,
preventing comparison with markerless historical reports under one policy.

## Change 0420: decoded payload retention

[0420](changes/0420-opc-owned-payload-reuse.md) selects equal existing payload
storage inside OPC after full target validation and decoding. Clean owned PPTX
cross-copy keeps the target archive's independent authority while sharing
payloads with the staged graph. Media-rich live-after falls 50.433 MB and RSS
about 9.1–9.3%; plain live-after falls 87,463 bytes. Requested allocation volume
remains effectively unchanged. The plain median is 1.22–1.69% slower and one
p99 pair is 5.215% slower, explicitly reviewed and accepted for this scoped
memory benefit. No latency speedup is claimed.

Full decompression, raw archive retention, and transient candidate work remain.
The [ownership review](results/change-0420/source-review.md) separates shared
holders from payload copies. Near-limit memory evidence and source-backed
lifecycle comparisons remain priorities in the [goal audit](results/change-0420/goal-audit.md).

## Change 0419: PPTX archive allocation requests

[0419](changes/0419-pptx-bounded-archive-growth.md) traced the large cumulative
allocation volume to exact per-chunk archive reservations. Capped geometric
growth plus a final fallible compact copy reduces media-rich operation request
volume by 98.995% and allocation calls by 3.288%. End-of-region live bytes are
unchanged; RSS is effectively unchanged. Plain median timing is 0.38–2.07%
slower in the 100-sample diagnostic, an accepted scoped tradeoff. No release
latency claim or physical-copy reduction is inferred from request accounting.

The remaining memory issue is decoded payload duplication alongside the
retained generated archive. The [source review](results/change-0419/source-review.md)
keeps storage adoption inside OPC and prohibits arbitrary source/graph
reauthorization. Large near-limit memory behavior and matched source-backed
lifecycles remain priorities in the [goal audit](results/change-0419/goal-audit.md).

## Change 0418: repeated owned PPTX candidate work

[0418](changes/0418-pptx-cross-copy-candidate-reuse.md) retains a freshly proven
reopened candidate for unmodified owned destinations. Required source, graph,
patch, physical and publication checks remain; dirty/custom destinations keep
their fallback. Media lifecycle p50 improves 38.27–38.69%, with about 8% more
whole-process peak RSS, explicitly accepted as a scoped tradeoff.

The paired whole-command profiles have zero lost samples. Deflate-family leaf
weight falls from 71.89% to 57.32%; SHA-family weight changes from 20.69% to
32.20%, with nearly unchanged absolute sampled weight. These include untimed
setup and verification, so they do not assign elapsed lifecycle phase shares.
PMU counters run 83% of the time; a zero cache-reference alias remains
unvalidated and cannot establish a cache rate.

Priorities are retained-artifact memory and near-limit behavior, attribution of
the large cumulative allocator request volume, and a matched source-backed
media lifecycle with the same corpus and timer boundary. The existing plain
source-backed phase case is not a comparable media lifecycle. Additional
changed/mixed OPC and scalar XLSX experiments are listed in
[next-batches.md](results/change-0418/next-batches.md).

## Change 0417: media-rich PPTX investigation

The [0417 representative baseline](changes/0417-representative-crud-baseline.md)
records media-rich PPTX copy p50 at 1,151.843 / 1,152.214 ms. Its timer sums
owned planning, commit and final publication. The [static audit](results/change-0417/pptx-static-audit.md)
traces repeated candidate serialization and image compression through those
phases. Any optimization must preserve equivalent graph, physical-source and
publication checks. The existing source-backed API supports a matched media-rich
experiment, but this matrix's plain source-backed corpus is not comparable.

The broader 30-selector baseline also identifies measurement gaps: 28 selectors
lack operation allocation attribution, five have >5% repeat drift on at least
one quantile, and timer boundaries exclude different portions of their complete
workflows. These gaps constrain optimization claims and remain open.

The separate whole-command profile records 33,347 user-cycle stacks, zero lost
samples and 0.479% unresolved leaf weight. Deflate contributes 71.929% of leaf
weight and SHA-256 20.66%; call chains reach candidate construction, re-planning,
physical fingerprinting and final publication. The observation includes untimed
setup/verification and therefore does not isolate elapsed phase shares. This
supports eliminating equivalent repeated work before changing compression loops.

## Change 0413: exact CFB scratch reservation

[0413](changes/0413-cfb-chain-scratch-reservation.md) removes a redundant
per-sector reservation check after the CFB collector's existing fallible exact
reservation. The scoped warm XLS ABBA comparison covers nine matched
open/list/one-cell workflows. Plain-source p50 improves 1.65–3.18% across
paired comparisons; eager p50 improves 6.95–8.54%. All nine selectors pass the
p50/mean/p95/p99 direction and drift policy. Allocation counts/bytes and
logical source I/O remain unchanged.

The few-large CFB open guard is about 3% slower. An initial p99 +5.02% review
trigger led to a retained follow-up with p99 +3.40%/+3.03%; both original and
follow-up evidence remain visible. No CFB speedup is claimed. All paired peak
RSS changes remain below 0.4%. The private change passes 4,004 legacy-format
tests and 305 CFB default-feature tests/doctests. Six unchanged Rust 1.98
Clippy chunk-iteration findings require a command-scoped lint exemption.

The strict registry now contains eight claims; `claim-0413-cfb-chain-scratch`
is limited to the documented corpus, machine, builds and warm serial lifecycle.
The residual profile still identifies mandatory CFB chain/FAT/ownership work;
remaining dynamic helper callers require attribution before another change.
The complete non-iWork goal remains open.

## Current non-iWork completion audit

The [goal audit](GOAL_AUDIT.md) distinguishes the implemented paths from the
remaining scenario, profiling, and validation requirements. The broad goal
remains open. Hardware counters are available on the current Linux host;
previous unavailable-counter observations must not prevent a fresh profile.

[Change 0404](changes/0404-zip64-preservation-integration.md) connects the
existing ZIP64 preservation writer to public ZIP and OPC publication. This
closes a capability gap rather than establishing a latency improvement.
Generated ZIP64 output promotion and representative large-file measurements
remain separate work.

[Change 0405](changes/0405-opc-cache-lock-observation.md) adds opt-in direct
cache/flight mutex acquisition observations to the contention harness. These
include observer overhead and pre-admission work, and exclude condition-variable
wait/reacquire time. The validated smoke establishes instrumentation coverage;
it does not establish a latency or contention improvement.

[Change 0406](changes/0406-current-hardware-profile.md) captures current
OPC materialization on the available hardware: p50 2.215 ms for four 4 MiB
synthetic Parts. Whole-process samples spend 65.50% in deterministic payload
generation and 26.20% in SHA-256. These surrounding harness costs must be
reduced or separated before using process profiles to select library work.
The normal run lacks operation-local allocator counters and the CPU call
chains contain unresolved frames; no speedup or complete phase attribution
is claimed.

[Change 0408](changes/0408-opc-materialization-evidence.md) removes repeated
expected-payload generation from the sample loop and restores useful caller
chains with a frame-pointer build. Whole-process verification remains dominant
(71.82% inclusive); materialization is 16.34%. Operation-local allocation volume
is 16,861,253 bytes for 16,777,216 decoded bytes, with 55 allocations and one
serial decoder session. This is descriptive allocation evidence, not a proof
about copy volume or a justification for a broad layout/SIMD rewrite. The next
profiling work should extend important CRUD scenarios and validate native cache
PMU events; the L1 aliases returned unvalidated zeroes and LLC aliases were
unsupported. No production speedup is claimed by this batch.

[Change 0409](changes/0409-xlsx-profile-and-range-accounting.md) extends the profile to the XLSX
selected-cell path. Normal query p50 is 3.528 ms, with 81,918 allocation calls and
10,690,444 allocated bytes in the separate counting-allocator run. The exact
`SelectedWorksheet::cell` ancestor contains 2,528 sampled stack blocks;
`clone_bounded_name_part` is 16.22% of their period-weighted leaf attribution.
Most of that helper's cost comes through expanded-name cloning (75.50%) and
namespace expansion (24.26%). `parse_element` is 23.31% inclusive but only
6.08% self in the selected subset. These overlapping sampled CPU percentages
are not wall-clock phase measurements.

Lexical borrowed `ElementData` is a plausible narrow experiment, but lexical
copy helper attribution does not establish it as the main bottleneck. Review
expanded-name/frame ownership before broader changes and retain bounds,
normalization, HRTB callback lifetime, recovery and namespace semantics.
Native L2 request counters are now validated locally; exact LLC events remain
unavailable in this guest and all-zero generic L1 aliases are unusable. The
source edit/save harness also fixes previously unconfigured member ranges;
corrected compressed overlap includes untouched raw publication and must not
be read as semantic decoding. No speedup is claimed by 0409.

## Change 0412 update

[0412](changes/0412-xls-observer-isolation.md) commits the XLS observer
correction and three explicit opt-in plain-source selectors:
`xls_owned_source_open`, `xls_owned_source_open_list_worksheets`, and
`xls_owned_source_open_one_cell`. The XLS observer now sorts its category
ranges and coalesces only exactly adjacent spans; overlap multiplicity remains
unchanged, and the unused generic repeated-read union is disabled. The plain
selectors time the same source-backed lifecycle through `OwnedSource` and
retain operation/allocation metrics with no source summary. A separate
instrumented replay retains the logical locality observations.

The focused XLS observer/owned/allocator tests and registry test pass, as do
the scoped format, boundary, coverage-index, seven-claim, and classification
checks, along with all 91 comparator tests. Plain one-cell p50 is
0.166–0.169 ms with 126 allocation calls / 223,774 allocated bytes in separate
allocator captures. The observer correction preserves all logical I/O and
locality counters; it is not a production speedup.

The plain-source profile has 1,718 observed source-open stack blocks.
`SectorChainScratch::collect_exact` and `try_push<u32>` account for 24.97% and
24.87% of that subset's leaf weight. The scratch collector already reserves
the expected chain length fallibly before walking; test removing its repeated
per-push reservation while retaining cycle, marker, allocation and source
checks. Whole-process copy/setup work still dominates, and the selected-cell
subset has only 21 blocks. These are CPU-attribution observations, not phase
latencies or a completed optimization. The full goal remains open.

## Change 0411 update

[0411](changes/0411-xls-read-allocation-baseline.md) supplies a current matched
XLS open/list/one-cell baseline with operation allocation observations. The
source one-cell path records 488 calls / 280,542 bytes, compared with the
separate eager observation of 7,260 calls / 1,015,781 bytes. Four normal repeats
place source one-cell p50 at 6.016–6.052 ms and eager at 0.510–0.515 ms; these
instrumented families are not a source/eager speedup comparison.

The source CPU profile assigns 87.35% of whole-process leaf weight to
`InstrumentedSource::read_at`, including 96.81% of source-open subset leaf
weight and 99.58% of selected-cell subset leaf weight. This makes diagnostic
range accounting the next measurement problem to isolate. Preserve full
locality replay while testing a cheaper observer or a matched plain source;
then re-rank production XLS/CFB costs. The eager whole-process profile also
contains substantial setup/copy/hash work. PMU counts are whole-process and
multiplexed; exact L1/LLC, cold input, remote input and scaling remain open.

## Change 0410 update

The MCE stream ownership change in
[0410](changes/0410-mce-attribute-name-reuse.md) reuses validated expanded
attribute names for the selected XLSX cell path. On the fixed 9,216-cell,
17-member corpus, the primary ABBA p50 candidate-minus-control deltas are
`-3.9447%` and `-4.1373%`; operation-local allocation calls fall from 81,918
to 77,212 and allocated bytes from 10,690,444 to 10,309,094. The seventh
strict-registry claim entry is present, with the strict checker passing all
seven claims.

The residual profile places `clone_bounded_name_part` at 10.91% of selected
leaf weight and `parse_element` at 5.89% self / 23.42% inclusive. These are
sampled selected-path attribution figures, not paired CPU evidence. The eager
guard remains descriptive: its initial p50 movement is adverse at +0.53% /
+5.66%, while a repeat is lower by 1.106% / 0.505% but both roles drift by
roughly 5–6%; the edit claim is withheld. The eager fixture does not traverse
the changed MCE stream, so this review cannot establish an eager no-regression
result. The broad goal remains open.

## Change 0402 update

The selected-source validation loop in
`SourceBackedPackage::write_part_overlays_to_stream` was the decoder
allocation hotspot. Candidate `51964019db3f6b0787645e3a56c2ecb83bdca65c` now
creates one indexed-read session for unmanaged packages and reuses its Deflate
decoder over the sequential selected-Part validation reads. Stored members
bypass the decoder and cache hits remain cache-only; managed packages retain
the one-shot path to keep decoder workspace within their existing budget
boundary. Control is `46ef44966d5be16f153b1f3375ac14401b7139ac`.

The opt-in `opc_source_overlay_multi_part_noop` selector uses equal-payload
non-empty replacement plans over three fixed shapes at counts 2, 8, and 32.
Normal stable-1.98.1 CPU-2 evidence uses one worker, 20 warmups, and 500
retained in-process samples per A1/B1/B2/A2 leg. The validator summarizes only
`source.opc_source_overlay.publication_ns`; top-level elapsed is checked only
for the preparation + open + planning + publication phase-sum identity. The
global `["warm", "cold-requested"]` setting does not establish cold evidence,
and fresh-child/process-isolated semantics are not claimed.

The accepted publication matrix is deliberately partial:

| Shape/count | Accepted statistics |
| --- | --- |
| `overlay-small / 2` | none |
| `overlay-small / 8` | p50, mean, p95, p99 |
| `overlay-small / 32` | p50, mean, p95, p99 |
| `overlay-large / 2` | none |
| `overlay-large / 8` | none |
| `overlay-large / 32` | p50 only |
| `overlay-media-incompressible / 2` | p50, mean, p95, p99 |
| `overlay-media-incompressible / 8` | none |
| `overlay-media-incompressible / 32` | none |

The exact allocator observation per count is −2/−2/0/0/−80,320/−80,320,
−14/−14/0/0/−562,240/−562,240, and
−62/−62/0/0/−2,489,920/−2,489,920, respectively, in calls/calls/reallocs/
failed/allocated bytes/deallocated bytes order. No overall matrix, top-level
latency, RSS/peak, cold, physical-I/O, throughput, or general OPC hotspot
claim follows. See the [0402 change record](changes/0402-opc-overlay-decoder-reuse.md)
and retained [evidence bundle](results/change-0402/).

## Change 0401 update

The selected-cell scanner now borrows numeric lexical validation and avoids
constructing an owned Number for an unselected, non-formula, non-inline
numeric/untyped cell. The guard does not apply to selected values, formula
cached values, or inline text; unselected inline text still follows the
existing validation path. Control `0859063be5a67bd2aafb3531f2126020b2b5000d`
is compared with production candidate
`87f26d5ee02a1903e668bf7f60fa3ef954a0c3fb`.

The fixed `xlsx_file_selected_cell` oracle is the medium
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` corpus with four
48×48 worksheets and 9,216 numeric cells. It prepares `bEnCh01` for
`Bench01`/position 1 and reads `M29`, expecting Number lexical `1028012`.
On the selected 48×48 sheet, one cell is selected and 2,303 numeric cells are
unselected. Each fixed-oracle lexical allocation is 7 bytes in the allocator
run; this is not a general numeric-size claim.

The CPU-2 normal release ABBA used one worker, warm fresh children, 20
warmups, and 500 samples per leg. Mean/p95/p99 reductions are
`+0.099577940251% / +0.625379111895% / +1.170167332729%` in A1→B1 and
`+0.026562239637% / +0.198122423529% / +0.045344544337%` in A2→B2. These
three statistics are retained; p50 is adverse in both directions
(`-0.012690677428%` / `-0.035254218167%`) and rejected, so no median claim is
made. The 3-warmup/30-sample allocator ABBA independently records −2,303
allocation calls, −2,303 deallocation calls, and −16,121 allocated and
deallocated bytes, with reallocations and failed allocations unchanged.

This is a narrow warm selected-cell result. Eager full-worksheet work,
non-selected/all-cell queries, other cell types, ranges, queries, corpora,
cold/cache, throughput, physical I/O, live/peak/RSS, allocator elapsed time,
and general XLSX behavior remain outside the claim. See the [0401 change
record](changes/0401-xlsx-selected-numeric-elision.md) and [evidence
bundle](results/change-0401/).

## Change 0400 update

The first Change 0400 hypothesis was numeric-value scratch reuse alone. It was
disproved by the pinned corpus: `Bench01` begins with
`<dimension ref="A1:AV48"/>`, which the selected scanner initially classified as
unsupported and therefore sent through the eager worksheet fallback. The
scratch path was never reached; a diagnostic allocator screen measured the
same **100,992 allocation calls** and **13,925,077 allocated bytes** for the
candidate and control. That screen is rejected and excluded from the evidence
bundle.

The production change pivots to cumulative **dimension-bearing selected-cell
streaming plus reusable numeric scratch**. Candidate
`f159c0aed603672aacee8e5923586ce4aa8753f7` is compared with control
`2e47ccebf449ef88943c0abcecd32bd9141eb520`. Valid transitional and strict
SpreadsheetML `dimension` metadata is now validated and discarded without
bounding the query, so the fixed worksheet stays on the selected streaming
path. Unknown attributes or nested content still conservatively fall back to
the eager parser. Untyped/numeric cell values reuse a private scratch buffer;
oversized capacity is not retained.

The existing `xlsx_file_selected_cell` selector uses the fixed medium
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` corpus: four
48×48 worksheets (9,216 numeric cells), 17 ZIP members, 4,226,429 archive
bytes, and source SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`. Each
fresh child prepares mixed-case `bEnCh01` for canonical `Bench01` (position 1)
and reads `M29`; the typed oracle is Number lexical `1028012`, with selected
cell digest
`36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`.
`litchi::Workbook::open(path)` and query preparation are outside the timer;
only case-insensitive sheet selection and the exact cell read are timed.

Normal release-binary CPU-2 A1/B1/B2/A2 ABBA used one worker, 20 warmups, and
500 retained warm samples per leg under Rust/Cargo/Rustdoc 1.98.1. Positive
values mean the candidate is faster:

| Pair | p50 | mean | p95 | p99 |
| ---: | ---: | ---: | ---: | ---: |
| A1→B1 | `+27.775881%` | `+27.728990%` | `+27.657228%` | `+28.150563%` |
| A2→B2 | `+27.711459%` | `+27.691341%` | `+27.705070%` | `+27.835790%` |

Same-implementation drift stayed within the harness ceilings (maximum
`0.72%`). Separate warm allocator ABBA used three warmups and 30 samples per
leg. Exact candidate-minus-control reductions were **16,771 allocation calls
(-16.6063%)**, **14,436 deallocation calls (-14.6347%)**, **26 reallocations
(-68.4211%)**, **3,218,512 allocated bytes (-23.1131%)**, and **2,706,847
deallocated bytes (-20.1822%)**. Allocator elapsed time and global live/peak
snapshots are not claim metrics.

The remaining high-impact opportunities are the eager full-worksheet path,
non-selected/all-cell queries, workbook-open/decompression work outside this
timer, and additional value or metadata forms that remain correctly gated to
fallback. Any extension of the streaming grammar needs the same eager-parser
parity for errors, namespaces, unknown content, limits, and cancellation.
This result is limited to the normal, non-allocator exact selected-cell query
on the one fixed medium corpus. It makes no cold/cache, throughput,
physical-I/O, RSS/peak-memory, broad XLSX/facade, or generalization claim. See
the [0400 change record](changes/0400-xlsx-selected-dimension-streaming.md)
and [evidence bundle](results/change-0400/).

## Change 0399 update

0399 adds a descriptive, opt-in selected-cell baseline through the existing
unified XLSX facade; it does not accept a new hotspot ranking or production
optimization. `xlsx_file_selected_cell` raises the selectable registry from
**419** to **420**, while the default remains **36 cases / 198 rows**. The
fixed corpus is the medium
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` shape with four
48×48 worksheets, 9,216 numeric cells, 17 ZIP members, 4,226,429 bytes, and
source SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.

The fixed query uses non-first canonical `Bench01`/position `1`, prepared as
`bEnCh01`, and reads non-first `M29` (zero-based `28 / 12`) with expected
stored-number lexical value `1028012`. The dedicated selected-cell digest is
`36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`.
Open and query preparation are outside each fresh-child timer; only
case-insensitive sheet selection and exact cell read are timed, with both
selected handles retained through operation snapshots. Independent semantic,
source, and eager typed-oracle checks are untimed. Logical source counters
are `not_applicable_filesystem_xlsx`; no physical-I/O or locality claim
follows. Cold-requested samples are advisory, and prepared-query
`cold-verified` is explicitly ineligible.

Stable 1.98.1 validation passed the harness check, focused registry/scope
tests, all four XLSX filesystem integration tests, the **77/77** strict Python
schema suite, normal warm and cold-requested CLI oracles, explicit
cold-ineligibility, and an allocator warm oracle. `performance_claim: none`;
`claim_authorized: false`; no A/B speedup, latency, allocation, RSS, cache,
throughput, or generalization claim is accepted. See the [0399 change record](changes/0399-unified-xlsx-selected-cell-baseline.md).

## Change 0398 update

0398 adds no measured performance result or accepted hotspot ranking. The
xlsx-gated unified `litchi::sheet::Workbook::sheet` selector returns
`Result<Option<SelectedWorksheet>>` by case-insensitive name or zero-based
position, with `Ok(None)` for a missing/out-of-bounds selector. Its private
eager/source wrapper is lifetime-free and `Clone + Send + Sync`; selected
`cell` and sparse `cells` reads return owned exact XLSX views while preserving
the `Missing`/`Covered`/`Stored(Cell)` and `Empty`/`Formula`/`Unknown` states.

Source selection remains catalog-only and XLSX owns the selected scanner,
fallback, cache, freshness, limits, cancellation, and typed source-change
behavior. Eager bridge parity, typed XLSX `NotWorksheet` for charts/non-grid
sheets, core `Unsupported` for non-XLSX runtimes, and unchanged legacy 1-based
dynamic traits are correctness coverage. The registry remains **419** and the
default remains **36 cases / 198 rows**; focused validation is four XLSX
public tests, five with XLSB, one owned/source bridge, one non-XLSX runtime,
and feature checks. `performance_claim: none`; `claim_authorized: false`.
See the [0398 change record](changes/0398-unified-xlsx-sheet-selectors.md);
no latency, allocation, RSS, physical-I/O, cache, throughput, or
generalization claim follows.

## Change 0397 update

The owning OPC open hotspot was one redundant eager ZIP validation/index pass
before the real `PhysPkgReader`. Production commit `f275d4566` is measured as
candidate `f20d3f417edc3f3da07bf515676b8e71285ad76f` against control
`6e98db9ece29c1e50241cf3e84c9410ce71dd748`. The authorized path is only
normal `OpcPackage::from_vec(owned)`; `authorize_owned_source` still performs
preservation work, so the result does not imply that only one ZIP index exists
overall. Public `OwnedPhysPkgReader` remains eager-validating. Path,
`from_reader`, and session behavior are correctness-covered, not separately
timed, with limits, error ordering, session charge, and exact mixed-storage
byte preservation retained.

The opt-in `opc_casefold_owned_open` selector moves the selectable registry
from **418** to **419**; the default remains **36 cases / 198 rows**. Fixed
stored corpora contain 256, 2,047, 2,048, and 16,384 ordinary 32-byte Parts.
CPU-2 A1/B1/B2/A2 ABBA used one worker, five warmups, and 30 samples under
rustc/Cargo/Rustdoc 1.98.1 because pinned 1.95 lacks Cargo. Normal,
non-allocator release-binary p50 speedups are positive-faster A1→B1 / A2→B2,
in corpus order:
`+8.617829% / +8.204676%`, `+8.298670% / +8.719476%`,
`+8.945417% / +8.268274%`, and `+4.648655% / +4.348226%`; pooled p50 is
`+8.452941% / +8.356702% / +8.490980% / +4.645459%`.

Allocator elapsed time is observational only. Exact allocation/deallocation
call reductions on each ABBA leg are `-1,038 / -8,202 / -8,206 / -65,550`,
and allocated/deallocated-byte reductions are `-152,024 / -1,212,620 /
-1,212,888 / -9,699,800`, in the same corpus order. Per-sample net-live
after-before bytes and reallocations are exactly unchanged; raw global
live-before/after baselines are not cross-run metrics. The accepted claim is
p50 only, with no p99 claim. See the [0397 change record](changes/0397-opc-owned-open-validation-index.md)
and [evidence bundle](results/change-0397/). No RSS, peak operation-memory,
physical-I/O, cold/cache, throughput, format/facade, or generalized
constructor claim follows.

## Change 0396 update

The exact-name lookup hotspot was investigated after expanding the existing
four opt-in case-fold selectors to seven by adding three class-isolated source
selectors for exact, ASCII-case-alias, and genuine-miss lookup. The corpus
coverage adds 2,047 Parts to the prior 256-, 2,048-, and 16,384-Part stored
OPC corpora. The 2,047 corpus is the below-threshold boundary control; exact,
ASCII-case-alias, genuine-miss, and combined 144-query vectors are
independently oracle-checked. The selectable registry rises from **415** to
**418**, while the default remains **36 cases / 198 rows**.

The latency deltas reported here are derived only from normal,
non-allocator release-binary p50 evidence. Source-open measurements time normal unmanaged
`SourceBackedPackage::from_read_at`; lookup measurements time fixed pre-open
unmanaged packages. Values are in **2,048 / 16,384 Parts** order;
allocator-enabled latency is observational only. Validation-constructor
coverage is correctness-only. Mapless exact regressed approximately
`+2,750% / +3,500%`. The preliminary scalar-`Vec` exact probe measured
`+14.85% / +20.96%`; the full inlined linear-probe ABBA measured
`+20.22% / +12.42%` while saving `N` source-open allocator allocation calls.
`std` prehashed exact
`+13.42% / +15.88%`; and direct `HashTable` exact `+14.66% / +13.36%`, with
a high-sample follow-up still around `+14.7%`–`+15.6%`. The final pooled
`Arc<str>` experiment regressed exact `+6.09% / +6.96%`, source-open
`+3.38% / +4.30%`, and mixed lookup `-0.59% / -0.50%`; its allocator result
was three extra allocation calls, approximately `N` extra deallocation calls,
and net-live reductions of 65,536 / 524,288 bytes. This is exact allocator and
net-live footprint evidence, not an RSS, total-memory, or system-footprint
claim. All candidates were rejected for lifecycle or latency regressions, so
no new production hotspot ranking or optimization claim is accepted.

The control is `c0ca6cb5f22ddc68d827b743018855f6b9dc89bd` and the final pooled
candidate is `8f7714ee011b170d938f2532fdd385fb2b61cd32`. See the [0396 change
record](changes/0396-opc-exact-lookup-index-experiments.md) and [evidence
bundle](results/change-0396/). No RSS, total-memory, physical-I/O,
decompression, cold-cache, throughput, scaling, eager/managed/mutable, or
general OPC/OOXML claim follows.

## Change 0395 update

The measured OPC hotspot was the bounded linear `eq_ignore_ascii_case` scan
after an exact `PackURI` hash miss in `SourceBackedPackage::part_index`. The
optimization retains an immutable position order sorted by an
allocation-free ASCII-fold comparator and binary-searches it only for
unmanaged catalogs with at least 2,048 ordinary Parts. Small catalogs and
managed opens retain the linear path; the index is fallibly reserved and
stores no folded names. It preserves source iteration order and all existing
freshness, ownership, resource, and public-API boundaries.

The initial unthresholded probe rejected indexing 256 Parts: normal,
non-allocator `from_read_at` lookup p50 was `+31.31%`/`+33.45%` in the
matched directions, despite improvements of
`-74.50%`/`-74.81%` at 2,048 and `-96.50%`/`-96.62%` at 16,384. The final
thresholded normal, non-allocator `from_read_at` run measured source-lookup
p50 deltas of
`-74.23%`/`-74.37%` and `-96.60%`/`-96.45%` at the two indexed sizes. Open
p50 overhead stayed below 5% (`+4.50%` maximum) in that normal binary, with
exact allocator and retained-vector footprint growth of one call and
`8 * parts` bytes for indexed opens. Allocator-enabled latency is observational
only. Eager lookup timings are not decision-quality because randomized
`HashMap` traversal caused large control/candidate drift; no eager hotspot
claim is made.

The fixed vector, source-counter replay, CPU-2 ABBA protocol, stable 1.98.1
provenance, exact binary/source/patch hashes, 282/282 library tests, focused
managed fallback/cancellation tests, and independent SAFE/pass reviews are
recorded in the [0395 change record](changes/0395-opc-source-casefold-index.md)
and [evidence bundle](results/change-0395/). `performance_claim: scoped`;
only normal, non-allocator unmanaged packages opened through
`SourceBackedPackage::from_read_at`, with source-lookup p50 at 2,048 and
16,384 Parts, is authorized. Validation-constructor coverage is
correctness-only, and allocator-enabled latency is observational only. Means,
tails, source-open latency, eager/managed/mutable/default/general behavior,
RSS, physical I/O, decompression, cold cache, throughput, and scaling remain
withheld.

## Change 0393 update

The selected-picture metadata hotspot was one full descriptor vector per
`SourceSlide::image` or `read_image` call. The selected mode now retains one
descriptor and the final count without weakening the complete-scene grammar,
relationship, target, cancellation, freshness, or error-ordering validation.
Test counters independently prove zero descriptor-vector reservation and all-
target resolution for selected queries.

Matched release evidence accepts the exact eight-call/1,831-byte allocation
reduction for both selected paths and the `image` p50 improvement of
12.404%–12.434% on the fixed eight-picture corpus. `images` allocation is
unchanged. Its favorable timing is not attributed to this mechanism, and
`read_image` timing plus all means and tails remain observations; one
`read_image` control p99 drifted 5.053%. The complete evidence and claim
boundary are in [Change 0393](changes/0393-pptx-selected-image-query.md).

`performance_claim: scoped`; `claim_authorized: true`.

## Change 0390 update

The source-backed OPC full-materialization decoder-construction hotspot now
uses one reusable operation-scoped session per unmanaged materialization;
stored members bypass it and managed reservation-bearing handles still refuse
escape. Operation allocator vectors show 4/160,640, 510/20,481,600, and
6/240,960 fewer calls/bytes for tiny, many-small, and few-large synthetic
corpora, while logical reads, returned bytes, and Part counts stay invariant.
Default owning OPC open remains eager and broad/default OPC behavior is not
changed. Latency, operation-local peak/RSS, copied/decompressed/physical I/O,
and broad performance claims remain withheld; see [Change
0390](changes/0390-opc-materialization-decoder-session.md).
`performance_claim: none`; `claim_authorized: false`.

## Change 0387 update

The source-backed OPC owning-conversion hotspot was confirmed at one
`Arc -> Vec -> Arc` handoff per admitted Part. The unmanaged conversion now
adopts the cache's immutable payload Arc directly. Operation-scoped allocator
evidence is exact across 15 samples/shape: 3-Part tiny removes 6 calls/1,656
bytes, 256-Part many-small removes 512 calls/272,384 bytes, and four-Part
16 MiB few-large removes 8 calls/16,777,376 bytes. Logical Part counts and
source read calls/bytes remain invariant. Pointer-identity and mutation tests
prove independent owning lifetime and copy-on-write separation.

This closes only the duplicate materialization handoff for explicitly selected
unmanaged source-backed conversion. Default owning OPC open still eagerly
loads all admitted Parts, and conversion remains proportional to every Part.
Managed reservations still refuse escape. Broad lazy-default OPC ownership is
therefore still the higher-impact unresolved hotspot. Timing, peak RSS,
copied-byte, physical-I/O, decompression, and real-producer conclusions are
withheld; see [Change 0387](changes/0387-opc-source-materialization-shared-payload.md).
`performance_claim: none`; `claim_authorized: false`.

## Change 0382 update

Change 0382 is a PPTX correctness, CRUD-completeness, and bounded-resource
batch, not a measured hotspot. The source-backed cross-slide copy now accepts
a nonempty caller-bounded set of direct `p:pic` leaves under exactly one
direct `p:spTree`; each selected blip targets one internal,
relationship-free `/ppt/media/` `image/*` leaf. Distinct media are copied
once, destination media URIs are deterministic, and selected image
relationship IDs are allocated and rewritten without XML normalization.

Semantic picture parsing preserves bounded foreign, non-MCE,
non-relationship `a:blip` attributes opaquely and accepts one valid
unqualified `cstate` token. Namespace-safe copy rewrites only the full-slide
resolved relationship-namespace `r:embed`; a full-slide unbound lexical
`r:embed` returns `UnsupportedRelationship`. `r:link`, unknown relationship
attributes, MCE, and duplicate or ambiguous resolved embeds refuse. Wrong-type
non-selected and non-anchor slide bindings fail at open, and planning still
revalidates every binding defense-in-depth. No malformed-object planner test
is claimed.

The destination anchor preserves other valid existing relationships while
anchoring exactly one dialect-correct internal `slideLayout`. A full-slide,
namespace-aware `SourceSlide::images` inventory fences the selected image set;
source catalog relationship reconstruction is fallible and physical ZIP media
deduplication is asserted. Strict XML end-name and unresolved-prefix fences
remain part of the refusal boundary.

Unselected XML and members remain exact, with freshness, signature,
cancellation, partial-sink, and resource fences retained. Layout,
non-selected, and unsupported relationship-ID collisions, broader dependency
graphs, MCE, malformed or ambiguous blips, external/linked/missing/mistyped
media, outbound media, unreferenced image relationships, and unsupported
topology refuse before output. No image decode, conversion, rendering,
durable inverse, or broad media-rich copy is added.

The focused suite passed `41/41`; the default-feature library passed `531`
with the exact pre-existing
`stale_and_unsupported_raw_xml_fail_before_publication` exclusion; the
all-features library passed `533/533` and all integration binaries passed
with the exact three exclusions recorded in the [0382 change
record](changes/0382-pptx-source-backed-cross-slide-image-batch.md). Doctests
passed `6` with `2` ignored. Strict Clippy passed with warnings denied and
the existing `clippy::nonminimal_bool`, `clippy::clone_on_copy`, and
`clippy::needless_lifetimes` allowances; the 64-package/240-declaration
crate-boundary gate passed with 14 existing debt entries.

The validation procedure used one Cargo invocation at a time,
`CARGO_BUILD_JOBS=1`, a 6 GiB per-process virtual-memory cap, a dedicated
target, serial test threads, and a `>=10 GiB` available-memory launch
threshold. These are OOM-mitigating, resource-capped operating constraints,
not evidence that OOM is prevented.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no measured hotspot, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, image-processing, broad media-rich, real-producer, or
system-level OOM-prevention result.

## Change 0381 update

Change 0381 is a DOCX correctness, CRUD-completeness, and bounded-resource
batch, not a measured hotspot. A general nonempty caller-bounded batch within
one glossary performs one topology resolution/materialization and one
inventory pass, resolves canonical source-order semantic selectors, refuses
alias duplicates/overlaps and duplicate paragraph intents, and stages selected
paragraphs under aggregate selector, entry, replacement, and output limits.
Every replacement size is measured before materialization and only one
temporary wrapper is staged at a time. Exact one-Part publication, no-op,
source-bound inverse, cancellation atomicity, and root/sibling/opaque XML
preservation remain in force. Glossary create/delete/rename/reorder and
metadata edits, cross-part or general-story batching, managed editing,
durable patch wire, and broad DOCX remain outside the tranche.

The focused glossary-batch suite passed `18/18`, the existing story-text suite
passed `11/11`, the default-feature library passed `926/926`, the
all-features library passed `935/935` with all integration binaries passing,
and DOCX doctests passed `74` with `31` ignored. Strict Clippy and the
64-package/240-declaration crate-boundary gate passed; the latter reports 14
existing debt entries.

The validation procedure used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, a 6 GiB per-process virtual-memory cap, and a `>=10 GiB`
available-memory launch threshold. These are OOM-mitigating,
resource-capped operating constraints, not evidence that OOM is prevented.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no hotspot rank, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, broad-DOCX, or system-level OOM-prevention result.

## Change 0380 update

Change 0380 is a PPTX correctness, CRUD-completeness, and bounded-resource
batch, not a measured hotspot. Source-backed cross-slide copy now accepts
exactly one direct embedded picture whose target is an internal,
relationship-free `/ppt/media/` image leaf. Planning preflights the declared
image size before reading its payload, reserves the complete staged candidate
with checked arithmetic, verifies the actual size after reading, and binds the
bytes, content type, source/destination identities, topology, and freshness in
the copy plan. Publication deterministically allocates a collision-free media
URI while preserving unrelated members and the existing signature,
cancellation, stale/foreign-source, and failure-atomicity fences.

Relationship-ID collisions are refused rather than remapped. Missing,
mistyped, non-leaf, multiple, unreferenced, shared, or external image
topologies; duplicate shape trees; misplaced or ambiguous blips; MCE; and
broader dependency graphs also fail closed. The operation does not decode,
convert, or render images and adds no durable or inverse patch surface.

The focused suite passed `22/22`; the filtered default library passed `531`
tests; the filtered all-features library passed `533/533` and all integration
binaries passed; and doctests passed `6` with `2` ignored. The exact three
audited unrelated test exclusions and three pre-existing Clippy allowances
are recorded in the [Change 0380
record](changes/0380-pptx-source-backed-cross-slide-image-copy.md). Strict
Clippy, the 64-package/240-declaration boundary gate with 14 existing debt
entries, and three independent source-only reviews passed.

The validation procedure used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, a 6 GiB per-process virtual-memory cap, a dedicated
target, serial test threads, and a `>=10 GiB` available-memory launch
threshold. These are OOM-mitigating, resource-capped controls, not evidence
that OOM is prevented.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no hotspot rank, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, image-processing, broad media-rich, real-producer, or
system-level OOM result.

## Change 0379 update

Change 0379 is a DOCX correctness, CRUD-completeness, and bounded-resource
batch, not a measured hotspot. Source-backed story text now covers one
existing glossary entry selected by unique Unicode-caseless name, canonical
ID, combined name and ID, or checked source-order index. The scanner binds an
exact part/body span and fingerprint, validates relationship/content-type
ownership, inbound closure, dialect, namespace, entry, output, freshness, and
signature boundaries, and publishes without re-resolving semantic identity.
Unselected glossary metadata, sibling entries, opaque XML, unrelated parts,
and package topology remain exact.

The focused glossary suite passed `12/12`, the prior source-backed story suite
passed `11/11`, the default-feature library passed `926/926`, the final
all-features library passed `935/935` with all integration binaries passing,
and DOCX doctests passed `74` with `31` ignored. Strict Clippy, the
crate-boundary gate, and independent source-only reviews passed.

The validation procedure used one Cargo process at a time,
`CARGO_BUILD_JOBS=1`, a 6 GiB per-process virtual-memory cap, a dedicated
target, serial test threads, and a `>=10 GiB` available-memory launch
threshold. These are OOM-mitigating controls, not evidence that OOM is
prevented.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no hotspot rank, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, broad-DOCX, or system-level OOM result. See
[Change 0379](changes/0379-docx-source-backed-glossary-entry-text.md).

## Change 0378 update

Change 0378 is a DOCX correctness, CRUD-completeness, and resource-boundary
batch, not a measured hotspot. Source-backed selectors now cover bounded text
read and exact replacement for individual footnote, endnote, and comment
entries, with entry ownership, package dialect, relationship/content-type,
freshness, preservation, inverse, signature, and failure-atomicity checks.
Glossary entry text remains explicitly deferred.

The focused secondary-story suite passed `25/25`, the existing story-text
suite passed `11/11`, the default-feature library passed `926/926`, the
all-features library passed `935/935` with all integration binaries passing,
and DOCX doctests passed `74` with `31` ignored. The crate-boundary policy
passed. Strict Clippy passed with `-D warnings`.

The validation procedure used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, a 6 GiB per-process virtual-memory cap, and a `>=10 GiB`
available-memory launch threshold. These are OOM-mitigating,
resource-capped operating constraints, not evidence that OOM is prevented.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no hotspot rank, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, or system-level OOM-prevention result.

## Change 0377 update

Change 0377 is a correctness and CRUD-completeness batch, not a measured
hotspot. The XLSX guarded source-backed cell-value editor now has an explicit
numeric-only absent-owner `Insert` operation with existing/new-row ordering,
dimension expansion, calculation invalidation, exact inverse restoration, and
bounded refusal semantics. Existing-owner updates remain separate.

Focused, broad, and doctest validation passed subject only to the four exact
pre-existing row-visibility exclusions recorded in the [0377 change
record](changes/0377-xlsx-source-backed-missing-numeric-insert.md). Production
Clippy passed with one named pre-existing allowance; crate boundaries and
independent production/test reviews passed.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no hotspot rank, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, or system-level OOM result.

## Change 0376 update

Change 0376 is a correctness and CRUD-completeness batch, not a measured
hotspot. Opened XLSB XML Maps transactions can now create the first canonical
Single Cell Tables binding owner and remove that owner after its final binding,
while preserving exact no-ops and refusing noncanonical, opaque, shared,
orphan, foreign, external, malformed, or dangling topology.

Focused and broader serialized XLSB tests passed subject only to the two exact
pre-existing exclusions recorded in the [0376 change
record](changes/0376-xlsb-table-single-cells-lifecycle.md). Strict Clippy,
the crate-boundary gate, and independent topology, safety, and test reviews
also passed.

`performance_claim: none`; `claim_authorized: false`. This change establishes
no hotspot rank, latency, allocation-volume, RSS, I/O, throughput,
fixed-memory, or system-level OOM result.

## Change 0375 update

The source-backed PPTX selected-slide publication path now carries validated
semantic snapshots from planning into publication. It validates current raw
selected bytes, source execution/version/lineage, URI and limits, and the
complete selected-slide closure before applying against the retained snapshot.
Focused SlidePart/Scene counters stay 1/1 through one-slide publication and
2/2 through the multi-slide batch; the foreign-identical-source case
recaptures once, rejects as StaleSource, and writes zero bytes.

The three existing selected-slide selectors were measured in clean serial
ABBA, but paired directions and/or stability gates fail for every package-level
timing interpretation. Numerically passing tail cells remain descriptive only.
The raw reports contain no semantic/refusal booleans,
so output-hash equality is not a correctness or reopen proof and
performance_claim is none. Validation passed 533/533 library tests and
21/21 source-backed edit tests, with exactly these three unrelated
pre-existing stale expectation exclusions:
opened::tests::stale_and_unsupported_raw_xml_fail_before_publication,
pptx_malformed_presentation::malformed_presentation_children_are_reported_by_their_owner,
and pptx_table_styles::noncanonical_style_target_survives_transactional_raw_save.
The exclusions concern stricter direct sldIdLst owner validation and never
enter Change 0375 publication. No resource, I/O, cold-cache, throughput,
fixed-memory, general-OOM, broad-PPTX, real-producer, topology/media, or
parallel claim is made. See [Change 0375](changes/0375-pptx-selected-slide-retained-snapshot.md)
and [the compact ABBA result](results/pptx-selected-slide-retained-snapshot-0375-abba.json).

## Change 0374 update

Change 0374 is a narrowly measured DOCX source-backed story-hyperlink
publication optimization. The planned validated `Snapshot` is retained in
`ForwardOnlyPatch`; publication verifies execution, source version, lineage,
the complete artifact fingerprint, the post-fingerprint source version, and
equality with the snapshot's stored fingerprint before reusing it. The
focused counter proof reports one `capture_source` and one `load_story` from
planning through publication, independently of timing.

Clean CPU-affinity-2 release ABBA used one logical CPU, one worker, 20 warmups,
and 500 samples per case over the fixed seven-story corpus. No-op p50/mean and
redaction p50/mean/p95/p99 are accepted for these exact cases and protocol;
no-op p95/p99 are withheld for control tail drift. No reads, decompression,
materialization, allocation, RSS, physical-I/O, cold-cache, throughput,
fixed-memory, general OOM, all-DOCX, unmeasured-selector, or parallel claim
follows. Serial validation, the `935/935` library result, `23/23` integration
result, and the exact pre-existing exclusion are recorded in the [0374
record](changes/0374-docx-story-hyperlink-retained-snapshot.md).

## Change 0373 update

Change 0373 is a correctness and resource-safety hardening batch, not a
measured hotspot. It removes ordinary `Vec` growth from the declared-size
overrun sentinel in ZIP materialization and ODF Deflate decryption by using
checked `size + 1` bounds and fallible full-capacity reservation. Encrypted
ODF members now fail manifest, Store, password, and plaintext-size preflight
before payload reads.

Source-backed ODP and ODS owners now reject `content.xml` above the shared
256 MiB family limit using metadata-only materialized size before reading the
payload. This includes encrypted manifest plaintext size. ODP final freshness
reconciliation precedes exposure of secondary parse errors. Existing CRC,
size, ZIP, MIME, and publication checks are retained.

Focused and broader release suites passed for all four touched crates;
scoped Clippy passed with six named pre-existing allowances, crate boundaries
passed, and independent reviewers accepted the batch. See the [0373
record](changes/0373-odf-source-allocation-preflight.md).

`performance_claim: none`; no hotspot rank, latency, allocation-volume, RSS,
physical-I/O, cold-cache, throughput, fixed-memory, or system-level OOM claim
follows.

## Change 0372 update

The fresh ODP source-backed catalog open is now a narrowly measured hotspot.
One borrowing XML pass replaces sequential validation and catalog scans while
preserving error precedence, source freshness, ZIP/MIME verification,
publication fences, media locality, and the existing content, depth,
namespace, page, and name bounds. Input-dependent namespace-tracker growth is
fallible, and oversized materialized content is rejected before allocation.

Clean CPU-2 ABBA used one worker, 30 warmups, and 500 samples per leg on the
fixed 12-slide, eight-media corpus with SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.
Control `32290f7ce`/`7291b2...` and candidate
`922eb5e2c`/`9a26cd...` improved fresh-open p50 by 15.627% and 17.327%, with
both p50 same-side drifts below 5%. See the [0372 record](changes/0372-odp-source-catalog-fused-parse.md)
and [ABBA result](results/odp-source-catalog-0372-abba.json).

Only fresh `SourceBackedPresentationCatalog::from_read_at` p50 on this corpus
is accepted. Mean and tail metrics are withheld; list is unstable and query
regressed. No broad open, list, query, RSS, allocation, fixed-memory,
physical-I/O, cold-cache, throughput, or OOM claim follows. Focused tests,
scoped Clippy, crate boundaries, and independent static reviews passed with
the recorded pre-existing exclusions and allowance.

## Change 0371 update

Change 0371 is a correctness and security-hardening prerequisite, not a
measured hotspot. It centralizes the plain-`Reader` ODF namespace tracker and
content validator in `litchi-odf-common`, migrates `litchi-odt` to the shared
substrate, and removes the local duplicate. Checked `u32` depth, a 4,096
content-depth limit, a 256-declarations-per-element limit, and the existing
256 MiB input limit bound hostile XML without entering quick-xml's `u16`
namespace-depth path.

Focused locked/offline release tests passed apart from two named pre-existing
writer-test failures in unmodified code. Scoped Clippy passed with only one
pre-existing `large_enum_variant` lint allowed, and crate boundaries passed.
`performance_claim: none`; no hotspot rank, latency, allocation, RSS,
physical-I/O, fixed-memory, or OOM-prevention claim follows. See the [0371
record](changes/0371-odf-shared-content-validation.md).

## Change 0370 update

Change 0370 establishes an ODP source-backed catalog measurement boundary,
not an accepted hotspot. The three opt-in selectors cover fresh catalog open,
catalog listing, and a selected slide query at index 6 over a fixed 12-slide,
13-member corpus with eight deterministic 2 MiB `Pictures/*` members. The
archive is 16,785,912 bytes with SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.

The open, list, and query timers isolate fresh source-backed catalog
construction, `catalog()` after preparation, and the selected-slide projection
after owner/index preparation respectively. Semantic, source-replay, and
media-locality checks are untimed. The retained [dirty control report](results/odp-source-catalog-0370-control.json)
uses CPU 2, 30 warmups, and 500 samples; revision is
`f35486fb7085bb128eb89a4d2e9edd3ad1065f02` and binary SHA-256 is
`08594839ede39d7f2ed0c143d818e41de0b7cdb77bc92fbcdd2a96083ca9966a`.
Timings are open `57,538/61,057.616/76,884/88,020` ns, list
`31/63.854/161/200` ns, and query `60,062/64,323.154/83,354/101,659` ns for
p50/mean/p95/p99.

The selectable registry is **407** and the default remains **36 cases / 198
rows**. `performance_claim: none`; `claim_authorized: false`. The dirty
control is not clean A/B evidence, so no hotspot ranking, latency, RSS,
allocation, physical-I/O, or OOM-prevention claim follows. Focused selector
and enumeration tests passed `1/1` each; no full suite was run.

## Change 0369 update

The source-backed ODT catalog open path is now a measured, narrowly bounded
hotspot result. One borrowing XML pass replaces sequential `content.xml`
validation and text-block-kind scanning, using a private handler and retaining
the former validation-before-scan error precedence. Source freshness, ZIP,
cancellation, and 256 MiB limits fences remain before catalog publication;
the 1,000,000-block and 4,096-depth ceilings remain. Styles, media, and
semantic payloads stay cold, with no public API or ownership-boundary change.

Clean CPU-2 ABBA evidence used 30 warmups and 500 samples per leg on the
corpus SHA-256
`d63726138d0a50c8ff7e150af4a86385df1a34d886bb5f61f985c78ac79b0220`.
Control `bf1cb55c6`/`a7991b...` and candidate `b712aafbf20e`/`1a75eb...`
had non-overlapping confidence intervals and same-side drifts below 15%.
Open reductions were 53.560%/53.116%/51.008%/49.508% for p50/mean/p95/p99
in A1/B1 and 56.320%/56.078%/54.542%/54.304% in A2/B2. See the [0369
record](changes/0369-odt-source-catalog-fused-parse.md) and [ABBA
result](results/odt-source-catalog-0369-abba.json).

The accepted claim is fresh `SourceBackedDocumentCatalog::from_read_at`
latency on this deterministic large media-rich ODT corpus only. The list
projection is rejected as a tens-of-nanoseconds unstable result; the query
projection is rejected as a 2.9%-3.4% below-materiality result. No all-ODT,
RSS, allocation, fixed-memory, physical-I/O, cold-cache, throughput, or OOM
claim follows. Exact rustfmt, the 557-library/all-integration-target ODT test
run (926 total), scoped `-D warnings` Clippy, and independent code/resource
review all passed or accepted.

## Change 0368 update

Change 0368 establishes an ODT catalog measurement boundary rather than an
accepted hotspot. The three opt-in selectors isolate source-backed catalog
open, catalog projection, and selected-block projection over one fixed
10,008-entry, 13-member corpus with eight 2 MiB `Pictures/*` members. The
control report records the source replay and media-locality gates needed for a
future validation/block-scan fusion experiment, including zero Pictures reads
and zero post-preparation reads for the list projection.

The retained [control report](results/odt-source-catalog-0368-control.json)
comes from dirty revision `14884ced9d8b29b7d2155134025986e9315ac771`, so its
timings are descriptive and do not establish a hotspot ranking, speedup, or
production baseline. The focused catalog and count tests passed `1/1` each;
the initial full harness context was `233 passed, 7 failed, 1 ignored`, with
the count assertion subsequently corrected and six unrelated failures
remaining. The selectable registry is **404** and the default remains **36
cases / 198 rows**. `performance_claim: none`; no latency, physical-I/O,
allocation, RSS, fixed-memory, or OOM-prevention claim follows.

## Change 0367 update

Change 0367 is a correctness and fallback-boundary update, not a measured
hotspot (`performance_claim: none`). The selected worksheet scanner supports
valid direct active `mergeCells` globally through verified EOF, with exact
count, nonempty `ref`, reference-grid, singleton, placement/direct-child, and
overlap validation. The canonical transient `merge::Index` lets a fenced
single-cell non-anchor return `Covered`; anchors remain `Stored`/`Missing`.

Range `cells` and `visit` retain sparse physical records, including merge
followers, and never synthesize covered cells; `stored_extent` is unchanged.
The range cap is 16,384 retained merges with `try_reserve`; 16,385+ drains to
verified EOF before mandatory eager fallback. Unknown merge attributes,
children, or payload use that fallback, while malformed structure is a hard
typed error. Eligible cold paths retain no `Store`, `PartData`, or semantic
caches.

The transient index's internal `BTreeMap` and heap allocations are bounded by
the cap but not individually fallible, so there is no fixed-memory, RSS, or
OOM claim. Focused validation passed `14/14`, full `litchi-xlsx` library
validation passed `906/906`, and scoped Clippy passed with only the unrelated
`clippy::useless-asref` issue allowed. No latency claim follows.

## Change 0366 update

Change 0366 extends the selected worksheet single/range scanner's eligible
scalar payloads with bounded XML general-reference decoding through the
canonical helper. Predefined `amp`, `lt`, `gt`, `quot`, and `apos` references
are eligible in formula, value, and inline payloads. Decimal and hexadecimal
numeric references are eligible only in formula/value payloads when ASCII and
the complete token is at most 12 bytes. Numeric inline references, overlong or
non-ASCII numeric spellings, and numeric scalars outside the XML 1.0 `Char`
production return `NotEligible` and use verified eager fallback; malformed,
custom, and out-of-range references remain MCE/typed errors.

The scanner continues to drain XML/MCE/x14ac and the OPC reader to verified EOF
before publishing or invoking callbacks. Eligible cold `cell`, `cells`, and
`visit_cells` paths retain no `Store`, `PartData`, or semantic caches. No API or
public accepted-input change is made. The pre-existing eager/shared-string
XML-legality residual remains out of scope. Focused `9/9`, full `litchi-xlsx`
library `892/892`, and scoped Clippy with `-D warnings` passed, with only the
unrelated `clippy::useless-asref` issue allowed. No latency, RSS, fixed-memory,
or OOM claim follows; `performance_claim: none`.

## XLSX source-worksheet range streaming boundary (change 0365)

Change 0365 is a correctness/ownership integration, not a measured hotspot
(`performance_claim: none`). Cold `SourceWorksheet::cells(area)` and staged
`visit_cells(area)` use the verified sparse raw range scan only for eligible
worksheets. Dependency scans reach XML/MCE/x14ac EOF, and ZIP CRC/size plus
source/execution fences complete before publication or callbacks. The result
is sparse physical output: missing coordinates are omitted and explicit empty
cells remain. A multi-index SST stream and direct style-count stream avoid
worksheet `Store`, `PartData`, and semantic dependency-cache publication;
warm `Store` remains fast.

`NotEligible` falls back to the eager reader only after verified-reader
completion. Merges, shared/array/data-table formulas, row/column styles, rich,
phonetic, extension, foreign, and general-reference cases remain eager;
`stored_extent` is unchanged. `visit_cells` stages an owned `Vec` that scales
with selected physical output, so this is not fixed-memory, OOM, latency, or
RSS evidence. Focused validation passed `27/27`, full `litchi-xlsx` library
validation passed `883/883`, and package Clippy passed with `-D warnings`
apart from the unrelated `clippy::useless-asref` issue. See [Change 0365](changes/0365-xlsx-source-worksheet-range-streaming.md).

## XLSX selected-cell dependency streaming boundary (change 0364)

Change 0364 extends the selected-cell path with sequential, verified
dependency reads, not a measured hotspot (`performance_claim: none`). The
scan tracks maximum shared-string and direct cell-style references across all
cells plus the target SST index, then streams canonical `sharedStrings`
followed by `styles`. Plain selected SST and direct `c@s` values avoid
`Store`, worksheet `PartData`, a full text `Vec`, a style `Catalog`, and
semantic dependency-cache publication. Warm semantic caches no longer
rematerialize evicted `PartData`; public signatures remain unchanged.

Dependency readers reach XML EOF and CRC, size, source, and cancellation
fences before returning a value or eager fallback. Invalid, missing, or
out-of-range references and unsupported or oversize parts retain established
eager diagnostics after readers close. Rich, phonetic, extension, and foreign
SST entries, row or column styles, merges, shared, array, and data-table
formulas remain eager. The final cell source/cancellation fence also runs on
parser errors.

Focused validation passed `28/28`, library validation passed `856/856`, and
scoped Clippy passed apart from the known unrelated pre-existing `hyperlinks`
`useless_asref` issue. Quick-XML and current-item allocations remain bounded
only by documented limits. No latency, RSS, OOM, or fixed-memory evidence
follows. See [Change 0364](changes/0364-xlsx-selected-cell-dependency-streaming.md).

## XLSX source-worksheet selected-cell routing boundary (change 0363)

Change 0363 is a correctness/ownership integration, not a measured hotspot
(`performance_claim: none`). Cold `SourceWorksheet::cell` uses
`PartView::with_verified_decoded_reader` and the raw selected-worksheet
scanner only for eligible simple scalar worksheets, without publishing full
worksheet `PartData`, `Store`, or cache state; repeated cold queries rescan.
Warm `Store` queries retain the existing fast path. Every `NotEligible` result
falls back to the eager store only after verified-reader return and
CRC/size/source/context checks complete, preserving merge, shared-string,
style, shared-formula, and rich-inline semantics. Source/cancellation/ZIP
errors remain primary, and final outer fences run before the value. Zero,
unrepresentable, and greater-than-2-GiB declared parts retain existing eager
behavior by bypassing the scanner. Public signatures, `cells`, `visit`, and
`stored_extent` are unchanged.

Focused/source/library evidence is `7/7`, `16/16`, and `828/828`; scoped Clippy
passed apart from the known unrelated pre-existing `hyperlinks` `useless_asref`
issue. Single-job capped validation observed no OOM as a protocol fact only.
No latency, RSS, fixed-memory, OOM-safety, or dependency-streaming claim
follows.

## XLSX selected-worksheet raw scan boundary (change 0362)

Change 0362 is a correctness-only raw capability, not a measured hotspot
(`performance_claim: none`). The public
`litchi_xlsx::raw::selected_worksheet::{scan, ScanOutcome, SelectedCell,
NotEligibleReason, StreamResult}` path performs one-pass MCE+x14ac active
selection through XML EOF for an eligible single-cell subset. It distinguishes
`Missing` from explicit `Empty`, validates strict row/cell order and scalar
lexical forms, and lets x14ac `ValidateOnly` parse descent without retaining a
row `BTreeMap`.

Merges, styles, shared strings, shared or array formulas, rich inline values,
and unknown valid structures return typed `NotEligible` only after XML/MCE/raw
EOF. Callers MUST fall back to the eager parser because `NotEligible` is not
worksheet semantic validity. Focused, worksheet-module, and library evidence
is `8/8`, `43/43`, and `821/821`. quick-XML, observer, and conversion
allocations are outside the accounting boundary. No latency, RSS, OOM,
source-worksheet, OPC verified-reader, CRC/size/source-fence, or
full-worksheet-streaming claim follows.

## XLSB source-ingress hard-probe boundary (change 0354)

Change 0354 is correctness/admission closure, not a measured hotspot
(`performance_claim: none`). Private XLSB source ingress keeps non-ZIP,
no-match, and missing-manifest outcomes on the compatibility fallback, while
hard ZIP/OPC/classifier failures return `OpcError` without an eager
`Workbook::from_bytes` retry. Path `FileSource` enforces the caller's exact
input limit before fallback allocation, drops the catalog first, and moves
retained `Bytes` without a clone, preserving pointer/capacity ownership;
known non-XLSB variants return `NotOfficeFile` without pathname reopen.
Evidence is private filter `7/7` within XLSB lib `51/51`, XLSB facade `23/23`,
and successful `xlsx`/`xlsx,xlsb` checks under serial 8 GiB-constrained
execution. The 564 MiB target and approximately 14 GiB post-run available
memory with saturated swap are observations only. No latency, RSS, OOM,
constant-memory, allocation, physical-I/O, or broad XLSB claim follows.
Explicit eager/public smart detection and the positive non-ZIP fallback remain
unchanged; PPTX hard-probe fallback and full selected worksheet materialization
remain outside this slice.

## XLSB source-backed fallback admission boundary (change 0353)

Change 0353 is correctness/admission closure, not a measured hotspot
(`performance_claim: none`). After source-owner admission, dynamic/source-backed
XLSB text no longer retries through an eager full-workbook fallback; the eager
adapter reader/caches/state and private detector-side duplicate source/limits
state are gone. Recognized nonworksheet tabs are skipped via filtered worksheet
positions, while direct nonworksheet and sparkline/pivot/slicer/timeline
selections remain typed refusals. Explicit eager APIs and `DetectedFormat::Xlsb`
remain unchanged, and pre-admission recoverable probes may retain the existing
`Workbook::from_bytes` fallback. Evidence is `23/23` and `40/40` under serial
8 GiB-constrained execution; the 647 MiB target and approximately 15 GiB
post-run available memory with saturated swap are run observations only. No
latency, RSS, OOM, constant-memory, allocation, physical-I/O, or broad XLSB
claim follows. A selected worksheet and required dependencies still
materialize.

## Selected-story text lifecycle boundary (change 0352)

Change 0352 is a DOCX source-backed correctness/CRUD closure, not a measured
hotspot (`performance_claim: none`). A selected `Main`, `Header(index)`, or
`Footer(index)` story supports bounded snapshot/text streaming, direct
paragraph edits, source-bound patch/inverse, and a same-topology one-part
overlay. Exact no-op/trailing-byte copying, canonical relationship/content-
type/external/shared-target checks, namespace/MCE resolution, freshness,
lineage, fingerprints, signatures, cancellation, and failure atomicity remain
explicit; unsupported XML and managed edits are typed refusals. The new test
target is `11/11` and existing `source_backed` is `16/16` under serialized
constrained execution. The 347 MiB target and approximately 15 GiB post-run
available memory with saturated swap are run observations only. No latency,
throughput, RSS, allocation, physical-I/O, benchmark, or broader DOCX claim
is made. Footnotes, endnotes, comments, and glossary remain outside this
slice. Focused coverage includes strict duplicate/end-tag validation, inverse
hostile-writer refusal, decoded namespaces, and actual emitted-byte bounds.

## Indexed-stream validation boundary (change 0351)

Change 0351 is not a measured hotspot (`performance_claim: none`). No artifact
supports the rejected compressor/zlib `~65%` premise; existing `read_to*` is
already fixed-buffer. The strict sink and locator now have structural checks
for encryption/method, single-disk ZIP64 provenance, complete local/central
metadata and descriptor agreement, all physical spans, counts/offsets/
adjacency, short buffers, fallible growth, retryable layout single-flight, and
`ReaderAt` byte stability. Store uses an exact range and Deflate exact
`total_in`; strict CRC-zero is sink-only, with ordinary owned/borrowed fallback
compatibility unchanged. Boundedness is one 16 KiB scratch buffer for one
active member, excluding source/index, sink/output, cache, process memory, and
concurrency. Evidence is `315/315` soapberry-zip and `13/13` litchi-opc under
serialized constrained execution. No latency, throughput, RSS, allocation,
syscall, physical-I/O, decompression, concurrency, selector, or artifact claim
is made.

The final successful package/scenario-scoped commands, not workspace-wide, are
recorded in [Change 0351](changes/0351-indexed-stream-validation.md): `cargo
fmt --package soapberry-zip -- --check`; `cargo test -p soapberry-zip --lib --
--test-threads=1` => `315/315`; and `cargo test -p litchi-opc --test
operation_accounting -- --test-threads=1` => `13/13`, with the record's exact
`ulimit`, target, serialized job/thread, and debug/incremental environment.

## Verified-streaming hardening boundary (change 0350)

Change 0350 is not a measured hotspot (`performance_claim: none`). Shared
overreported-read checks cover `ReaderAt` loops, ZIP verification, streaming,
and the OPC `BorrowedReaderAt` boundary; offsets/counters are checked, bounded
sink `read_to*` / `read_entry_to*` use strict CRC equality, owned reads retain
zero-CRC compatibility, and borrowed nonempty zero-CRC returns `None` for the
owned fallback. Deflate extra output retains `InvalidSize` precedence. The
only bounded statement is one fixed-size scratch buffer for one active member,
excluding source/archive/index, sink/output, cache, and aggregate process
memory. Final evidence is `287/287` soapberry-zip, `4/4` litchi-opc with `261`
filtered, and successful package formatting under one job/thread, disabled
incremental/debug compilation, one disk target, and an 8 GiB limit. No
latency, throughput, RSS, allocation, syscall, decompression, or concurrency
claim follows.

## PhysPkgReader stored-Part borrow boundary (change 0349)

Crate-scoped formatting evidence: `cargo fmt --package soapberry-zip --package litchi-opc -- --check` passed after formatting.

Change 0349 removes one logical destination `Vec` allocation and payload
`memcpy` for an eligible validated Store Part consumed through an immutable
slice, and avoids the materialization budget/cache charge. CRC and local/
central ZIP layout validation remain. Encrypted Store/Deflate members keep
typed errors before owned fallback; nonempty CRC-zero members return `None`.
The evidence is correctness/ownership only (`8/8`, `10/10`, `281/281`) with
serialized jobs/test threads and an 8 GiB ceiling. This is not a measured
hotspot: `performance_claim: none`, with no timing/RSS/throughput,
physical-I/O, decompression, or allocator claim. Source-backed positional
owners and weak mixed stored corpora remain outside scope.

## 2026-08-25: unchanged freshness-session replication closed

- Change 0280 failed its binary identity gate before smoke; do not retry the unchanged change 0279 candidate or reinterpret the prior rejected tails.
- Move the XLS source-backed program to a different bounded design or hotspot while preserving per-read freshness behavior.

## 2026-08-25: XLS freshness probes remain a measured hotspot

- Change 0279 confirmed that successful-read version probes dominate the selected FileSource path, but the production session candidate was rejected because four same-side p95/p99 drift checks exceeded 5%.
- The prior per-read freshness behavior remains in production; do not cite the descriptive 48.44-56.31% direct reduction as an accepted claim.
- A future retry must predeclare its tail-stability policy and satisfy it without retrospective narrowing.

## XLS FileSource freshness attribution (change 0278)

[Change 0278](changes/0278-xls-source-attribution.md) closes the attribution
gap left by change 0277. Exact-equal atomic/FileSource logical work shows
FileSource central gaps of 180-267 microseconds across open/list/one-cell; its
measured version probes consume 46.83%-49.68% of mean elapsed time and closely
explain the gap. The tracked wrapper is separately dominated 74.44%-77.86% by
range-union bookkeeping and is not a production proxy. The next bounded
hotspot is therefore a private operation-scoped freshness session, followed by
clean A1/B1/B2/A2 keep/revert evidence before any span batching. This remains
single-revision diagnostic evidence with `performance_claim: none`; the matrix
stays **398 names** and **36 cases / 198 default records**.

## CFB monotonic cursor result and next XLS hotspot (change 0277)

[Change 0277](changes/0277-cfb-monotonic-cursor-abba.md) removes repeated
in-memory FAT/MiniFAT chain walks and small header allocations from the XLS
source-backed forward scans. Strict A1/B1/B2/A2 evidence accepts the three
source p50/mean cells and open p95, but rejects list/one-cell tails and retains
the 31.64% and 5.31% adverse p99 review triggers. Logical ranges, bytes,
version probes, and locality are exact-neutral, while the tracked source still
runs roughly 11-12x eager. The next hotspot/evidence boundary is an owned vs
atomic-only vs tracked vs FileSource/facade phase breakdown, followed only then
by bounded safe span batching. `performance_claim: none`; the matrix remains
**398 names** and **36 cases / 198 default records**.

## XLS source-global coalescing result and next hotspot (change 0276)

[Change 0276](changes/0276-xls-source-global-coalescing.md) removes the
header-plus-frame duplication from source-backed Workbook globals and reuses
one facade CFB catalog. Global calls fall `136 -> 69`; open/list calls fall
`401 -> 334`, with all 24 source-backed statistics improving across the two
paired 30-sample legs. Equal logical bytes and zero opaque/unselected-sheet
overlap are retained. The instrumented path remains roughly 11-12x eager, so
the next hotspot is a bounded monotonic CFB stream reader or validated span
batcher for repeated chain walks, source-version probes, and selected-sheet
frames. `performance_claim: none`; the matrix remains **398 names** and
**36 cases / 198 default records**.

## Source-backed XLS selective-read boundary (change 0275)

Change 0275 closes the production ownership and matched-measurement boundary
for BIFF8 open/list/one-cell. The 16,995,840-byte opaque-heavy corpus records
138,459 logical source bytes for open/list and 138,593 for one-cell, with zero
opaque or unselected-sheet overlap, but the dirty five-sample release p50 is
roughly 11.7x/12.0x/13.4x eager. The next hotspot is the 401-429 fine-grained
CFB/global reads and associated freshness fences; parser-owned global/SST
materialization is a secondary deferral candidate. The matrix is **398
names**, the default remains **36 cases / 198 records**, and
`performance_claim: none`.

## Rejected DOC owner-public-phases hypothesis (change 0274)

Change 0274 does not establish a hotspot or production optimization. Removing
the public-reader `Vec` clone produced one accepted large lifecycle p50
direction pair, adverse tiny p50, and disagreeing payload-heavy directions;
means and tails were noisy/rejected. The candidate was reverted, so no
latency, allocator, RSS, physical-I/O, or broad DOC hotspot claim follows.
The current matrix remains 393 names with the default 36 cases / 198 records.

## Latest DOCX section-layout closure (change 0273)

[Change 0273](changes/0273-docx-source-backed-section-layout.md) adds one
opt-in typed existing-main-story section-layout selector and takes the current
selectable matrix to **393 names**; the default remains **36 cases / 198
records**. It is correctness/CRUD coverage with `performance_claim: none`.
The dirty five-sample whole-process profile is prioritization evidence only;
clean retained performance and any production hotspot claim remain open.

## Current count and gap (change 0272)

The three opt-in source-overlay multi-part selectors add 27 benchmark records
across changed, equal-payload no-op, and mixed paths. The selectable matrix is
now **393 names** while the default remains **36 cases / 198 records**.
`performance_claim: none`: the dirty five-sample profile is only a
prioritization observation. The open gap is clean retained evidence and
explicit-context/scaling evidence before considering recompression, parallel,
or compression-policy changes.

## Latest evidence boundary and allocator probe (changes 0270-0271)

[Change 0270](changes/0270-opc-relationship-open-timing.md) corrects the OPC
relationship-open timer to production open only, with a `black_box` result
fence and post-timing relationship/package oracles. It supplies no latency or
resource claim.

[Change 0271](changes/0271-xlsx-repeated-store-allocator-probe.md) adds no
hotspot claim. Its exploratory operation-scoped allocator observations are
identical in A1/B1 and A2/B2: medium `568 -> 560` calls and `225206 -> 81224`
bytes; oversized `816 -> 560` calls and `271112552 -> 81224` bytes. Latency,
operation-local peak/RSS, physical-I/O, decompression, copy, and broad XLSX
claims remain withheld; the default remains 36 cases / 198 records.

Status: source-audited; initial ZIP/OPC and CFB substrate measurements captured
Branch: `feat/office-format-completeness`
Evidence through:
[change 0269 — XLSX repeated-store cache ABBA](changes/0269-xlsx-repeated-store-cache-abba.md)
(the direct same-selector release ABBA comparison accepts all eight
p50/mean/p95/p99 cells across the medium and oversized primary selectors,
with zero adverse-both cells. The semantic-query-only interval repeats four
queries eight times in fresh warm children; structural reacquisition controls
remain excluded. This is a latency-only XLSX claim with no resource,
allocation/RSS, physical-I/O, cold-cache, publication/save, producer, or broad
XLSX scope.)

[change 0268 — XLS owned-source publication ABBA](changes/0268-xls-owned-source-publication-abba.md)
(the direct CPU-2 release A1/B1/B2/A2 comparison accepts all eight
p50/mean/p95/p99 cells for the Number and RK/MulRK source-backed publication
selectors, with zero adverse-both cells. This is a latency-only XLS claim
with no resource, allocation/RSS, physical-I/O, cold-cache, producer, or broad
XLS scope.)

[change 0267 — XLSX repeated-store strict schema and harness](changes/0267-xlsx-repeated-store-strict-harness.md)
(the four opt-in selectors bring the current selectable matrix to 389 names
while the default remains 36 cases / 198 records. A pinned medium/oversized
XLSX corpus runs four semantic queries eight times in fresh warm children under
the explicit semantic-query timing boundary. Primary selectors are reserved
for same-selector ABBA comparison; reacquisition controls prove medium-cache
eviction and oversized bypass but are structural-only and excluded from
candidate elapsed comparison. Strict corpus, semantic, cache/read/Budget,
child-process, allocator, and result-channel schemas fail closed. This is
neutral correctness/evidence-boundary coverage only; no latency, allocation,
RSS, physical-I/O, or production claim is made.)

[change 0266 — fail-closed historical REPORT classification](changes/0266-report-claim-classification.md)
(the sidecar/checker binds the two audited historical REPORT tables' headings,
headers, row order, labels, and digests. It classifies 167 rows as 145
historical, 14 descriptive, 8 withheld, and 0 strict claims, with no strict
links. This is report-integrity evidence only and makes no latency or
production claim.)

[change 0265 — PPTX slide-boundary publication selectors](changes/0265-pptx-slide-boundary-publication.md)
(the opt-in remove and move selectors use a deterministic four-slide,
dependency-free PPTX corpus with 45 ZIP members, 32,396 source bytes, and
SHA-256 `685a1805ad291e8f9852d3ccd584320f20847bd0ac8fdf29857f96efe1109477`.
Removal covers first/middle/last positions and a final-only refusal; move
covers both boundaries and the `from == to` no-op. Production
`Snapshot`/`Transaction`, semantic reopen, phase vectors, serialized
forward/inverse patches, source immutability, deterministic rebuilds, strict
untouched raw local/offset-normalized central records, and dependency,
unknown-member, MCE, signed, limits, stale/foreign, partial, and zero-sink
gates are covered. Move requires strict `[Content_Types].xml` identity. This
is correctness and phase evidence only; no latency, allocation/RSS,
physical-I/O, or broad PPTX claim is made.)

[change 0264 — real-producer security correctness corpus](changes/0264-real-producer-security-corpus.md)
(the ignored locked gate covers eight pinned POI/OOXML fixtures: signed
DOCX/XLSX/PPTX, protected DOCX, two encrypted DOC files, a macro XLS, and an
external-link XLSX. Signature/protection/password, inert-VBA, external
inventory, one-under input, zero-output publication, and RAII release checks
pass. This is bounded correctness-only evidence with no selector/default
count, latency, allocation/RSS, physical-I/O, or resource-performance claim.)

[change 0263 — DOCX story-hyperlink publication selectors](changes/0263-docx-story-hyperlink-publication.md)
(the opt-in no-op and shared-target redaction selectors cover seven Word story
kinds, 15 Parts, 24 ZIP members, and a pinned 9,900-byte source. Their timer
covers open, strict planning, commit, and sequential publication while the
independent story XML/`.rels`, ZIP locality, deterministic-output,
source-immutability, and refusal oracles stay outside. This is correctness and
phase evidence only; no latency, allocation/RSS, physical-I/O, or broad DOCX
claim is made.)

[change 0262 — XLSX vendor-extension preservation corpus](changes/0262-xlsx-vendor-extension-preservation.md)
(the explicit CLI-only `vendor-extension` shape extends the deterministic
four-sheet, 48-by-48, eight-media XLSX cell corpus with orphan XML/BIN Parts
and one XML-local relationship. It is excluded from `XlsxCellCrudShape::ALL`,
adds no Case, and at landing left 381 selectable names plus the default 36
cases / 198 records unchanged. Exact no-op/edit/lifecycle, topology/content-type/
relationship, managed-Budget, typed sink-refusal, and raw untouched-member
identity gates pass; no latency, allocation/RSS, decompression,
physical-I/O, or producer claim is made.)

[change 0261 — strict claim canonical recomputation](changes/0261-strict-claim-canonical-recomputation.md)
(the strict verifier now has `_project_report` sequentially validate raw
samples, recompute bounded elapsed statistics and identity projections without
retaining elapsed sample values, and discard the raw report/sample payload
before the next leg. It recomputes elapsed cells and the complete canonical
summary from the four raw ABBA projections, rejects mixed report profiles, and
derives resource A1/B1/B2/A2 values and paired deltas from parsed leg sources.
`time`/`heaptrack` run, status, artifact, and parser identities fail closed;
exact resource variant, revision, binary, harness tool, and profile binding is
required, and raw projection-marker fields are ignored. Public projection
helpers are not exposed; the module-private `_ValidatedProjection` trust carrier is
created by the public verifier path only after raw validation, while plain
mapping inputs and mutations fail before summarization. The four-report package
is bounded at 512 MiB per member, 2 GiB total decompressed input, and 64 MiB
summary size. This is evidence-integrity and verifier-memory hardening only; no
speedup or library-memory claim follows.)

[change 0260 — fresh-child XLSX filesystem roots](changes/0260-xlsx-fresh-child-filesystem-roots.md)
(the existing path-open and open-plus-names/count/full-text selectors now run
each sample in a fresh child over one pinned medium XLSX. Warm and explicit
cold-verified modes retain exact source/semantic hashes and validate the exact
timed object against typed XLSX and OPC/property oracles after operation-only
snapshots. This is input-mode/cache-proof coverage only; no latency,
allocation/RSS, physical-I/O, storage-media, or broad XLSX claim.)

[change 0259 — shared lazy OPC structural members](changes/0259-opc-shared-structural-members.md)
(private OPC catalog parsing now retains the lazy ZIP reader's existing shared
decompression allocation for deflated content-types and relationship
manifests. Stored members remain borrowed and indexed sources remain owned.
Pointer-identity and complete OPC/ZIP gates establish the ownership handoff;
no latency, allocation-count, peak-memory, RSS, I/O, or broad OOXML claim is
made before a relationship-heavy controlled profile.)

[change 0258 — byte-native unified RTF ingress](changes/0258-rtf-byte-native-facade.md)
(the facade now passes owned RTF bytes directly to the native parser and
recognizes literal CP-1252, LZFu, and stored MELA transports without changing
ZIP/OLE2 precedence. Back-to-back CPU-2 release captures reuse identical
binaries and workloads but do not reproduce the same accepted latency cells,
so no speedup statistic is retained. Both compact packages are preserved as
non-reproducibility evidence; the landing is correctness-only.)

[change 0254 — DOCX story-hyperlink planning ABBA evidence](changes/0254-docx-story-hyperlink-index-abba.md)
(the independently audited `fcb3104a5` release package accepts all p50/mean/p95/p99
statistics for eight repeated `Snapshot::plan_target_urls` calls on a prepared
immutable 49-story/1,152-link snapshot. The candidate replaces per-selector
repeated story/relationship scans with preindexed targets; paired reductions
range from 88.225897% to 91.592577% under CPU-0 release A1/B1/B2/A2 with 20
warmups and 500 retained samples per leg. This is plan-only evidence, not an
end-to-end DOCX, I/O, allocation, RSS, resource, or general speedup claim.)

Current implementation-only status:
[change 0257 — source-backed DOCX owned-byte ingress](changes/0257-docx-owned-bytes-source-backed-ingress.md)
removes eager materialization of ordinary DOCX Part payloads from the normal
owned-byte facade while retaining catalog admission, typed limits, ODT/OOXML
arbitration, and truthful historical eager harness controls. It adds no
latency, allocation, RSS, physical-I/O, decompression, or broad DOCX claim;
pinned 20/500 release ABBA evidence remains open.

[change 0252 — XLSX page-break projection ABBA evidence](changes/0252-xlsx-pagebreak-projection-abba.md)
(the fixed media-rich package records accepted cells for historical candidate
`e619debe`. Current production contains later hardening and projection-cache
commits, and that measured candidate is not its ancestor. Focused current-head
correctness tests and audit pass, but the historical latency cells are not
attributable to current code. Fresh edit/save and repeated-cache ABBA evidence
is required; no current or broad XLSX claim.)
[change 0251 — XLSX borrowed XML worksheet parsing ABBA evidence](changes/0251-xlsx-xml-borrowed-abba.md)
(three selectors accept all four statistics; `xlsx_first_cell` accepts p50,
mean, and p99, while p95 is drift-rejected at +10.941931% candidate drift.
No adverse-both cell is present. Latency supports the listed cells, but
the strict 0.1.6 resource ABBA accepts all six required metrics in both
pairings: allocation calls fall 50.98%, allocated bytes fall 23.70%, peak
heap is unchanged, and the largest positive delta is temporary allocations at
1.83748%, below the 5% ceiling. The byte-identical production patch landed as
`ecb6b9429`; the full XLSX suite passes. No broad XLSX claim.)
[change 0250 — ZIP ordering/index ABBA evidence](changes/0250-zip-ordering-abba.md)
(the eight-row `zip_index` package accepts no statistic; adverse-both counts
are p50=5, mean=3, p95=3, and p99=3. The compressible tiny, many-small, and
wide-root rows are adverse at all four statistics. The production
monotonic-offset sort fast path is not landed; the candidate is retained only.
This is short in-memory evidence with no physical-I/O, allocation/RSS, or
cold-cache claim. The independent ZIP-index count oracle landed in
`5eb8c1959`/`197bd3645`; it is a correctness change, not fast-path approval.
No broad ZIP/OPC claim.)
[change 0249 — ODS known-change source publication ABBA](changes/0249-ods-known-change-source-publication-abba.md)
(neither selector has an accepted statistic: `one_edit` is adverse-both at
p95/p99, and `one_percent` disagrees in paired direction at all four. Drift
and identity gates pass; the production candidate is not landed on latency
evidence, and the logical replay counters are not physical-I/O evidence. No
broad ODS claim.)
[change 0248 — CFB streaming release ABBA evidence](changes/0248-cfb-streaming-abba.md)
(the independently audited 24-row package accepts only p50=1, mean=2,
p95=2, and p99=2 cells; adverse-both counts are p50=14, mean=10, p95=7,
p99=4. The selectors exercise `SharedOleFile`, not the direct `OleFile`
payload-read code changed by the candidate, so neither population adjudicates
that optimization. With no applicable measurement, `67a37235c` rolls it back
pending a direct/native-path experiment. Actual-output oracle correctness
landed in `66bb83abb`, and source-identity measurement projection was fixed in
`cd21f7670`. No broad CFB claim.)
[change 0247 — XLSX bytes-facade ABBA evidence (0230 package)](changes/0247-xlsx-bytes-abba-evidence.md)
(the independently audited `0230-xlsx-bytes-20260821` package measures only
in-memory `Workbook::from_bytes(Vec<u8>)` and its named lifecycle). The
candidate source tree equals landed `9aaf9d136`; `xlsx_bytes_open` accepts
exact p95/p99 readings only because p50/mean are drift-rejected, and
`xlsx_bytes_open_lifecycle` accepts all four statistics. No source-file,
filesystem/cold-cache, physical-I/O, preservation, output-byte, allocation,
RSS, or broad XLSX claim is made.)
[change 0229 — DOCX text-path hand-rolled binding tracker](changes/0229-docx-text-binding-tracker.md)
(full-text latency readings are recorded for the semantic and ordinary-root
selectors, but the overall verdict is provisionally withheld under the 0228
pre-floor rule: `docx_file_source_open` and the `xlsx_file_open`
cross-guardrail retain adverse-both uncalibrated readings. Guardrails are
explicitly mixed/failed. The corrected resource report was reprocessed by
`litchi-resource-profile` 0.1.1 with raw print/histogram hashes verified;
temporary allocations are effectively neutral (paired deltas −2/−6), and
RSS is mixed within <2.4% heaptrack and <0.4% time RSS, with no broad DOCX
claim.)
[`change 0228`](changes/0228-docx-floor-calibration.md)
(0228 is a methodology calibration, not a code change — the DOCX-family
analog of 0223/0226, calibrating the likely guardrail phases BEFORE the
first DOCX optimization. Three probe binaries (banked post-0227 tree plus
never-executed parser-shaped padding of +3,872/+5,984/+7,824 B text in
litchi-docx, seeds 2281-2283) measured pure layout noise over 72 legs
(6 phases × 3 probes × A1/B1/B2/A2): the four docx open/lifecycle
guardrails plus the `xlsx_file_open` / `pptx_file_source_open`
cross-guardrails. Five phase/probe pairs failed control-leg drift AND
failed again on their one permitted rerun — four with a cold-requested
page-cache bimodality signature (whole-distribution 2-7x shifts when the
eviction request genuinely misses; pptx cold p50 6.1-6.4 ms vs 41.0 ms),
xlsx with persistent machine noise — and are drift-rejected-excluded,
contributing nothing. Effective floors from the 13 accepted pairs:
`docx_file_eager_open` p99 **3.5%**,
`docx_file_eager_open_full_text_lifecycle` p99 **2.7%**; every other
statistic on the six phases stays uncalibrated (pre-floor rule).
Implication recorded: cold-requested rows are currently uncalibratable
at the 5/5/10/15% ceilings on this machine. The 0229 text-path change is
the first consumer of these floors.)
[`change 0227`](changes/0227-odt-text-binding-tracker.md)
(0227 removes `NsReader`'s per-event `process_event` binding maintenance
(profiled at 9.1%-18.1% of timed on the text-path phases) from the 0217
discard-but-validate text path: the 0224 `BindingTracker` is lifted to a
shared `pub(crate)` module `litchi-odt::binding_tracker` (plus a new
`resolve_attribute`), and `parse_text_block_texts` drives a plain
borrowing `Reader` with hand-maintained push/pop — byte-identical
tokenization, namespace errors, and resolutions by construction, the
0225 memo intact. Attribute validation is generalized over a private
`TextAttributeResolver` trait so the retained/selected paths keep their
`NsReader` untouched. Differential oracles: the lockstep
tracker-vs-`NsReader` per-event replay across the corpus, an
adversarial reserved-prefix/declaration-limit battery with
byte-identical error strings (901 litchi-odt tests). **Banked**:
`odt_semantic_full_text` p50 **9.44%-13.32%**, `odt_repeated_text_cached`
p50/mean/p95 **17.57%-18.52% / 17.66%-18.11% / 13.60%-17.56%**,
`odt_repeated_text_uncached` ALL FOUR (rerun 0227r, superseding the
primary's anomalous b1 leg) **18.79%-20.32% / 18.45%-20.37% /
17.43%-21.68% / 8.65%-18.10%**,
`odt_file_source_open_full_text_lifecycle` p50/mean
**10.37%-13.34% / 9.10%-13.15%**; guardrails clean, within-floor, or
cleared-by-rerun. The ODT open/text paths are now floor-fighting; the
calibration-first DOCX pivot ran as 0228 (above).
Candidate `1d503363…` is the control for the next change.)
[`change 0226`](changes/0226-odt-source-open-floor-calibration.md)
(0226 is a methodology calibration, not a code change — the 0223 analog
completing the `odt_file_source_open` floor set after 0225's verdict was
blocked by an adverse-both p50 reading (max 6.65%) on that byte-identical
guardrail's uncalibrated statistic. Three probe binaries (banked 0224 tree
plus never-executed parser-shaped padding of +3,872/+5,984/+7,744 B text
in litchi-odt, seeds 2261-2263, bracketing 0225's −6,064 B shift in
magnitude) measured pure layout noise under the full A1/B1/B2/A2 protocol
(12 legs): probe a adverse-both p50 3.50%/mean 5.75%, probe b none,
probe c mean 6.09%/p95 45.02%/p99 38.69%. Folding in the 0225
byte-identical-phase history per the 0223 rule (worst observed
adverse-both magnitude, no margin), the `odt_file_source_open` floor is
now p50 6.7% / mean 6.1% / p95 45.0% / p99 38.7% — p95/p99 supersede
0223's 2.5%/28.0% (this session's machine was far noisier in the
500-sample tails); the 0223 lifecycle/eager floors are unchanged. All
0225 source-open adverse readings fall within floor — layout readings —
and 0225 was re-verdicted **banked** — see below.)
[`change 0225`](changes/0225-odt-text-resolution-memo.md)
(0225 adds a last-prefix namespace resolution memo (`TextNamespaceMemo`)
to the 0217 discard-but-validate text path `parse_text_block_texts`
(litchi-odt): `read_event_into` plus a content-versioned memo replaces
per-event `resolve_event` reverse scans over ~37 live bindings, with
provably exact invalidation (`xmlns` memmem prefilter for pushes,
scope-tracking for the deferred pops) — infallible lookups only, error
stream unchanged, panic-free. Differential oracles: 8 synthetic
rebinding fixtures with pinned text, a per-event classification replay,
and corpus-wide parity (900 litchi-odt tests). **Banked** — after the
0226 calibration cleared the residual source-open guardrail reading:
`odt_semantic_full_text` p50/mean **15.79%-17.44% / 19.03%-20.70%**,
`odt_repeated_text_cached` p50/mean/p95
**20.10%-20.24% / 19.85%-20.68% / 19.02%-21.93%**,
`odt_repeated_text_uncached` ALL FOUR
**21.79%-23.51% / 22.56%-23.42% / 23.58%-24.96% / 26.73%-26.96%**,
`odt_file_source_open_full_text_lifecycle` p50 **15.96%-16.50%**
(lifecycle mean/p95/p99 favorable but control-drift-rejected);
guardrails clean, within-floor, or cleared-by-rerun. Candidate
`ec3dc81d…` is the control for the next change.)
[`change 0224`](changes/0224-odt-openparse-binding-tracker.md)
(0224 replaces `NsReader` in `OpenParse::run` with a plain `Reader` plus a
hand-rolled `BindingTracker` — the binding push/error stream replicated
byte-exactly (real `NamespaceError` values, so messages are identical by
construction; silent-break push scan, 256-declaration limit, unbinding
asymmetry, deferred pop, error preemption) plus an `xmlns` memmem
prefilter with a length-gated inline fast path and a flat-buffer binding
layout mirroring quick-xml's allocator pattern. Differential oracles pin
per-event resolutions at every depth across 24 adversarial cases and the
full corpus (898 litchi-odt tests). v1 showed a reproduced tiny-shape
(24-paragraph) p50 regression — diagnosed as a real fixed ~1.5 µs/open
overhead vs ~17 ns/paragraph saving (crossover ~70 paragraphs) and fixed
in v2; the harness's reported `results[0]` is the tiny shape, so the
medium/large wins (−11.6%/−24.1% p50) are documented as analysis
evidence. **Banked**: `odt_file_eager_open` p50/mean/p99
**9.41%-13.07% / 9.86%-14.48% / 14.01%-25.39%** (0223 floors),
`odt_file_source_open` p50 **21.85%-23.05%** (pre-floor),
`odt_file_source_open_full_text_lifecycle` mean **9.39%-11.77%** (0223
floor); `odt_semantic_open` no claim (drift/floor) but no adverse.
Candidate `48bd4072…` is the control for the next change.)
[`change 0223`](changes/0223-odt-source-path-floor-calibration.md)
(0223 is a methodology calibration, not a code change — the 0218 analog
extended to the ODT source-path and eager-open phases 0218 did not cover.
Three probe binaries (banked 0221 tree plus never-executed parser-shaped
padding of +6.1KB/+12.1KB/+14.6KB text in `litchi-odt::document` and
`elements`, retained via `#[used]`, .text deltas bracketing 0222's
−10,016) measured pure layout noise under the full A1/B1/B2/A2 protocol.
Effective floors, folding in the 0222 historical adverse-both readings:
file-source-open p95 2.5%, p99 28.0% (p50/mean uncalibrated — no
adverse-both evidence); file-source-lifecycle p50/mean/p95/p99
3.8%/2.5%/4.0%/6.5%; file-eager-open 5.6%/5.7%/9.3%/9.2% (eager is the
most layout-sensitive ODT open phase — probe a was adverse-both on all
four statistics with zero changed code). Under this rule 0222 was
re-verdicted **banked** — see below.)
[`change 0222`](changes/0222-odt-owned-open-fused-parse.md)
(0222 promotes the fused `OpenParse` to the owned ODT open path
(`from_owned_package`): one borrowing, depth-gated pass replaces the
standalone validator scan + `StyleElements::parse_styles(content.xml)`
rescan, with stage-by-stage error precedence preserved and pinned by 2
new cross-stage parity tests against the cfg(test) sequential oracle
(895 litchi-odt tests). Provisionally withheld under the pre-floor rule
(lifecycle p50 reproduced at max 1.76%), re-verdicted **banked** under
the 0223 floors. Claims: `odt_semantic_open` p50/mean/p95
**6.33%-6.45% / 10.02%-11.55% / 31.02%-33.20%** lower (0218 floors) and
`odt_file_eager_open` ALL FOUR statistics **12.05%-17.05% /
12.29%-16.82% / 11.46%-18.86% / 13.40%-17.27%** lower (0223 floors).
Re-applied bit-exact — rebuilt harness matches the measured candidate
`f53a43f1…`, the control for the next change.)
[`change 0221`](changes/0221-odt-openparse-borrowing-reads.md)
(0221 replays the banked 0220 transformation on the fused source-backed
ODT open (`OpenParse::run`, litchi-odt): borrowing `read_event()` +
depth-gated `resolve_element()` (Start/Empty, depth ≤ 2 — verified
arm-by-arm; `StyleHandler` never resolves). Pre-change loop retained as a
cfg(test) oracle; 2 new parity tests (26 synthetic edge cases + 69 ODT
fixtures, byte-identical errors; 893 litchi-odt tests). **Banked**:
`odt_file_source_open` p50/mean/p95 **43.28%-46.30%** lower and
`odt_file_source_open_full_text_lifecycle` p50/mean/p95
**20.55%-23.22%** lower, both directions, clean drifts — the largest
executed-phase wins of the series alongside 0217. All guardrails clean or
within-floor: the `odt_semantic_open` p99 primary above-floor pattern did
not reproduce at magnitude in its rerun, and the 0220 watch-listed
`ods_file_source_open` p95 reads within floor under this layout — flag
CLEARED. Candidate `93c2279b…` is the control for the next change.
Selected next: 0222 — promote the fused parse to the owned open path
(0219 candidate B, re-quantified post-0220 at ~27% of timed
`odt_semantic_open`).)
[`change 0220`](changes/0220-odf-validator-borrowing-reads.md)
(0220 rewrites the shared ODF content validator
`validate_content_document_part` — ~69% of the timed `odt_semantic_open`
call — with borrowing `read_event()` reads (no per-event buffer copy) and
depth-gated namespace resolution (only Start/Empty at depth ≤ 2, the only
arms whose resolved value is observable); the pre-change body survives as
a cfg(test) oracle cross-checked on 41 synthetic edge cases plus the full
ODF corpus with byte-identical error messages (248 litchi-odf-common,
891 litchi-odt tests). Sole production caller is the ODT owned open path.
**Banked**: `odt_semantic_open` p50 **5.99%-7.30%** lower (over the 0218
floor 3.3%) and `odt_file_eager_open` p95 **19.12%-21.14%** lower claimed;
the `odt_semantic_full_text` p95 primary adverse was cleared by its
rerun; `ods_file_source_open` p95 reproduced marginally above the 0205
floor (max 4.90% vs 4.5%) — recorded as a flagged above-floor layout
reading on a zero-changed-code phase with bit-identical read evidence,
floors unchanged, watch-listed for the next ODF change. Candidate
`1971c3ad…` is the control for the next change.)
[`change 0219`](changes/0219-odf-validator-reprofile.md)
(0219 is an analysis, not a code change: post-0217 `perf record` profiles
of the three ODF family open workloads attributed
`validate_content_document_part`'s internals — per-event buffer copies
and per-event namespace resolution whose result is consumed only at
depth ≤ 2 — re-tested the "shared across families" premise (only the ODT
owned open path calls it), and selected the borrowing-reads +
depth-gated-resolution target implemented as 0220. No measurement legs;
profiling data under `/tmp/0219-prof/`.)
[`change 0218`](changes/0218-odt-layout-floor-calibration.md)
(0218 is a methodology calibration, not a code change — the litchi-odt
analog of 0205/0213: three probe binaries — banked 0215 tree plus
never-executed parser-shaped padding (+5.8KB to +14.6KB text, in
`elements::text`, `document`, and `parser`) retained via a `#[used]`
table — measured under the full A1/B1/B2/A2 protocol, so every
paired-direction reading is pure per-binary-pair layout noise. The
effective floor for adverse both-directions patterns on unexecuted
phases, folding in historical 0217 byte-identical-phase readings: open
p50/mean 3.3%/7.2% (p95 27.6%, p99 28.2%), list-paragraphs p50/mean
5.2%/6.7% (p95 22.4%), one-paragraph p50/mean 4.8%/8.3% (p95 52.6%),
full-text p50 4.1% (p95 16.1%, p99 9.3%), repeated-text-cached p50/mean
7.1%/7.4% (p95 7.5%, p99 29.2%), repeated-text-uncached p50/mean
4.8%/4.1% (p95 3.2%, p99 8.2%). ODT tails are far more layout-sensitive
than ODS/ODP; p50/mean are the operative banking statistics. The 0205
banking rule extends unchanged to litchi-odt. Under this rule 0217 was
re-verdicted **banked** — see below.)
[`change 0217`](changes/0217-odt-discard-validate-text.md)
(0217 gives the ODT `extract_text` path a discard-but-validate parse mode:
`parse_text_block_texts` runs the identical event loop, suppression rules,
depth accounting, and `make_text_block_element`-ordered attribute
validation, but never builds the retained `Element` (no QualifiedName
triple-alloc, no attributes HashMap, no owned attr strings) — the banked
in-file precedent `parse_selected_text_block_element(retain=false)` made
this low-risk; 3 new parity/precedence/limit tests against the
cfg(test)-gated pre-change path as a live oracle (891 litchi-odt tests).
**Banked** under the 0218 floor rule after re-verdict: all three executed
workloads accept ALL FOUR statistics in both directions —
`odt_semantic_full_text` **42.12%-52.07%**,
`odt_source_backed_repeated_text_cached` **51.91%-57.18%**,
`odt_source_backed_repeated_text_uncached` **40.31%-53.85%** lower; the
reproduced guardrail layout readings (open p50 max 3.26% vs floor 3.3%;
list-paragraphs mean max 6.67% vs floor 6.7%) are within-floor and no
longer block. Re-applied bit-exact — the rebuilt harness matches the
measured candidate `8425066a…`, which is the control for the next
change.)
[`change 0216`](changes/0216-odt-query-reprofile.md)
(0216 is an analysis, not a code change: post-0215 `perf record` profiles
of the four ODT semantic workloads refuted the double-tokenization
hypothesis (each query call tokenizes `content.xml` exactly once) and
located the real `full_text` cost — per-block retained `Element`
materialization immediately discarded by `into_text` — selecting the
discard-but-validate extraction (implemented as 0217). Open's dominant
cost (69% of the timed call) is `validate_content_document_part`, a full
`content.xml` tokenization in litchi-odf-common shared with all ODF
families — rejected as a target for rerun-exposure reasons. No
measurement legs; profiling data under `/tmp/0216-prof/`.)
[`change 0215`](changes/0215-odp-shape-attr-harvest.md)
(0215 folds the ODP `drawing_attributes` fresh attribute re-scan
(12.7%-13.7% inclusive of every ODP workload post-0212) into the shared
`ElementAttrs` incremental scan: a new `ElementAttrs::drawing_attributes`
replays the cached prefix and continues the shared iterator, harvesting
non-modeled DRAW/SVG/DR3D/TABLE attributes from both halves, and
`shape_builder` drops its per-element `element.attributes()` re-scan (the
deleted fresh-scan body survives verbatim as a test oracle). Exactness
pinned by 7 new tests (157 lib + 111 integration): document order,
error-message identity by first reach (`"invalid XML attribute"` vs
`"invalid ODP shape attribute"`), duplicate detection, decode positions.
**Banked** under the 0213 floor rule: `odp_semantic_list_slides` p50/mean
**6.53%-9.97%**, `odp_semantic_one_slide` p50/mean **8.80%-15.10%**,
`odp_semantic_full_text` p50/mean **5.44%-11.32%** lower, all far above
floor; executed-phase p95/p99 withheld (corpus-tail drift). The
byte-identical open workload's adverse p50 reading (max 2.48%) is a
within-floor layout reading (floor 3.1%; floor invoked for a −2.1KB text
delta, sign-agnostic reasoning documented in the change doc). The 0215
candidate binary `6c7fcfb9…` is the control for the next change.)
[`change 0214`](changes/0214-odp-post-0212-reprofile.md)
(0214 is profiling/analysis, not a code change — dwarf call-graph profiles
of all four litchi-odp semantic workloads on the post-0212 banked tree.
Post-0212 the attribute cluster still dominates: `ElementAttrs::get` self
6.4%-7.8% / inclusive 25.3%-27.7% on every workload;
`Parser::drawing_attributes` inclusive 12.7%-13.7% is the largest removable
block — every shape element is attribute-scanned twice (shape_builder's
lazy lookups, then a fresh `element.attributes()` pass for unmodeled
attributes). `resolve_prefix` fell from 24.56% to 5.5%-7.0% inclusive after
0212. `odp_semantic_open` has no unique hotspot — same cluster plus inflate
≈3.1%, memchr ≈4.5%, allocator ≈5-6% self. Selected next target: fold
`drawing_attributes` into the shared `ElementAttrs` scan (single-scan
shape-attribute harvest), expected 6-10% on all four workloads, with
document-order and error-message-identity exactness constraints recorded in
the change doc.)
[`change 0213`](changes/0213-odp-layout-floor-calibration.md)
(0213 is a methodology calibration, not a code change — the litchi-odp
analog of 0205: three probe binaries — banked 0211 tree plus
never-executed parser-shaped padding (+5.5KB to +14.5KB text, in
`codec::xml`, `model`, and `package`) retained via a `#[used]` table —
measured under the full A1/B1/B2/A2 protocol, so every paired-direction
reading is pure per-binary-pair layout noise. The effective floor for
adverse both-directions patterns on unexecuted phases, folding in
historical 0211/0212 byte-identical-phase readings: open p50/mean
3.1%/2.5% (p95 7.8%, p99 17.2%), list-slides p50/mean 2.0%/3.6% (p95
6.3%, p99 10.1%), one-slide p50/mean 2.5%/3.2% (p95 17.8%, p99 14.4%),
full-text p50/mean 0.1%/0.5% (p95 1.8%, p99 19.4%). The 0205 banking
rule extends unchanged to litchi-odp: within-floor adverse readings on
phases executing no changed code are layout readings and do not block;
adverse readings on executed phases still block unless cleared by the
single rerun; accepts are claimed only above the floor. Under this rule
0212 was re-verdicted **banked** — see below.)
[`change 0212`](changes/0212-odp-cached-attr-resolution.md)
(0212 caches attribute namespace resolution in the litchi-odp slide
parser: `ElementAttrs::get` was 39.66% inclusive of
`odp_semantic_full_text`, dominated by replaying
`resolver().resolve_attribute(key)` for every cached attribute on every
lookup (`NamespaceResolver::resolve_prefix` 24.56% inclusive). Each
attribute key is now resolved ONCE when the incremental scan first
reaches it, storing a `ResolvedAttributeNamespace` snapshot
(`Bound`/`Unbound`/`Unknown`) plus the borrowed local name; lookup replay
compares the snapshot via `matches()` — semantically identical to
`is_namespace` on the live `ResolveResult` — with no resolver calls, and
value decoding still happens per-match at lookup time. The exactness
invariant (namespace stack constant between `ElementAttrs::new` and last
use) was audited at all 12 call sites; two new tests pin cached-vs-fresh
parity across an 8-target matrix and nested prefix shadowing; 150 lib +
111 integration litchi-odp tests pass. Originally withheld under the
pre-floor rule for a reproduced adverse open p50 reading (max 1.63%) on
a phase executing no changed code; **banked** after the 0213 floor
(open p50 3.1%) reclassified it as a layout reading. Claim scope:
`odp_semantic_full_text` p50/mean/p95/p99 **20.82%-29.50%**,
`odp_semantic_list_slides` p50/mean **19.15%-25.51%**,
`odp_semantic_one_slide` p50/mean/p99 **17.48%-31.16%** lower.
Re-applied bit-exact — the rebuilt harness matches the measured
candidate `246c6b1f…`, which is the control for the next change.)
[`change 0211`](changes/0211-odp-fused-query-parse.md)
(0211 fuses the per-query double scan of `content.xml` in litchi-odp's
`parse_pages_with_styles`: the transition-style definitions pass (14.95%
inclusive of `odp_semantic_full_text`) becomes an event-fed
`TransitionStyleCollector` inside the single slide-parse tokenization,
with record-first-error and deferred-error plumbing preserving the
historical two-pass error precedence exactly (all six nested parsers
feed the collector; the standalone transition shell stays byte-identical
as the oracle anchor). Equivalence oracle: verbatim sequential reference
cross-checked on 19 fixtures across all four entry points plus 14
synthetic precedence pins; 148 lib + 111 integration litchi-odp tests
pass. **Banked** under pre-floor acceptance: `odp_semantic_list_slides`
p50/mean **15.56%-17.84%**, `odp_semantic_one_slide` p50/mean
**18.88%-19.50%**, `odp_semantic_full_text` p50 **15.85%-16.43%** lower;
the byte-identical open workload's adverse mean/p95 primary reading did
not reproduce in the single permitted rerun. The 0211 candidate binary
`ceba155b…` — tree rebuilds bit-exact — is the control for the next
change.)
[`change 0210`](changes/0210-odp-lazy-element-attrs.md)
(0210 eliminates the O(n·k) attribute re-scan in the litchi-odp slide
parser: `get_attr` re-iterated `element.attributes()` from scratch per
lookup (24.61% self of `odp_semantic_full_text`, plus ~18% quick-xml
attribute-iteration machinery), and handlers like `shape_builder` make
14 sequential lookups per element. The new `ElementAttrs` lazy
incremental cache parses each element's attributes once into raw
zero-copy entries; namespace resolution and value decoding still happen
at lookup time with the current reader, so error messages, malformed/
duplicate-attribute positions, and decode behavior are per-lookup
identical (pinned by 6 new tests; 245 litchi-odp tests pass). Ten
multi-lookup handlers migrated. **Banked** under pre-floor acceptance —
every executed workload accepts ALL FOUR statistics in both directions:
`odp_semantic_list_slides` **16.77%-47.07%**, `odp_semantic_one_slide`
**7.45%-19.70%**, `odp_semantic_full_text` **13.11%-21.71%** lower; the
non-executed open p50/mean (1.01%-3.92%) is layout-favorable and
recorded as such. No adverse pattern anywhere; no rerun needed. The 0210
candidate binary `c4b2b568…` — tree rebuilds bit-exact — is the control
for the next change.)
[`change 0209`](changes/0209-odt-fused-open-parse.md)
(0209 fuses the two complete `content.xml` tokenizations on the ODT
source-backed open path — `validate_content_document_part` (26.23%
inclusive of the workload) and the `StyleRegistry::from_xml`
content-styles scan (6.58%) — into one `NsReader` pass with two
handlers in `litchi-odt::document::open_parse`, the litchi-odt analog of
0201. Validation errors keep their historical early-return (styles.xml
is never fetched on a validation failure); the style handler replicates
`from_xml` byte-exactly including the literal `b"style:style"` raw-qname
match and raw undecoded attribute values; error precedence
validate → styles.xml → content-styles → `try_extend` is exact. An
equivalence oracle cross-checks fused vs sequential on 69 fixtures plus
15 synthetic malformed cases; 888 litchi-odt tests pass (+3). **Banked**
under pre-floor acceptance (the 0205 floor is ODS-only):
`odt_file_source_open` p50/mean/p95 **14.63%-18.99% lower** in both
directions with clean drifts; source-open full-text lifecycle mean
0.71%-4.53% lower; the sub-1.2% eager-open p50/mean adverse reading on
the byte-identical phase did not reproduce in the single permitted
rerun. The 0209 candidate binary `41b5f923…` — tree rebuilds bit-exact —
is the control for the next change.)
[`change 0208`](changes/0208-ods-borrowed-validate-decode.md)
(0208 tried borrowed `decode_borrowed`/`resolve_namespace_borrowed`
decoding in the ODS commit validate-reparse (`worksheet/package.rs`),
with a URI-equality `NamespaceKind` partition replacing owned-String
namespace comparisons. The deterministic allocation win was large —
counting-allocator driver measured 36,616 → 20,863 allocations per
commit transaction (-43.0%; commit-only 29,473 → 13,720, -53.5%) — but
the latency gate failed: both executed-phase workloads with an adverse
both-directions pattern **reproduced** it in their single permitted
rule-2 reruns (one-edit commit p50/mean/p95 -2.12% to -6.94%;
repeated-edit commit all-four -2.40% to -5.84%, clean drifts). Per
banking rule 2, **withheld regardless of floor** and reverted; the tree
rebuilds bit-exact to the banked 0207 control `57270d24…`. Lesson: on
this corpus the commit phase is layout-sensitive enough that an
allocation-only change can produce a reproducible sub-5% adverse
reading; allocation-only commit changes should expect this outcome
unless paired with an above-floor latency mechanism. The banked control
for the next change remains `57270d24…`.)
[`change 0207`](changes/0207-ods-byte-matched-attributes.md)
(0207 rewrites `worksheet::codec::Attributes::from_resolved` to
byte-match resolved namespace/local names and allocate owned strings only
for consumed attribute values, while still decoding EVERY attribute so
the historical per-attribute error order (syntax → decode/normalize →
unknown prefix) and messages are preserved exactly — pinned by five new
error-order tests; the 0200 shell-oracle pair shares one implementation.
**Banked with an allocation claim only**: counting-allocator driver
measured 10,727 → 8,665 allocations per source-open (-19.22%) and
1,928,711 → 1,829,857 bytes (-5.13%); cumulative with 0206, source-open
allocations are 14,891 → 8,665 (-41.81%). Latency neutral everywhere
under the 0205 floor rule — source-open p50 accepted at 5.46% against
the 5.5% floor, everything else below floor, no adverse both-directions
pattern on any workload. The retained per-attribute decode scan is
semantics-mandated. The 0207 candidate binary `57270d24…` is the control
for the next change.)
[`change 0206`](changes/0206-ods-lazy-settings-qname.md)
(0206 removes dead per-element work in the ODS settings location scan:
the qualified element name was materialized (alloc + copy + UTF-8
validation + free) for EVERY Start/Empty element, but spans are recorded
only for the spreadsheet host and calculation-settings element, and the
only qname consumer is the empty-spreadsheet `replace()` expansion. The
name is now materialized only for recorded kinds, identically in the
locate shell and the fused handler; the dropped UTF-8 error is
structurally unreachable (ASCII-delimited subslices of a `&str` source).
**Banked with an allocation claim only**: a counting-allocator driver
measured 14,891 → 10,727 allocations per source-open (-27.96%) and
1,974,783 → 1,928,711 allocated bytes (-2.33%) — deterministic counts,
not timing. All accepted latency statistics landed below the
0205-calibrated layout floor (neutral; no latency claim); the one
adverse both-directions reading (one-percent commit p99 -12.18%) is
within its 13.5% floor. Lesson: profile-attributed ~4-6% latency
expectations can vanish under the layout floor; allocation-count evidence
is the robust claim channel for work elimination at this scale. The 0206
candidate binary `8e17ab3e…` is the control for the next change.)
[`change 0205`](changes/0205-layout-noise-floor-calibration.md)
(0205 is a methodology calibration, not a code change: three probe
binaries — banked 0202 tree plus never-executed parser-shaped padding
(+5.5KB to +16.5KB text) retained via a `#[used]` table — measured under
the full A1/B1/B2/A2 protocol, so every paired-direction reading is pure
per-binary-pair layout noise. The measured floor for adverse
both-directions patterns on unexecuted phases: source-open 5.5% (p99 up
to 35%), eager-open 3.0%, one-edit lifecycle 2.8%, one-edit commit
5.8% (p99 17%), one-percent lifecycle 2.6%, one-percent commit 4.6%
(p99 13.5%), repeated-edit stage 5.3% (p99 11%), repeated commit 6.7%,
totals/publication ~1-2.5%. The refined banking rule: within-floor
adverse readings on phases executing no changed code are layout readings
and do not block; adverse readings on executed phases still block unless
cleared by the single rerun; accepts are claimed only above the floor.
Under this rule 0204 was re-verdicted **banked** — see below.)
[`change 0204`](changes/0204-ods-protection-fused-parse.md)
(0204 fuses the ODS protection double-parse of `content.xml`
(`Location::parse` + `parse_protection`, two `NsReader` passes per fresh
owner's first edit) into one tokenization with the 0200-0202
handler/driver pattern, error selection preserving the historical
interleave exactly and standalone shells kept byte-identical as
equivalence oracles (369 tests, +13). **Banked** under the 0205 refined
rule — originally withheld for reproduced adverse patterns on
non-executed phases, all later shown to be within the calibrated layout
floor. Claim scope: repeated-edit stage p50/mean/p95/p99 20.01%-25.86%
lower (floor 3.8%/5.3%/11%); one-percent lifecycle p50/mean/p95 marginal
(≤0.25pp over floor, not claimed); repeated-edit total and one-edit
lifecycle within floor (neutral). Harness rebuild bit-identical to the
measured candidate `5f0dab64…`; this binary is the control for the next
change.)
[`change 0203`](changes/0203-ods-open-namespace-classify-memo.md)
(0203 explored memoizing the fused open driver's element-name namespace
classifications — the quick_xml `resolve_event` reverse bindings scan,
6.25% of post-0202 source-open self samples — first with a SipHash
HashMap (v1: mechanism-confirmed regression on every workload, e.g.
source-open 5.05%-15.22% slower in both directions) and then with a
direct-mapped 64-slot generation-tagged array cache with conservative
`xmlns` mutation tracking (v2: targeted source-open neutral, guardrails
broadly adverse in both directions including source-identical phases).
Both withheld and reverted per the no-regression-pattern standard; the
tree is restored byte-exact to the 0202 state (harness SHA match, 356
litchi-ods tests). A content-fingerprint invalidation scheme was rejected
by counterexample (`xmlns:p="xy"` vs `xmlns:px="y"`). The remaining
namespace-machinery seam is `NsReader::process_event`'s per-tag attribute
scan, likely requiring a litchi-owned incremental resolver.)
[`change 0202`](changes/0202-ods-single-tokenization-open.md)
(0202 folds the pass-2a calculation-settings parse
(`litchi_odf_common::calculation::parse`) into the fused open tokenization
as a fifth handler, so the source-backed open tokenizes `content.xml`
exactly once; the standalone parse keeps its original inline loop
byte-identical for the eager facade and commit-side callers, the 64 MiB
size limit moves to the deferred record-first-error pattern, and the
two-phase driver preserves the exact validate → calculation → locate →
names → worksheet pass order and error selection. Frozen cross-binary
CPU-2 A/B/B/A accepts the source-backed open p50/mean/p95/p99
(17.41%-23.21% lower, stacking on 0201), the one-edit lifecycle
p50/mean/p95/p99 (0.81%-10.47% lower) and commit p50/mean, the
one-percent lifecycle p50/mean/p95 in both the primary run and the
single rerun (whose commit accepts all four statistics), the
repeated-edit total p50/mean, stage p50/mean/p95, and commit
p50/mean/p95/p99, and the eager-open mean. The one-percent lifecycle p99
primary adverse reading did not reproduce in the single permitted rerun;
a sub-0.5% repeated-edit publication p50/mean reading on
source-identical phases is documented as code-layout wobble.
Allocation/RSS, physical-I/O, cold-cache, producer, and broad ODF claims
remain withheld.)
[`change 0201`](changes/0201-ods-fused-open-validate-fold.md)
(0201 folds the ODS `content.xml` structural validation pass into the 0200
fused open tokenization as a fourth handler, cutting the source-backed open
from three tokenizations to two; the two-phase driver preserves the exact
validate → styles/meta/metadata → semantic-calculation-parse interleave and
the pass-ordered error selection, and the standalone validator keeps its
original inline loop for its five other call sites. Frozen cross-binary
CPU-2 A/B/B/A accepts the source-backed open p50/mean/p95/p99
(9.32%-17.88% lower, stacking on 0200), the one-edit lifecycle
p50/mean/p95/p99 (1.87%-5.23% lower) and commit p50/mean, the one-percent
lifecycle p95, and the repeated-edit total p95/p99 plus publication
p50/mean/p95/p99. The eager-open primary adverse reading did not reproduce
in the single permitted rerun; a sub-1.5% repeated-edit commit p50/mean
reading on source-identical phases is documented as code-layout wobble.
Allocation/RSS, physical-I/O, cold-cache, producer, and broad ODF claims
remain withheld.)
[`change 0200`](changes/0200-ods-fused-open-parse.md)
(0200 fuses the three litchi-ods-owned open passes over `content.xml` —
settings locate, named definitions, and the worksheet parse — into one
shared `NsReader` tokenization with per-pass handler state machines and
pass-ordered error selection; the standalone shells keep the original inline
loops byte-identical so eager-open, commit-readback, and settings-edit paths
are unchanged. Frozen cross-binary CPU-2 A/B/B/A accepts the source-backed
open p50/mean/p95/p99 (19.72%-24.55% lower), the one-percent lifecycle
p50/mean/p95/p99 (3.26%-5.72% lower), the one-edit lifecycle p50/mean
(2.81%-5.38% lower), the eager-open p99, and the repeated-edit total
mean/p95/p99, stage p99, and publication mean/p95/p99. A reproduced
both-directions sub-1% slower reading on repeated-edit total p50 and
publication p50 — phases source-identical to control whose p95/p99 tails
accept — is documented as per-binary-pair code-layout wobble, not a
regression pattern. Allocation/RSS, physical-I/O, cold-cache, producer, and
broad ODF claims remain withheld.)
[`change 0199`](changes/0199-ods-parse-event-copy-elision.md)
(0199 removes the per-event `Event::into_owned()` deep copies from the two
full-document `NsReader` parse loops in litchi-ods (`worksheet::codec::parse`
and `settings::codec::locate`); both loops only borrow each event within its
iteration, so the elision is byte-exact. Frozen cross-binary CPU-2 A/B/B/A
accepts the source-backed open p50/mean/p95/p99 (6.42%-9.70% lower) and the
eager open p50/mean/p95/p99 (9.46%-15.17% lower); the one-cell guardrail
lifecycle p50/mean/p95 (0.97%-3.42% lower) and commit p50/mean/p95
(6.62%-9.43% lower); the one-percent guardrail lifecycle p50/mean/p95
(0.05%-2.71% lower) and commit p50/mean/p95 (4.16%-7.52% lower); and the
four-transaction repeated-edit total, commit, and publication
p50/mean/p95/p99 (2.14%-5.28%, 8.61%-14.80%, and 0.17%-3.03% lower
respectively) plus stage p50/mean/p95. The lifecycle/commit p99 tails and
stage p99 are withheld as neutral and no regression pattern fired.
Allocation/RSS, physical-I/O, cold-cache, producer, and broad ODF claims
remain withheld.)
[`change 0198`](changes/0198-ods-content-layout-topology-cache.md)
(0198 extends the 0195 per-owner ODS content-layout cache with the derived
table/row topology, eliminating the per-commit span-vector re-scans for the
table inventory and per-sheet row lists on cached-layout commits. Frozen
cross-binary CPU-2 A/B/B/A accepts the four-transaction repeated-edit total
p50/mean/p95/p99 (0.19%-5.77% lower), commit p50/mean/p95/p99 (2.20%-7.72%
lower), stage mean/p95/p99 and publication p50; the one-cell guardrail
lifecycle mean/p95 and commit p99; and the one-percent guardrail lifecycle
p50/mean/p95 plus commit p50/mean/p95; all other statistics are withheld as
neutral and no regression pattern fired. Allocation/RSS, physical-I/O,
cold-cache, producer, and broad ODF claims remain withheld.)
[`change 0197`](changes/0197-ods-batched-row-window-reparse.md)
(0197 explored batching the per-row synthetic-document reparse in the ODS
`changed_row_edits` validation into one reparse per changed window. The
commit-phase win reproduced (one-percent commit p50/mean/p95 0.10%-2.10%
lower, accepted in two independent runs), but the one-percent lifecycle
measured the candidate slower in both paired directions on all four
statistics in both runs, and the repeated-edit stage phase showed the same
adverse pattern — a systematic per-binary code-layout effect on untouched
phases. Withheld and reverted per the no-regression-pattern standard; the
two added multi-row window regression tests remain. The seam stays open:
a future attempt should eliminate the synthetic reparse entirely rather
than amortize it.)
[`change 0196`](changes/0196-xml-escape-byte-scan.md)
(0196 replaces the Aho-Corasick automata behind
`litchi_core::xml::{escape_xml, unescape_xml}` with exactly equivalent
left-to-right byte scans (fuzz-verified byte-identical over 2M randomized
cases against the original automata) and drops the dependency from
litchi-core. Frozen cross-binary CPU-2 A/B/B/A accepts the escape-heavier
one-percent ODS lifecycle p50/mean (0.44%-3.02% lower) and commit
p50/mean/p99 (1.53%-20.15% lower); the repeated-edit and one-edit selectors
are neutral-withheld throughout. Allocation/RSS, physical-I/O, cold-cache,
producer, and broad ODF claims remain withheld.)
[`change 0195`](changes/0195-ods-source-content-layout-cache.md)
(0195 caches the row-local edit layout scan of `content.xml` at most once
per source-backed ODS owner through a private success-only `OnceLock`,
removing three of four full-document scans in the four-transaction
repeated-edit selector. Frozen cross-binary CPU-2 A/B/B/A accepts
four-transaction total p50/mean/p95/p99 at 5.55%-12.49% lower and
commit-phase p50/mean/p95/p99 at 25.13%-36.98% lower, plus the one-cell
guardrail commit p50/mean/p95/p99 and the one-percent guardrail lifecycle
p99 and commit p50/mean/p99; all other statistics are withheld as neutral.
Single-transaction lifecycles still pay one scan by design; allocation/RSS,
physical-I/O, cold-cache, producer, and broad ODF claims remain withheld.)
[`change 0194`](changes/0194-ods-validate-text-byte-scan.md)
(0194 rewrites the litchi-ods worksheet `validate_text` forbidden-character
check from a per-`char` scan to an exactly equivalent per-byte scan; profiling
attributed 9.15% of source-backed commit-phase samples to it. Frozen
cross-binary CPU-2 A/B/B/A accepts the four-transaction repeated-edit total
p50/mean and commit-phase p50/mean/p99, the one-cell guardrail commit p50,
and the one-percent guardrail lifecycle p50/mean plus commit p50/mean/p95;
all other statistics are withheld as neutral. Allocation/RSS, physical-I/O,
cold-cache, producer, and broad ODF claims remain withheld.)
[`change 0193`](changes/0193-ods-source-edit-protection-cache.md)
(0193 computes the ODS source-backed edit protection parse at most once per
owner through a private `OnceLock`, removing three of four complete
protection-domain parses in the new four-transaction repeated-edit selector.
Frozen cross-binary CPU-2 A/B/B/A accepts four-transaction total
p50/mean/p95/p99 at 9.31%-10.68% lower and stage-phase p50/mean/p95/p99 at
67.87%-71.61% lower; commit and publication phases are neutral-withheld.
Single-transaction lifecycles are unchanged by design; allocation/RSS,
physical-I/O, cold-cache, producer, and broad ODF claims remain withheld.)
[`change 0192`](changes/0192-odt-open-only-rerun-evidence.md)
(0192 repeats only the withheld change 0191 ODT open-only workload on clean
current HEAD with a bit-identical release binary. Warm open-only p50 and p99
are now accepted at 49.01%-59.56% lower than eager byte ownership across both
paired directions; mean and p95 remain withheld on eager same-implementation
drift (5.10% and 11.58% against 5%/10% ceilings). This is evidence closure,
not a production change; allocation/RSS, physical-I/O, cold-cache, edit/save,
producer, and broad ODF claims remain withheld.)
[`change 0191`](changes/0191-odt-unified-source-ingress.md)
(0191 routes high-level ODT filesystem opening through one retained
source-backed ODF owner. Fixed-corpus open-only samples remain withheld because
same-implementation drift fails each tier. Open-plus-full-text p50/mean/p95/p99
reductions of 30.02% to 35.36% pass both paired-direction and drift gates. An
untimed replay reads 29,080 logical bytes and zero picture range bytes.
Physical-I/O, cold-cache, resource, producer, edit/save, broad ODF and iWork
claims remain withheld.)
[`change 0190`](changes/0190-cfb-stream-chain-scratch.md)
(0190 reuses fallible MiniFAT/FAT stream-chain validation scratch. On its exact
two-shape process profile, allocation calls fall 48.44% and Heaptrack temporary
allocations fall 98.94%; accepted release timing is limited to many-small
p95/p99 and wide-root p50/mean/p95. Operation-local bytes, peak-RSS
improvement, cold/physical I/O, concurrent contention, save/edit and native
DOC/XLS/PPT semantic claims remain open.)
[`change 0188`](changes/0188-ooxml-root-lifecycle-evidence.md)
(0188 adds matched fresh-open-plus-query DOCX/PPTX lifecycle controls. Every
warm CPU-2 release direction is descriptively lower, but source-backed PPTX
and paragraph-count p50/mean drift plus eager full-text drift miss the
predeclared gates, so no latency statistic is accepted. Resource, physical-I/O,
cold-cache, producer, edit/save, broad OOXML and iWork claims remain open.)
[`change 0187`](changes/0187-xlsx-unified-source-ingress.md)
(0187 routes high-level XLSX filesystem opening through the existing
source-backed OPC/workbook owner. On one generated four-sheet media-rich
corpus, open-only p50/mean/p95/p99 are 91.59%-93.10% lower across both clean
paired directions; open plus names/count/full text is 14.35%-18.30% lower.
The evidence is warm and in-process; physical-I/O, cold-cache, resource,
producer, edit/save, broad OOXML and iWork claims remain withheld.)
[`change 0186`](changes/0186-opc-eager-shared-payloads.md)
(0186 carries one immutable decompressed allocation from serial eager ZIP
ingress through OPC XML/binary Part construction, removing one full payload
copy per admitted Part. The 16 MiB owned-open diagnostic peak heap changes
71.72M -> 55.02M. Few-large p50 directions are both materially lower but fail
the control-drift gate; only owned-open p99 is accepted. Ordinary open still
inflates every admitted Part, and many-small, I/O, decompression, scaling,
producer, broad OOXML and iWork claims remain withheld.)
[`change 0185`](changes/0185-opc-shared-source-overlay.md)
(0185 adds an additive Arc-owned source-overlay handoff in OPC and migrates
eligible DOCX/PPTX/XLSX publishers. One complete selected-Part ownership copy
is removed while selected source comparison, XML validation, compression,
signatures, budgets, freshness and sink semantics remain. Matched XLSX results
are mixed and scenario-scoped; no allocation/RSS, I/O, decompression,
topology-changing or broad OOXML claim is made.)
[`change 0184`](changes/0184-xlsx-row-visibility-store-reuse.md)
(0184 removes one complete scalar-cell parse from each changed source-backed
existing-row visibility commit through a private lifetime/source-bound rewrite
proof. Large commit statistics and large-batch complete lifecycle pass the
paired gates; medium total and medium hide-one latency remain withheld, as do
resource, physical-I/O, producer, structural-row, formula and broad XLSX
claims.)
[`change 0183`](changes/0183-ods-one-percent-release-evidence.md)
(0183 closes the previously withheld fixed ODS 21-existing-cell result. A
clean current-HEAD A/B/B/A rerun accepts complete source-backed lifecycle p50
at 72.07%-72.61% lower than eager ownership; all distribution and stability
gates pass. This is evidence closure, not a production change, and resource,
physical-I/O, producer, structural and broad ODS claims remain withheld.)
[`change 0182`](changes/0182-pptx-validation-catalog-graph-fusion.md)
(0182 fuses the bounded PPTX validator's catalog and graph traversal. Package
relationship passes change 2 -> 1 and per-Part relationship passes 4 -> 1.
The clean large semantic-corpus validation p50 result is accepted at
7.08%-11.50% lower in paired directions; tiny/medium latency plus resource,
physical-I/O, cold-cache, scaling, producer, and broad PPTX claims remain
withheld.)
[`change 0181`](changes/0181-xls-source-policy-reuse.md)
(0181 reuses immutable native-XLS snapshot policy facts in the existing
plan-only fixed-width numeric path. Number total and commit p50/mean/p95/p99
pass the clean paired-direction/stability gates; RK/MulRK latency and all
publication/resource/I/O claims remain withheld.)
[`change 0180`](changes/0180-odt-source-text-cache.md)
(0180 retains one bounded full-text projection on the first successful parse
after the two-call source-backed ODT threshold. Four candidate public calls
reduce complete `content.xml` projection phases from four to two. Two clean balanced cycles accept p50 and
mean only; p95/p99 and resource/I/O claims remain withheld.)
[`change 0179`](changes/0179-pptx-source-catalog-reuse.md)
(0179 retains the source-backed PPTX editor's validated presentation catalog.
One-slide workflows remove two full 200-slide graph builds; the eight-slide
batch removes nine. Exact materializations and logical source reads are
unchanged. Clean paired latency directions disagree and stability gates fail,
so only deterministic catalog-build/allocation work is accepted.)
[`change 0178`](changes/0178-cfb-owned-planning-fingerprint.md)
(0178 removes one final complete logical fingerprint scan from sealed owned
CFB planning after candidate reopen and optional format-owner validation.
Generic positional sources retain their hostile stable-token fence. The fixed
XLS corpora remove 16,995,840 bytes/17 reads or 202,752 bytes/one read per
effective plan; clean paired latency directions are lower but stability drift
withholds every workload-level latency claim.)
[`change 0172`](changes/0172-cfb-owned-numeric-publication.md)
(0172 preserves immutable owned-byte provenance into native XLS numeric plans
and removes the two redundant outer fingerprint scans from direct publication.
Both measured complete workflows and publication through p95 pass the clean
paired-direction and drift gates; atomic save and resource claims are
unchanged/withheld.)
([`0171`](changes/0171-cfb-owner-validation-fusion.md) fuses semantic owner
validation into the existing final CFB fingerprint fence for source-backed DOC
paragraph, PPT shape-text, and XLS visibility
transactions. It removes one complete source scan per effective transaction;
the measured XLS batch total and scalar/batch plan phases pass paired-direction
and drift gates, while narrower totals, tails, publication and resource claims
are withheld.)
([`0170`](changes/0170-xlsx-streaming-escape-runs.md) batches the measured
one-sheet XLSX streaming XML text path while
preserving exact output. Large all-statistic, medium through-p95, and tiny p50
directions are accepted; remaining tiny statistics and medium p99 are withheld,
and branch misses regress in the descriptive process-counter run.)
(0169 retains a shared hierarchical-budget charge optimization selected by
the one-sheet XLSX streaming writer. Medium/large and tiny-through-p95 latency
directions agree; tiny p99 is withheld, process Heaptrack allocation calls fall,
peak heap is flat, and RSS directions disagree.)
(0168 retains a narrow native-XLS validation-fusion mechanism and raw release
evidence but withholds an acceptance-grade latency claim because the same-
implementation drift gate failed. 0167 retains an XLSX publication work-
elimination mechanism and raw
release evidence but withholds an acceptance-grade latency claim because the
same-implementation drift gate failed. 0165 is an accepted native-DOC
lifecycle/workflow result for the exact
deterministic corpus; it does not complete the native owner/public-reader or
CRUD matrix.)
(0152 is a final clean release ABBA for same-target MiniFAT single-flight;
all correctness/source-event invariants passed and the existing concurrent
scenarios recorded 6,473 versus 8,000 logical source calls, but only that
source-event/correctness scope is accepted. Change 0151 is a production
managed-XLSX correctness/resource freeze and adds no performance result; the
latest accepted CFB timing result remains the configured-simulator CFB repeat
evidence in [`change 0149`](changes/0149-cfb-same-target-repeat-release-abba.md)).
(the newest accepted semantic-format optimization remains limited to four repeated
source-backed ODP full-text projections; the latest release CFB selective-range
evidence is the configured simulator result in
[`0144`](changes/0144-cfb-simulated-range-source-evidence.md), while
[`0146`](changes/0146-cfb-open-stream-evidence.md) adds `open_stream`-specific
correctness/counter instrumentation and
[`0147`](changes/0147-cfb-open-stream-release-abba.md) accepts only its
configured-simulator one-shot result while retaining the repeat tradeoff, and
[`0148`](changes/0148-cfb-same-target-repeat-policy.md) adds correctness/source-
event coverage for different-SID A-B-A, public bulk A-B-A, and overlapping
same-target calls, and
[`0149`](changes/0149-cfb-same-target-repeat-release-abba.md) accepts only the
configured simulator's aggregate same-target repeat result while withholding
local/per-invocation/bulk/concurrent/resource claims;
[`0094`](changes/0094-cfb-selective-read-evidence.md) retains the non-simulated
exact-range result, and the latest accepted
generic multi-format filesystem result remains
[`0089`](changes/0089-filesystem-release-repeated-evidence.md); Change 0143 is
the latest accepted before/after CFB filesystem result)

This document records facts established by source inspection. It is not a
performance-results report. A path is called a bottleneck only after the
process benchmark and profiler evidence in `BASELINE.md` confirms its effect
on a named corpus and scenario.

## Current resource observations (change 0115)

The [current-HEAD resource profile](results/resource-profile-current-head-0115.json)
adds process-total evidence for a narrow set of named paths.  It does not
promote any observation to a production bottleneck because heaptrack includes
startup and synthetic corpus construction, while `strace` covers the whole
process.

- The managed XLSX batch run recorded 6,130,956 allocation calls and
  1,026,348,498 allocated bytes in one heaptrack process profile.  This is a
  strong candidate for operation-attribution work, not proof that the timed
  edit itself owns all those allocations.
- The OPC source one-Part profile recorded 549 logical source reads and
  16,785,201 logical source bytes, while the CFB save profile recorded 1,825
  logical reads and 84,838,500 bytes before publishing 16,913,408 bytes.  The
  CFB read/output ratio is a concrete measurement target; it is not a physical
  disk-I/O claim.
- The existing bounded RTF stream retained zero output bytes and a 37-byte
  authoring window in the harness.  Heaptrack still observed 450,852 process
  allocation calls. Change 0256 now brackets each timed RTF creation sample
  with aligned allocator/procfs observations while keeping corpus setup and
  correctness checks outside the interval. A pinned release resource report
  is still required before treating those counters as accepted evidence or
  changing the streaming path.
- Explicit 1/2/4/8/available execution-context runs on many-small OPC and CFB
  corpora were classified `nonideal_or_measurement_noise`: raw p50 showed no
  measured speedup and out-of-range Amdahl fractions are invalidated rather
  than treated as serial-fraction estimates.  The result supports investigating
  task granularity and serial work, but does not justify adding parallelism or
  changing the execution API.

The CFB read-amplification breakdown is now captured and its bounded
fingerprint-request hypothesis is accepted in Change 0143: logical bytes stay
84,838,500 while calls fall from 1,825 to 857, with both clean ABBA directions
improving warm and advisory-cold p50/p95/mean. The next evidence-oriented
priorities are operation-scoped allocation profiles for CFB and managed XLSX,
block-backed physical-cold/high-latency CFB evidence, and a CPU-pinned repeated
scaling run with uncertainty. None should be treated as an optimization
acceptance gate until matched controls and preservation gates exist.

## CFB fingerprint read coalescing (change 0143)

Complete CFB overlay fingerprints now use a right-sized request window capped
at 1 MiB; comparison and publication remain at 64 KiB, the buffers do not
overlap, and no fingerprint or stable-token validation stage is removed. A
clean CPU-2 `A1, B1, B2, A2` release run with 200 samples per warm and
advisory-cold state reduces exact logical requests 53.0411% (1,825 -> 857) with
unchanged logical bytes, output hash and one-span publication. Warm p50 improves
3.3327%/1.3163% and advisory-cold p50 10.7679%/9.4641%; p95 and mean agree in
both directions. A matched whole-process RSS boundary found no candidate
increase, but operation-only allocation/peak memory, physical I/O and proven
cold-storage claims remain open.

## ODT repeated source-backed full-text projection (change 0180)

The production threshold-two cache removes two of four complete semantic text
projections for the exact prepared `SourceBackedDocument` workload. It retains
at most one 16 MiB string, returns a fresh owned string per call, refuses
retention on size/allocation pressure, retries parse errors, and checks source
freshness around publication and cache hits. Concurrent initial publishers may
duplicate the bounded parse; this is safe but not single-flight. Two clean
CPU-2 A/B/B/A cycles
accept p50 reductions of 47.01%-50.95% and mean reductions of
46.83%-51.29%. First-cycle candidate p95/p99 drift fails policy, so tails are
withheld despite the clean retry. Every replay records zero post-preparation
reads; allocation/RSS, physical I/O, single-call/open, producer, and broader
ODF evidence remain open.

## XLS immutable source policy reuse (change 0181)

`Snapshot::from_bytes` already creates a complete validated public Workbook
model before retaining the private fixed-width BIFF inventory. The plan-only
numeric path formerly reopened the same immutable source to repeat worksheet
coverage, protection, and macro checks. Those content-free facts now travel
with the snapshot; target semantic validation, CFB verification, fingerprints,
and publication are unchanged.

The exact Number workload passes all clean total and commit
p50/mean/p95/p99 gates, with 1.92%-5.91% lower total p50 and 3.95%-8.27% lower
commit p50 across paired directions. RK/MulRK is directionally lower but fails
stability, so its latency is withheld. The next native-XLS work should target a
measured larger owner/candidate path rather than remove additional validation
fences speculatively.

## ODP repeated full-text projection (change 0140)

The production threshold-two cache removes two of four complete semantic text
projections in the matched `SourceBackedPresentation` selector shape. A clean
CPU-2 `A1, B1, B2, A2` release run accepts p50 reductions of 45.80%/46.32% and
p95 reductions of 45.25%/45.83%; p99 and mean agree. Whole-process Heaptrack
allocation calls fall 14.31% and temporary allocations 17.25%, but peak heap
is unchanged at 89.22M and process VmHWM is near-neutral. The prepared-source
replay performs zero post-preparation reads, so this is parse/projection/cache
work rather than physical-I/O or decompression evidence. Broader slide-object,
single-call, open, edit/save, real-producer, and generic ODF work remains in the
ODF queue.

## Rejected XLSX publisher provenance reuse (change 0141)

A private lineage/version fast path was tested across calculation metadata,
defined names, page breaks, page margins, page setup, print options, and sheet
protection. It skipped the publication-time semantic reload but left the raw
ZIP publication path unchanged. Clean CPU-2 `A1, B1, B2, A2` evidence found a
1.04% regression in the pooled seven-case p50 geometric mean; calculation
metadata regressed 3.84% p50, and paired directions were mixed. Whole-process
allocation calls fell only 2.84%, temporary allocations 2.12%, peak heap was
unchanged, and VmHWM moved less than 1%. The production change was fully
reverted. Future XLSX publication work should target physical output work,
broad graph validation, whole-Part reconstruction, or a materially larger
semantic parse instead of reintroducing generic provenance fields to these
seven snapshots.

## Shared OOXML data path

```text
path / Read / Vec / &[u8]
  -> litchi-opc physical reader
     -> complete source Vec for path and generic Read ingress
     -> soapberry-zip central-directory index
     -> content types and package relationships
     -> relationship-graph validation
     -> classify every physical member
     -> decompress every admitted Part
  -> OpcPackage
     -> HashMap<PackURI, Box<dyn Part>>
     -> second source-XML index
  -> DOCX / PPTX / XLSX mandatory catalog
  -> lazy format-owned semantic parse of a selected Part
  -> Edit plan and dependency validation
  -> candidate Part reconstruction and readback
  -> PackageWriter
     -> exact owned-source copy when no mutable API was entered
     -> otherwise regenerate manifests and relationship Parts
     -> build, audit, and retain one deterministic publication plan
     -> Deflate every Part into a sequential sink
```

Current work shape:

- Legacy path and generic-reader OPC ingress still has a contiguous-buffer
  path, while source-backed ingress uses an immutable positional source with
  source versions and a validated ZIP index.
- `PackageReader::load_parts_eager` is not physically lazy: it classifies and
  decompresses every admitted Part, including unreferenced Parts that must be
  preserved.
- Ordinary bulk opens are serial. Explicit eager opens opt into local bounded
  ZIP sessions through `litchi-core::ExecutionContext` and OPC `OpenSession`;
  there is no hidden global Rayon pool.
- `OpcPackage` retains every inflated Part. XML Parts also participate in a
  second source-XML map. Part lookup has an exact hash-map fast path and a
  linear ASCII-case-insensitive fallback.
- Exact unchanged owned OPC output reuses the complete source archive. Owned,
  same-topology mutation now retains private provenance and raw-copies every
  semantically unchanged ZIP member; topology changes, borrowed ingress, and
  unsupported ZIP layouts still use the complete rewrite.
- The immutable source-backed package now also has one consuming low-level
  same-topology publisher. It accepts at most 64 unique selected existing
  Parts, validates/materializes that bounded set, regenerates only changed
  selected members, raw-copies every other physical member, and
  monitors source version through bounded sequential output. Signed real
  changes and unsupported layouts return typed zero-output refusals. DOCX now
  exposes a guarded exact-source main-document transaction over that publisher:
  raw-MCE identity and main-Part-only operations are required, transfers are
  refused. PPTX now exposes the analogous guarded exact-source selected-slide
  transaction: its raw package/presentation/slide relationship closure is
  bound into the snapshot, MCE-rewritten slides and more than one shape edit
  per selected slide are refused. A bounded outer batch now composes up to 32
  exact slide snapshots into one atomic multi-Part publication. XLSX now has a
  narrower guarded transaction for typed
  calculation properties/features or the direct defined-name catalog in
  `xl/workbook.xml`; cells, formulas, chains and topology remain outside those
  capabilities. Selected-worksheet
  variants now bind the workbook relationship and worksheet owner for direct
  typed page breaks, page margins, print options, relationship-free page
  setup, complete sheet-protection metadata, or typed core/Office 2010 data
  validations; all materialize two Parts and refuse wider worksheet/topology
  changes. Auto-filter and core conditional-formatting variants additionally
  bind styles/DXF state and materialize three Parts.
- `PackageWriter` previously reconstructed generated XML and Part order during
  emission. The measured `PublicationPlan` change now constructs, audits, and
  reuses that state once. It reduced allocation calls by 37.0% in the profiled
  256-Part save and mean latency by 5.49% in the 2,048-Part compressible save;
  full-Part recompression remains unchanged on the fallback path. Targeted raw
  publication separately improves p50 by 58-96% across the synthetic cells,
  while retained-source peak heap originally rose 37%; see change 0008. The
  changed Part now shares its existing immutable logical payload with the ZIP
  regeneration layer, removing one measured 4.19 MiB copy and reducing the
  matched peak by 3.42%; see change 0021. After validation, the ZIP layer also
  moves that entry's generated local span instead of cloning it, removing a
  second 4.20 MiB allocation and reducing matched peak heap another 3.20%; see
  change 0022. The source-backed one-Part path then removes three unselected
  Part materializations/recompressions on the four-Part corpus, reducing p50
  73.12%, instructions 65.42% and peak heap 3.20%; see change 0037. Complete
  physical archive input/output and the selected-Part compressor buffer remain.
  The DOCX facade integration then removes eager ownership and recompression of
  16 unselected Parts in the media-rich one-edit/save case: p50 falls 97.43%,
  instructions 74.91%, and semantic materializations 17 -> 1 while the eager
  DOCX guard remains neutral; see change 0039. The PPTX facade integration
  removes eager ownership and recompression of 227 unselected Parts in the
  fixed media-rich one-slide edit/save case: p50 falls 97.12%, instructions
  67.91%, and semantic materializations 229 -> 2 with byte-identical output;
  see change 0044. The atomic eight-slide follow-up regenerates eight selected
  slide members in one plan, cuts p50 95.78%, allocations 32.54%, peak heap
  8.94%, and materializations 229 -> 9 while preserving byte-identical output;
  see change 0077. The XLSX calculation-metadata integration removes eager
  ownership and recompression of 11 unselected Parts in its fixed media-rich
  edit/save case: p50 falls 99.2519% (133.67x), instructions 77.78%, and
  semantic materializations 12 -> 1 with byte-identical output; see change
  0046. The defined-name variant also materializes only the workbook Part and
  cuts p50 97.84% (46.32x), instructions 78.45% and materializations 12 -> 1;
  see change 0076. The selected-worksheet page-break, page-margin, print-options and
  relationship-free page-setup, sheet-protection and data-validation
  variants each materialize only the workbook catalog and target worksheet; on
  their matched media-rich controls, p50 falls 97.86%, 97.93%, 97.87%, 97.78%,
  97.75% and 97.75%, respectively; see changes 0061, 0067, 0070, 0073, 0078
  and 0079.

The managed XLSX source-editor freeze in [change 0151](changes/0151-xlsx-managed-source-editors.md)
extends the same selected-Part boundary to calculation properties, defined
names, tab state, print options, page breaks, page margins, page setup, sheet
protection, data validation, auto filter, and conditional formatting. A private
`Managed(PartData)`/`Owned(Arc<Vec<u8>>)` payload owner keeps managed cache
reservations attached; managed-to-owned `Arc` escape is typed and fallible.
Managed constructors hand off a checked `SourceBackedPackage`, and direct
publication materializes only the proven selected Part(s) while raw-copying
unselected members. Exact no-op/signed, MCE/unknown-owner, stale, cancellation,
and one-byte-under Budget gates remain. This is correctness/resource-accounting
evidence only: no latency, allocation, RSS, copied/decompressed-byte, cold-I/O,
or total-memory claim is attached.

The committed managed source-cache change (`f8d417ac3`) charges exact
physical `InputBytes`, cumulative declared cold-load `Work`, retained
catalog/flight/payload `Objects`, and retained/in-flight payload `Memory` to the
caller's hierarchical `Budget`; compatibility constructors retain the finite
unmanaged `SourceCacheLimits` behavior. Focused correctness tests cover these
resource charges, retained-resource releases, flights, waiters, pinning, eviction,
cancellation, sibling competition, and contention invariants. Release
contention accepts no managed-versus-control speedup; allocation/peak-memory/
RSS, hardware, copied/decompressed-byte, CPU-utilization, and production-performance
evidence remain open.

## XLSX selective read and edit path

```text
whole-package OPC materialization
  -> workbook catalog and relationship parse
  -> worksheet handles with OnceLock<Store>
  -> first cell/range query
     -> parse the complete selected worksheet XML
     -> materialize and sort the complete sparse Store
  -> targeted edit commit
     -> compare against complete Store
     -> scan complete lossless worksheet layout
     -> allocate/copy complete replacement worksheet XML
     -> compact and reparse complete replacement for publication proof
     -> clone shared OPC graph and replace changed Part
     -> hand off the validated Store only below the cell/XML retention bounds
        and only when final Part and style/shared-string identities still match
  -> save
     -> recompress complete package
```

Confirmed source facts:

- `litchi-xlsx` now exposes bounded forward-only creation for one-sheet
  workbooks. Change 0169 adds scoped warm in-memory tiny/medium/large latency
  evidence and descriptive whole-process allocation profiles for the exact
  inline-scalar corpus. Operation-local allocation, total/peak-memory and RSS
  attribution, physical/cold I/O, richer authoring and producer evidence remain
  pending. This is distinct from the source-backed existing-cell publication
  result below.
- The legacy eager path still materializes all admitted Parts. The additive
  source-backed XLSX facade avoids timed source reads while listing after open;
  managed source-backed OPC caches charge exact physical `InputBytes`,
  cumulative declared cold-load `Work`, retained catalog/flight/payload
  `Objects`, and retained/in-flight payload `Memory` to a caller's hierarchical
  `Budget`, preserve externally pinned handles, and coordinate same-Part cold
  loads through one flight. Compatibility opens remain finite under the
  unmanaged `SourceCacheLimits` path. Correctness tests cover these resource
  charges, retained-resource releases, budget hierarchy, eviction, pinning, sibling
  competition, cancellation and failure; the release contention ABBA adds
  structural/distribution evidence but accepts no speedup. Allocation,
  peak-memory/RSS, hardware, copied/decompressed-byte, CPU, and
  production-performance evidence remain missing.
- Managed direct `SourceBackedPackage` sequential publication now charges
  `Resource::OutputBytes` per sink write and commits only exact accepted bytes;
  exact/no-op and changed overlays retain typed refusal, partial-output,
  cancellation and source-freshness behavior. This is a bounded correctness
  accounting change only and excludes `OpcPackage` atomic saves, `to_bytes`,
  and unmanaged compatibility sinks; no performance result is claimed.
- The additive source-backed calculation-metadata editor loads only the
  workbook Part, stages existing typed `calcPr`/feature edits, reparses the
  complete candidate workbook XML, and consumes the commit into the accepted
  one-Part publisher. It recaptures owner/content-type/URI/XML/source-version
  identity before output; MCE projection and changed signed sources refuse.
  The fixed eight-media case improves p50 99.2519% and materializations 12 ->
  1. This does not authorize cell, formula, cached-result, relationship or
  calculation-chain edits.
- The additive source-backed defined-name editor likewise loads only the
  workbook Part. It binds the exact workbook owner/XML and ordered sheet
  catalog, validates global/local name scope, reparses the complete candidate,
  and refuses protected or MCE/unknown catalogs and changed signed sources.
  The media-rich control improves p50 97.84% and materializations 12 -> 1.
- The additive source-backed sheet-protection editor binds the exact workbook,
  selected worksheet and complete outbound worksheet-relationship set. It
  atomically replaces the complete direct core/Office 2010 protection state,
  reparses the result and refuses MCE-selected protection or changed closure.
  The media-rich control improves p50 97.75%, instructions 77.87%, and
  materializations 12 -> 2.
- The additive source-backed data-validation editor binds the same exact
  workbook, selected worksheet and complete outbound relationship closure. It
  atomically replaces complete typed direct core/Office 2010 collections,
  consumes checked post-write readback, and refuses MCE-selected collections
  or changed closure. The media-rich control improves p50 97.75%, instructions
  73.43%, and materializations 12 -> 2; allocation calls remain within policy.
- The additive source-backed auto-filter editor additionally binds the styles
  relationship and differential-format count so value/color/DXF filters and
  sorts cannot publish dangling style references. It replaces one direct
  worksheet filter/sort subtree, refuses MCE-selected or protected state, and
  raw-copies all unrelated Parts. The media-rich control improves p50 97.75%,
  instructions 73.57%, and materializations 12 -> 3.
- The additive source-backed conditional-formatting editor reuses that exact
  workbook/worksheet/relationship/styles closure and atomically replaces the
  complete direct core owner collection. Its matched selectable case uses the
  same typed values and worksheet rewriter in eager and source-backed paths,
  proves byte-identical publication, and records materializations 12 -> 3.
  Balanced ABBA evidence has not yet been retained, so this is not a latency,
  instruction, or allocation result.
- Change 0151 freezes managed ownership for all eleven focused source-backed
  editors listed above, including tab-state publication and the conditional-
  formatting closure. Representative one-byte-under `Resource::Memory`,
  cancellation, typed Arc-escape, exact no-op/signed, stale/foreign, and
  unknown-owner preservation/refusal checks pass. Parsed stores, staging,
  rewritten candidates, and sink buffers remain outside that accounting; no
  performance or total-memory claim is made.
- In the eager path, and for source-backed inputs that are ineligible for
  selected streaming or take the unsupported-structure fallback, one first
  cell access parses the entire selected worksheet. The non-evicting `OnceLock`
  retains it for the snapshot lifetime. Eligible source-backed scalar/range
  queries use the streaming scanner; after Change 0400 this includes
  worksheets with validated `dimension` metadata.
- The sparse cell store is row-major and supports binary-search point lookup.
  A compact immutable row-start index now skips preceding rows for narrow
  ranges. The measured range query improves about 80%; full scan and first-cell
  guardrails remain near neutral.
- A targeted cell edit performs a semantic parse, an independent lossless
  layout scan, full replacement-byte construction, and a full changed-sheet
  semantic readback before publication.
- Source-backed scalar-cell publication now carries a tri-state source
  provenance proof from the checked snapshot into the publisher. Matched
  lineage/version avoids a second publication-time semantic worksheet reload;
  mismatched sources refuse and unavailable provenance retains the prior full
  reload/readback path. Balanced release ABBA across one-cell, `ceil(1%)` and
  exact-256 batches accepts p50 geomean improvements of 21.66%/22.65% and p95
  improvements of 21.38%/22.70%. Physical source reads and successful
  materializations are unchanged, so this is a semantic-reparse result rather
  than an I/O claim. See [`change 0096`](changes/0096-xlsx-source-provenance-publication.md).
- Change 0163 adds four opt-in eager/source-backed scalar-cell lifecycle
  selectors over the existing medium and dense/sparse numeric four-sheet
  corpora. The eager `WorksheetEdit` and positional source editor each clear or
  remove `Sheet1!A1`; clear retains an empty owner and remove deletes it. Open,
  plan/stage, commit, publication and lifecycle phases are separate, and a
  fixed 64-KiB hashing sink retains zero output. Generic logical source and
  materialization counters, semantic/package/no-op/volatile-patch/stale/
  foreign gates, and source-backed raw preservation of unselected members are
  recorded outside timing. This is correctness/phase/counter evidence only;
  the source-backed patch has no durable wire contract and no latency,
  allocation/RSS, physical-I/O, cold-cache, decompression or real-producer
  claim is made. See [`change 0163`](changes/0163-xlsx-cell-clear-remove-evidence.md).
- Plain worksheets previously ran a separate namespace-aware x14ac collection
  before every complete semantic parse even when no `dyDescent` token existed.
  Successful no-token reads now skip that pass; rejected inputs rerun it to
  preserve error precedence. Medium changed commits improve about 20% and cold
  reads about 35%; dense-wide 1% commit improves 19.62% p50, allocation calls
  fall 25.24%, and peak heap remains flat. Direct x14ac/MCE paths are unchanged.
- Eligible changed sheets now adopt that exact commit-validated store into the
  target snapshot. Medium commit plus first read improves 23.23% p50 and
  allocation calls fall 21.01%. The handoff is capped at 4,096 cells and 1 MiB
  XML; the unrestricted dense-wide prototype was rejected at +8.99% peak heap.
- Bulk cell actions are held in address order, then regrouped into nested
  row/cell `BTreeMap`s during worksheet emission. A direct owned-stream
  replacement removed that regrouping but improved formal 1% commit/save by
  at most 1.61% p50, so it was fully reverted in change 0030.
- An empty edit returns the original immutable workbook allocation. When the
  workbook came from owned ingress, saving that no-op snapshot now preserves
  the exact validated OPC source; borrowed ingress still performs a rewrite.

## DOCX and PPTX paths

DOCX format views are borrowed after eager OPC materialization. Repeated
`paragraphs`, `tables`, and `blocks` queries rescan and allocate result vectors.
Single-index paragraph lookup now scans the complete bounded XML but constructs
only the selected shared range; the 10,000-paragraph cell improves 4.72% p50
and removes ten collection-growth allocations per call. Table lookup still
builds the complete collection. Canonical direct-body paragraph batches now
plan every replacement against one snapshot, emit the disjoint ranges in one
forward pass, parse one candidate, and read back every selected paragraph. The
10,000-paragraph / 100-edit save improves 94.99% p50 (19.97x) and allocation
calls fall 94.11%. Scalar edits, unordered/nested selections, structural edits,
and complete transaction-capture costs are unchanged.

Change 0263 adds two opt-in end-to-end story-hyperlink publication selectors.
The fixed corpus covers main, header, footer, footnotes, endnotes, comments,
and glossary stories, with one selected shared target, one unselected target,
media, and an opaque member per preservation path. The no-op requires exact
archive bytes; redaction requires the exact story XML/`.rels` changes and raw
ZIP local plus offset-normalized central identity for untouched members. Stale,
foreign, signed, unknown-owner, partial-sink, and zero-sink refusals are
checked. Source/sink preparation and all independent oracles are outside the
open + plan + commit + sequential-publication timer. This is correctness and
phase evidence only, not a speedup, allocation, RSS, physical-I/O, or broad
DOCX claim.

Change 0265 adds two opt-in whole-slide boundary publication selectors over a
45-member, 32,396-byte, four-dependency-free-slide PPTX corpus. Removal proves
first/middle/last positions plus final-only refusal; move proves both boundary
directions and the exact `from == to` no-op. Production opened-presentation
`Snapshot`/`Transaction` plans and commits are timed as separate plan, commit,
sequential-publication, and reopen phases. Semantic reopen, twice-built
determinism, source immutability, durable serialized forward/inverse patches,
stale/foreign, dependency, unknown-member, MCE, signed, limits, partial, and
zero-sink gates remain untimed. Untouched raw local and normalized-central
records must match; move requires strict `[Content_Types].xml` identity. This
is correctness and phase evidence only, not a latency, allocation/RSS,
physical-I/O, or broad PPTX claim.

PPTX ordinary reads defer slide payload parsing, but repeatedly parse the
presentation slide-reference list. Exact-name slide lookup resolves and parses
all candidate slide names. The opened-transaction snapshot is deliberately
stronger and more expensive: it resolves every slide, notes graph, Part and
relationship fingerprint, and retains a cloned shared OPC graph. Commit
recaptures and re-fingerprints the candidate after readback. Shape-text edits
now reuse the selected scene when mapping its raw span, removing one redundant
scene parse per change. The 100-edit cell improves 9.37% p50/mean and allocation
calls fall 11.67%; the single-edit end-to-end guardrail remains neutral because
complete capture/commit work dominates.

The additive source-backed PPTX editor instead snapshots one selected slide
and its exact package/presentation/slide relationship closure. One operation
may replace one shape or atomically replace up to 256 unique, nonoverlapping
shape texts in a single bounded scan/emission. It consumes the commit into the
source-backed one-Part publisher. The other 199 slides, all eight 2 MiB media
Parts, and every other unselected physical member remain on the raw-copy path.
MCE preprocessing that changes raw slide bytes, duplicate/overlapping batch
selectors, stale or foreign patches, topology changes, and changed signed
sources are refused before publication. The original one-shape case improves
97.12% p50; the matched eight-shape batch improves 97.45%, reduces allocation
calls 39.80%, and retains the 229 -> 2 materialization reduction.

These paths have strong preservation and atomicity tests plus generated-text
timing/allocation evidence. Real-producer, media/dependency, malformed,
security, copied-byte and cold-source matrices remain missing.

Change 0120 adds eight filesystem-isolated ordinary-root PPTX controls over
the 200-slide/eight-text-box/eight-2 MiB-media corpus. The source candidate
uses `litchi::Presentation::open(path)` and the eager control uses a prepared
byte root for query phases; `list_slides` materializes all owned slides and
`selected_slide` uses the selector-first `Presentation::slide(100)` API. A
separate untimed source replay classifies exact compressed ZIP payload-range
overlap: open/count are catalog-only, selected reads only slide 100, and list
reads all slides without media. This establishes a useful correctness and
logical-read guard for the unified facade. Change 0188 adds fresh-open-plus-
count/selected-slide lifecycles over the same corpus. Its final-source warm
ABBA retains no latency claim because same-implementation drift gates fail;
tails, allocation, RSS, decompression, physical-I/O and cold-cache remain
withheld. Eager controls explicitly have no source replay.

## ODF paths

ODS and ODP ordinary opens eagerly read and parse their ZIP packages. The
high-level ODT filesystem path now retains a source-backed package and owner;
byte-backed ODT opening remains eager.
The opt-in public semantic matrix now measures owned open, listing, one object,
full text, small creation, exact no-op and one supported edit/save across all
three owners. ODT indexed paragraph lookup still scans complete XML for
validation, but retains only the requested paragraph. ODP indexed slide lookup
likewise validates styles and content through EOF while retaining semantic text
and completed shapes only for the requested slide; repeated independent ODP
queries still rescan both XML inputs.
ODP content-only rich-object operations and ODT content-only paragraph
replacement, line-break, inline-run, hyperlink, insertion, and removal
operations now reuse checked raw preservation. On the fixed eight-by-2 MiB ODT
corpus, paragraph edit/save p50 falls 95.58%; the
matched line-break path falls 98.17% (54.59x), instructions fall 78.34%, and
allocation calls fall 6.90%. The matched inline-run path falls 98.39% (62.01x),
instructions fall 78.48%, and allocation calls fall 7.00%, with flat peak
heap/RSS. Structural insertion/removal fall 98.20%/98.27% p50 (55.55x/57.86x)
with exact member preservation. Oversized ODT content and resource-adding,
new-style, or richer structural ODF publication retain the established rebuild.
Generic packaged ODF chart-definition replacement now uses the same raw
publisher with an opt-in full payload preflight, retaining the former logical
writer's malformed-member rejection while preserving eligible unchanged ZIP
frames. Existing shared ODT/ODS/ODP raw paths remain lazy. This generic
integration is correctness-only pending matched release evidence; see
[`change 0101`](changes/0101-generic-odf-verified-raw-publication.md).

Existing ODT embedded-resource replacement now has matched selectable evidence
for 64 fixed existing package-backed image owners. The scalar control repeats
`replace_embedded_image` 64 times in one transaction; the bounded batch resolves
the same base-snapshot positions and publishes them through one
`edit_embedded_resources` call. Both reopen to the same complete
paragraph/image projection and retain exact frame names, paths, media types,
payload digests, retained media and untouched raw ZIP members. Case-specific
physical hashes are recorded without requiring scalar/batch byte identity. ODT
exposes no positional-source or logical-Part materialization diagnostics; the
record reports real bounded sink counters only and makes no performance claim
before frozen CPU-pinned balanced ABBA evidence. See
[`change 0085`](changes/0085-odt-embedded-resource-batch-evidence.md).

ODT one-paragraph lookup now has an additive public indexed selector. It keeps
the complete namespace-aware, resource-bounded EOF scan while retaining one
paragraph rather than the 10,000-paragraph collection. Large middle-paragraph
p50 falls 48.56%, allocation calls fall 27.05%, peak heap falls 24.74%, and
uninstrumented RSS falls 10.93%. The established list path remains neutral; a
shared-mode prototype that regressed it was removed. See
[`change 0047`](changes/0047-odt-indexed-paragraph-selector.md).

ODP one-slide lookup now uses a compile-time-specialized selector so the
established full-list parser does not carry a runtime mode. Large middle-slide
p50/mean/p95 improve 4.09%/4.20%/5.18%, whole-process allocation calls fall
3.86%, and the list, full-text, no-op, edit/save and media-save guards remain
within thresholds. Style inheritance, namespaces, shape/animation limits and
tail errors are still checked before return. See
[`change 0049`](changes/0049-odp-indexed-slide-selector.md).

ODP editing snapshots now pass their already validated slide projection into
private transaction staging instead of parsing every slide again from the same
immutable package bytes. Package/security reopening, settings, declarations,
page metadata, raw source-page coverage, isolated draft clones, changed
publication and complete final reopen/readback remain. Large exact no-op
edit/save improves 59.96% p50 and large changed edit/save improves 20.78%;
allocation calls fall 20.13% with flat peak heap/RSS. See
[`change 0060`](changes/0060-odp-snapshot-slide-projection-reuse.md).

Exact slide-only commits now keep the already mandatory parsed candidate until
final publication and move that projection into the immutable snapshot instead
of parsing the same bytes a second time. The independent final package reopen,
raw/compact XML audits and staged-media check still run; any RDF, chart, design,
annotation or rich-content operation retains the ordinary final parse. Large
one-slide edit/save improves 32.35% p50/32.92% mean, allocation calls fall
16.71%, and peak heap/RSS stay flat. See
[`change 0065`](changes/0065-odp-final-snapshot-handoff.md).

Direct ODT transaction snapshots now adopt the exact package allocation
created by validation and share it with staging rehydration. This removes two
complete archive copies while retaining both complete semantic parses. On the
same media-rich paragraph case, p50 falls 75.84% and peak heap/RSS remain flat;
the compactness audit formerly retained further archive-sized copies.

The changed-operation compactness audit now clones the validated predecessor's
private immutable package and borrows the validated candidate package. This
removes three complete archive copies (50.36 MB on the fixed media-rich case)
without removing archive/manifest parsing or compact XML/splice validation.
Edit/save p50 falls 30.44%, mean 31.36% and p95 32.41%; allocation calls fall
0.57% and peak heap/RSS remain flat. Final transaction materialization,
envelope classification and independent reopen/readback remain. See
[`change 0041`](changes/0041-odt-compact-audit-package-sharing.md).

Envelope classification now clones the immutable snapshot package handle
instead of allocating/copying another complete archive. ZIP validation and
manifest/signature/encryption inspection still run. Across two balanced ABBA
cycles on the same media-rich case, p50 falls 11.40%, mean 11.95%, and p95
12.19%; Heaptrack removes exactly two allocations per changed commit with flat
peak heap/RSS. The independent reopen/readback remains.
See [`change 0042`](changes/0042-odt-envelope-package-sharing.md).

Final changed-result publication now clones the already validated document's
private immutable package bytes into the byte-only snapshot. This removes one
16.79 MB copy and one redundant parse while retaining a fresh complete
`after.document()` reopen. Media-rich edit/save p50/mean/p95 improve
22.74%/22.56%/21.48%; allocation calls fall 3.46%, and peak heap/RSS remain
flat. The earlier parsed-final-document retention stays reverted; the guarded
medium one-paragraph path remains within 3% p50/mean and improves p95. See
[`change 0052`](changes/0052-odt-final-result-byte-handoff.md).

ODS unified snapshot construction previously cloned package bytes and parsed
the same ODS package twice: once for package/resource validation and again for
complete `Spreadsheet` readback. It now moves the one validated package into a
crate-private facade constructor. Large no-op edit/save p50 falls 11.78%; the
large changed case improves 2.06% because full spreadsheet rewrite/readback
dominates.

Eligible same-topology worksheet commits now reuse the bounded flat-ODS row
splicer: only changed modeled rows are serialized and untouched source spans
are copied exactly. Large/medium one-cell edit-save p50 falls 9.54% / 7.22%,
allocation calls fall 5.85%, and peak heap falls 27.18%. Structural changes
fall back to full-table replacement; an opaque untouched row is preserved
byte-for-byte, while touching it refuses publication. Compactness, package
reopen, snapshot parsing and complete typed-sheet readback remain mandatory.

The row-local editor now carries its exact checked source ranges through
package emission instead of flattening them and asking the package layer to
rediscover one maximal diff. On the fixed 2,048-cell plus 16 MiB-media case,
this avoids the full-package fallback that recompressed unchanged media.
Edit/save p50/mean/p95 improve 74.16%/74.17%/74.11%; instructions fall 69.04%
and matched peak heap/RSS remain flat. Foreign provenance and unexpected
assembled content refuse. Signatures, encryption-sensitive inputs,
unsupported ZIP layouts, structural edits and every unproved case retain the
established logical rebuild/signature policy. See
[`change 0057`](changes/0057-ods-row-splice-raw-publication.md).

The remaining unified-to-worksheet path formerly copied the exact archive at
each ownership boundary even after row publication stopped recompressing
media. Worksheet snapshots and patches now retain `Arc<Vec<u8>>`, the private
ODS package adopts that owner, and the unified worksheet handoff moves its
source and target allocations through validation with exact failure rollback.
On the same media-rich case, p50/mean/p95 improve
21.32%/21.30%/21.15%; peak heap falls 22.03% and uninstrumented RSS 20.57%.
The durable unified patch boundary and other semantic domains are unchanged.
See [`change 0068`](changes/0068-ods-shared-worksheet-archive-handoff.md).

A media-rich ODS publication case now adds eight deterministic 2 MiB opaque
resources. Eligible compact `content.xml` replacements raw-copy every other
validated ZIP member; exact local/central-member comparison skips unchanged
payload inflation only when the manifest is also exact. The media-rich
one-cell edit/save falls 4.73% p50, 5.73% mean and 7.65% p95, with peak heap
down 8.78%. The existing medium no-media p50 falls 0.77%. Encryption,
signatures, unsupported layouts and every unproved member retain established
logical rebuild/comparison. See
[`change 0031`](changes/0031-ods-unchanged-media-preservation.md).

ODS durable-patch construction formerly copied both exact package archives
into semantic blob bundles even though the outer patch already retained the
same immutable `Arc<[u8]>` owners, then hashed both packages again for
operation preconditions. The bundles now retain those existing allocations
and the preconditions reuse their content addresses. Media-rich one-cell
edit/save p50/mean/p95 improve 8.80%/9.07%/13.85%; the 33.58 MB payload-copy
site disappears and matched peak heap falls 1.92%. ZIP publication,
comparison, compact audit, final reopen and media verification remain. See
[`change 0054`](changes/0054-ods-shared-durable-patch-blobs.md).

ODS content-validation catalog CRUD is separately correctness-covered and
unmeasured. The clone-staged owner supports add/set/update/same-name
replace/remove/clear/rollback, exact no-op and source-checked reversible patch;
the unified document transaction publishes only `content.xml`, raw-preserves
untouched members and fully reopens the result. Referenced removal/clear,
unrepaired dangling references on changed commit, duplicate names, unsafe
rename, opaque/MCE/DTD owners, operation/output bounds and changed signed
packages refuse atomically. This exact closure does not establish a
performance hotspot or broader ODS cell/formula/style/structural capability.

The format-owned validation tranche now has bounded DOCX, PPTX, RTF and XLS
semantic reports in addition to the CFB, OPC and ODF reports. These paths are
finite correctness boundaries, not profiled hotspots. ODF repair remains one
typed non-destructive plan for removing a recognized local-header extra from a
first stored `mimetype` member. One opt-in selector now exercises its bounded
preflight, exact forward/inverse, refusal and zero-retained-output publication
contract, but supplies no latency or total-memory claim. Encrypted, signed,
macro, structural and semantic repairs refuse rather than widening the
preservation boundary.

A matching media-rich ODP case now adds one source-backed text box beside
eight deterministic 2 MiB opaque resources. Reusing the same accepted common
checked-splice/raw-copy primitive cuts edit/save p50 94.44%, mean 94.43%, and
p95 94.29%; allocation calls move +0.52% and peak heap/RSS stay flat. Exact
patch/inverse behavior, complete slide/rich-content/media readback, and every
common security/layout fallback remain. Resource-adding operations still use
the complete rebuild. See
[`change 0034`](changes/0034-odp-unchanged-media-preservation.md).

Existing ODP whole-model replacement now has matched selectable evidence for
eight fixed-name text boxes distributed across eight of 12 slides. The scalar
control repeats candidate staging eight times in one transaction; the bounded
batch resolves and publishes the same set once. Both reopen to the same full
slide/text/rich-content projection and retain exact auxiliary/media payloads.
The batch raw-preserves the manifest, while repeated scalar staging regenerates
it, so their physical output digests differ. ODP exposes no positional-source
or logical-Part materialization diagnostics; the record reports real bounded
sink counters only and makes no performance claim before frozen CPU-pinned
balanced ABBA evidence. See
[`change 0084`](changes/0084-odp-cross-slide-text-box-batch-evidence.md).

Change 0122 adds four opt-in matched selectors over the same 12-slide/eight-
2 MiB `Pictures/` ODP corpus: eager/source-backed open and eager/source-backed
one-middle-slide query. Source timing uses an uninstrumented `OwnedSource`,
while each measured sample has a separate `InstrumentedSource` replay for
exact calls, bytes, coalesced prior-range overlap, and compressed Pictures
overlap (`pictures_read_compressed_range_bytes`), distinct from prior-read
overlap (`source_read_range_overlap_bytes`). Open and one-slide query remain
distinct from a further explicit selected-media replay, which must cover one
complete selected compressed Pictures range and reports bytes outside
Pictures. The summary names compressed ZIP range totals separately from
uncompressed payload bytes/digests; the eager one-slide parity and selected
media checks run outside its timed query. A ZIP-tail catalog request may
physically touch the final Pictures range during open; that overlap is
retained as physical-range evidence and is not treated as media
materialization. Full eager/source semantic parity and deterministic media
digests remain outside timing. The selectors bring the matrix to 233 names
while leaving the default 36 cases / 198 records unchanged. This is
correctness/logical-read evidence only; no latency, decompression, allocation,
RSS, or release-ABBA claim is made.

Change 0123 adds four opt-in unified-root ODP filesystem selectors over the
same media-rich fixture: eager/source-backed open and eager/source-backed
middle-slide query. A temporary corpus is created and written before the
measurement; open timing covers only matching root owner construction, while
query timing covers only an already-open root query. Post-timing gates compare
full root semantics and metadata, source archive/member/hash identity, and
selected media payloads. Source controls pair each sample with a separate
direct typed `SourceBackedPresentation` instrumented replay, so catalog/query
media laziness and exact selected compressed-range coverage are evidence
domains distinct from root timing. This brings the matrix to 237 names while
leaving the default 36 cases / 198 records unchanged. Production routing
tests cover the filesystem handoff; no latency, physical-I/O, decompression,
allocation, RSS, or release-ABBA claim is made.

Change 0124 adds six opt-in ODS unified-root/source selectors over the existing
two-sheet media-rich ODS fixture: eager/source-backed root open, typed
selected-cell, and typed selected-media controls. Corpus and file publication,
eager cloning, and typed owner construction are outside the corresponding
timers. Each sample checks root names/count/text, complete cell and metadata
parity, exact source/archive/member/media identity, and typed ODS readback.
Independent `InstrumentedSource` replays report logical positional calls and
compressed-range overlap separately from uncompressed payload bytes; open
replays avoid unrelated media, selected-cell replay adds no reads after
content preparation, and selected-media evidence pairs an all-Pictures replay
with a selected-range-only replay, requiring both to cover exactly one
compressed member range and excluding other media. Eager source vectors are
empty. The six selectors bring the matrix to 243 names while leaving the
default 36 cases / 198 records unchanged. This is correctness/logical-range
evidence only, with no latency, physical-I/O, decompression, allocation, RSS
or release-ABBA claim.

Repeated `Spreadsheet::cell` scans previously linearly walked physical row and
cell runs for every coordinate. A new opt-in lookup-only sweep attributes the
cost, and the immutable facade now builds a sheet-aligned locator only after 64
successful queries. The large sweep falls 81.74% p50 and the existing
full-cell-text aggregate falls 52.65%; the dense locator requests 3,216 bytes,
is capped at 4 MiB, and peak heap/RSS stay flat. Repeated runs use cumulative
endpoints without expanding logical cells; point queries remain on the linear
path. See [`change 0027`](changes/0027-ods-adaptive-cell-locator.md).

Adopting an already parsed target package directly into the worksheet snapshot
was measured separately and fully reverted: large one-cell edit/save p50
improved only 0.44%, while p95 regressed 0.30%. Package/readback work remains a
hotspot, but that ownership handoff is not a material optimization.

ODT transaction snapshots created from an already validated `Document`
previously allocated and copied the complete package solely to establish the
snapshot owner. They now clone the package's private immutable `Arc` after the
same transaction size check. Large no-op edit/save p50 falls 18.51%, and
Heaptrack attributes exactly two fewer allocations and no package copy to each
snapshot; changed edit/save and unrelated open guardrails remain within 3%.
Direct snapshot byte ingress, full changed-package publication/readback, and
signed/encrypted envelope behavior are unchanged.

ODT full-text extraction now selects a private consuming parser mode: each
parser-created validated block string moves into its element, then into the
final text instead of being cloned at both boundaries. Repeated large-corpus
ABBA improves 3.25% p50 and 4.81% mean; process allocation calls fall 15.48%
and temporary allocations 45.52%, with peak heap and uninstrumented RSS flat.
Public structured block/list queries keep their original path and remain near
neutral. The unchanged open guard moves +3.94% p50/+4.17% mean; its +10.95%
p99 trigger is retained in change 0023. Repeated semantic scans, source-backed
reads and changed-member publication remain.

## Legacy CFB data path

```text
Read + Seek or positional `ReadAt`
  -> header and complete FAT
  -> complete directory bytes
     -> structural validation pass
     -> public entry decoding pass
  -> complete MiniFAT metadata
  -> validate every stream allocation chain
  -> semantic DOC / XLS / PPT owner
     -> lookup a child by cached validated sibling-tree keys
     -> materialize selected stream Vecs or a bounded caller-owned range
  -> edit/rebuild
     -> retain all output stream Vecs
     -> copy borrowed stream slices into OleWriter
     -> assemble MiniFAT/FAT/directory sector buffers
     -> Write + Seek output
```

Confirmed source facts:

- `SharedOleFile` provides positional CFB access and explicit bounded bulk
  operations. Four 4 MiB streams reach 5.93x p50 at 12 visible CPUs, while 256
  1 KiB streams regress at high worker counts; thresholds remain essential.
- Public `SharedOleFile::read_stream_range` now has pinned release ABBA evidence
  against legacy full-stream materialization. For the final 36-byte MiniFAT
  target, one physical source request falls from 261,184 to 36 bytes among 256
  siblings and from 2,096,192 to 36 bytes among 2,048 siblings. Read-stage p50
  improves 95.1%/94.8% and 99.2%/99.2% across the two ABBA directions; p95
  improves 94.4%/94.8% and 98.9%/99.1%. Total p50 moves 8.4%/14.2% and
  6.6%/11.9%. The 4 MiB FAT controls retain one request and one call; paired
  read and total p50 changes stay within 5% control drift. FAT p95/p99 and all
  MiniFAT p99 tails are not accepted. p99, cold
  filesystem, simulated high-latency range, allocation, and peak-RSS claims
  remain open. This is substrate evidence only, not DOC/XLS/PPT semantic
  adoption.
- Change 0125 adds a distinct 4095-byte MiniFAT boundary pair over the same
  256- and 2,048-sibling shapes. The target occupies 64 logical 64-byte
  mini-sectors (eight regular 512-byte sectors);
  the matched legacy/positional controls record separate open/read/total
  timing, exact source calls/bytes/range sizes, and payload hashes. The focused
  gate requires legacy source-byte amplification and one exact positional
  4095-byte request, exposing physical-run coalescing without
  making a latency or resource claim. Release ABBA, tails, cold/high-latency,
  allocation/RSS, and native semantic consumers remain open.
- Change 0148 extends the current `open_stream` harness with correctness/source-
  event selectors for different-SID A-B-A, public bulk A-B-A, and overlapping
  same-target calls at 36- and 4095-byte MiniFAT targets. The selectors cover
  source ranges, ordered workload outputs, source-version stability, and typed
  refusal only; failure/retry, ineligible-root, FAT, native semantic, resource,
  and performance acceptance for those extended selectors remain open. See
  [`0148`](changes/0148-cfb-same-target-repeat-policy.md).
- Change 0149 compares the target-aware repeat policy against the immediate
  pre-change policy with four clean CPU-2 release legs and 28,800 retained
  samples. Under the named 100 us + 25 us/request, 50 MiB/s, 4 KiB-ceiling
  simulator, repeat-3 aggregate total p50/p95/p99/mean improves about 60-64%
  and repeat-8 about 56-64% in both adjacent ABBA directions. Same-target work
  changes from `[L,R,0...]` to `[L,L,...]`; later calls are direct reads rather
  than zero-source cache hits. Local bulk/concurrent cells contain >5% review
  triggers and substantial control drift, so local, per-invocation, bulk,
  concurrent, allocation/RSS, physical-I/O, and native-format claims remain
  withheld. See
  [`0149`](changes/0149-cfb-same-target-repeat-release-abba.md).
- Change 0152 compares the final same-target MiniFAT single-flight revision
  (`c270c8f3b` plus `f46381c6f`) with clean control `e486e4b1` in a CPU-2
  release ABBA: 20 warmups and 500 samples across 24 records per leg (48,000
  retained samples). All correctness/source-event invariants passed; existing
  concurrent scenarios recorded 6,473 candidate versus 8,000 control logical
  source calls (19.09% fewer). Only this source-event/correctness scope is
  accepted. At that revision the 291-name matrix was unchanged; change 0153
  adds four RTF selectors measured at the pre-staged publication-call interval,
  making that matrix 295. Change 0154 adds six ODF publication selectors,
  making that matrix 301; change 0159 later made it 302, change 0160 made it
  303, change 0162 made the matrix 305, change 0163 made it 309, and change
  0164 made that matrix 311; change 0166 made it 315, change 0174 made it 319,
  and change 0175 made the then-current matrix 320.
  Only `cfg(test)` source-event acceptance and tests changed in 0152. Root
  MiniStream cache and resource-accounting boundaries, broader performance
  gaps, and local/generic latency, allocation/RSS/peak-memory, physical
  I/O/syscall, cold-cache/device/network, decompression, native semantic,
  OOXML/ODF/RTF/iWork claims remain outside scope. See the
  [`0152` release record](changes/0152-cfb-same-target-singleflight-release-abba.md)
  and [summary](results/cfb-singleflight-abba-0152-summary.json).
- Change 0126 adds eight ordinary-root DOCX filesystem selectors over the
  unchanged 200-paragraph/eight-incompressible-2 MiB-media corpus. The eager
  control times `fs::read` plus `Document::from_bytes`; the source control
  times `Document::open(path)`; prepared-root query selectors time only their
  exact query. Untimed parity covers semantic projections and metadata; exact
  source SHA plus logical OPC part/relationship/content-type/blob-hash gates
  cover package preservation, including media hashes and source immutability.
  A separate typed source replay classifies zero payload overlap at open,
  complete compressed main-document range coverage during query-selector preparation, and
  zero main/media/unselected/core overlap during the query, while recording
  calls, bytes, request sizes, coverage and materializations. This is
  correctness/logical-range evidence only; latency, physical-I/O,
  decompression, allocation, RSS, cold-cache, ABBA, broad-security and
  Markdown-performance claims remain open.
- Open eagerly materializes FAT, directory, MiniFAT, and allocation topology,
  while ordinary large stream payloads remain lazy.
- MiniFAT now parses directly into its final `Vec<u32>`; FAT/DIFAT/MiniFAT use
  one bounded sector buffer and directory sectors batch into the final buffer.
- Child lookup now descends the validated sibling tree with SID-aligned cached
  comparison keys. The 2,048-root-stream measurement improves about 94%.
- Fresh XLS and PPT writers move generated stream buffers into `OleWriter`.
  PPT improves about 20%; XLS peak heap falls about 9.5%. DOC retains the
  exact-sized copy because moving its spare-capacity buffer regressed 58%.
- Directory writing allocates scratch structures proportional to every entry
  for each storage, and duplicate checks scan existing siblings.
- `HashMap`/`HashSet` iteration in fresh CFB directory construction requires a
  separate determinism audit; it is not treated as a performance result.
- A new source-backed same-length overlay substrate resolves selected existing
  streams through validated FAT/MiniFAT chains, derives bounded sorted physical
  spans, reopens the complete composed CFB and checks every selected stream
  before output. It rechecks source version and exact source/target fingerprints
  around 64 KiB sequential publication. Direct sinks receive typed partial
  progress; path publication uses synced sibling staging and atomic rename. The
  common wrapper retains signed/encrypted/DRM refusal and never falls back to a
  topology-changing render. No DOC/XLS/PPT end-to-end consumer or speed claim
  is adopted yet; generic CFB substrate correctness is not semantic format
  coverage.
- Atomic CFB overlay save now skips only the duplicate post-emission complete
  fingerprint scan. The saved path is mechanically `4N -> 3N`; direct
  `write_to` retains its post-emission scan. A pinned warm release ABBA run on
  the four-megabyte/five-entry corpus reduces exact logical source reads from
  101,751,908 bytes and 2,084 calls to 84,838,500 bytes and 1,825 calls:
  16,913,408 bytes (16.6222%) and 259 calls (12.4280%). All four legs publish
  the same 16,913,408-byte digest. The paired p50 directions are +3.7963% and
  -10.0141%, so no latency/speedup, RSS/allocation, physical-cold, or storage
  claim is accepted. Parent-wall and warm process-I/O counters remain
  descriptive only; see [change 0103](changes/0103-cfb-atomic-save-scan-evidence.md).
- Complete CFB overlay fingerprint scans now coalesce positional requests with
  a right-sized window capped at 1 MiB; comparison/emission stay at 64 KiB and
  no scan is removed. Clean balanced release evidence reduces calls from 1,825
  to 857 with unchanged 84,838,500 logical bytes and accepts p50/p95/mean in
  both directions for warm and advisory-cold states. The maximum code-local
  fingerprint buffer grows by 983,040 bytes; whole-process RSS is neutral in
  the matched boundary, while operation-only allocation and physical-I/O
  claims remain open. See [change 0143](changes/0143-cfb-fingerprint-read-coalescing.md).
- PPT root slide-order capture now passes its package-owned validated
  `OleFile` to independent live-document inspection instead of rebuilding the
  CFB index. Large root-open p50 improves 8.78% and allocation calls fall
  5.01%; the stream/current-user/live-persist and higher-level snapshot checks
  remain.
- Direct PPT text editing now holds its semantic selector result until the
  complete protection/editor preflight succeeds, then uses that editor for
  persisted-record resolution instead of opening the CFB editor a second time.
  Large direct edit/save p50 improves 14.12%; commit-time fresh-editor source
  comparison, publication and complete readback remain.
- The native PPT root now adopts a just-validated private text publication
  only after exact source and slide persist-ID checks. Default-limit root
  one-shape edit/save p50 improves 18.59%; custom limits and every structural
  path retain the complete root reopen.

The harness now measures native DOC/XLS/PPT open/list/one/full/no-op/one-edit
flows over deterministic writer artifacts. From the original baseline, large
one-edit/save p50 was 1.722 ms for XLS, 1.416 ms for DOC, and 0.357 ms for PPT.
XLS changed commit now reuses its already validated CFB editor instead of
discarding one BIFF parse and repeating the CFB open/capture; p50 improves
7.72%. Same-family fixed-width numeric commits now also certify that only the
requested Number/RK/MulRK fields changed, retain untouched worksheet
inventories, and clone only the edited sheet instead of rebuilding the complete
private offset inventory. The large 8,192-cell one-edit/save p50 improves a
further 7.83%; the complete independent public Workbook open/readback remains.
DOC publishes its ordinary WordDocument and table-stream replacements
as one failure-atomic object-editor batch instead of rendering/reopening the
CFB after each stream; p50 improves 10.52%. Both retain their final owner and
independent public-reader reopens. PPT root snapshot capture separately reuses
its first validated CFB open and improves p50 8.78%. Direct text-edit setup now
reuses its full editor preflight for record resolution and improves 14.12% p50.
Checked adoption of that result then improves root one-shape edit/save 18.59%
p50. The previous
spare-capacity DOC move remains rejected and must remain an independent writer
guardrail.

Change 0160 adds the missing direct attribution boundary rather than another
shortcut. The opt-in `doc_owner_public_phases` case observes the exact same
strict owner and complete public-reader work on tiny, large, and payload-heavy
writer artifacts, then independently records edit creation, replacement
staging, in-memory owner rendering, candidate owner/public validation,
source retention, patch construction, and output materialization. The format
crate emits content-free ordered events and owns no clock. A focused test and
three-shape unoptimized smoke establish event, arithmetic, semantic, patch,
hash, refusal, and untouched-stream correctness. The clean revision
`ab333008d3` release distribution, pinned to CPU 2 across four fresh processes
per shape and 800 retained samples per shape, now shows that combined
initial/final complete public-reader validation is the largest grouped named
phase for large (0.598 ms p50 of a 1.157 ms lifecycle) and payload-heavy
(20.721 ms of 44.227 ms), while patch fingerprinting is largest for tiny
(0.026 ms of 0.081 ms). Replacement staging is also material at 7.470 ms p50
in payload-heavy. Lifecycle p50/mean spread across processes remained below
3.0%/3.8%; two tiny subphase means crossed the 5% review trigger without
changing rank. This is an attribution baseline, not an accepted optimization
or speedup. The next native mechanism must keep both validation layers and pass
balanced clean release comparison.

Change 0161 tested and rejected the narrowest copy-elision interpretation of
that result. Borrowing the source/candidate bytes only for complete
public-reader validation improved tiny lifecycle p50 3.20%/3.24%, but large
regressed 3.06%/7.31% and payload-heavy directions disagreed. The complete
public-reader interval itself regressed 8.18%/11.54% p50 for large. The removed
clone may have preconditioned cache/allocation state, but no cache or allocation
claim is made. The candidate is fully absent from production. Further native
DOC work needs a measured shared physical/parsed substrate or fused proof, not
the same naked borrow substitution.

Change 0162 closes the immediate selectable-evidence gap for the existing RTF
standalone-picture APIs. Two opt-in selectors exercise bounded same-length
PNG/JPEG payload replacement and exact picture-group removal over 2/8/64-group
ASCII corpora whose hexadecimal transports mix case, spaces and newlines. The
harness builds an independent exact splice, leaves an unselected replacement
group in every shape, reports open/stage/commit/publication/lifecycle vectors,
and publishes through the public sequential writer into a zero-retained-output
hashing sink. Volatile and deterministic durable patch round trips, no-op,
stale/foreign/refusal and partial/zero-sink gates are untimed. This is phase and
correctness evidence, not an optimization result; `Edit::commit` still owns a
complete candidate and reparses it. Allocation/RSS and a clean balanced release
comparison are required before choosing a picture-path optimization. See
[`0162`](changes/0162-rtf-picture-crud-evidence.md).

Change 0164 closes the adjacent ordinary-paragraph split/merge evidence gap.
The two opt-in selectors reuse the exact plain lifecycle corpus at
tiny/medium/large sizes and call the committed `Edit::split_paragraph` or
`Edit::merge_paragraphs` API once per transaction. Split inserts one canonical
five-byte `\\par ` boundary and merge removes only the authenticated adjacent
boundary. Independent raw splice, semantic reopen, exact no-op, volatile and
deterministic durable forward/inverse, stale/foreign, bounded refusal,
partial/zero-sink and hash gates are outside the named phase intervals. The
public writer publishes to a fixed 16-KiB windowed sink retaining zero output;
the candidate transaction remains outside that sink bound. This is a
correctness/phase baseline only. It makes no latency, speedup, allocation/RSS,
transaction-memory, physical-I/O, cold-cache, source-backed, real-producer,
or rich-RTF claim. The selector summary exposes
`forged_result_artifact_refusal_verified`; the existing focused RTF tests
remain the authority for exact boundary-byte restoration and forged-boundary
precondition refusal.
See [`0164`](changes/0164-rtf-paragraph-split-merge-evidence.md).

Change 0165 records the narrow native-DOC lazy-fingerprint implementation found
by the owner/public-reader phase attribution plus a bounded descriptive comparison. `Snapshot` caches its FNV-1a
diagnostic fingerprint in an inline `OnceLock`; patch construction no longer
eagerly scans before/after artifacts, and `Arc` identity plus length provides
a same-lineage no-op/apply fast path. Independently reopened sources still
perform lazy fingerprint comparison followed by exact bytes, preserving the
full collision/stale/inverse/refusal boundary. The fingerprint accessors are
non-`const` because their first call may initialize the cache.

The established DOC lifecycle timer remains comparable. Same-lineage apply and
first fingerprint demand are explicit post-lifecycle workflow vectors. The
final clean control/candidate revisions are
`d6818e290aa77fd7666b7b16ee6908319d0f332b` and
`5dd813b1e108e253457ccb6c504c125c2becc1c6`; their binary SHA-256 values are
`344c0504c254109ee6b4361e375599d187f8a12333abb44f207d837af259ef8c` and
`c95e6c6004cbd725c789597566a81c0897ab6915ecd7c274deab222d134b3fd3`.
Clean CPU-2 release ABBA used 20 warmups and 500 retained samples per tiny,
large, and payload-heavy shape in each leg. Lifecycle p50/mean/p95 positive-faster deltas are
`+33.77/+35.19/+38.94` and `+33.21/+34.76/+39.67` tiny,
`+12.28/+12.59/+17.53` and `+13.81/+13.55/+11.68` large, and
`+17.33/+17.09/+16.58` and `+17.82/+17.75/+16.25` payload-heavy. With
immediate fingerprint demand, workflow p50/mean/p95 positive-faster deltas are
`+14.56/+16.34/+22.24`, `+13.89/+15.80/+21.90`,
`+4.50/+4.82/+10.24`, `+5.83/+5.64/+4.26`,
`+6.55/+6.41/+6.26`, and `+7.08/+7.08/+6.33` in shape/direction order.
The isolated patch/apply extension spans about 99.6-99.99% across the reported
p50/mean/p95 deltas, while its
deferred first scan is visible at roughly 25.7 us, 164 us, and 8.37-8.39 ms
for the three candidate shapes.

Same-implementation lifecycle p50/mean drift is control
`-1.18%/-1.41%`, `+0.25%/-0.42%`, `+0.48%/+0.72%` and candidate
`-0.34%/-0.75%`, `-1.50%/-1.51%`, `-0.12%/-0.08%` in tiny/large/
payload-heavy order. The positive paired directions are not generalized beyond
the named host/corpus. Final heaptrack records 50,677 allocation calls and
128.28M peak heap for both revisions, with profiler RSS 145.14M versus
142.81M; `/usr/bin/time` A1/B1/B2/A2 maximum RSS is
138160/138024/138028/138032 KiB. These are descriptive whole-process probes
only. No speedup, physical-I/O, cold-cache, real-producer, total-memory, operation-only
allocator/RSS, generic-DOC, or CRUD-completion claim is attached. Remaining DOC
work must attribute distinct validation or publication work without removing
either independent validation layer.

Change 0167 confirms one narrower XLSX publication hotspot: source-backed row
visibility was reparsing the selected worksheet and rescanning row tags after
the commit had already retained an exact source-bound cell-values patch and row
snapshot. The publisher now reuses the established matched/mismatched/
unavailable provenance boundary and still enters the complete OPC overlay
publisher. A >8 MiB read trap proves that only the mandatory selected-member
publication read remains. Descriptive release publication reductions are
50.42%-68.23% across p50/mean/p95/p99 in both paired directions, but control
and candidate drift exceed the 5% stability gate and complete-workflow medium
p99 directions disagree. Treat the redundant parse/scan as removed production
work, not as an accepted end-to-end or resource result; stable ABBA, allocator,
RSS, and physical-I/O evidence remain open. See
[`0167`](changes/0167-xlsx-row-visibility-provenance-reuse.md).

Change 0168 confirms a separate native XLS validation hotspot. The plan-only
Number/RK/MulRK path was reconstructing a fingerprint-checked composed source
twice after the common CFB planner had already reopened and fenced the same
candidate. An additive callback now runs BIFF semantic validation on the exact
composed view inside CFB's existing final fingerprint bracket. This removes two
complete source scans without weakening structural, selected-range, semantic,
security, stale-source, no-op, or publication checks. The deterministic work
delta is 33,991,680 bytes/34 one-MiB reads for Number and 405,504 bytes/two
reads for RK/MulRK per effective sample. Clean release pairs observe lower
complete-workflow and semantic-commit distributions in both directions, but
control and candidate drift exceed the 5% stability gate. Treat the scans as
removed production work, not as an accepted latency or physical-I/O result;
stable ABBA, allocator/RSS, cold-device, and producer evidence remain open. See
[`0168`](changes/0168-xls-numeric-validation-fusion.md).

Change 0169 confirms that hierarchical accounting itself was a measurable XLSX
streaming hot path. The large one-sheet shape performs one Work charge per row,
one per cell, and one releasable Objects reservation per accepted row. The old
cumulative charge path built and dropped an owned ancestor vector each time.
`Budget::consume` now walks immutable ancestors by reference; reservations keep
four charged nodes inline and spill beyond that without bounding caller-defined
hierarchy depth. Clean paired release runs accept medium/large
p50/mean/p95/p99 and tiny p50/mean/p95 reductions. Tiny p99 regresses and is
withheld. Matched whole-process Heaptrack allocation calls fall 48.81% and
temporary allocations 69.77%, while peak heap is unchanged and RSS directions
disagree. Deflate remains the dominant sampled cost. Treat this as accepted
warm in-memory one-sheet XLSX creation and shared-accounting allocation evidence,
not total-memory, physical-I/O, cold-cache, richer XLSX, producer, or universal
budget evidence. See
[`0169`](changes/0169-xlsx-streaming-budget-charge.md).

Change 0170 addresses the next measured encoder slice without changing the
larger Deflate policy. The control profile attributes 4.35% of process samples
to `push_escaped`, including 2.01% reaching `memmove` through per-scalar
extension. Ordinary UTF-8 is now appended by run between the five XML entities;
text length avoids a second scalar pass when bytes prove the character bound,
and each row number is formatted once. Clean paired release runs accept large
p50/mean/p95/p99, medium p50/mean/p95, and tiny p50 improvements. Exact output
and the 4 KiB window remain fixed. Process-wide instructions/branches decrease,
but branch misses regress and no allocation/RSS/total-memory/I/O claim follows.
Deflate remains the dominant sampled cost and requires an explicit cross-format
compression-policy study. See
[`0170`](changes/0170-xlsx-streaming-escape-runs.md).

Change 0171 confirms another legacy CFB validation hotspot. Source-backed DOC
paragraph, PPT shape-text, and XLS visibility paths were asking the validated
plan to reconstruct and fingerprint a composed source after the common planner
had already created, reopened, and fenced that exact view. Owner validation now
runs in the existing callback before the final fingerprint fence. Each
effective transaction removes one complete artifact scan and one source/target
digest pair; the measured 2,135,552-byte XLS corpus avoids three logical
one-MiB reads. Clean release ABBA accepts 12.51%-15.38% lower p50/mean/p95 for
the 64-worksheet complete workflow and 31.44%-33.16% lower p50/mean/p95 for
scalar/batch semantic staging/plan. Scalar total, p99, publication, resource,
physical-I/O, cold-cache and DOC/PPT latency claims remain open. See
[`0171`](changes/0171-cfb-owner-validation-fusion.md).

Change 0172 confirms that generic mutable-source publication fences were also
dominant on the native XLS plan-only path after semantic validation fusion.
The CFB owner now seals immutable provenance only from `Arc<[u8]>`; direct
`write_to` skips its redundant complete pre/post scans but retains the 64 KiB
emission pass and source/target hashes. Number removes 33,991,680 logical bytes
and 34 one-MiB reads; RK/MulRK removes 405,504 bytes and two reads. Clean
20-warmup/500-sample release ABBA accepts 37.54%-39.00% lower complete-workflow
statistics and 64.44%-66.76% lower direct-publication statistics through p95
for both families (plus Number p99). RK/MulRK publication p99, atomic-save,
resource, physical-I/O, cold-cache and producer claims remain open. See
[`0172`](changes/0172-cfb-owned-numeric-publication.md).

Change 0173 closes the same duplicated validation/publication work for native
XLS existing comments. The format owner validates the exact composed candidate
inside the planner's final fingerprint bracket, and immutable `Arc<[u8]>`
provenance removes direct `write_to`'s two outer mutation preflights. On the
16,995,840-byte corpus this eliminates three complete scans, 50,987,520
logical bytes, 51 one-MiB reads, and three source/target digest pairs per
effective transaction. Emission hashing, generic-source mutation defenses,
and atomic-save fences remain.

Clean 20-warmup/500-sample release ABBA accepts 45.54%-47.19% lower scalar
complete p50/mean/p99, 30.78%-32.42% lower scalar semantic staging/plan,
59.15%-61.03% lower scalar direct publication, and 30.53%-32.57% lower batch
semantic staging/plan. Scalar complete p95 and batch total/publication remain
withheld for drift/guard instability; resource, physical-I/O, cold-cache and
producer claims remain open. See
[`0173`](changes/0173-cfb-comment-publication-fusion.md).

The source-backed XLS worksheet-visibility overlay landed in committed
production change `bac279116`. Change `0091` adds four opt-in eager/source-backed
scalar and bounded-batch selectors for one-owner and 64-owner visibility edits.
They verify complete worksheet/catalog/opaque-stream readback, exact overlay
bytes, patch/inverse, source fingerprints/spans, and cap/protection refusals.
Change `0095` makes the existing comment and visibility source-backed owners
submit only their exact NOTE/TXO or `BoundSheet8` byte ranges to the common CFB
splice planner. Replacement staging falls from 80,946 bytes to 109/27,904 for
one/256 comments and from 18,166 to 1/64 for one/64 visibility owners. Balanced
ABBA accepts no latency speedup: all source-backed p50 directions remain inside
1.5%, and each workload's largest absolute source-backed delta is below its
largest absolute eager-control delta. Allocation, RSS, peak-memory and physical I/O
remain open, and the complete candidate snapshot/readback remains.

Change `0136` now measures the fixed-width Number/RK/MulRK source-backed path
directly before further production work. On one pinned release process, the
source-backed Number p50 is 146.410 ms versus 31.492 ms eager (4.65x), and the
source-backed RK/MulRK p50 is 1.627 ms versus 0.100 ms eager (16.25x). Both
paths retain complete 16,995,840-byte or 202,752-byte targets and produce
byte-identical family outputs. Source-backed commit and publication phase
medians are 101.618/44.783 ms for Number and 1.117/0.509 ms for RK/MulRK.
This confirms complete target capture/publication as a high-value attribution
boundary; it does not yet isolate an allocation, reopen, hashing, or emission
substage and is not an accepted speedup/regression claim. Change `0137` now
adds two opt-in plan-only selectors that validate a composed target without
retaining a second complete CFB artifact. Their forward-only API deliberately
does not expose patch/inverse; full reopen, fingerprint, sink-failure,
no-op and exact source/target fingerprint preflights, topology and security
proofs remain, and complete bytes are still emitted at publication. Composed
semantic validation may allocate/read a candidate Workbook model, so zero
target-artifact bytes is not a bounded total-memory claim. No latency, memory,
allocation, RSS, or I/O claim is accepted until matched release ABBA evidence
is captured.

An XLS-only immediate handoff of the first validated terminal CFB rendering was
also measured and fully reverted. Tiny changed save improved 7.55%, but large
changed save was neutral at -0.39% p50 and four repeated large exact-no-op
cycles regressed 22.00% p50 / 16.69% mean. Allocation calls fell 0.33% and peak
heap stayed flat, which confirms that work was removed but not that the public
operation improved safely. See
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

## Native OLE2 semantic path

The native semantic matrix separates ordinary reader facades from exact-source
transaction owners. Open is timed explicitly; list/one/full operations start
from opened ordinary models; no-op and one-edit publication start from opened
DOC body-text, XLS cell-value, or PPT slide-order snapshots and include owned
output materialization. Complete semantic verification and patch/inverse
checks stay outside timing.

Measured large-corpus priorities:

1. XLS one-cell publication originally measured 1.722 ms p50. Reusing the
   rendered/reopened CFB editor removes a discarded BIFF parse and redundant
   package capture. Fixed-width numeric inventory carry-forward then reduces
   the current large path another 7.83% p50 while retaining exact byte-range
   proof and independent public readback. Complete Workbook validation, common
   CFB publication, patch construction and output materialization remain.
2. DOC one-paragraph publication originally measured 1.416 ms p50. Batching
   its ordinary two-stream replacement removes one intermediate CFB
   render/reopen, while complete revision, style/property and independent
   document readback remain. A later profile found repeated linear physical
   PieceTable scans at 36.89% of large-open self cycles. The accepted FC index
   reduces large open from 790.727 to 348.679 us p50 and the changed edit/save
   path from 1.379 to 0.950 ms p50 while retaining all FKP/readback validation.
   A subsequent profile found another 6.94% self cycles in repeated paragraph
   style resolution and validation. The accepted one-entry resolved-baseline
   cache reduces the current large open from 343.503 to 304.199 us p50 and
   allocation calls 18.61%, while every direct PAPX and style switch remains
   scalar and independently validated. A later CHPX profile attributed 7.56%
   of process self cycles to paragraph character-run extraction. The accepted
   monotonic range slice reduces the 512-paragraph list from 454.100 to 358.414
   us p50 and the frame to 1.23%, without adding storage or allocations. The
   next exact-source profile found two ordered containment tables restarted
   from the beginning for every paragraph terminator. Predecessor binary
   searches reduce the already-open 512-paragraph snapshot list from 206.644
   to 168.142 us p50 and the full one-edit/save path from 888.602 to 817.424 us
   p50; allocation calls and peak heap remain flat.
3. PPT one-shape publication (0.357 ms original p50) retains its complete
   text-owner commit and public readback. Root snapshot capture improves from
   37.522 to 34.227 us p50, the direct text-edit transaction improves from
   206.209 to 177.089 us, and checked root adoption reduces the full operation
   from 352.306 to 286.805 us p50.

Change 0117 adds pinned balanced release probes for native PPT lazy `Pictures`.
On the generated eight-slide/32-image corpus, independent untimed replay proves
that source-backed open reads 79,265 metadata/mandatory-stream bytes with zero
`Pictures` overlap, the cold all-images query reads the complete 8,389,408-byte
stream once, and a cached query adds no reads. A directly timed fresh
open-plus-all-images pair prevents misleading sums of phase medians. Both the
100-sample preflight and 200-sample/cooldown attempt failed the fixed
same-implementation drift gates, so no latency result is accepted. Allocation,
RSS attribution, cold-cache, producer-breadth, and save-path evidence remain
open.

An additive source-backed DOC owner now covers one ordinary Word97+ main-story
paragraph when its text and terminating mark are contained in one uncompressed
Unicode piece. Positional selection uses bounded chunks and same-width
`WordDocument` splicing, with exact no-op/source/fingerprint/stale checks,
candidate reopen/readback, inverse and typed partial-output coverage. Complete
artifact fingerprints and CFB validation/publication scans remain mandatory;
this is correctness/selector coverage only, with no end-to-end latency,
physical-I/O/range, allocation/RSS, cold/high-latency, real-producer or broad
DOC CRUD claim. See [`change 0105`](changes/0105-doc-source-backed-paragraph-splice.md).

See [`change 0015`](changes/0015-native-ole2-semantic-baseline.md),
[`change 0016`](changes/0016-xls-commit-editor-reuse.md), and
[`change 0017`](changes/0017-doc-batched-stream-publication.md), and
[`change 0050`](changes/0050-doc-piece-table-physical-index.md), and
[`change 0051`](changes/0051-doc-adjacent-style-baseline-cache.md), and
[`change 0053`](changes/0053-doc-chpx-range-index.md), and
[`change 0056`](changes/0056-doc-papx-containment-index.md), and
[`change 0105`](changes/0105-doc-source-backed-paragraph-splice.md), and
[`change 0024`](changes/0024-ppt-slide-order-open-reuse.md), and
[`change 0026`](changes/0026-ppt-text-edit-resolver-reuse.md), and
[`change 0062`](changes/0062-ppt-root-text-publication-adoption.md), and
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

The retained opaque-heavy common case now isolates editor open, candidate
publication, changed final rendering and the chained control at 1.382, 7.979,
5.473 and 26.086 ms p50. The stages are not additive: their sum is only 56.86%
of the end-to-end p50. A narrowly scoped inline recapture-allocation reuse
improved candidate publication 6.49% p50/5.95% mean but the complete operation
only 2.61%/2.30%, with p95 +0.54%; it was fully reverted. See
[`change 0036`](changes/0036-ole-common-stage-attribution.md).

## RTF path

Change 0258 closes the unified owned-byte transport gap: the facade no longer
forces RTF through `String::from_utf8`, and all three detector seams admit
native compressed/stored framing after container precedence checks. New
`rtf_file_open` and `rtf_file_open_lifecycle` selectors isolate this adapter
boundary. Their two matched controlled captures disagree on the accepted
cell set, so this removes a correctness/compatibility hotspot without an
accepted latency result. See
[`change 0258`](changes/0258-rtf-byte-native-facade.md).

`litchi-rtf` now also exposes bounded forward-only authoring. Escape-free
printable ASCII is emitted in direct spans capped at 32 bytes, without adding a
retained writer buffer and while retaining Work/Output reservations and
per-write cancellation checks. Balanced release ABBA accepts p50 geomean
improvements of 76.41%/76.47% and p95 improvements of 75.23%/75.76%; the large
case drops from 7,208,970 to 1,441,802 sink calls with exact bytes and hashes.
This is fresh creation only, and allocation, peak-memory/RSS and cold-I/O
evidence remain open. See
[`change 0097`](changes/0097-rtf-bounded-ascii-streaming.md).

Existing-document logical-tail append is now a separate, opt-in harness path
from streaming creation. Tiny/medium/large plain corpora append 4/64/256
bounded one-run paragraphs and verify candidate reopen, exact sequential bytes,
durable patch/inverse and foreign-source refusal. The fixed 16 KiB hashing-sink
window caps accepted bytes per write and retains zero output, but does not bound
the transaction's validated candidate snapshot. This is correctness/coverage
evidence only; no release latency, allocation, RSS, or speedup claim exists.
See [`change 0090`](changes/0090-rtf-logical-tail-append-evidence.md).

Change 0153 adds matched Commit-versus-PublicationPlan append and exact-no-op
selectors over the same three plain shapes. The timer covers only publication
of pre-staged objects to the fixed sink; planning and publication vectors are
per-sample, while reopen and lifecycle vectors are one-element preflight-only
gates run once outside the sample loop. Durable patch, cancellation,
sink-failure/partial-progress, limits, and source-version gates are separate
untimed evidence. Results report retained source, complete-candidate, and
publication-window bytes explicitly. This is a
publication-boundary tranche only: no end-to-end, rich-format,
allocation/RSS, physical-I/O, or ABBA latency claim is accepted. See
[`change 0153`](changes/0153-rtf-tail-publication-plan-evidence.md).

Change 0154 closes the evidence gap for the generic source-positional ODF
`content.xml` publisher at its prepared publication boundary. Across matched
media-rich ODT/ODS/ODP corpora, a clean CPU-2 A/B/B/A run accepts p50
improvements of 96.35%-96.63% in both pair directions; p95, p99, and mean
agree, while maximum absolute same-implementation p50 drift is 1.441%. The
positional path retains no complete output and preserves every untouched raw
member plus physical and central order. This does not close end-to-end edit,
archive-open/indexing, allocation/RSS, physical-I/O, decompression,
cold-cache, filesystem, real-producer, or richer structural/resource-adding
ODF gaps. See
[`change 0154`](changes/0154-odf-content-cow-publication-evidence.md) and the
[summary](results/odf-content-cow-abba-0154-summary.json).

## PPTX additive-topology publication (change 0158)

The owned-source OPC publisher now preserves raw unchanged physical members
while appending generated Parts and supports exact physical-suffix omission
for the PPTX slide-copy inverse. Clean CPU-2 release ABBA accepts total p50
improvements of 29.643%/26.196% for the plain corpus and 43.294%/43.604% for
the approximately 16 MiB media-rich corpus. Media-rich publication p50
improves 49.321%/49.680%; p95/p99/mean agree. Plain publication tail claims
are withheld after same-implementation drift triggers.

Process-wide profiles attribute the media-rich result to less CPU work:
task-clock falls 42.399%/43.122%, cycles 42.583%/43.116%, and instructions
46.686%/46.775%, while maximum RSS is within 0.5% and peak heap is effectively
unchanged. The path still copies the complete owned source, ordinary OPC open
is still eager, and the writer performs unconditional plan-container reserves
that deserve a separate measured no-op/small-mutation tranche. Source-backed
topology mutation, physical/cold I/O, decompression/recompression-byte counts,
real producers, and general OPC/PPTX remain open. See the
[`0158` record](changes/0158-pptx-additive-topology-release-abba.md) and
[summary](results/pptx-additive-topology-abba-0158-summary.json).

The standalone harness now records a fixed six-bucket distribution for every
serialized sink summary. It counts logical `Write::write` calls at the point
where bytes are accepted, includes zero-length calls, excludes rejected calls,
and checks that the bucket total equals `write_calls` at the exact inclusive
boundaries. This is reporting evidence only: it does not measure syscalls, disk
I/O, memory copies, compression, latency, allocation, RSS, or performance.
See [`change 0107`](changes/0107-output-write-size-evidence.md).

The existing seven native public cases cover owned open, lazy paragraph listing, one
paragraph, first complete text, exact stream save, exact empty-edit save, and
one checked paragraph edit/save over 24/200/10,000-paragraph corpora. The
unified root path remains intentionally outside this evidence.

`RtfDocument` now retains the total block-text byte length during its existing
owned-detach pass, so first full-text materialization performs one exact
allocation and one block pass instead of allocating and joining a temporary
fragment vector. The large full-text p50 improves 27.08%. Canonical text
emission now writes contiguous ASCII spans rather than formatting one character
per sink call, and text-only commits skip unused paragraph-property vectors and
scans. Together with early successful paragraph selection, large one-edit/save
p50 improves 25.79%. Full-text caching, forward-only sink errors, exact no-op
identity, opaque refusal, validation, and complete reopen/readback remain.

The next parser profile found a full `State::clone` on every ordinary body-text
flush at 8.53% exclusive samples. Ordinary flushes now borrow the state and
copy only effective encoding, formatting and paragraph properties; the full
state is retained only for insertion/deletion metadata. Large open p50 improves
20.09% and large one-edit/save improves 11.54%, with flat allocation count,
peak heap and RSS. Code-page selection, revision ranges and deletion behavior
have focused and complete-suite coverage.

The following profile attributed 15.37% of large-open and 14.46% of large
one-edit/save samples to extending parser transport buffers one byte at a time.
All-ASCII source tokens now enter those buffers in one extension; byte-valued
non-ASCII and invalid-Unicode input retain the checked per-character fallback.
Large open p50 improves 26.67% and one-edit/save 6.26%. Instructions fall
18.40%; allocation count, peak heap and RSS remain flat.

The next matched profile retained 17.36% exclusive large-open cycles in
`Lexer::tokenize_with_spans`: ordinary text decoded and advanced over every
UTF-8 scalar twice merely to find five ASCII delimiters. One checked byte scan
now finds those delimiters while retaining UTF-8 boundaries and exact source
spans. Large open p50 improves 17.23%, one-edit/save 14.65%, instructions fall
21.27%, and the lexer frame falls to 11.06%. Medium/large plain, raw CP-1252
and LZFu opens all improve; the prepared LZFu no-op microsegment exception is
disclosed in change 0040.

A follow-up attempted to move decoded block ownership directly into the final
document. The broad version removed 20.15% of process allocation calls and
improved raw CP-1252 open 3.08% p50, but moved ordinary ASCII allocation into
the parser loop and regressed plain large open 25.53% p50. Owned-only variants
measured -1.41% and +1.02% p50 across separate 4,000-sample/state runs. The
production parser was restored exactly; do not revisit this copy in isolation.
See change 0043.

The next changed-commit profile isolated a separate, short-lived owner: after
the initial complete parse, `ordinary_body_source_span` cloned the 540,051-byte
ASCII source, tokenized it again and scanned root depth before the required
candidate parse/readback. Direct uncompressed ASCII parses now retain a compact
range proven inside the parser's existing structural preflight. Ambiguous,
empty, binary, non-ASCII, compressed and over-32-bit ranges keep the established
locator/refusal. Large one-edit/save improves 10.72% p50 and 10.11% mean;
instructions fall 10.64%, and the 588 before-only locator allocation calls
over 20 edits disappear. Peak heap and uninstrumented RSS remain flat. See
change 0048.

The next large-open profile attributed 25.65% of cycles to `memmove`; the
10,000 retained `StyleBlock` values grew through 12 vector allocations to a
16,384-element / 16.12 MiB capacity. The existing structural preflight now
counts root text tokens and passes a bounded hint to the first retained block.
One 9.84 MiB exact reserve replaces those growth/copy steps. Large open p50
improves 21.17%, mean 21.00%, cycles 14.91%, and cache misses 32.20%; peak heap
falls 29.73%. Table/deletion-heavy and sub-64-KiB sources retain lazy growth.
Medium plain/CP-1252 p50 movements of +0.49%/+2.84% are disclosed. See change
0055.

Formatting/media, malformed/security, broader real-producer, cold-source,
broad edit and conversion matrices remain missing. Compressed LZFu and raw
CP-1252 open/read/no-op coverage is now measured but remains narrow.

## Source and detector path

`litchi-core::ReadAt` provides immutable positional reads and source versions.
Source-backed OPC and positional CFB now consume it; IWA also consumes it,
though current IWA physical ingress snapshots the complete source with one
full-range request.

Generic smart detection may scan one ZIP package through multiple format
owners and then discard the prepared parse before the selected owner opens it.
The focused iWork route has already disproved that this duplication is
architecturally necessary: `litchi-iwa-detect::PreparedSource` retains an
opaque classified physical catalog without exposing archive types in the root
facade. The generic detection path must be measured before adapting that
pattern elsewhere.

## Initial hypotheses

| # | Source-audit disposition | Measurement needed |
|---:|---|---|
| 1 | Refined: legacy OPC path and `Read` ingress slurp the source; source-backed ingress is positional. Five filesystem cases now record process-isolated warm/cold-requested counters and atomic-save hashes, including a repeated release tmpfs capture. | Repeat on a controlled block-backed filesystem/cache host. The release run's accepted cold advice and zero process `read_bytes` on tmpfs are counters/output evidence only, not physical cold-cache behavior. |
| 2 | Confirmed: ordinary OPC open inflates every admitted Part. Change 0186 removes the subsequent full-payload ownership copy, but not decompression. | Open/list/one-object scaling against total uncompressed bytes and member count; migrate selective facade opens onto deferred catalogs. |
| 3 | Implemented for managed source-backed OPC: finite weighted eviction, pinned-handle preservation, per-entry single-flight, exact physical `InputBytes`, exact accepted direct-sink `OutputBytes`, cumulative declared cold-load `Work`, retained catalog/flight/payload `Objects`, retained/in-flight payload `Memory` Budget charging and content-free diagnostics exist; compatibility/unmanaged opens retain finite `SourceCacheLimits`, and legacy eager open does not use that managed cache. | Correctness tests cover all managed resource dimensions and charging/release invariants. Release contention ABBA covers structural/distribution counters but accepts no speedup. Add allocation, peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance evidence. |
| 4 | Measured: ordinary OPC open is serial and explicit eager open has a local bounded session. Six large ZIP tasks reach 4.52x p50 at 12 CPUs; small tasks regress. | Broader real-package scaling and threshold tuning. |
| 5 | Confirmed: stored entries are CRC-checked then copied. | Stored-media one-Part read and package-open copied-byte/RSS deltas. |
| 6 | Refined by measurement and implementation: exact unchanged saves copy the source; owned same-topology mutations raw-copy unchanged entries; changed Parts share their immutable logical payload and validated generated local span without extra copies; the bounded source-backed publisher materializes only selected targets and raw-copies the rest; guarded DOCX, atomic same-slide and multi-slide PPTX shape-text batches, and eleven managed XLSX source-editor closures consume it; borrowed/topology-changing paths rewrite fully, while unsupported source-backed layouts refuse. | Real-producer media-heavy multi-Part updates, broader semantic closures, signature/topology policies, and attribution of the remaining selected-Part/compressor-buffer memory cost. Change 0151 adds no latency/resource-performance evidence. |
| 7 | Confirmed structurally: duplicate indexes, boxed Parts, source-XML map, and linear fallback exist. | Allocation profiles, type sizes, cache counters and repeated noncanonical lookup. |
| 8 | Refined: source-backed XLSX structural open/list avoids timed reads; selected first/range reads physically overlap only the selected worksheet; guarded calculation-metadata, defined-name, tab-state, worksheet page-break, page-margin, print-options, relationship-free page-setup, sheet-protection, data-validation, auto-filter and conditional-formatting edits materialize only their one- to three-Part semantic closures, with managed `PartData` retention/refusal checks frozen in change 0151. | Broader source-backed selectors, general cell/formula edits and real workbook matrices; change 0151 adds no latency, RSS, allocation, I/O, copy, decompression or total-memory evidence. |
| 9 | Refined by measurement: small XLSX edits scan/rebuild/reparse the complete touched sheet; bounded commits can reuse the validation store for first read, while large sheets fall back cold. Direct writer-local action regrouping was immaterial and reverted. | Attribute larger semantic-planning/emission/readback passes, first/middle/last cells, distinct bulk actions, structural edits, large-sheet retention and commit-versus-save separation without reviving direct regrouping alone. |
| 10 | Plausible but unmeasured: per-cell semantic ownership and transient parse duplication may dominate large stores. | Allocation count/bytes, type sizes, peak RSS and cache-miss profiles. |
| 11 | Refined by implementation and measurement: CFB has positional `SharedOleFile`, bounded bulk reads, exact-range reads and exact-range splice publication; MiniFAT parsing and sector reads no longer require the former temporary buffers; child lookup descends the validated tree; and stream-allocation validation now reuses bounded MiniFAT/FAT chain scratch. Native DOC/XLS/PPT semantic baselines, XLS editor and inventory reuse, DOC batched publication and indexes, PPT root-open reuse, text-edit resolver reuse, and checked root text-publication adoption are accepted. Change 0094 accepts only the generic MiniFAT read-stage/source-byte and modest total-p50 evidence. Change 0095 adds native XLS comment/visibility semantic splice consumers with exact replacement-byte evidence but no latency speedup. Change 0102 range-resolves the native PPT one-shape selector but keeps full source fingerprint/publication checks and makes no end-to-end performance claim. Change 0105 adds a correctness-only Word97+ DOC main-story one-paragraph Unicode-piece splice with bounded positional selection, same-width replacement, candidate readback and source/inverse checks; complete CFB fingerprints and validation/publication scans remain. Change 0190 removes 48.44% of whole-process allocation calls and 98.94% of Heaptrack temporary allocations on the exact many-small plus wide-root open profile. Its accepted release timing is limited to many-small p95/p99 and wide-root p50/mean/p95; many-small p50/mean and wide-root p99 are withheld by their predeclared drift gates. The XLS terminal-render handoff was neutral on large changed saves and regressed exact no-op. The opaque-heavy common case rejected direct shared writer payloads, an editor-wide validated-render cache, and inline recapture-allocation reuse; its open/publication/finish/end-to-end stage split is non-additive. | Attribute materially different final owner/public-reader work without reviving the rejected handoffs or recapture reuse; add deep-directory, mixed MiniFAT/FAT, concurrent-open, real-producer, and security scenarios beyond generated corpora. Cold/high-latency sources, operation-local allocated bytes, and allocator-contention evidence remain open; the DOC owner still needs matched source/read-range and real-producer breadth before any performance claim. |
| 12 | Confirmed for generic detection; disproved for focused prepared iWork detection. | Generic detect-then-open versus prepared-source handoff. |
| 13 | Measured for ODS snapshots: one package clone and duplicate package parse were removable. Same-topology ODS row-local publication retains exact range provenance through raw ZIP emission, its unified worksheet handoff now shares/moves the exact archive allocation through the nested snapshot and package validation, and compact ODS/ODP/ODT content publication avoids rebuilding untouched data; repeated ODS cell lookup uses a bounded lazy locator. ODT existing-document/direct-byte/final-result snapshots, changed-operation compact audits and envelope classification share exact validated package allocations, consuming full-text block strings and an indexed one-paragraph retention path are accepted, consecutive plain-text replacements publish one candidate, and scalar line-break/run/hyperlink plus plain paragraph insertion/removal use the accepted content-only publisher. A matched release ODT mixed model-content case now measures one staged publication against 49/193 scalar publications over 80/320 operations, with 96.8685%–96.8695% medium and 99.2289%–99.2381% large p50 reduction and equal per-shape hashes; this is a narrow repeated-publication result, not a general ODT/resource/I/O/memory claim. ODP one-slide lookup retains only its requested semantic projection while validating through EOF, ODP transaction staging reuses its snapshot-validated complete slide projection, and exact slide-only commits adopt that already validated candidate only after final package audits. Parsed final-document adoption remains reverted. All accepted paths retain readback and source lineage. | Broader ODF source-backed reads, repeated independent ODP semantic scans, formatted/non-text bulk edits, resource-adding/richer structural publication, real-producer media, and structural-edit profiles. |
| 14 | Confirmed for DOCX direct-body batches: repeated full XML rebuild/parse work was removable while retaining ordinary durable operations and complete readback. | Real-producer/extension/security corpora and broader structural/bulk edit semantics. |
| 15 | Measured and implemented for RTF full-text, text-only edit/save, ordinary parser and already-open story-query paths: temporary fragment/property vectors, per-character writer calls, unconditional full-state cloning, per-character ASCII transport-buffer extensions, twice-decoded ordinary-text delimiter traversal, the second ordinary-body source lexer and repeated full-block length scans were removable. Raw CP-1252, LZFu and a real-producer watermark have capability-bounded read/no-op coverage, and `relsize` has checked native semantic readback. | Extend the accepted native matrix to formatting/media, malformed/security, more real producers and broad edit scenarios; attribute a distinct remaining frame before another specialization. |

## Ranked work queue

The order below is provisional until baseline measurements are recorded.

| Rank | Candidate | Expected CRUD reach | Risk | ADR fit |
|---:|---|---|---|---|
| 1 | Extend source-backed OPC from selective reads and the bounded consuming publisher to broad query/edit/patch coverage. | All OOXML selective read/query/edit paths; offsets eager full-package work. | High | Positional source/descriptors, low-level one-Part/bounded multi-Part publication and managed cache charging across physical `InputBytes`, cumulative declared cold-load `Work`, retained `Objects`, and `Memory` are implemented and correctness-tested; broader semantic CRUD and controlled cache acceptance remain. |
| 2 | Broaden the accepted source-backed DOCX/PPTX and XLSX calculation-metadata/defined-name/tab-state/page-break/page-margin/print-options/page-setup/sheet-protection/data-validation/auto-filter/conditional-formatting transactions only where complete semantic closures can be proved, with real media/signature/topology matrices. | Targeted OOXML save, especially media-heavy packages; avoids eager all-Part inflate/recompression where the same-topology proof applies. | High | DOCX is accepted in change 0039, guarded same-slide PPTX in 0044/0063 and bounded multi-slide PPTX in 0077, XLSX calculation metadata in 0046, page breaks in 0061, page margins in 0067, print options in 0070, relationship-free page setup in 0073, defined names in 0076, sheet protection in 0078, data validation in 0079, and auto filters in 0080. Change 0151 freezes managed constructors/ownership and correctness gates for those closures plus tab state and conditional formatting, but adds no performance result. Change 0120 adds ordinary-root PPTX open/list/count/selected-slide logical-read controls and complete parity gates, but no speedup/resource claim. General XLSX cells/formulas/chains, table filters, printer settings and structural PPTX edits require wider closures; all accepted facades still need real-producer and broader topology/signature policy matrices. |
| 3 | Tune explicit bounded-session thresholds and complete remaining I/O budget policy. | Large multi-Part open/save/validation. | Medium-high | 1/2/4/8/12 evidence exists; large tasks scale, small tasks regress; no hidden Rayon path remains. |
| 4 | Build one validated OPC publication plan and reuse its generated XML and Part order during emission. | Every rewritten OPC save. | Low-medium | Implemented; see `changes/0001-opc-publication-plan.md`. |
| 5 | Exact owned-source OPC no-op publication. | Owned DOCX/PPTX/XLSX open/read/no-op save. | Medium | Implemented; same-topology mutations now use targeted preservation. See changes 0004 and 0008. |
| 6 | Move already-owned XLS/PPT writer buffers into `OleWriter`. | Legacy fresh creation and some rebuilds. | Low | Implemented for XLS/PPT; DOC rejected by measurement. See `changes/0003-legacy-owned-stream-handoff.md`. |
| 7 | Use validated cached CFB sibling-tree descent and reusable sector/chain-validation buffers. | Legacy stream-heavy open/rebuild workflows. | Medium | Implemented; see changes [0002](changes/0002-cfb-lookup-and-sector-buffers.md) and [0190](changes/0190-cfb-stream-chain-scratch.md). Change 0190 accepts allocation-call reductions plus many-small p95/p99 and wide-root p50/mean/p95 latency reductions for its exact profiles; mixed-table and concurrent-open evidence remain. |
| 8 | Extend the accepted XLSX row-start index and bounded validated-store handoff to broader selector and edit matrices. | Sparse range queries and first reads after eligible changed-sheet commits. | Low-medium | Narrow ranges and bounded commit/read reuse are accepted in changes 0006 and 0025; dense-wide handoff is intentionally excluded, and preservation/readback gates and broad CRUD coverage remain unchanged. |
| 9 | Coalesce DOCX same-structure paragraph replacements and measure PPTX capture/fingerprint reuse. | 1% semantic document/presentation edits. | Medium-high | Implemented for canonical direct-body DOCX batches and PPTX selected-scene reuse; complete source validation and candidate readback remain. See changes 0010 and 0012. |
| 10 | Measure and tune the managed source-backed cache under controlled contention. | Concurrent repeated Part reads. | Medium-high | Hierarchical charging across physical `InputBytes`, cumulative declared cold-load `Work`, retained `Objects`, and `Memory`, plus pinned-aware eviction and per-entry single-flight, are implemented and correctness-tested in change 0086; release ABBA in 0088 covers structural/distribution counters but accepts no speedup. Allocation, peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance evidence are open. |
| 11 | Extend ODF beyond accepted ODS snapshot, row-local provenance reuse/shared worksheet ownership, ODS/ODP/ODT unchanged-member publication, adaptive cell lookup, ODP indexed-slide retention/snapshot handoffs and ODT byte/full-text/indexed-query/audit/envelope/batch/final-byte ownership: positional source-backed reads, repeated independent ODP scans, richer non-text/bulk edits, resource-adding/richer structural publication and real-producer media. | ODT/ODS/ODP open/query and changed save. | High | Same-topology ODS row splicing now carries exact range proofs through raw ZIP emission and the adjacent nested worksheet/package owners share and move their archive allocation; compact ODS/ODP/ODT content raw preservation, bounded facade lookup, direct/existing/final-result ODT byte sharing, consuming full-text blocks, indexed paragraph/slide retention, ODP staging and final slide-only snapshot projection reuse, matched ODP text-box and ODT embedded-resource scalar/bounded evidence, compact-audit/envelope sharing, consecutive paragraph coalescing and scalar line-break/run/hyperlink plus plain paragraph insertion/removal publication are accepted. Change 0122 adds matched ODP eager/source-backed media-rich open and middle-slide logical-read selectors with explicit selected-Pictures replay; change 0123 adds matched unified-root eager/source-backed filesystem open and middle-slide controls with complete post-timing semantic/metadata/media/member/hash parity plus direct typed replay evidence; change 0124 adds matched unified-root ODS eager/source-backed open plus typed selected-cell/media controls with complete untimed root/typed/archive/member/hash parity and direct positional-read evidence. These are correctness/range evidence only. ODS content-validation catalog CRUD is correctness-covered but unmeasured. Parsed final-document adoption remains reverted for a read regression; other structural fallback, exact no-op and full readback remain. See changes 0011, 0014, 0018, 0019, 0020, 0023, 0027, 0031, 0034, 0035, 0038, 0041, 0042, 0045, 0047, 0049, 0052, 0057, 0060, 0065, 0068, 0071, 0072, 0074, 0075, 0084, 0085, 0122, 0123 and 0124. |
| 12 | Extend accepted native RTF work beyond the capability-bounded variant matrix after parser-state, transport batching, byte-delimiter scanning, retained ordinary-body ranges, retained story-length/cardinality handoffs and sparse paragraph selection. | RTF formatted/media, malformed/security, broader real-producer and broad edit paths. | Medium | Plain, raw CP-1252, LZFu and producer-watermark read/no-op inputs plus a narrow native shape-text chain are covered; plain generated paragraph queries and editing are timed, public paragraph cardinality is parser-retained, and explicit sparse `nth` no longer constructs discarded paragraph views. Cached full text, byte-valued fallback, revisions, candidate readback and native forward-only output contracts remain. See changes 0013, 0019, 0020, 0029, 0040, 0048, 0064, 0066 and 0069. |
| 13 | Remove the second complete target artifact from fixed-width native XLS publication, then continue attributing remaining OLE2 final-owner/public-reader work. | OLE2 spreadsheet/document/presentation edit publication rather than substrate-only insertion. | Medium-high | Changes 0136/0137 established the source-backed baseline and forward-only plan; change 0138 accepts complete-operation latency for Number and RK/MulRK after strict CPU-2 release A1/B1/B2/A2 (p50/p95/p99/mean agree in both directions). Number process VmHWM also agrees (-10.73%/-10.66%), while RK/MulRK RSS directions disagree and valid heaptrack A/B profiles show descriptive whole-process allocation reductions with identical peak heaps. The accepted result is limited to these deterministic fixed-width families; composed validation may allocate/read a candidate Workbook model, so zero target-artifact bytes is not a bounded total-memory claim. No physical-I/O, cold-cache, operation-only allocation or broad-producer claim is made. |
| 14 | Share existing ODT transaction bytes when a validated document creates a snapshot. | ODT no-op and changed edit/save. | Low-medium | Implemented with private `Arc` identity proof; no-op p50 -18.51% large, guardrails within 3%. See change 0014. |
| 15 | SIMD or lock-free work. | Unknown. | High | Deferred until remaining hot loops/locks are measured after work elimination. |

## Evidence still missing

The deterministic harness now records warm latency distributions, confidence
intervals, corpus hashes, complete output validation, and sequential-write
call/byte counts. Targeted `heaptrack` runs also cover allocation count,
temporary allocation count, peak heap, and peak RSS for the implemented
changes. Remaining gaps are:

- Reproducible physical cold-cache distributions on a controlled host. Change
  0087's one-sample debug warm/cold-requested run and change 0089's repeated
  tmpfs release run are correctness/counter and descriptive distributions only;
  neither proves a cold device or storage result.
- CFB selective-range acceptance is bounded to exact source-byte counters,
  MiniFAT read-stage p50/p95, and the modest total-p50 direction in change
  0094. Change 0144 additionally accepts p50/p95 only for the named configured
  simulated range source: both MiniFAT targets reduce to one exact request and
  improve in both ABBA directions, while the exact-work FAT control stays near
  neutral. Real cold/network/device range sources, FAT tail behavior, p99,
  allocation, and peak-RSS evidence remain open, and no DOC/XLS/PPT semantic
  consumer is covered.
- CFB `open_stream` now has direct one-shot and sequential repeat selectors for
  36-byte and 4,095-byte MiniFAT targets. Change 0147 accepts the configured
  simulator's one-shot timing and exact source-work result, but repeats add one
  target-sized request before root-cache materialization. Change 0148 adds
  correctness/source-event coverage for different-SID A-B-A, public bulk A-B-A,
  and overlapping same-target calls. Change 0149 accepts the target-aware
  policy only for aggregate repeat-3/repeat-8 totals under the configured
  simulator; it explicitly withholds local/per-invocation/bulk/concurrent and
  resource claims after noisy >5% review triggers. Failure/retry,
  ineligible-root, FAT, native semantic, and complete resource acceptance
  remain open.
- Decompressed and recompressed byte observers. Positional range-request
  distributions now exist for OPC and XLSX, but not yet for every format/source.
- Broad hardware-counter evidence. A matched targeted-OPC run is committed now
  that the environment reports `perf_event_paranoid=1`; stage-1 remains without
  counters and no claim is generalized from the one measured save workload.
- Cache contention acceptance: change 0088 has release structural/distribution
  ABBA evidence, but no accepted speedup, allocation, RSS, hardware,
  copied/decompressed-byte, or CPU-utilization result.
- XLS visibility overlay performance/resource evidence. Change 0091 is
  correctness/coverage only: it has no release ABBA, speedup, allocation, RSS,
  peak-memory, or physical-I/O claim, and its complete source-backed candidate
  snapshot is not bounded by the 64 KiB publication sink.
- Format-semantic preservation evidence beyond the generated
  DOC/XLS/PPT/DOCX/PPTX/RTF/ODT/ODS/ODP slices and native targeted-OPC raw
  passthrough corpus.

## Change 0175-0176 update

Owned CFB atomic save no longer repeats the two complete fingerprint scans
whose source is sealed as `Arc<[u8]>`. The 16.9 MiB control drops 33,826,816
logical source bytes and 34 large fingerprint reads while retaining the
emission hashes and atomic durability sequence. This closes the safe owned
variant of the atomic-publication hotspot; generic stable-token `ReadAt`
sources remain deliberately fenced. Latency is withheld because control drift
exceeded 5%, despite both candidate directions being lower.

Two smaller duplicate-work hypotheses are now negative results. ODS retained
content proof added a second hash and regressed both measured source-backed
workflows; XLSX conditional-formatting readback reuse was directionally
inconsistent. Both production experiments are reverted. The next ODF/OOXML
work should target a larger retained allocation or complete physical/semantic
pass rather than reviving either micro-handoff unchanged.

## Change 0177 update

The accepted source-backed ODS existing-cell closure now has clean release
evidence at the complete open/stage/commit/sequential-publication boundary.
One-cell p50 improves 75.03%/74.27% across paired directions with all four
distribution metrics inside the stability policy, so eager whole-package
ownership is no longer the preferred path for that bounded media-rich case.

The same conclusion is not generalized to the 21-cell 1% selector: mean and
tail same-path drift exceed policy despite a large apparent p50 advantage.
ODS structural cells/rows, formulas, merges, insert/delete, real producers,
resource profiles and physical I/O remain higher-value open work than another
retained-content proof micro-handoff.

## Change 0183 update

A clean current-HEAD rerun now closes the previously withheld 21-existing-cell
case. Both paired p50 directions are 72.07%-72.61% lower and all p50/mean/p95/
p99 drift gates pass. This promotes only the fixed 1% complete lifecycle from
correctness/phase evidence to accepted warm latency evidence; no production
mechanism changed.

The next larger ODS seams remain the entire touched-`Sheet` clone during a
21-cell transaction and the complete XML layout/semantic reparse after the
rewritten worksheet is assembled. Either needs separate implementation and
matched evidence; neither is implied solved by this rerun.

## Change 0184 update

The source-backed XLSX existing-row visibility path no longer reparses every
scalar cell after its own bounded direct-`hidden` rewrite. A private token keeps
the exact source slice borrowed and identity-checked, candidate XML grammar is
still validated, row visibility is rescanned, and generic worksheet editors
cannot fabricate the handoff. The removed work is one complete scalar-cell
parse per effective changed commit.

Large hide-one/unhide-256 commit distributions pass all paired-direction and
drift gates; large unhide-256 complete lifecycle also passes all statistics.
The medium batch retains only commit p50/p99, while medium totals and medium
hide-one remain withheld. The next XLSX row work should target measured
stage/publication attribution, structural row ownership, formulas and producer
matrices rather than widening this proof to arbitrary worksheet rewrites.

## Change 0345: OPC source-backed reader ingress

The public source-backed OPC reader has recorded bounded-ingress evidence: one input consumption, typed exact-maximum rejection with actual = maximum + 1 asserted for the overrun, zero ordinary cold payload loads during open, and one selected cold/successful load. Relative to compressed-plus-all-decompressed eager retention, it retains one compressed buffer plus indexed metadata and deferred selected payloads. ReadLimits and try_reserve_exact bound logical input/local admission work, not total RSS or aggregate concurrent opens. This is not a hotspot or optimization claim (performance_claim: none); no RSS or before/after latency was measured. The evidence came from 4/4 focused tests, including reader_ingress_retries_one_interrupted_read and reader_ingress_rejects_invalid_read_count_without_panicking, and four owner-library checks under one Cargo process/job on a dedicated disk target. Callers needing tighter host memory must provide a lower max_input_bytes, serialize opens, and account aggregate process memory externally. Arbitrary blocking Read cancellation is not provided, and no facade or iWork API is involved. See [Change 0345](changes/0345-opc-source-backed-reader-ingress.md).

## Change 0346 update

The current-head XLS control smoke and 40-block lock/fingerprint probe do not
support a production `FileSource` lock substitution. `parking_lot` measures
154.03 ns/call mean versus 155.47 for `std`, but the modeled whole-operation
gain is only 0.36-0.40%. No candidate was applied, no XLS/CFB freshness fence
changed, and `performance_claim: none`. Do not revive the unchanged 0279
freshness session; the next XLS work must use a different measured design. See
[Change 0346](changes/0346-file-source-lock-candidate-rejected.md).

## Change 0347 update

The XLSX cell-values calculation-closure oracle now accepts only the direct
`calcPr` rewrite and exact optional calc-chain removal closure, while raw
identity remains required for all other members. Direct medium one-edit and a
24-row serialized ABBA smoke are complete with zero failure rows, but timing
gates and claim authorization remain false. This is harness evidence only;
formula/date workloads are excluded, and source planning/commit dominate the
retained source phases. A shared publication-copy design remains deferred.
See [Change 0347](changes/0347-xlsx-cell-values-harness-calculation-closure.md).

## Change 0348 update

Stored ZIP borrowing now validates complete local and central metadata,
signed and unsigned 32-bit descriptor CRC/size forms, local ZIP64-extra
provenance, encryption/overlap/duplicate safety, and nonempty zero-CRC
refusal. The immutable-slice path preserves pointer identity without cache or
materialization charge; ZIP64 EOCD, Deflate, and generic positional sources
remain owned or streaming fallbacks. Serialized evidence was
`focused borrowed 10/10; full soapberry-zip lib 280/280` under one build job,
one test thread, and an 8 GiB process ceiling. Downstream
`litchi-opc borrowed 12/12` passed, and `cargo fmt --package soapberry-zip
-- --check` passed after formatting; this is not the full `litchi-opc` suite.
`performance_claim: none`:
no latency/RSS/copy claim, stored OOXML representativeness is weak, and
concurrency is unchanged. See [Change 0348](changes/0348-stored-zip-borrow-validation.md).

## Change 0355: PPTX source-probe fallback admission

Change 0355 records a correctness and ownership boundary
(`performance_claim: none`). The private PPTX bytes probe now returns typed
`OpcError` outcomes and terminal `OtherOoxml`/`DisabledOtherOoxml` classifier
outcomes. Only genuine non-ZIP, short-input, or missing `[Content_Types].xml`
inputs admit compatibility fallback and reclaim the original `Vec`
allocation. Hard ZIP, OPC, and classifier errors do not eagerly retry PPTX or
ODP. Public `DetectedFormat` and eager behavior are unchanged, and the
ordinary proven ODP native-owner handoff/reparse remains.

Path `FileSource` captures `SourceVersion`, preflights the caller's exact
`max_input_bytes`, and uses a same-source bounded `Bytes` fallback instead of
pathname re-open or unbounded `fs::read`. Semantic conversion failure
rechecks freshness first; `Presentation` consumes retained bytes with exact
limits. Input/part-limit, malformed-ZIP, missing-manifest allocation,
wrong-family/polyglot precedence, extensionless bounded path,
reserved-namespace, and freshness/cancellation regressions remain covered.

The constrained validation used an 8 GiB virtual-memory ceiling, one Cargo
job, disabled incremental/debug compilation, and one disk target. The PPTX
check passed, the combined PPTX/ODP library test passed `48/48` with one test
thread, and formatting passed. The final target was 674 MiB with
approximately 15 GiB host availability and saturated swap; no additional
pressure or OOM was observed. No speed, RSS, or OOM-prevention claim follows.
DOCX/non-Unix/ODT/ODP prepared-package/public eager-smart/selected-part
materialization seams remain open.
## Change 0356 update

Change 0356 closes the DOCX path-ingress and OPC typed-error hotspot for the
implemented scope. One `FileSource`/`SourceVersion` now spans ODT arbitration,
DOCX ownership, and bounded fallback on Unix and Windows; portable fallback is
length-checked and freshness-checked, and no pathname reopen or unbounded
`fs::read` remains. DOCX `OtherOoxml`/`DisabledOtherOoxml` are terminal, while
only genuine missing-manifest/no-match probes reclaim the fallback allocation.
The OPC boundary types allocation, six limit resources, raw I/O, and preserves
cancellation/execution/source freshness ordering across archive, catalog,
selected-stream, and preservation-index paths. Caller-sized physical result
buffers use typed fallible reservation and release the part reservation on
admission failure; this is correctness/resource safety only, with no
performance or OOM claim. `performance_claim: none`.

Residual hotspots are the eager public smart detector, the neutral 2 GiB
materializing fallback, non-Unix ODT policy differences, lower-family probe
input-limit behavior, ordinary ODT limits, `parts_by_name` casing, and
selected-Part materialization. Evidence is recorded in [Change 0356](changes/0356-docx-source-path-and-opc-errors.md).

## Change 0357 update

Change 0357 closes the implemented Workbook and Presentation filesystem
fallback policy boundary. OOXML and uncertain/polyglot candidates use caller
limits capped by a neutral 2 GiB ceiling, while ordinary canonical ODP/ODS,
content-derived renamed ODP, OLE, and generic non-ZIP fallback use the neutral
2 GiB policy. Unknown or missing-content-types ZIPs and uncertain ODF/OOXML
catalogs remain caller-limited. PPTX, ODP, native PPT, and bounded `Bytes`
arbitration share one `FileSource`/`SourceVersion`; freshness is checked on
that source, and wrong-family `OtherOoxml`/`DisabledOtherOoxml` is terminal.
The ODF catalog detector now has a neutral-budget helper for its checked input,
compressed, entry, and total ceilings. This is correctness/resource evidence
only (`performance_claim: none`).

Serial evidence is `15/15` focused ODF detection tests with `260` filtered,
`6/6` catalog arbitration tests, `82/82` `litchi` `pptx,odp,ppt` tests, and
`84/84` `litchi` `ods,xlsx` tests, with quiet `pptx`, `odp,ppt`, `odp`, `ppt`,
and `xls,xlsx` checks passing. Two initial `Arc<FileSource>` to
`Arc<dyn ReadAt>` compile errors were corrected before final validation. One
8 GiB process ceiling, one Cargo job, disabled incremental/debug compilation,
one disk target, and one test thread were used; the target's final/peak
observed footprint was 1.3 GiB, host availability approximately 14 GiB with
133 GiB disk free, and swap was exhausted. No parallel build or OOM occurred.

Residual hotspots are the eager public `DetectedFormat`, fully materialized
neutral fallback, flat ODF MIME decode before strict bounding, infallible
Presentation aggregate `Vec`/`join`, portable same-size identity, native PPT
probe-time mutation coverage, `Current User` plus `Workbook` OLE classifier
inconsistency, applicable prepared ODP reparsing, OPC case lookup, and
selected-Part materialization. Evidence is recorded in [Change 0357](changes/0357-workbook-presentation-two-ceiling-policy.md).

## Change 0358 update

Change 0358 is a rejected and reverted XLS worksheet-span hotspot experiment.
The candidate used stateless 64 KiB/1,024-item consecutive payload spans and
no CFB API. Correctness passed while present (`15/15` Python driver, `9/9`
span, `46/46` source-backed, `1021/1021` litchi-xls, `7/7` CFB cursor, and
`9/9` fragmentation). The serial six-selector ABBA retained 12,000 samples
across all 24 groups with 20 warmups per fresh child, one child at a time, CPU
2, a 2 GiB child cap, no retries, and one Cargo build lane.

The mechanism changed one-cell work by `+316` read bytes, `-79` reads, and
`-158` version calls, while the claim-bearing A1 -> B1 and A2 -> B2 p50/mean
deltas were `+4.984845886382849%`/`+4.771328093073383%` and
`+5.785582423178705%`/`+5.78027327071032%`. Exactly five p99 gates failed:
FileSource/list A1 -> A2 `+6.59542478684531%`, FileSource/list A2 -> B2
`-6.916640348285569%`, FileSource/one-cell B1 -> B2 `+8.748517200474495%`,
AtomicFile/one-cell A1 -> A2 `-5.699947129465672%`, and AtomicFile/one-cell
B1 -> B2 `+6.439283716879541%`. No gate narrowing or rerun was made; production
and candidate tests were reverted. The approximately 7.4M evidence is
retained in [Change 0358](changes/0358-xls-worksheet-span-batching-rejected.md)
and the OOM-bounded serial driver is reusable evidence infrastructure.

The target peak/final observed footprint was 1.9 GiB, host availability was
approximately 14 GiB with 132 GiB disk free, and swap was exhausted. No
latency, RSS, allocation, physical-I/O, or OOM-prevention claim follows.
The XLS freshness optimization queue remains open.

## Change 0359 update

Change 0359 implements the bounded transport foundation for the selected-cell
XLSX hotspot. `soapberry-zip` and `litchi-opc` now provide callback-scoped
verified decoded readers with fixed 16 KiB decoder scratch, finite interrupted
read retries, drain-to-EOF size/CRC/compressed-consumption verification, typed
callback-secondary errors, source/cancellation fences, and no `PartData` or
payload-cache admission. Archive-wide strict-layout proof allocation is
pre-existing indexed state and is not included in the 16 KiB scratch claim.

Serial validation passed ZIP `4/4` focused and `319/319` library tests, plus
OPC `6/6` focused, `277/277` library, `13/13` accounting integration, and
`6/6` source-reader integration. The on-disk target was 381 MiB; no parallel
build ran. This foundation does not yet replace XLSX worksheet/store
materialization. Streaming MCE/x14ac, full-EOF worksheet semantics, bounded
shared-string/style lookup, and measurement remain open. See [Change 0359](changes/0359-callback-scoped-verified-decoded-readers.md);
`performance_claim: none`.

## Change 0360 update

Change 0360 removes the next architectural blocker for selected-cell XLSX
streaming by adding a bounded MCE event processor. Raw observers see inactive
branch elements needed for later x14ac marker checks, while active observers
receive only selected semantic events. The parser reaches EOF after callback
errors and applies finite token, event, attribute, context, namespace, depth,
choice, and name bounds without building a normalized document.

Focused `11/11`, library `223/223`, and existing integration `1/1` tests passed
in one serial build lane. The target was 267 MiB; no parallel build ran. This
is bounded streaming rather than fixed-memory or OOM-safe processing because
quick-XML internals, decoded values, collection overhead, and callback-owned
state remain outside the strict byte estimate. XLSX x14ac/worksheet consumers
and selected-cell scanning remain open. See [Change 0360](changes/0360-bounded-streaming-mce-events.md);
`performance_claim: none`.

## Change 0361 update

Change 0361 implements the bounded streaming x14ac raw and active observer
foundation in `litchi-xlsx`. The MCE raw observer sees ordinary and alias
duplicates before generic duplicate validation. Semantic `NonConformant` or
`MustUnderstand` results can use raw-only one-pass recovery; a later XML,
input, or limit error becomes primary while the typed prior semantic error is
retained. The MCE and `AlternateContent` x14ac byte-compatibility branch now
streams while the plain fast path remains unchanged.

The branch uses a fixed 8 KiB `InterruptedRetryReader`, at most eight
interrupted-read retries, and the existing bounded stream limits. MCE recovery
`7/7`, raw attributes `4/4`, x14ac focused `12/12`, worksheet `35/35`,
`litchi-ooxml-common` library `234/234`, and `litchi-xlsx` library `813/813`
tests passed. x14ac `capture_rows=true` can retain a `BTreeMap` up to
configured `ROWS`; quick-XML and observer allocations remain outside the
fixed input-buffer claim. This is not selected-cell or full-worksheet
streaming and makes no latency, RSS, or OOM-safety claim. See [Change 0361](changes/0361-bounded-streaming-x14ac-observers.md);
`performance_claim: none`.
## Change 0383 update

This change is correctness and bounded-resource evidence for the PPTX
source-backed cross-slide copy hotspot, not a performance optimization. The
new chart leaf is admitted only as a direct ordinary graphic-frame leaf with a
single internal canonical `/ppt/charts/` relationship-free target, matching
dialect/content type/root, and no chart dependency graph. Distinct chart parts
are copied once while each source binding remains represented; deterministic
canonical part/rId allocation and exact namespace-resolved rewriting preserve
the rest of the slide and package byte structure. The image boundary from 0382
is unchanged.

ChartEx, workbooks, `externalData`, style/color, chart drawing/userShapes,
outbound or external links, malformed/ambiguous/nested hosts, stray chart
namespace content, MCE/DTD/PI, unresolved namespaces, stale/foreign/signed or
limit-invalid inputs, cancellation failures, and unsupported collisions are
refused before publication. Source/destination preservation, partial sinks,
and no durable inverse remain part of the contract.

The evidence is focused `52/52`, isolated cancellation `1/1`, default lib
`531` plus one named filtered test, all-features primary lib `533` plus one
named filtered test, integrations green with three exact exclusions, doctests
`6` passed/`2` ignored, Clippy green with the inherited three allowances, and
boundary `64/240/14`. No latency, RSS, allocation, throughput, or other
performance measurement is available or authorized. The serial run used one
Cargo process/job, a 6 GiB virtual-memory cap, and a 10 GiB `MemAvailable`
gate; these are resource-capped/OOM-mitigating controls, not proof of OOM
prevention.

`performance_claim: none`; `claim_authorized: false`.

The change-0484 focused metadata2 diagnostics now confirm syscall
amplification: authored-heavy file input makes 3,735,939 `statx` calls
versus 12 for owned input, with 3,735,927 attributed to the exact source
descriptor in the separate raw trace. Counts include setup, one warmup and
one measured sample. `FileSource::len` and `version` query metadata on each
call; next work must locate redundant callers and preserve freshness policy.
No production optimization or before/after speedup is claimed by this batch.
