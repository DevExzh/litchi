# Performance program phase report

Date: 2026-08-17
Branch: `feat/office-format-completeness`
Historical production base for the original measured tranche:
`6df5d4a1fbe53a8216e63f24cc1392be60b714a8`

This report summarizes the measured implementation tranches to date. It is not a
claim that the end-to-end performance program or CRUD scenario matrix is
complete. The reproducible environment, original substrate baseline, corpus
definitions, commands, and profiler limitations are in
[`BASELINE.md`](BASELINE.md); raw reports are under [`results/`](results/).

## Current-HEAD resource evidence (change 0115)

The [current-HEAD resource record](results/resource-profile-current-head-0115.json)
adds process resource and physical-observation context for five representative
CRUD paths plus the two existing explicit scaling selectors.  It measures
revision `be500459961471659f65c180de0e5fe98bc14e3a` with the release harness
identified by SHA-256
`1cbb2340eae13f4ed49d5baa27532e1f9b31d5781036bb2a302837bcd2210f5c`.
The result is deliberately current-HEAD evidence only; it does not compare two
revisions or claim an optimization.  The dirty-worktree release build completed
successfully, so the exact binary hash/size and successful build are recorded.
Because the original run retained only a post-build dirty snapshot and no
pre-build or bounded untracked-content identity, it is explicitly
`build_succeeded_source_snapshot_only`, not a complete or cryptographic
source-to-binary binding; a clean pre/post-snapshot rerun is required.

The three-sample harness p50s were 59.68 ms for OPC source one-Part
publication, 33.26 ms for managed XLSX batch edit/save, 10.02 ms for medium RTF
streaming, 0.141 ms / 0.375 ms for the paired CFB MiniFAT/FAT selective reads,
and 156.31 ms for the CFB same-length atomic-save operation.  `/usr/bin/time`
maximum RSS was respectively 118,176, 66,132, 30,080, 30,336, and 110,884
KiB.  Heaptrack process totals are retained in the JSON; they include startup,
corpus construction, and profiler overhead and are not per-operation allocator
attribution.

The harness counters confirm the measured logical scope: OPC used 549 source
reads / 16,785,201 bytes and one ordinary payload materialization; managed
XLSX used 225 reads / 4,230,793 bytes and six materializations; RTF retained
zero output bytes with a 37-byte authoring window; CFB save used 1,825 logical
reads / 84,838,500 bytes and published 16,913,408 bytes with one changed span.
The aggregate also records all six perf counters, read/write syscall-size
histograms, tool versions, artifact SHA references, commands, corpus hashes,
and null/unsupported states where a tool is unavailable.

The 1/2/4/8/available execution-context widths were classified
`nonideal_or_measurement_noise` for both many-small OPC and CFB: raw p50 showed
no measured speedup at the 12-worker endpoints (OPC 567,473 -> 789,610 ns;
CFB 224,090 -> 225,201 ns), and out-of-range Amdahl fractions are represented
as null estimates with validity flags.  This is a bounded host observation, not a
conclusion about all documents or hardware.  Syscall counts are whole-process
`strace` observations and must not be interpreted as logical source bytes,
decompressed bytes, recompressed bytes, or memory-copy volume.  Cold-cache,
remote-range, before/after, and allocation-attribution claims remain open.

## Current stable tranche

The original stage-1 results below remain historical evidence. The current
harness contains **320 selectable cases**; 200 was the count before the
opt-in ODF `mimetype` repair-plan selector was added. The measured 36-default-case,
198-default-record tranche remains historical evidence; newer selectable cases
do not inherit its performance results. That measured tranche includes six
opt-in simulated-range cases, two opt-in scaling cases, one opt-in XLSX
commit/read attribution case, four opt-in opaque-heavy common OLE2
stage/control cases, one opt-in source-backed OPC one-Part publication case,
one opt-in source-backed DOCX semantic publication case, one opt-in
source-backed media-rich PPTX semantic publication case, four opt-in matched
same-slide/multi-slide PPTX batch cases, six opt-in media-rich ODT
paragraph/line-break/inline-run/hyperlink/insertion/removal publication cases,
two opt-in matched cross-slide ODP text-box publication cases, 20 opt-in matched XLSX
calculation-metadata/defined-name/page-break/page-margin/print-options/page-setup/sheet-protection/data-validation/auto-filter/conditional-formatting publication cases, 16 opt-in DOCX/PPTX
semantic cases, 13 opt-in RTF semantic case names across four
capability-bounded variants (39 tiny / 70 tiny-plus-large rows), 24
shape-selected ODT/ODS/ODP semantic cases, twelve fixed media-rich ODF cases,
and 22 opt-in native DOC/XLS/PPT semantic/phase-attribution cases. It
is still not broad program or CRUD coverage.

## Native DOC owner/public-reader phase attribution (change 0160)

[Change 0160](changes/0160-doc-owner-public-phase-attribution.md) adds the
opt-in `doc_owner_public_phases` selector over the exact deterministic tiny,
large, and payload-heavy DOC writer corpora. A feature-gated, content-free
observer emits ordered boundaries for strict owner validation, complete public
reader validation, exact-source retention, in-memory owner rendering, exact
no-op detection, and patch construction; the format crate owns no clock.
The harness separately times edit construction, replacement staging, outer
operations, and final output materialization, then checks attributed plus
unattributed time against each measured lifecycle.

The focused two-shape test and three-shape debug smoke pass exact semantic
reopen, no-op, patch/inverse/stale, malformed/typed refusal, output-hash, event
order/cardinality, and untouched-CFB-stream gates. A subsequent clean release
distribution at exact revision `ab333008d3`, pinned to CPU 2, used four fresh
processes per shape with 20 warmups and 200 retained samples per process.
Tiny/large/payload-heavy lifecycle p50 was 0.081/1.157/44.227 ms.
Initial-plus-final complete public-reader validation p50 was
0.016/0.598/20.721 ms; patch p50 was 0.026/0.165/8.413 ms, and replacement
staging p50 was 0.014/0.174/7.470 ms. Every untimed case-level gate passed in
all 12 reports; all 2,400 timed samples passed arithmetic, event, and output
checks.
Across-process lifecycle p50/mean spread was at most 2.98%/3.76%; two tiny
subphase means crossed the 5% review trigger, but the phase rank did not change.

This names the largest measured phases in the exact deterministic distribution
but is not a control/candidate comparison. Synchronous observer overhead is
included, and no optimization, speedup, physical-I/O, allocation/RSS,
cold-cache, filesystem, or real-producer result is accepted. See the
[summary](results/doc-owner-public-phases-0160-summary.json) and
[raw-artifact manifest](results/doc-owner-public-phases-0160.sha256).

The immediate clone-elision experiment did not survive end-to-end guards.
[Change 0161](changes/0161-doc-public-validation-borrow-rejected.md) borrowed
the existing DOC byte vector for both public-reader validations without
removing either validation layer. In clean CPU-2 A/B/B/A release runs, tiny
lifecycle p50 improved 3.20%/3.24%, large regressed 3.06%/7.31%, and
payload-heavy directions disagreed at -0.18%/+2.52%. Large p95 regressed
37.52%/14.49%. All semantic and preservation gates passed, but the >5% p50
trigger and tails reject the candidate. Its temporary branch and code were
removed; only the [negative-result summary](results/doc-public-borrow-0161-summary.json)
and raw evidence remain.

## RTF standalone-picture CRUD evidence (change 0162)

[Change 0162](changes/0162-rtf-picture-crud-evidence.md) adds two opt-in
selectors for the already committed exact-source picture transaction APIs:
bounded same-length payload replacement and bounded complete-group removal.
Tiny/medium/large generated ASCII corpora contain 2/8/64 alternating PNG/JPEG
groups with deterministic mixed-case hexadecimal digits, spaces and newlines.
Replacement selects 1/7/63 pictures and leaves one group unselected; removal
selects 1/4/32 alternating positions.

The harness constructs expected bytes independently by replacing only source
hexadecimal digit positions or deleting exact group spans, then requires public
commit bytes to match. It reports open, one bounded staging call, commit,
fixed-memory hashing-sink publication and complete lifecycle vectors. Untimed
gates cover semantic reopen, visible text and unselected raw preservation,
same-payload no-op identity, volatile and deterministic-durable patch
forward/inverse, stale/foreign/refusal checks, partial/zero sinks and exact
digests. The six-record all-shape debug smoke passes every gate with exact sink
byte counts and zero retained output bytes.

This raises the selectable matrix from 303 to 305 without changing the default
36 cases / 198 records. It is not release latency evidence and makes no
speedup, allocation/RSS, total-memory, physical-I/O, real-producer, compressed/
binary/nested picture, image-rendering or general rich-media claim.

## XLSX scalar-cell clear/remove evidence (change 0163)

[Change 0163](changes/0163-xlsx-cell-clear-remove-evidence.md) adds four
opt-in eager/source-backed lifecycle selectors over the existing medium and
dense/sparse four-worksheet numeric corpus. Each targets `Sheet1!A1` and
separately reports open, planning/staging, commit, sequential publication, and
lifecycle phases. Eager uses `WorksheetEdit`; source-backed uses the
positional cell-values editor. Clear retains an empty `<c>` owner, while remove
deletes that owner. A fixed 64-KiB windowed hashing sink retains zero output
bytes; generic logical source/materialization counters are recorded.

Semantic/package/no-op, volatile source-patch, stale/foreign, and source-backed
raw-unselected-member gates are outside timing. The source-backed patch has no
durable wire contract, so durable-source-patch evidence is explicitly absent.
The four selectors raise the matrix from 305 to 309 without changing the
default 36/198 tranche. This is correctness/phase/counter evidence only, with
no latency, allocation/RSS, physical-I/O, cold-cache, decompression, or
real-producer claim.

## RTF ordinary paragraph split/adjacent merge evidence (change 0164)

[Change 0164](changes/0164-rtf-paragraph-split-merge-evidence.md) adds two
opt-in selectors, `rtf_semantic_split_paragraph_save` and
`rtf_semantic_merge_paragraph_save`, over the exact generated plain RTF
lifecycle corpus at tiny, medium, and large shapes (24, 200, and 10,000
ordinary paragraphs). Split inserts exactly one canonical five-byte `\\par `
boundary at a checked interior offset; merge removes only the authenticated
adjacent five-byte boundary. Expected output is an independent raw splice,
with unchanged surrounding bytes and complete semantic paragraph projection
checked separately.

Each selector reports open, stage, commit, fixed 16-KiB windowed sequential
publication, and complete lifecycle vectors. The publication sink retains zero
output bytes and hashes the complete accepted stream; the candidate transaction
is not covered by that sink window. Untimed gates include exact no-op/source
identity, volatile and deterministic durable forward/inverse, stale/foreign
source refusal, forged result-artifact refusal, bounded invalid/unsafe/
protected refusal, partial/zero sink failure, and deterministic source/output
hashes. The selectors raise the matrix from 309 to 311 while preserving the
historical default 36 cases / 198 records.

This is correctness, phase, and sequential-sink evidence only. It makes no
latency, speedup, transaction-memory, allocation/RSS, physical-I/O,
cold-cache, source-backed, real-producer, or general rich-RTF claim. The
literal-ASCII root-level ordinary-body closure and existing native focused
tests remain the authority for unsupported and forged-input boundaries.

## Native DOC lazy fingerprints and same-lineage patch replay (change 0165)

[Change 0165](changes/0165-doc-lazy-fingerprint.md) records the inline lazy
DOC snapshot fingerprint cache, immutable same-lineage patch fast path, and a
bounded descriptive comparison on
the exact deterministic tiny, large, and payload-heavy native-DOC lifecycle.
Patch construction no longer scans complete before/after artifacts solely to
populate the diagnostic FNV-1a values. Same-lineage no-op/apply first checks
`Arc` allocation identity and length; independently reopened sources still do
lazy fingerprint comparison followed by exact bytes, preserving collision,
stale-source, inverse, and failure-atomic semantics. The fingerprint
accessors are intentionally non-`const` because first demand may initialize
the cache.

The historical `measured_total_ns` lifecycle boundary remains unchanged.
Same-lineage apply and first fingerprint demand are explicit workflow
extensions. The final clean control source is
`d6818e290aa77fd7666b7b16ee6908319d0f332b`, the candidate is
`5dd813b1e108e253457ccb6c504c125c2becc1c6`, and their release binary
SHA-256 values are
`344c0504c254109ee6b4361e375599d187f8a12333abb44f207d837af259ef8c` and
`c95e6c6004cbd725c789597566a81c0897ab6915ecd7c274deab222d134b3fd3`.
Both builds were clean exact-revision builds. Clean CPU-2 release ABBA used
20 warmups and 500 retained samples per shape and leg, retaining 6,000
lifecycle samples. Descriptive lifecycle p50/mean/p95 positive-faster deltas were
`+33.77/+35.19/+38.94` and `+33.21/+34.76/+39.67` tiny,
`+12.28/+12.59/+17.53` and `+13.81/+13.55/+11.68` large, and
`+17.33/+17.09/+16.58` and `+17.82/+17.75/+16.25` payload-heavy. With
immediate fingerprint demand included, workflow p50/mean/p95 positive-faster deltas are
`+14.56/+16.34/+22.24` and `+13.89/+15.80/+21.90` tiny,
`+4.50/+4.82/+10.24` and `+5.83/+5.64/+4.26` large, and
`+6.55/+6.41/+6.26` and `+7.08/+7.08/+6.33` payload-heavy.

The isolated edit-patch/same-lineage-apply extension spans approximately
99.6-99.99% across the reported p50/mean/p95 deltas versus the eager-fingerprint
control, while the deferred
source-plus-target scan is explicit and lands at roughly 25.7 us, 164 us, and
8.37-8.39 ms for tiny, large, and payload-heavy candidates. Mandatory DOC
no-op, one-edit, and open guards remain within policy: no-op p50 improves
`+78.84%/+79.89%` tiny and `+71.08%/+70.40%` large, one-edit improves
`+37.23%/+40.81%` and `+20.45%/+19.79%`, and open is
`-3.52%/+0.13%` tiny and `+0.55%/-1.80%` large. Neighboring XLS one-edit
and open guards are mostly neutral or improved; XLS no-op remains noisy.
Representative final heaptrack and `/usr/bin/time` observations are
descriptive whole-process boundaries: allocation calls are 50,677 in both
revisions, peak heap is 128.28M in both, profiler RSS is 145.14M versus
142.81M, and A1/B1/B2/A2 maximum RSS is 138160/138024/138028/138032 KiB.
These are not operation-only attribution.

The final same-implementation lifecycle p50/mean drift is control
`-1.18%/-1.41%` tiny, `+0.25%/-0.42%` large, and `+0.48%/+0.72%`
payload-heavy; candidate drift is `-0.34%/-0.75%`, `-1.50%/-1.51%`, and
`-0.12%/-0.08%`, respectively. The paired directions remain positive, but
the result is still limited to the named host and corpus.

The same-implementation drift disclosure and all raw vectors are retained in
the [machine-readable summary](results/doc-lazy-fingerprint-0165-summary.json)
and [release manifest](results/doc-lazy-fingerprint-0165-manifest.json). The result
is limited to this native-DOC host/corpus/workflow. It makes no speedup, physical-I/O,
cold-cache, real-producer, generic-DOC, total-memory, operation-only
allocation/RSS, or CRUD-completeness claim; the wider native owner/public
validation and format matrix remain open.

## XLSX row-visibility provenance reuse (change 0167)

[Change 0167](changes/0167-xlsx-row-visibility-provenance-reuse.md) makes the
existing source-backed row-visibility publisher consume its embedded
cell-values patch through the established tri-state provenance proof. Matched
lineage/version publication no longer reloads the selected worksheet,
reparses its cell store, or rescans row tags. Mismatched provenance still
refuses, unavailable provenance still takes the conservative semantic fallback,
and every path retains mandatory OPC overlay validation and sequential
publication.

A >8 MiB worksheet regression permits exactly the one selected-member read
required by OPC publication; the former second semantic read would fail.
Focused row/cell suites, 765 XLSX library tests, strict warning/deprecation
Clippy, format/diff checks, and an independent adversarial review are green.

Clean CPU-2 `A1 control, B1 candidate, B2 candidate, A2 control` release runs
used 20 warmups and 500 retained samples for four medium/large hide-one and
unhide-256 records per leg. Every publication p50/mean/p95/p99 pair is lower
for the candidate, with descriptive reductions of 50.42%-68.23%. Logical
`ReadAt` topology is unchanged while source-version checks fall by 13 per
sample. The 5% same-implementation gate fails: maximum absolute drift is
34.80% for control large/unhide publication p99 and 10.23% for candidate
medium/hide complete-workflow p50; first-pair medium hide/unhide complete-
workflow p99 regresses 6.95%/2.69%. The production work
elimination is retained, but no acceptance-grade end-to-end latency, tail,
allocation/RSS, physical-I/O, decompression, cold-cache, or producer claim is
accepted. Raw vectors and sidecars are bound by the
[summary](results/xlsx-row-visibility-provenance-0167-summary.json) and
[manifest](results/xlsx-row-visibility-provenance-0167-manifest.json).

## XLS numeric validation fusion (change 0168)

[Change 0168](changes/0168-xls-numeric-validation-fusion.md) moves native XLS
Number/RK/MulRK semantic target validation onto the exact composed positional
view already created, reopened, and range-checked by the common CFB planner.
The owner callback runs before CFB's final complete source/target fingerprint
fence and preserves native semantic errors. No-op plans skip it. All
structural, source-precondition, protection, macro, encryption/signature,
numeric readback, stale-source, partial-output, and publication checks remain.

The former two post-plan `composed_source()` calls each performed a complete
source scan. Their removal saves 33,991,680 logical source bytes and 34
one-MiB reads per effective Number sample, or 405,504 bytes and two reads per
RK/MulRK sample. These deterministic counts describe the in-memory code path,
not physical I/O; the existing source counters cover owned-source ingress.

Clean CPU-2 `A1 control, B1 candidate, B2 candidate, A2 control` release runs
used 20 warmups and 500 samples per family. Complete-workflow p50/mean/p95/p99
values are descriptively 19.22%-28.16% lower and semantic-commit values
37.58%-48.04% lower in both paired directions. The 5% stability gate fails:
maximum absolute control drift is 10.56% and candidate drift is 9.81%. The
production work elimination and exact correctness artifacts are retained, but
no acceptance-grade latency, tail, allocation/RSS, physical-I/O, cold-cache,
or producer improvement is accepted. See the
[summary](results/xls-numeric-validation-fusion-0168-summary.json) and
[manifest](results/xls-numeric-validation-fusion-0168-manifest.json).

## XLSX streaming hierarchical-budget charges (change 0169)

[Change 0169](changes/0169-xlsx-streaming-budget-charge.md) removes a transient
owned ancestor vector from every cumulative `Budget::consume` call and stores
up to four charged nodes inline for releasable reservations. Deeper caller-
defined hierarchies still spill; public charge order, rollback, commit/drop,
errors, limits, and atomics remain unchanged.

The existing one-sheet `xlsx_streaming_create` selector supplied the measured
hot path and no selector/schema changed. Clean CPU-2 release A/B/B/A runs used
20 warmups and 200 samples for each shape. Medium and large p50/mean/p95/p99
improve in both paired directions by 1.05%-9.76%; tiny p50/mean/p95 also
improve, while tiny p99 regresses 1.81%/2.75% and is withheld. Matched
whole-process Heaptrack captures record 38,672,384 -> 19,794,608 allocation
calls and 22,545,902 -> 6,815,902 temporary allocations, with unchanged
225.45M peak heap. GNU Time RSS directions disagree, and branch misses increase
from a sub-0.25% absolute rate.

Archive and worksheet hashes, row/cell/text counts, logical sink counters,
zero retained output, and the 4 KiB row-authoring window remain exact. The
accepted scope is warm in-memory, synthetic, one-sheet inline-scalar XLSX
creation. It is not total-memory, physical-I/O, cold-cache, multi-sheet,
shared-string/style/formula/date, real-producer, or broad `Budget` evidence.
See the [summary](results/xlsx-stream-budget-charge-0169-summary.json) and
[manifest](results/xlsx-stream-budget-charge-0169-manifest.json).

## XLSX streaming UTF-8 escape runs (change 0170)

[Change 0170](changes/0170-xlsx-streaming-escape-runs.md) batches contiguous
ordinary UTF-8 between the five XML entities, skips scalar counting when byte
length proves the finite character bound, and formats each row number once.
The worksheet XML, ZIP write boundaries, archive bytes, limits, and public API
remain unchanged.

Clean CPU-2 release A/B/B/A used 20 warmups and 300 samples per shape. Large
p50/mean/p95/p99 improve by 5.02%-6.99% in both paired directions. Medium
p50/mean/p95 improve by 4.45%-5.52%, and tiny p50 by 5.03%/7.74%. Tiny
mean/p95/p99 and medium p99 are withheld because paired directions disagree.
Every accepted statistic passes its 5%/10%/15% same-implementation drift tier.
Archive/worksheet hashes, semantic reopen, rows/cells, sink topology, zero
retained output, and the 4 KiB authoring window remain exact.

Matched whole-process large-shape counters record 6.15%-6.19% fewer
instructions and 10.54%-10.57% fewer branches; branch misses regress
8.99%-14.37%. GNU Time RSS differs by at most 380 KiB and the candidate is
higher in both paired directions. These are descriptive process boundaries,
not operation-local allocation, CPU, memory, physical-I/O, or cold-cache claims.
See the [summary](results/xlsx-stream-escape-0170-summary.json) and
[manifest](results/xlsx-stream-escape-0170-manifest.json).

## Legacy CFB owner-validation fusion (change 0171)

[Change 0171](changes/0171-cfb-owner-validation-fusion.md) moves source-backed
DOC paragraph, PPT shape-text, and XLS worksheet-visibility semantic readback
onto the exact composed CFB view already owned by the common planner. The
existing callback remains inside the complete final fingerprint fence; CFB
reopen/range validation, no-op behavior, security checks, publication, and
atomic-save fences are unchanged.

Each effective transaction removes one complete source scan and one
source/target SHA-256 pair. The measured 2,135,552-byte XLS corpus therefore
avoids three logical one-MiB reads per scalar or 64-worksheet batch. Clean
CPU-2 release A/B/B/A runs used 20 warmups and 300 samples. The source-backed
64-worksheet complete workflow improves p50/mean/p95 by 12.51%-15.38% in both
directions. Scalar and batch semantic staging/plan p50/mean/p95 improve by
31.44%-33.16%.

The scalar complete workflow is withheld because its eager guard shifts by a
similar or larger amount. Batch p99 exceeds the control drift gate, and the
publication phase regresses or disagrees. DOC/PPT latency and allocation, RSS,
physical-I/O, cold-cache, and producer claims are also withheld. See the
[summary](results/cfb-owner-fusion-0171-summary.json) and
[manifest](results/cfb-owner-fusion-0171-manifest.json).

## Immutable CFB numeric-plan publication (change 0172)

[Change 0172](changes/0172-cfb-owned-numeric-publication.md) preserves the
native XLS plan-only snapshot's immutable `Arc<[u8]>` ownership through an
explicit CFB ingress. Only direct sequential publication consumes the private
proof. It skips the redundant complete pre/post fingerprint scans but retains
the 64 KiB emission read, source/target hashes, exact sink progress, partial
output and flush. Generic `ReadAt`, checked composed views and atomic save keep
their previous complete fences.

The deterministic reduction is two source scans: 33,991,680 logical bytes and
34 one-MiB reads for Number, or 405,504 bytes and two reads for RK/MulRK. Clean
CPU-2 release A/B/B/A used 20 warmups and 500 samples. Number complete-workflow
p50/mean/p95/p99 is 37.54%-39.00% lower and direct publication 64.44%-65.63%
lower. RK/MulRK complete workflow is 36.63%-38.96% lower and publication
p50/mean/p95 is 65.54%-66.76% lower. All accepted directions agree and pass
the 5% drift gate.

RK/MulRK publication p99 is withheld because control drift is 5.28%. Atomic
save is deliberately unchanged. The measured source is a complete owned
in-memory artifact, and process RSS is slightly higher in both paired
directions, so no allocation/RSS, physical-I/O, cold-cache, producer,
compression, or throughput claim is made. See the
[summary](results/cfb-owned-numeric-publication-0172-summary.json) and
[manifest](results/cfb-owned-numeric-publication-0172-manifest.json).

## Native XLS comment publication fusion (change 0173)

[Change 0173](changes/0173-cfb-comment-publication-fusion.md) validates
effective existing-comment edits on the exact composed CFB view already held
inside the final fingerprint bracket. It also preserves the snapshot's sealed
immutable ownership into direct sequential publication. The combined
code-derived reduction is three full scans of the 16,995,840-byte artifact,
50,987,520 logical bytes, 51 one-MiB reads, and three source/target SHA-256
pairs per transaction. The 64 KiB emission/hash pass and atomic-save fences
remain.

Clean CPU-2 release A/B/B/A used 20 warmups and 500 samples. Accepted paired
reductions are 45.54%-47.19% for scalar complete-workflow p50/mean/p99,
30.78%-32.42% for scalar semantic staging/plan, 59.15%-61.03% for scalar direct
publication, and 30.53%-32.57% for the 256-comment semantic phase. Scalar
complete p95 misses the matched eager guard by 0.027675 percentage points;
batch complete/publication has excessive candidate drift and is withheld.

The evidence is warm, in-memory, generated XLS. It does not establish
allocation/RSS, physical-I/O, cold-cache, independent-producer, compression,
throughput, atomic-save, or broader comment-lifecycle improvements. See the
[summary](results/cfb-comment-fusion-0173-summary.json) and
[manifest](results/cfb-comment-fusion-0173-manifest.json).

## Managed XLSX source-editor production freeze (change 0151)

[Change 0151](changes/0151-xlsx-managed-source-editors.md) freezes managed
source-backed constructors and ownership for eleven focused XLSX editors:
calculation properties, defined names, tab state, print options, page breaks,
page margins, page setup, sheet protection, data validation, auto filter, and
conditional formatting. The private `Managed(PartData)`/`Owned(Arc<Vec<u8>>)`
owner keeps managed cache reservations attached to snapshots and makes
managed-to-owned `Arc` escape a typed fallible boundary. Validated package
handoffs, direct selected-Part publication, raw preservation of unselected
members, exact no-op/signed/MCE/stale/cancellation/unknown-owner protections,
and representative one-byte-under `Resource::Memory` gates are retained.

This is a production correctness/resource-accounting freeze, not a measured
optimization. The validation run reported 765 green XLSX unit, integration,
and documentation tests, including 74 focused source-editor checks. It adds no
benchmark selector or performance artifact and makes no latency, allocation,
RSS/peak-memory, copy, decompression, cold-I/O, total-memory, hardware, or
real-producer claim.

## CFB same-target MiniFAT single-flight release ABBA (change 0152)

[Change 0152](changes/0152-cfb-same-target-singleflight-release-abba.md)
compares the final same-target MiniFAT single-flight revision `f46381c6f`
(introduced by `c270c8f3b`) with clean control `e486e4b1` in strict CPU-2
`A1 control, B1 candidate, B2 candidate, A2 control` order. Each leg used 20
warmups and 500 samples across 24 records, retaining 48,000 samples. All
correctness and logical source-event invariants passed. Existing concurrent
scenarios recorded 6,473 candidate versus 8,000 control logical source calls,
19.09% fewer.

Only this named source-event/correctness scope is accepted. At the 0152
revision the 291-name selector matrix was unchanged; change 0153 adds four RTF
selectors measured at the pre-staged publication-call interval, making that
matrix 295. Change 0154 adds six ODF publication selectors, making the current
matrix 301; change 0159 later made it 302, change 0160 made it 303, change
0162 made it 305, change 0163 made it 309, change 0164 made it 311, and
change 0166 makes it 315. No runtime
selector was added to 0152; only `cfg(test)` source-event acceptance and tests
changed. Local or generic latency,
allocation/RSS/peak memory, physical I/O/syscalls, cold-cache/device/network,
decompression, native semantic, OOXML, ODF, RTF, and iWork claims are
withheld. The root MiniStream cache and resource-accounting boundaries remain,
as do broader performance gaps. See the
[machine-readable summary](results/cfb-singleflight-abba-0152-summary.json).

## ODF content-COW publication release ABBA (change 0154)

[Change 0154](changes/0154-odf-content-cow-publication-evidence.md) adds six
matched ODT/ODS/ODP owned-rebuild and source-positional `content.xml`
publication selectors. A clean CPU-2 `A1 owned, B1 positional, B2 positional,
A2 owned` run used 20 warmups and 100 samples per record. Both pair directions
accept p50 improvements of 96.35%-96.63%; p95, p99, and mean agree, and the
largest absolute same-implementation p50 drift is 1.441%.

The timer covers a prepared owner/candidate publication call plus the same
fixed 16 KiB non-seek hashing sink. Semantic edit construction, archive
open/indexing, reopen, exact content/inventory, positional untouched-member raw
identity and order, no-op, limits, cancellation, source immutability, and
logical `ReadAt` replay are untimed. This is therefore a prepared in-memory
publication result, not end-to-end edit/save, allocation/RSS, physical-I/O,
decompression, cold-cache, filesystem, real-producer, or iWork evidence. See
the [summary](results/odf-content-cow-abba-0154-summary.json) and compressed raw
reports.

## PPTX additive-topology release ABBA (change 0158)

[Change 0158](changes/0158-pptx-additive-topology-release-abba.md) compares
clean control `e8a67b19e` with candidate `d900ae633` using the existing plain
and media-rich cross-presentation slide-copy selectors. The harness and locked
dependency graph are byte-identical. Four CPU-2 legs used 20 warmups and 200
retained samples per case, totaling 1,600 observations.

The candidate's owned-source OPC publisher raw-copies unchanged physical
members while appending generated Parts; the control rebuilds the complete
package. Total p50 improves 29.643%/26.196% on plain and 43.294%/43.604% on
media-rich in the two ABBA directions. Media-rich publication p50 improves
49.321%/49.680%, with p95/p99/mean agreement. Plain publication p50/mean is
accepted, but its p95/p99 claim is withheld because candidate same-revision
tail drift crossed the declared thresholds.

All semantic/topology/dependency/durable-patch/refusal gates pass. Matched
whole-process media-rich profiles show task-clock reductions of
42.399%/43.122%, cycle reductions of 42.583%/43.116%, and instruction
reductions of 46.686%/46.775%; maximum RSS is less than 0.5% higher and peak
heap is effectively unchanged. The result is restricted to canonical
generated owned-source prepared slide copy. Source-backed/cold-I/O,
decompression, generic OPC/PPTX, real-producer, and iWork claims remain open.
See the [summary](results/pptx-additive-topology-abba-0158-summary.json) and
[artifact manifest](results/pptx-additive-topology-0158.sha256).

| Change | Current evidence | Scope / limitation |
|---|---|---|
| XLSX row-start index | ABBA p50 geomean **-80.499%**, mean geomean **-79.962%**; full scan **+0.03%** mean; first cell **-1.31%** mean | Heap allocations **+17**, RSS **+0.25%**; narrow-range query only |
| Targeted OPC raw publication | Four-cell ABBA p50 geomean **-84.98%**; few-large/incompressible **-71.70%**; matched cycles **-69.21%** | Initial peak heap **+37.18%**, one-shot RSS **+22.26%** from retained source/provenance and a changed-payload copy; the copy is removed by the shared-payload follow-up below |
| Positional CFB/ZIP and explicit execution | Large-task p50 scaling at 12 CPUs: OPC **4.52x**, CFB **5.93x**; no hidden global Rayon | Many-small tasks regress at high worker counts; default/legacy paths remain serial |
| CFB selective exact-range read | MiniFAT source bytes **261,184 -> 36** (many-small) and **2,096,192 -> 36** (wide-root), one request in each; read-stage p50 **-95.1%/-94.8%** and **-99.2%/-99.2%** across ABBA directions; read-stage p95 **-94.4%/-94.8%** and **-98.9%/-99.1%**; total p50 **-8.4%/-14.2%** and **-6.6%/-11.9%** | FAT retains one 4 MiB request/call and paired read/total p50 changes stay within 5% control drift. That record makes no p95/p99 FAT, MiniFAT p99, cold-filesystem, simulated high-latency range, allocation, peak-RSS, or DOC/XLS/PPT semantic claim |
| CFB selective simulated-range read | With a harness-only 100 us + 25 us/request, 50 MiB/s, 64 KiB-ceiling model, both MiniFAT targets reduce selective-read work to one exact request. Total p50 improves **40.12%/39.99%** and **40.09%/39.82%** on many-small, and **41.96%/41.83%** and **42.00%/41.84%** on wide-root; p95 agrees in both directions | The 4 MiB FAT controls retain 64 requests / 4 MiB / 88 ms modeled read floor and stay near neutral. This accepts only configured simulator latency and request/byte/service-floor evidence; no real cold/network/device, production scheduling, allocation/RSS, or native semantic claim |
| CFB same-target MiniFAT repeat policy | Same-target source work changes from `[L,R,0...]` to `[L,L,...]`. In the 100 us + 25 us/request, 50 MiB/s, 4 KiB-ceiling model, aggregate total repeat-3 p50 improves **60.70-64.09%** and repeat-8 **55.86-63.67%** across both adjacent ABBA directions; p95/p99/mean agree | Local/per-invocation/bulk/concurrent distributions are withheld: later zero-source cache hits become exact target reads, and special-workload local tails contain reversing >5% review triggers with substantial control drift. No resource, physical-I/O, cold/network/device, or native semantic claim |
| CFB MiniFAT 4095-byte physical-run evidence | Focused matched legacy/positional controls record exact open/read/total timing, source calls/bytes/range sizes, 4095-byte payload hash, and request amplification across 64 logical mini-sectors | Correctness/request-amplification evidence only; no release latency, p99, cold/high-latency, allocation/RSS, physical-I/O, or native semantic claim |
| Ordinary-root DOCX source-path evidence | Eight opt-in eager/source filesystem selectors over the unchanged 200-paragraph/eight-incompressible-2 MiB-media corpus; untimed parity covers paragraphs, full text, tables, elements, and metadata, while exact source SHA plus logical OPC part/relationship/content-type/blob-hash gates cover package preservation, including all media hashes and source immutability; typed replays classify zero-payload-overlap open, complete main-range preparation for query selectors, and zero-overlap queries | Selectable matrix 245 -> 253; correctness/logical compressed-range evidence only; no latency, physical-I/O, decompression, allocation, RSS, cold-cache, ABBA, broad-security, or Markdown-performance claim |
| Source-backed OPC and DOCX/XLSX/PPTX facades | EOCD structural-open source bytes **-73.6% to -98.5%**; ordinary payload overlap zero | No latency claim: later EntryId/cache-diagnostic changes confound comparison and some cells exceed 5% variance |
| Source-backed PPTX selected-slide publication | Media-rich one-edit/save p50 **-97.12%**; atomic same-slide batch p50 **-97.45%**, materializations **229 -> 2**; atomic eight-slide batch p50 **-95.78%**, allocations **-32.54%**, materializations **229 -> 9**; byte-identical output | At most 32 existing slides with one bounded 256-selector shape-text operation each; MCE rewrites, relationships/topology changes and changed signed packages refuse before output |
| Source-backed XLSX calculation-metadata publication | Media-rich one-edit/save p50 **-99.2519%** (133.67x), mean **-99.2507%**; instructions **-77.78%**; materializations **12 -> 1**; byte-identical output | Existing `xl/workbook.xml` calculation properties/features only; cells, formulas, cached results, chains, relationships and topology remain outside the capability |
| Source-backed XLSX defined-name publication | Media-rich catalog edit/save p50 **-97.84%** (46.32x), mean **-97.81%**; instructions **-78.45%**; materializations **12 -> 1**; byte-identical output | Complete direct `definedNames` catalog only; protected/MCE/unknown catalogs, sheet topology, cells, formulas, relationships and changed signed sources remain outside the capability |
| Source-backed XLSX page-break publication | Media-rich one-edit/save p50 **-97.86%** (46.65x), mean **-97.86%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's page-break collections only; cells, formulas, styles, relationships, topology, and changed signed sources remain outside the capability |
| Source-backed XLSX page-margin publication | Media-rich one-edit/save p50 **-97.93%** (48.26x), mean **-97.93%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's direct six-value page-margin set only; cells, formulas, styles, relationships, topology, and changed signed sources remain outside the capability |
| Source-backed XLSX print-options publication | Media-rich one-edit/save p50 **-97.87%** (46.98x), mean **-97.88%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's direct five-flag print options only; cells, formulas, printer settings, relationships, topology, and changed signed sources remain outside the capability |
| Source-backed XLSX page-setup publication | Media-rich one-edit/save p50 **-97.78%** (45.10x), mean **-97.79%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's relationship-free typed settings only; printer settings, cells, formulas, relationships, topology, and changed signed sources remain outside the capability |
| Source-backed XLSX sheet-protection publication | Media-rich one-edit/save p50 **-97.75%** (44.54x), mean **-97.75%**; instructions **-77.87%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's complete direct core/Office 2010 protection state only; password verification, cells, relationships, topology, MCE-selected state and changed signed sources remain outside the capability |
| Source-backed XLSX data-validation publication | Media-rich one-edit/save p50 **-97.75%** (44.51x), mean **-97.75%**; instructions **-73.43%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's complete direct core/Office 2010 validation collections only; cells, formula evaluation, relationships, topology, MCE-selected state and changed signed sources remain outside the capability |
| Source-backed XLSX auto-filter publication | Media-rich filter/sort edit-save p50 **-97.75%** (44.40x), mean **-97.75%**; instructions **-73.57%**; materializations **12 -> 3**; byte-identical output | One existing normal worksheet's direct auto-filter and sort state only; cells, tables, formula evaluation, relationships, topology, MCE-selected state and changed signed sources remain outside the capability |
| Deterministic range simulation | XLSX listing has zero timed requests; selected reads have zero unselected-sheet overlap; full physical size distributions recorded | Synthetic latency model, not a cold filesystem or ambient network |
| [Filesystem cache-state smoke](changes/0087-filesystem-cache-state-evidence.md) | Schema 1 debug artifact completed 10 warm/cold-requested result records and five evidence records; source OPC open uses 13 logical reads/1,008 B and zero Part materializations versus four eager materializations; eager/source saves share the exact `f4bbe4...` output hash; CFB reports one changed span and `799475...` | One sample, no warm-up, dirty worktree, debug build and merely requested cold state. Counter/output correctness only; no latency, allocation, memory, throughput or warm/cold claim |
| [Filesystem repeated release evidence](changes/0089-filesystem-release-repeated-evidence.md) | Five cases, 30 fresh-child samples in each warm/cold-requested state (300 total), CPU-pinned release tmpfs run; logical/process I/O, materialization, span, hash and descriptive latency distributions retained | Accepted advisory cold request on tmpfs; process `read_bytes == 0` is only a process-I/O observation and gives no physical cold-cache or storage claim. No comparator, allocation, peak-memory, or production-performance acceptance |
| [Managed OPC source cache](changes/0086-opc-source-cache-budget-management.md) and [release contention](changes/0088-opc-source-cache-contention-evidence.md) | Managed source-backed OPC (`f8d417ac3`) charges exact physical `InputBytes`, cumulative declared cold-load `Work`, retained catalog/flight/payload `Objects`, and retained/in-flight payload `Memory` to hierarchical `Budget`; compatibility opens remain finite under unmanaged `SourceCacheLimits`; correctness tests cover resource charges, retained-resource releases, pinning, eviction, single-flight, cancellation, sibling competition and contention invariants | Release ABBA provides structural/distribution evidence only; no managed-versus-control speedup accepted. Allocation, peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance evidence remain missing |
| XLSX source-provenance publication reuse | Matched scalar-cell source-backed p50 geomean **-21.66%/-22.65%** and p95 **-21.38%/-22.70%** across ABBA directions; exact output hashes | Removes the repeated semantic worksheet reload/reparse only; physical read/materialization counters are unchanged, and allocation/RSS/cold-I/O claims remain open |
| Rejected generic XLSX publisher provenance reuse | Seven typed source-backed publishers: pooled p50 geomean **+1.04%**; individual pooled p50 **-1.52% to +3.84%**; whole-process allocation calls **-2.84%**; peak heap unchanged | Fully reverted by `a12387478`. The skipped reload usually hit retained cache state, source/materialization/sink/output evidence was unchanged, and the small allocation reduction did not justify the added snapshot/conflict complexity; see [change 0141](changes/0141-xlsx-source-provenance-negative-result.md) |
| Bounded forward-only XLSX/RTF creation | RTF streaming p50 geomean **-76.41%/-76.47%**, p95 **-75.23%/-75.76%**; large sink calls **7,208,970 -> 1,441,802**. XLSX change 0169 accepts hierarchical-budget improvements and descriptively records whole-process allocation calls **-48.81%**/temporary allocations **-69.77%**. Change 0170 additionally accepts large p50/mean/p95/p99 **-5.02% to -6.99%**, medium p50/mean/p95 **-4.45% to -5.52%**, and tiny p50 **-5.03%/-7.74%** from exact-output XML-run batching | RTF claim is escape-free ASCII with a hard 32-byte request ceiling and unchanged 37-byte retained encoder state. XLSX is synthetic warm in-memory one-sheet inline-scalar creation; change 0170 withholds tiny mean/tails and medium p99, branch misses regress, and total-memory/physical-I/O/cold-cache/richer authoring/producer claims remain pending |
| Bounded semantic validation and ODF repair | DOCX, PPTX, RTF and XLS reports now complement CFB, OPC and ODF reports; one opt-in selector exercises ODF's typed non-destructive `mimetype` local-extra repair plan with exact forward/inverse and zero-retained-output sink evidence | Correctness-only, finite and fail-closed. Planning still performs a bounded full-candidate preflight, so no memory or latency claim is made; structural, encrypted, signed, macro and semantic repairs remain unsupported |
| RTF logical-tail append | Two historical staging/commit/publication-timing cases plus four matched Commit/PublicationPlan controls cover tiny/medium/large append and exact no-op publication; the four new selectors use the pre-staged publication-call interval and a fixed 16 KiB non-seek counting sink, while separate planning/publication/reopen/lifecycle vectors and retained-byte fields are emitted. Planning/publication vectors are per-sample; reopen/lifecycle vectors are one-element preflight-only gates run outside the sample loop | Correctness/coverage only: candidate snapshot is not window-bounded, and no end-to-end, rich-format, release ABBA, allocation/RSS, physical-I/O, or speedup claim exists; see [changes 0090](changes/0090-rtf-logical-tail-append-evidence.md) and [0153](changes/0153-rtf-tail-publication-plan-evidence.md) |
| DOCX/PPTX semantic selectors and edits | DOCX one paragraph **-4.72%** p50; PPTX 1% edit/save **-9.37%** p50 and mean; PPTX one-edit guardrail +0.28% p50 (neutral) | Generated text corpora; complete transaction capture dominates one edit; no ODF/iWork implication |
| Coalesced DOCX paragraph edits | Large 100-edit/save p50 **-94.99% (19.97x)** and mean **-95.02%**; medium two-edit/save p50 **-12.98%**; scalar one-edit guardrail neutral | Direct-body, strictly ordered paragraph text replacement; generated corpus; scalar API remains separate |
| ODF semantic baselines and ODS snapshot reuse | Medium/large ODS no-op edit-save p50 **-7.45% / -11.78%**; one-cell edit-save **-3.57% / -2.06%** | Generated ODT/ODS/ODP baseline corpora; focused ODP/ODT publication follow-ups are listed below |
| ODP cross-slide text-box batch evidence | Same semantic projection for matched eight-call scalar and one-call bounded batch publication; deterministic case-specific output hashes and one-write sink counters | Selectable evidence only: physical outputs differ because repeated scalar staging regenerates the manifest; no latency/allocation/memory/materialization claim without frozen CPU-pinned ABBA, and owned ODP exposes no source/materialization diagnostics |
| RTF semantic baseline and text paths | Medium/large full-text p50 **-38.39% / -27.08%**; one-edit/save **-33.40% / -25.79%** | Generated native RTF text corpus; open guard +0.96% / +3.41%; formatting/media/security matrices remain missing |
| RTF retained story length | Large paragraph-list p50 **-15.04%**, mean **-13.71%**; middle-paragraph p50 **-27.19%**, mean **-25.23%** | Already-open generated 10,000-block story queries only; exact parser-derived length, all allocation/peak-heap/RSS metrics flat, and open/full-text/save/no-op guards remain within 5% |
| RTF sparse paragraph selection | Large middle-paragraph p50 **-47.87%**, mean **-47.95%**, p95 **-49.42%** | Explicit `Paragraphs::nth` only; remains linear and allocation-free, constructs the selected view once, preserves iterator state/formatting, and leaves open/list/full-text/save/edit guards within policy |
| RTF retained paragraph cardinality | Large public paragraph-count p50 **-99.93%**, mean **-99.91%**, p95 **-99.86%** | Exact visible body count is retained only after full parser validation; allocation calls and peak heap are flat, collection p50 improves 1.61%, and open/list/read/save/edit guard p50/mean stay within policy |
| ODT shared transaction bytes | Medium/large no-op edit-save p50 **-27.05% / -18.51%**; exactly two allocations and one archive copy removed per snapshot | Existing-document snapshot handoff only; changed edit/save and open guardrails remain within 3%; changed publication still rewrites the package |
| ODT consuming full-text blocks | Repeated large full-text p50 **-3.25%**, mean **-4.81%**; allocation calls **-15.48%**, temporary allocations **-45.52%** | Private full-text mode only; structured queries remain near neutral; unchanged open +3.94% p50/+4.17% mean and +10.95% p99 disclosed |
| ODT indexed paragraph selector | Large middle-paragraph p50 **-48.56%**, mean **-48.33%**; allocation calls **-27.05%**; peak heap **-24.74%**; RSS **-10.93%** | Complete XML/limit validation remains; retains one paragraph, excludes headings from the index, and leaves the established list path neutral |
| ODT content-only unchanged-media publication | Media-rich paragraph edit/save p50 **-95.58%**, mean **-95.63%**, p95 **-95.43%**; allocation calls **-6.71%**; peak heap flat and RSS **-0.59%** | Exactly one paragraph in a fixed 16 MiB-media package; ineligible/mixed operations and regenerated content over the common 16 MiB limit retain the established rebuild |
| ODT content-only line-break publication | Media-rich line-break edit/save p50 **-98.17% (54.59x)**, mean **-98.16%**; instructions **-78.34%**; allocation calls **-6.90%** | Exactly one appended line break in the same fixed 16 MiB-media package; only `content.xml` changes while untouched core/media members remain raw-identical; all ineligible operations retain the established rebuild/policy |
| ODT content-only inline-run publication | Media-rich append-run edit/save p50 **-98.39% (62.01x)**, mean **-98.38%**; instructions **-78.48%**; allocation calls **-7.00%** | One existing styled or unstyled run append in the same fixed 16 MiB-media package; exact no-op dispatch also avoids the changed-path stack frame; other/mixed/ineligible operations retain the established rebuild/policy |
| ODT content-only hyperlink publication | Media-rich append-hyperlink edit/save p50 **-98.20% (55.52x)**, mean **-98.18%**; instructions **-78.34%**; allocation calls **-6.99%** | One inert hyperlink in the same fixed package; exact URL/text reopen and raw-member identity remain, with no relationship or fetch semantics |
| ODT content-only structural paragraph publication | Media-rich insert/remove p50 **-98.20%/-98.27% (55.55x/57.86x)**; combined instructions **-82.14%**; allocation calls **-8.47%** | Existing bounded plain insertion/removal only; removal performs no resource GC, and richer structural/resource/security cases retain established behavior |
| ODT direct snapshot byte sharing | Media-rich direct paragraph edit/save p50 **-75.84%**, mean **-73.84%**, p95 **-75.41%**; peak heap/RSS flat | Removes two 16 MiB archive copies from direct snapshot validation/rehydration; complete XML parsing, publication, reopen/readback, patch and inverse remain |
| ODT compact-audit package sharing | Media-rich paragraph edit/save p50 **-30.44%**, mean **-31.36%**, p95 **-32.41%**; allocations **-0.57%**, peak heap/RSS flat | Removes three archive-sized audit copies (50.36 MB/operation); compact validation, final materialization and readback remain; exact no-op +39 ns p50 is disclosed |
| ODT envelope-classification package sharing | Media-rich paragraph edit/save p50 **-11.40%**, mean **-11.95%**, p95 **-12.19%**; two allocations/commit removed, peak heap/RSS flat | Removes one 16.79 MB envelope copy; archive/manifest and signed/encrypted classification remain; large exact no-op +152 ns p50 is disclosed |
| ODT final changed-result byte handoff | Media-rich paragraph edit/save p50 **-22.74%**, mean **-22.56%**, p95 **-21.48%**; final 16.79 MB copy removed; allocation calls **-3.46%** | Snapshot remains byte-only and one independent final reopen remains; medium one-paragraph +2.77% p50/+1.29% mean is within the 3% gate; peak heap/RSS flat |
| Coalesced ODT paragraph publication | Large 100-edit/save p50 **-98.28% (58.05x)**, mean **-98.27%**; medium two-edit/save p50 **-27.62%**; allocation calls **-96.13%** | Consecutive plain-text replacements only; ordinary durable operations, ordered duplicate semantics, atomic refusal, compact audit, full reopen and scalar path remain |
| Native DOC/XLS/PPT semantic baseline | Large one-edit/save p50: XLS **1.722 ms**, DOC **1.416 ms**, PPT **0.357 ms**; large XLS open **1.383 ms** | Generated writer corpora; accepted XLS and DOC follow-ups are listed below |
| Native XLS validated-editor reuse | Large one-cell edit/save p50 **-7.72%**, mean **-7.90%** | Final exact owner parse, public Workbook reopen and typed readback remain; peak heap/RSS flat |
| Native XLS fixed-width numeric inventory carry-forward | Large one-cell edit/save p50 **-7.83%**, mean **-7.37%**, p95 **-7.20%** | Exact byte-range proof plus complete public Workbook validation/readback remain; peak heap -5.54%, RSS flat; all nonnumeric/structural/resource edits retain full parse |
| XLS comment/visibility source-splice publication | Existing matched eager/source-backed controls cover one/256-comment and one/64-visibility edits; source-backed owners now submit exact NOTE/TXO or `BoundSheet8` ranges | Replacement staging is 109/27,904 bytes instead of an 80,946-byte Workbook and 1/64 instead of 18,166; balanced ABBA accepts no latency speedup, while complete semantic/readback/security gates remain; allocation, RSS and physical I/O remain open |
| XLS fixed-width numeric publication evidence | CPU-2-pinned current-revision p50/p95/p99: eager Number 31.492/34.116/35.916 ms, source-backed Number 146.410/149.108/150.693 ms; eager RK/MulRK 0.100/0.120/0.127 ms, source-backed 1.627/1.659/1.690 ms | Source-backed/eager p50 ratios are descriptively 4.65x/16.25x and both retain the complete 16,995,840/202,752-byte target. This is a before baseline only, with byte-identical family outputs and complete correctness gates; no optimization, regression, allocation/RSS, bounded-memory, physical-I/O, cold-cache, or broad-producer claim |
| XLS plan-only fixed-width numeric publication | Balanced CPU-2 release A1/B1/B2/A2: Number total p50 **-27.57%/-28.58%**, p95 **-27.52%/-28.75%**; RK/MulRK p50 **-24.90%/-24.56%**, p95 **-25.66%/-24.71%**; all paired p99/mean directions agree | Accepted for complete-operation latency in the two deterministic families. Number process VmHWM is **-10.73%/-10.66%** in matched three-warmup/30-sample legs; RK/MulRK RSS disagrees. Valid heaptrack A/B profiles show whole-process allocation reductions but identical peak heap; no operation-only allocation, bounded-memory, physical-I/O or cold-cache claim |
| Rejected XLS terminal-render handoff | Tiny changed save p50 **-7.55%**; large changed save **-0.39%** (neutral) | Fully reverted: repeated large exact no-op p50 **+22.00%**, mean **+16.69%** |
| Common OLE2 publication stages and rejected handoffs | Current open/publication/finish/end-to-end p50: **1.382 / 7.979 / 5.473 / 26.086 ms**; inline recapture prototype end-to-end **-2.61%** p50 | Stages are non-additive; shared-payload, validated-render and inline recapture prototypes are all fully reverted |
| Native DOC batched stream publication | Large one-paragraph edit/save p50 **-10.52%**, mean **-10.48%** | Ordinary two-stream replacement only; final strict revision and independent document reopens remain |
| Native DOC PieceTable physical index | Large open p50 **-55.91%**, mean **-55.78%**; changed edit/save p50 **-31.08%** | Private FC-ordered/prefix-max index only; exact scalar mapping, full FKP validation and strict/public reopens remain; peak heap/RSS flat |
| Native DOC paragraph-style baseline cache | Large open p50 **-11.44%**, mean **-11.87%**; changed edit/save p50 **-4.01%** | One private resolved baseline only; direct PAPX, piece modifiers, direct style switches and complete readbacks remain; allocation calls -18.61%, peak heap/RSS flat |
| Native DOC CHPX range index | Large paragraph-list p50 **-21.07%**, mean **-20.93%**, p95 **-20.00%** | Private monotonic slice query only; exact run identity/order, property cascading and complete readbacks remain; allocations and peak heap/RSS flat |
| Native PPT root snapshot CFB reuse | Repeated large root open p50 **-8.78%**, mean **-10.58%**; allocation calls **-5.01%** | Reuses only the validated CFB index; independent stream/current-user/live-document, slide-order, review-history and public-reader checks remain |
| Native PPT text-edit resolver reuse | Direct large edit/save p50 **-14.12%**, mean **-15.39%**; allocation calls **-3.53%** | Reuses the full editor preflight for persisted-record resolution; exact error precedence, fresh commit editor and complete readback remain; minor-fault increase disclosed |
| Native PPT root text-publication adoption | Large root one-shape edit/save p50 **-18.59%**, mean **-17.83%**, p95 **-16.58%**; allocation calls **-6.54%** | Exact source and selected-slide persist identity gate a private output-Arc handoff; custom limits and structural edits retain complete root reopen; peak heap/RSS flat |
| Bounded XLSX validated-store handoff | Medium one-cell commit + first read p50 **-23.23%**, mean **-23.15%**; allocation calls **-21.01%** | At most 4,096 cells / 1 MiB XML with exact byte and lineage identity; peak heap +4.29%; unrestricted dense-wide candidate rejected at +8.99% peak heap |
| Rejected direct XLSX action-plan flattening | Best formal p50 **-1.61%**; dense commit **-0.27%** p50 with mean interval crossing zero | Fully reverted; process allocation calls -0.0623%, peak heap flat, medium commit p99 +4.33% |
| XLSX no-extension worksheet scan | Medium commit/save p50 **-19.31% to -20.74%**; cold reads about **-35%**; dense 1% commit p50 **-19.62%** | `dyDescent`-free success path only; rejected inputs rerun the original collector for error precedence; allocation calls -25.24%, peak heap flat |
| ODS row-local publication | Large/medium one-cell edit-save p50 **-9.54% / -7.22%**; allocation calls **-5.85%**, peak heap **-27.18%** | Same-topology modeled rows only; structural edits fall back and touched opaque rows refuse |
| ODS unchanged-media publication | Media-rich one-cell edit/save p50 **-4.73%**, mean **-5.73%**, p95 **-7.65%**; peak heap **-8.78%** | Compact `content.xml` replacements in ordinary unsigned/unencrypted ZIPs; every unproved layout/member retains logical rebuild or comparison fallback |
| ODS shared durable-patch blobs | Media-rich one-cell edit/save p50 **-8.80%**, mean **-9.07%**, p95 **-13.85%**; 33.58 MB copy site removed; peak heap **-1.92%** | Shares only already retained immutable source/target package bytes with the forward/reverse semantic bundles; patch wire, limits, final reopen and media verification remain |
| ODS row-splice raw publication | Media-rich one-cell edit/save p50 **-74.16%**, mean **-74.17%**, p95 **-74.11%**; instructions **-69.04%**; peak heap/RSS flat | Same-topology compact row replacements only; exact checked range provenance reaches raw ZIP emission, while structural, signed/encrypted and unsupported layouts retain established fallback/policy |
| ODS shared worksheet archive handoff | Media-rich one-cell edit/save p50 **-21.32%**, mean **-21.30%**, p95 **-21.15%**; peak heap **-22.03%**, RSS **-20.57%** | Private nested worksheet snapshot/package/unified staging only; exact source lineage, failure rollback, durable patches and final readback remain |
| ODP content-only unchanged-media publication | Media-rich text-box edit/save p50 **-94.44%**, mean **-94.43%**, p95 **-94.29%**; allocation calls **+0.52%**; peak heap/RSS flat | Source-backed content-only operations reuse accepted checked-splice/raw-copy publication; resource additions and unsupported/security-sensitive layouts retain logical rebuild |
| ODS exact no-op handoff | Large exact-no-op p50 **-23.26%**, mean **-23.21%**; instructions **-10.54%**; peak heap flat | Exact no-op only; changed commits retain complete audit, preservation and readback paths; read-only link-layout trigger disclosed |
| ODP indexed slide selector | Large middle-slide p50 **-4.09%**, mean **-4.20%**, p95 **-5.18%**; allocation calls **-3.86%**; peak heap/RSS flat | Full style/content EOF validation remains; tiny is neutral, medium p50 -1.55%, and unchanged list/save guards remain within thresholds |
| ODP snapshot slide-projection reuse | Large exact-no-op edit/save p50 **-59.96% (2.50x)**, mean **-59.92%**; large changed edit/save p50 **-20.78%**; allocation calls **-20.13%** | Reuses only the snapshot-validated slide projection for detached staging; package/security reopen, auxiliary parsing, raw page coverage, complete publication/readback and peak heap/RSS remain |
| ODP final slide-snapshot handoff | Large one-slide edit/save p50 **-32.35%**, mean **-32.92%**, p95 **-35.95%**; allocation calls **-16.71%** | Exact slide-only commits move the already parsed candidate projection only after the independent final package/audit/media pipeline; compound domains retain ordinary final parsing; peak heap/RSS flat |
| ODS adaptive cell locator | Large public cell sweep p50 **-81.74%**, mean **-80.72%**; full cell text p50 **-52.65%** | Builds lazily at 64 calls, requests 3,216 bytes on the dense corpus and is capped at 4 MiB; peak heap/RSS flat |
| RTF parser-state specialization | Large open p50 **-20.09%**; large/medium one-edit-save **-11.54% / -14.16%**; cycles **-10.50%** | Ordinary body text only; insertion/deletion metadata retains the full state; allocation count, peak heap and RSS flat |
| RTF ASCII transport batching | Large open p50 **-26.67%**; large/medium one-edit-save **-6.26% / -10.07%**; instructions **-18.40%** | ASCII source tokens only; byte-valued non-ASCII and invalid-Unicode fallback unchanged; allocation count, peak heap and RSS flat |
| RTF byte delimiter scanning | Large open p50 **-17.23%**, mean **-17.99%**; one-edit/save p50 **-14.65%**, mean **-14.84%**; instructions **-21.27%** | Ordinary-text lexer only; plain/CP-1252/LZFu opens improve; prepared LZFu no-op segment +0.290 us/+6.41% p50 is disclosed while complete open improves 19.39%; peak heap/RSS flat |
| RTF retained body source span | Large one-edit/save p50 **-10.72%**, mean **-10.11%**, p95 **-8.76%**; instructions **-10.64%** | Direct uncompressed ASCII ordinary bodies only; cached range is proven during full parser preflight, while ambiguous/binary/non-ASCII/LZFu inputs keep the established locator/refusal and candidate parse/readback |
| RTF bounded body-block reservation | Large open p50 **-21.17%**, mean **-21.00%**, p95 **-21.04%**; one-edit/save p50 **-1.46%**; peak heap **-29.73%** | Sources >=64 KiB only; exact root-text count, token/source/16 MiB caps, lazy fallible allocation, and table/deletion fallback retain semantic behavior; medium plain/CP-1252 +0.49%/+2.84% p50 disclosed |
| Rejected RTF decoded-body ownership | Broad raw CP-1252 open **-3.08% p50 / -3.28% mean**; allocation calls **-20.15%** | Fully reverted: plain large open **+25.53% p50 / +22.45% mean**; owned-only variants were compiler-layout sensitive at -1.41% and +1.02% p50 |
| OPC shared changed-Part payload | Few-large compressible targeted save **-20.73%** p50 / **-18.49%** mean; cache misses **-31.12%** | Removes one 4.19 MiB handoff copy; peak heap -3.42%, uninstrumented RSS +0.22% (flat); the remaining local-span copy is removed by the follow-up below |
| ZIP generated local-span move | Few-large compressible/incompressible targeted save **-4.09% / -2.70%** p50; means **-4.08% / -2.25%** | Removes the separate 4.20 MiB post-validation local-span copy; peak heap -3.20%, uninstrumented RSS -0.10% (flat); required compressor/archive buffer remains |
| Source-backed OPC one-Part publication | Fixed four-Part save p50 **-73.12%**, mean **-73.58%**; semantic materializations **4 -> 1**; instructions **-65.42%** | Low-level consuming same-topology replacement only; raw-copies all unselected ZIP members; signed real changes and unsupported layouts refuse before output; complete physical input/output bytes remain |
| Source-backed DOCX semantic publication | Fixed media-rich one-edit/save p50 **-97.43%**, mean **-97.41%**, p95 **-97.27%**; materializations **17 -> 1**; instructions **-74.91%** | Exact raw main-document transactions only; MCE rewrites, dependency transfers and signed real changes refuse; physical archive input/output remains and eager DOCX guard p50 is +0.25% |

Raw evidence: [`XLSX before A`](results/abba-xlsx-range-before-a.json),
[`after A`](results/abba-xlsx-range-after-a.json),
[`before B`](results/abba-xlsx-range-before-b.json),
[`after B`](results/abba-xlsx-range-after-b.json); [`EOCD before A`](results/abba-eocd-before-a.json),
[`after A`](results/abba-eocd-after-a.json), [`before B`](results/abba-eocd-before-b.json),
[`after B`](results/abba-eocd-after-b.json); and
[`source-versus-eager`](results/stage3-source-vs-eager-many-small.json). The
committed positional XLSX record is
[`xlsx-source-positional.json`](results/xlsx-source-positional.json): p50 open
is 33.881 us/56.493 us/139.897 us (tiny/medium/dense), listing after open has
zero timed source reads, and first/range reads physically overlap only the
selected worksheet member (zero unselected worksheet read calls). These are
physical-overlap counts, not materialization counts.

The semantic raw evidence is [`before A`](results/abba-semantic-before-a.json),
[`after A`](results/abba-semantic-after-a.json),
[`after B`](results/abba-semantic-after-b.json), and
[`before B`](results/abba-semantic-before-b.json). The dedicated 60-sample
one-edit guardrail is linked from
[`change 0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md).

The ODF/ODS raw evidence is [`before A`](results/abba-odf-before-a.json),
[`after A`](results/abba-odf-after-a.json),
[`after B`](results/abba-odf-after-b.json), and
[`before B`](results/abba-odf-before-b.json). Pooled statistics and the
allocation/RSS guardrails are in
[`change 0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md).

The coalesced-DOCX raw evidence is
[`before A`](results/abba-docx-batch-before-a.json),
[`after A`](results/abba-docx-batch-after-a.json),
[`after B`](results/abba-docx-batch-after-b.json), and
[`before B`](results/abba-docx-batch-before-b.json). Pooled statistics and the
allocation/RSS guardrails are in
[`change 0012`](changes/0012-docx-coalesced-paragraph-edits.md).
That record also links the dedicated four-leg large-corpus scalar one-edit
guardrail (p50 -1.28%, mean +0.79% with overlapping intervals), which is
treated as neutral.

The RTF raw evidence is
[`text before A`](results/abba-rtf-text-before-a.json),
[`text after A`](results/abba-rtf-text-after-a.json),
[`text after B`](results/abba-rtf-text-after-b.json), and
[`text before B`](results/abba-rtf-text-before-b.json). The independent open
guard, complete seven-case matrix, allocation/RSS evidence, and rejected first
candidate are in
[`change 0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md).

The ODT shared-snapshot raw evidence is
[`before A`](results/abba-odt-shared-before-a.json),
[`after A`](results/abba-odt-shared-after-a.json),
[`after B`](results/abba-odt-shared-after-b.json), and
[`before B`](results/abba-odt-shared-before-b.json). Allocation attribution,
pooled statistics, open/changed-publication guardrails, and RSS evidence are in
[`change 0014`](changes/0014-odt-shared-snapshot-bytes.md).

The ODT media-publication raw evidence is
[`before A`](results/abba-odt-media-paragraph-before-a.json),
[`after A`](results/abba-odt-media-paragraph-after-a.json),
[`after B`](results/abba-odt-media-paragraph-after-b.json), and
[`before B`](results/abba-odt-media-paragraph-before-b.json). The ordinary ODT
guard ABBA, allocation/RSS/counter profiles, binary identity and common-limit
fallback proof are indexed in
[`change 0035`](changes/0035-odt-content-only-paragraph-publication.md).

The matched ODT line-break publication evidence is
[`before A`](results/abba-odt-line-break-before-a.json),
[`after A`](results/abba-odt-line-break-after-a.json),
[`after B`](results/abba-odt-line-break-after-b.json), and
[`before B`](results/abba-odt-line-break-before-b.json). The pooled
distribution, isolated regression guards, allocation/RSS/counter attribution,
raw-member identity, and exact output digest are indexed in
[`change 0071`](changes/0071-odt-content-only-line-break-publication.md).

The matched ODT inline-run publication evidence is
[`before A`](results/abba-odt-append-run-before-a.json),
[`after A`](results/abba-odt-append-run-after-a.json),
[`after B`](results/abba-odt-append-run-after-b.json), and
[`before B`](results/abba-odt-append-run-before-b.json). The pooled
distribution, exact-no-op and changed-path guards, allocation/RSS/counter
attribution, styled/unstyled raw-member identity, and exact output digest are
indexed in
[`change 0072`](changes/0072-odt-content-only-run-publication.md).

The matched ODT hyperlink evidence is indexed in
[`change 0074`](changes/0074-odt-content-only-hyperlink-publication.md). The
structural insert/remove raw ABBA JSON, pooled distributions, allocation/RSS/
counter attribution, guard reruns, exact output digests, and raw-member proof
are indexed in
[`change 0075`](changes/0075-odt-structural-paragraph-publication.md).

The native OLE2 semantic baseline is
[`ole2-semantic-baseline-a57506d23-2026-08-11.json`](results/ole2-semantic-baseline-a57506d23-2026-08-11.json).
Its complete latency table, Heaptrack/RSS evidence, hardware counters, and
ranked next target are in
[`change 0015`](changes/0015-native-ole2-semantic-baseline.md).

The accepted native XLS follow-up reuses its validated object editor; its
primary raw reports are
[`before A`](results/abba-xls-commit-reuse-one-edit-before-a.json) and
[`after A`](results/abba-xls-commit-reuse-one-edit-after-a.json), with pooled
statistics and all four legs in
[`change 0016`](changes/0016-xls-commit-editor-reuse.md). The later fixed-width
numeric follow-up carries the private BIFF inventory only after exact
field-range certification and keeps the complete public Workbook validation
boundary. Its record and pooled evidence are
[`change 0059`](changes/0059-xls-fixed-numeric-inventory-carry.md) and the
[`primary summary`](results/xls-inventory-carry-primary-summary.json). The DOC follow-up
batches ordinary stream replacement; its primary raw reports are
[`before A`](results/abba-doc-stream-batch-one-edit-before-a.json) and
[`after A`](results/abba-doc-stream-batch-one-edit-after-a.json), with the
complete record in
[`change 0017`](changes/0017-doc-batched-stream-publication.md).

The later XLS terminal-render handoff was measured and fully reverted. Its
large changed-save p50 improved only 0.39%, while four repeated exact-no-op
cycles regressed 22.00% p50 and 16.69% mean. The profile, equality prototype,
allocation evidence and rejection gate are in
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

The committed source-backed XLS visibility and comment owners retain their
complete semantic candidates, exact fingerprints, patch/refusal contracts and
opaque-stream readback. Change 0095 replaces only their prior complete
`Workbook` replacement handoff with bounded CFB splices: 109/27,904 bytes for
one/256 comments versus 80,946, and 1/64 bytes for one/64 visibility owners
versus 18,166. In balanced CPU-pinned ABBA, all source-backed p50 directions
stayed inside 1.5%; each workload's largest absolute source-backed delta stayed
below its largest absolute eager-control delta, so no speedup or material
regression is accepted. Allocation, RSS and physical/source-I/O evidence remain open. See
[`change 0091`](changes/0091-xls-visibility-source-overlay-evidence.md)
and [`change 0095`](changes/0095-xls-semantic-splice-publication.md).

The native PPT root-snapshot evidence is retained as four short ABBA cycles
under `results/abba-ppt-slide-order-root-repeat-*.json`. Reader/edit guards,
allocation attribution, RSS, counters, the disclosed initial selected-shape
tail and its neutral repeat are summarized in
[`change 0024`](changes/0024-ppt-slide-order-open-reuse.md).

The later native PPT root text-publication adoption is summarized in
[`change 0062`](changes/0062-ppt-root-text-publication-adoption.md), with exact
pooled latency, guard, allocation, RSS, and counter values in its linked JSON.

The bounded XLSX commit/read evidence is retained under
`results/abba-xlsx-store-handoff-*.json`; the exact identity gates, primary
latency, allocation/RSS/counter attribution and rejected unrestricted
dense-wide prototype are summarized in
[`change 0025`](changes/0025-xlsx-validated-store-handoff.md).

The direct XLSX writer-regrouping prototype was also fully reverted. Its
medium and dense-wide 1% commit/save ABBA reports are under
`results/abba-xlsx-action-plan-*.json`; matched allocation evidence and the
rejection rationale are in
[`change 0030`](changes/0030-xlsx-action-plan-flattening-rejected.md).

The accepted XLSX no-extension scan evidence is under
`results/abba-xlsx-x14ac-*.json`. Medium and dense-wide latency, read/no-op
guards, allocation/RSS/counter attribution and malformed-input precedence are
summarized in
[`change 0032`](changes/0032-xlsx-no-extension-scan.md).

The new common OLE2 publication evidence is under
`results/abba-ole-common-*.json`. It retains the deterministic opaque-heavy
case, but both measured production handoffs were reverted: direct shared
writer payloads regressed 32.02% p50, while retaining the validated render
improved the target 34.06% but regressed DOC open 21.64%. The full rationale
and DOC/XLS guards are in
[`change 0033`](changes/0033-ole-common-publication-handoffs-rejected.md).

The common OLE2 stage/recapture reports are
[`before A`](results/abba-ole-recapture-before-a.json),
[`after A`](results/abba-ole-recapture-after-a.json),
[`after B`](results/abba-ole-recapture-after-b.json), and
[`before B`](results/abba-ole-recapture-before-b.json). The stage profile,
non-additivity finding and fully reverted inline allocation-reuse prototype are
documented in
[`change 0036`](changes/0036-ole-common-stage-attribution.md).

The ODS row-local publication evidence is
[`before A`](results/abba-ods-row-splice-one-edit-before-a.json),
[`after A`](results/abba-ods-row-splice-one-edit-after-a.json),
[`after B`](results/abba-ods-row-splice-one-edit-after-b.json), and
[`before B`](results/abba-ods-row-splice-one-edit-before-b.json). Medium,
guardrail, allocation, RSS and hardware-counter evidence is summarized in
[`change 0018`](changes/0018-ods-row-local-publication.md).

The adaptive ODS cell-locator ABBA evidence starts at
[`before A`](results/abba-ods-cell-locator-before-a.json) and
[`after A`](results/abba-ods-cell-locator-after-a.json); the complete profile,
guard, memory and counter record is
[`change 0027`](changes/0027-ods-adaptive-cell-locator.md).

The ODS unchanged-media publication evidence is
[`before A`](results/abba-ods-media-preservation-before-a.json),
[`after A`](results/abba-ods-media-preservation-after-a.json),
[`after B`](results/abba-ods-media-preservation-after-b.json), and
[`before B`](results/abba-ods-media-preservation-before-b.json). Raw-member
proofs, the no-media guard, fallback semantics, memory and counter attribution
are summarized in
[`change 0031`](changes/0031-ods-unchanged-media-preservation.md).

The ODS durable-patch ownership evidence starts with the balanced primary
[`before`](results/ods-shared-patch-blobs-primary-forward-1-before.json) and
[`after`](results/ods-shared-patch-blobs-primary-forward-1-after.json) legs.
All four primary pairs, medium/large guards, profiles, counters, memory, RSS,
wire-identity checks and binary provenance are indexed in
[`change 0054`](changes/0054-ods-shared-durable-patch-blobs.md).

The ODS row-splice raw-publication evidence retains all 300 samples per state
in the
[`primary summary`](results/ods-row-splice-raw-publication-primary-summary.json).
Tiny, medium and large ordinary CRUD distributions, matched profiles,
counters, Heaptrack, GNU Time and exact binary provenance are indexed in
[`change 0057`](changes/0057-ods-row-splice-raw-publication.md).

The ODS shared worksheet-ownership evidence pools 2,000 ABBA/reverse-BAAB
samples per state in the
[`summary`](results/ods-worksheet-shared-ownership-summary.json). Large
ordinary guards, matched Heaptrack/GNU Time/perf evidence, exact binary
provenance, and the rejected 4.01% intermediate are indexed in
[`change 0068`](changes/0068-ods-shared-worksheet-archive-handoff.md).

The ODP content-only publication evidence is
[`before A`](results/abba-odp-media-textbox-before-a.json),
[`after A`](results/abba-odp-media-textbox-after-a.json),
[`after B`](results/abba-odp-media-textbox-after-b.json), and
[`before B`](results/abba-odp-media-textbox-before-b.json). Raw-member proofs,
ordinary ODP guards, patch/inverse checks, memory, and hardware counters are
summarized in
[`change 0034`](changes/0034-odp-unchanged-media-preservation.md).

The ODP snapshot-projection evidence pools 4,000 large exact-no-op samples per
state in the
[`primary summary`](results/odp-slide-projection-primary-summary.json).
Tiny/medium scaling, large changed-edit and read/media guards are in the
[`guard summary`](results/odp-slide-projection-guard-summary.json); matched
profiles, counters, Heaptrack, GNU Time and exact binary provenance are indexed
in [`change 0060`](changes/0060-odp-snapshot-slide-projection-reuse.md).

The ODP final-snapshot evidence uses a drift-gated warmed 2,000-sample pool per
state in the [`summary`](results/odp-final-snapshot/summary.json). Tiny/medium
scaling, ineligible no-op/media guards, repeated read-only tails, matched
Heaptrack, GNU Time, counters and binary provenance are indexed in
[`change 0065`](changes/0065-odp-final-snapshot-handoff.md).

The RTF parser-state follow-up evidence is
[`before A`](results/abba-rtf-state-clone-one-edit-before-a.json),
[`after A`](results/abba-rtf-state-clone-one-edit-after-a.json),
[`after B`](results/abba-rtf-state-clone-one-edit-after-b.json), and
[`before B`](results/abba-rtf-state-clone-one-edit-before-b.json). Open/save
guardrails, profiles, hardware counters, memory results and the rejected ODS
candidate are summarized in
[`change 0019`](changes/0019-rtf-parser-state-specialization.md).

The RTF transport-batching evidence is
[`before A`](results/abba-rtf-ascii-transport-primary-before-a.json),
[`after A`](results/abba-rtf-ascii-transport-primary-after-a.json),
[`after B`](results/abba-rtf-ascii-transport-primary-after-b.json), and
[`before B`](results/abba-rtf-ascii-transport-primary-before-b.json). Medium,
save-only, profile, counter and memory guardrails plus the rejected ODT
candidate are summarized in
[`change 0020`](changes/0020-rtf-ascii-transport-batching.md).

The RTF byte-delimiter evidence is
[`before A`](results/abba-rtf-byte-delimiter-final-before-a.json),
[`after A`](results/abba-rtf-byte-delimiter-final-after-a.json),
[`after B`](results/abba-rtf-byte-delimiter-final-after-b.json), and
[`before B`](results/abba-rtf-byte-delimiter-final-before-b.json). Plain,
CP-1252 and LZFu guards, the prepared LZFu no-op disclosure, profiles,
counters, memory and complete correctness gates are summarized in
[`change 0040`](changes/0040-rtf-byte-delimiter-scanning.md).

The rejected RTF decoded-body ownership evidence includes two broad-prototype
ABBA cycles, plain/CP-1252/LZFu and prepared-operation guards, two owned-only
refinements, Heaptrack, `perf record`, and GNU Time summaries. The raw JSON
digests and full rejection rationale are in
[`change 0043`](changes/0043-rtf-decoded-body-ownership-rejected.md).

The retained RTF body-source-span evidence is
[`before A`](results/abba-rtf-body-span-before-a.json),
[`after A`](results/abba-rtf-body-span-after-a.json),
[`after B`](results/abba-rtf-body-span-after-b.json), and
[`before B`](results/abba-rtf-body-span-before-b.json). Tiny/medium scaling,
open/list/no-op guards, allocation attribution, counters, RSS, capability
smoke and artifact hashes are summarized in
[`change 0048`](changes/0048-rtf-retained-body-source-span.md).

The existing-document RTF logical-tail evidence is a correctness/coverage
tranche rather than a performance result. Its two selectors cover tiny,
medium, and large appends plus an exact no-op through a fixed 16 KiB hashing
sink window that caps accepted bytes per write and retains zero output; complete
reopen, sequential bytes, in-memory and durable patch/inverse, and foreign-source
refusal remain untimed gates. The sink window does not bound the transaction's
validated candidate snapshot. See
[`change 0090`](changes/0090-rtf-logical-tail-append-evidence.md).

Change 0153 adds matched Commit-versus-PublicationPlan append and exact-no-op
selectors. Their `elapsed_ns` is the pre-staged publication-call interval
around the respective public write call; `planning_ns` and `publication_ns`
are per-sample vectors, while `reopen_ns` and `lifecycle_ns` are one-element
preflight-only vectors for expensive gates run once outside the sample loop.
The planning, reopen, lifecycle, durable patch, cancellation, sink
failure/partial progress, limits, and source-version checks are separate
untimed evidence.
Source-retained, complete-candidate-retained, and publication-window bytes are
reported explicitly. This tranche makes no end-to-end, rich-format,
allocation/RSS, physical-I/O, or ABBA latency claim; no release measurement is
run without root approval after code review. See
[`change 0153`](changes/0153-rtf-tail-publication-plan-evidence.md).

The source-backed PPTX selected-slide publication evidence is
[`before A`](results/abba-pptx-source-edit-before-a.json),
[`after A`](results/abba-pptx-source-edit-after-a.json),
[`after B`](results/abba-pptx-source-edit-after-b.json), and
[`before B`](results/abba-pptx-source-edit-before-b.json). The eager semantic
guard, CPU/allocation/RSS attribution, exact preservation/refusal matrix and
frozen binary hashes are summarized in
[`change 0044`](changes/0044-pptx-source-backed-semantic-publication.md).

The source-backed PPTX multi-slide batch evidence is
[`before A`](results/abba-pptx-multi-slide-batch-before-a.json),
[`after A`](results/abba-pptx-multi-slide-batch-after-a.json),
[`after B`](results/abba-pptx-multi-slide-batch-after-b.json), and
[`before B`](results/abba-pptx-multi-slide-batch-before-b.json). Counters,
profiles, allocation/RSS attribution, raw member preservation, refusal coverage
and frozen binary hashes are summarized in
[`change 0077`](changes/0077-pptx-source-backed-multi-slide-batch-publication.md).

The source-backed XLSX calculation-metadata publication evidence is
[`before A`](results/abba-xlsx-calculation-metadata-edit-before-a.json),
[`after A`](results/abba-xlsx-calculation-metadata-edit-after-a.json),
[`after B`](results/abba-xlsx-calculation-metadata-edit-after-b.json), and
[`before B`](results/abba-xlsx-calculation-metadata-edit-before-b.json).
Counters, allocation/RSS attribution, exact workbook/media preservation,
refusal coverage and frozen binary/input/output hashes are summarized in
[`change 0046`](changes/0046-xlsx-source-backed-calculation-metadata-publication.md).

The source-backed XLSX defined-name publication evidence is
[`before A`](results/abba-xlsx-defined-names-stable-before-a.json),
[`after A`](results/abba-xlsx-defined-names-stable-after-a.json),
[`after B`](results/abba-xlsx-defined-names-stable-after-b.json), and
[`before B`](results/abba-xlsx-defined-names-stable-before-b.json).
Counters, profiles, allocation/RSS attribution, exact workbook/media
preservation, refusal coverage and both frozen binary hashes are summarized in
[`change 0076`](changes/0076-xlsx-source-backed-defined-names-publication.md).

The source-backed XLSX sheet-protection publication evidence is
[`before A`](results/abba-xlsx-sheet-protection-before-a.json),
[`after A`](results/abba-xlsx-sheet-protection-after-a.json),
[`after B`](results/abba-xlsx-sheet-protection-after-b.json), and
[`before B`](results/abba-xlsx-sheet-protection-before-b.json). Complete typed
protection readback, counters, profiles, allocation/RSS attribution,
relationship/media preservation, refusal coverage and frozen binary hashes are
summarized in
[`change 0078`](changes/0078-xlsx-source-backed-sheet-protection-publication.md).

The source-backed XLSX data-validation publication evidence is
[`before A`](results/abba-xlsx-data-validation-before-a.json),
[`after A`](results/abba-xlsx-data-validation-after-a.json),
[`after B`](results/abba-xlsx-data-validation-after-b.json), and
[`before B`](results/abba-xlsx-data-validation-before-b.json). Complete typed
core/Office 2010 readback, counters, allocation/RSS attribution,
relationship/media preservation, refusal coverage and the matched binary hash
are summarized in
[`change 0079`](changes/0079-xlsx-source-backed-data-validation-publication.md).

The source-backed XLSX auto-filter publication evidence is
[`before A`](results/abba-xlsx-auto-filter-before-a.json),
[`after A`](results/abba-xlsx-auto-filter-after-a.json),
[`after B`](results/abba-xlsx-auto-filter-after-b.json), and
[`before B`](results/abba-xlsx-auto-filter-before-b.json). Complete typed
filter/sort readback, style-DXF validation, counters, allocation/RSS
attribution, relationship/media preservation, refusal coverage and the final
binary/input/output hashes are summarized in
[`change 0080`](changes/0080-xlsx-source-backed-auto-filter-publication.md).

The coalesced ODT paragraph-publication evidence is
[`before A`](results/abba-odt-paragraph-batch-before-a.json),
[`after A`](results/abba-odt-paragraph-batch-after-a.json),
[`after B`](results/abba-odt-paragraph-batch-after-b.json), and
[`before B`](results/abba-odt-paragraph-batch-before-b.json). Scalar/no-op
guards, CPU/allocation/RSS attribution, durable replay, media preservation,
over-limit fallback and frozen binary hashes are summarized in
[`change 0045`](changes/0045-odt-coalesced-paragraph-publication.md).

The matched ODT mixed model-content publication evidence is
[`scalar A`](results/odt-mixed-model-scalar-a-0112.json),
[`batch A`](results/odt-mixed-model-batch-a-0112.json),
[`batch B`](results/odt-mixed-model-batch-b-0112.json), and
[`scalar B`](results/odt-mixed-model-scalar-b-0112.json). The medium and
large shapes preserve their logical/output hashes while reducing the measured
publication count from 49/193 to one. The exact p50 extraction, binary and
raw-report hashes, and exclusions are in
[`change 0104`](changes/0104-odt-mixed-model-publication-evidence.md) and its
[`compact summary`](results/odt-mixed-model-publication-0112-summary.json).

The ODT compact-audit package-sharing evidence is
[`before A`](results/abba-odt-compact-audit-final-before-a.json),
[`after A`](results/abba-odt-compact-audit-final-after-a.json),
[`after B`](results/abba-odt-compact-audit-final-after-b.json), and
[`before B`](results/abba-odt-compact-audit-final-before-b.json). Ordinary
open/edit/no-op guards, the dedicated 10,000-sample/state no-op disclosure,
profiles, counters, memory, allocator policy and complete correctness gates are
summarized in
[`change 0041`](changes/0041-odt-compact-audit-package-sharing.md).

The ODT envelope-sharing evidence comprises two balanced ABBA cycles:
[`cycle 1 before A`](results/abba-odt-envelope-sharing-rerun-before-a.json),
[`after A`](results/abba-odt-envelope-sharing-rerun-after-a.json),
[`after B`](results/abba-odt-envelope-sharing-rerun-after-b.json),
[`before B`](results/abba-odt-envelope-sharing-rerun-before-b.json), plus the
four matching `final2` reports. Ordinary edit/open/no-op guards, the discarded
exploratory run, profiles, counters, memory and complete correctness gates are
summarized in
[`change 0042`](changes/0042-odt-envelope-package-sharing.md).

The ODT final changed-result byte-handoff evidence comprises two balanced
execution cycles with four 500-sample legs per state. Primary raw reports use
the `odt-final-handoff-cycle*` prefix; the matched medium/large read/no-op/edit
matrix uses `odt-final-handoff-guards*`. Profiles, counters, allocation/RSS,
the byte-only ownership distinction and complete correctness gates are
summarized in
[`change 0052`](changes/0052-odt-final-result-byte-handoff.md).

The OPC shared-payload evidence is
[`before A`](results/abba-opc-shared-regeneration-primary-before-a.json),
[`after A`](results/abba-opc-shared-regeneration-primary-after-a.json),
[`after B`](results/abba-opc-shared-regeneration-primary-after-b.json), and
[`before B`](results/abba-opc-shared-regeneration-primary-before-b.json).
No-op/edge guardrails, allocation attribution, RSS and hardware counters are
summarized in
[`change 0021`](changes/0021-opc-shared-regenerated-payload.md).

The generated-local-span evidence is
[`before A`](results/abba-opc-local-span-move-before-a.json),
[`after A`](results/abba-opc-local-span-move-after-a.json),
[`after B`](results/abba-opc-local-span-move-after-b.json), and
[`before B`](results/abba-opc-local-span-move-before-b.json). Repeated small,
edge, tiny and exact-no-op guardrails, allocation attribution, RSS and hardware
counters are summarized in
[`change 0022`](changes/0022-zip-generated-local-span-move.md).

The source-backed overlay evidence is
[`before A`](results/abba-opc-source-overlay-before-a.json),
[`after A`](results/abba-opc-source-overlay-after-a.json),
[`after B`](results/abba-opc-source-overlay-after-b.json), and
[`before B`](results/abba-opc-source-overlay-before-b.json). Source/sink
counters, CPU and memory attribution, failure boundaries and binary/evidence
digests are summarized in
[`change 0037`](changes/0037-opc-source-backed-one-part-publication.md).

The ODT full-text ownership evidence is retained as four short ABBA cycles
under `results/abba-odt-full-text-single-repeat-*.json`. Structured-query,
open, size, exact-no-op and edit guardrails, rejected broad-parser evidence,
allocation attribution, RSS and hardware counters are summarized in
[`change 0023`](changes/0023-odt-full-text-owned-blocks.md).

The ODT indexed-selector evidence is retained as four headline ABBA cycles
under `results/abba-odt-indexed-paragraph-repeat-*.json`, with separate
size/guard reports, Heaptrack attribution, GNU Time RSS and `perf stat`
counters. The complete validation contract and rejected shared-parser design
are summarized in
[`change 0047`](changes/0047-odt-indexed-paragraph-selector.md).

The RTF block-reservation evidence pools six balanced pairs and retains every
sample in the [`primary summary`](results/rtf-body-block-reservation-primary-summary.json).
The [`medium guard summary`](results/rtf-body-block-reservation-medium-guards-summary.json)
covers plain, raw CP-1252 and LZFu with the same six-pair protocol. Allocation,
RSS, profile, counter, tiny-variant and binary-provenance artifacts are indexed
in [`change 0055`](changes/0055-rtf-body-block-reservation.md).

The retained RTF story-length evidence pools two 1,000-sample legs per state
for the paragraph-list and middle-paragraph queries. A reverse-order
2,000-sample pool covers open, full-text, exact stream-save and no-op guards;
allocation, RSS and process-wide profile records are indexed in
[`change 0064`](changes/0064-rtf-retained-story-length.md).

The sparse RTF paragraph-selection evidence pools two 1,000-sample legs per
state for the already-open middle-paragraph query. Reverse-order read/save,
4,000-sample no-op and changed-edit guard pools, iterator-equivalence tests,
variant verification, allocation, RSS and process-wide profile records are
indexed in [`change 0066`](changes/0066-rtf-sparse-paragraph-nth.md).

The retained RTF paragraph-cardinality evidence pools two 1,000-sample legs
per state for a cold public count query and separately guards complete
collection. Seven large read/save/edit cases use 1,000 samples per state;
allocation, heap, RSS, `perf stat`, `perf record`, variant verification and
binary provenance are indexed in
[`change 0069`](changes/0069-rtf-retained-paragraph-count.md).

The DOC PAPX-containment evidence pools five balanced pairs for both the
already-open snapshot paragraph list and complete one-edit/save path, retaining
every sample in the
[`primary summary`](results/doc-papx-containment-primary-summary.json).
Ordinary-reader/no-op and tiny direct distributions are retained in the
[`guard summary`](results/doc-papx-containment-guards-summary.json); profiles,
counters, Heaptrack and GNU Time artifacts are indexed in
[`change 0056`](changes/0056-doc-papx-containment-index.md).

Managed source-backed OPC (`f8d417ac3`) charges exact physical `InputBytes`,
cumulative declared cold-load `Work`, retained catalog/flight/payload
`Objects`, and retained/in-flight payload `Memory` to a caller's hierarchical
`Budget`; compatibility opens remain finite under unmanaged `SourceCacheLimits`.
Resource charges, retained-resource releases, hierarchy, pinning, eviction, sibling competition,
cancellation, single-flight and failure are correctness-tested. The committed
release contention ABBA adds structural and distribution evidence, but no
managed-versus-control speedup is accepted; allocation, peak-memory/RSS,
hardware, copied/decompressed-byte and CPU-utilization evidence remain missing.
Raw ZIP preservation is integrated for
owned same-topology OPC mutations and the bounded consuming source-backed
multi-Part publisher; broad source-backed semantic editing remains pending.
The release cache capture retains balanced control/managed ABBA distributions
and invariant recomputation, but its fixed source delay is only a coordination
instrument; no production speedup, allocation, RSS, hardware, or CPU-utilization
claim is made. The repeated filesystem release capture retains 300 fresh-child
tmpfs samples for the five file cases; accepted cold advice and zero process
`read_bytes` keep it descriptive rather than a cold-storage result.
See [`0005`](changes/0005-xlsx-row-start-index.md),
[`0006`](changes/0006-positional-containers-and-explicit-execution.md), and
[`0007`](changes/0007-source-backed-opc-and-facades.md),
[`0008`](changes/0008-targeted-opc-preservation.md), and
[`0009`](changes/0009-range-source-and-scaling.md), and
[`0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md), and
[`0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md), and
[`0012`](changes/0012-docx-coalesced-paragraph-edits.md), and
[`0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md), and
[`0014`](changes/0014-odt-shared-snapshot-bytes.md), and
[`0015`](changes/0015-native-ole2-semantic-baseline.md),
[`0016`](changes/0016-xls-commit-editor-reuse.md),
[`0017`](changes/0017-doc-batched-stream-publication.md), and
[`0018`](changes/0018-ods-row-local-publication.md), and
[`0019`](changes/0019-rtf-parser-state-specialization.md), and
[`0020`](changes/0020-rtf-ascii-transport-batching.md), and
[`0021`](changes/0021-opc-shared-regenerated-payload.md), and
[`0022`](changes/0022-zip-generated-local-span-move.md), and
[`0023`](changes/0023-odt-full-text-owned-blocks.md), and
[`0024`](changes/0024-ppt-slide-order-open-reuse.md), and
[`0025`](changes/0025-xlsx-validated-store-handoff.md).

Consolidated OPC, PPTX and performance-harness tests passed, along with
warning-denied changed-crate Clippy, formatter, workflow, JSON and final-diff
checks. Warning-denied ODF-common Clippy and rustdoc also pass, revalidating the
GenericArray deprecation fix. The broad crate-boundary checker retains existing
unclassified workspace edges; no manifest or dependency edge changed. A
workspace all-target/all-feature gate was not run because iWork was explicitly
excluded while its crates are changing independently.

## Accepted results

All latency figures below are warm-memory release-build p50 results from
matched before/after binaries. Each linked change record contains raw-sample
counts, ABBA ordering, mean or interval context, hashes, and memory profiles.

| Workload group | Before | After | Result | Memory result |
|---|---:|---:|---:|---|
| Targeted OPC mutation, four synthetic cells | individual rows in record | individual rows in record | **-84.98% p50 geometric mean**; range -58.24% to -96.41% | Few-large/incompressible peak heap +37.18%; one-shot RSS +22.26% |
| Shared changed-Part handoff, few-large compressible | 1.342 ms | 1.063 ms | **-20.73% p50 / -18.49% mean** | One 4.19 MiB allocation removed; peak heap -3.42%; uninstrumented RSS +0.22% (flat) |
| Exact owned OPC no-op, 16.78 MB incompressible archive | 211.531 ms | 3.443 ms | -98.37% | Peak heap +22.6%; profiler RSS +25.5% because the compressed source is retained alongside eagerly inflated Parts |
| Exact owned OPC no-op, six named many-Part/large-Part cells | individual rows in record | individual rows in record | -99.93% p50 geometric mean | Many-small allocation calls -93.7%; large memory tradeoff above |
| CFB final-root-stream lookup, four 256/2,048-sibling cells | 1.067-7.596 us | 0.451-0.486 us | -84.70% p50 geometric mean | Wide-root peak heap +1.5%; profiler RSS +7.6% for retained exact comparison keys |
| CFB open, four 256/2,048-stream cells | 141.1-963.1 us | 136.8-974.9 us | -1.42% p50 geometric mean | Allocation calls -6.1% to -8.8%; temporary allocations -20.6% to -27.7% |
| Rejected common OLE2 inline recapture allocation reuse, 16 MiB opaque streams | 26.086 ms | 25.404 ms | **-2.61% p50 / -2.30% mean** | Fully reverted as immaterial; p95 +0.54%; isolated publication p50 -6.49% but stages are non-additive |
| OPC rewritten publication, eight named cells | individual rows in record | individual rows in record | -1.65% mean geometric mean; best intended cell -5.49% | Allocation calls -37.0%; peak heap -2.3% |
| Payload-heavy PPT fresh writer | 6.312 ms | 5.035 ms | -20.23% | Peak heap -12.4%; profiler RSS -12.9% |
| Payload-heavy XLS fresh writer | 4.126 ms | 4.065 ms | -1.48%, treated as latency-neutral | Peak heap -9.5%; profiler RSS -12.6% |
| DOCX one paragraph, 10,000-paragraph corpus | 2.945 ms | 2.805 ms | -4.72% p50 / -4.99% mean | 10 collection-growth allocations removed per selector invocation; process peak unchanged |
| DOCX 1% edit/save, 10,000 paragraphs / 100 edits | 487.542 ms | 24.418 ms | **-94.99% p50 (19.97x) / -95.02% mean**; scalar one-edit neutral | Allocation calls -94.11%; peak heap flat; uninstrumented RSS +0.37% (flat) |
| PPTX 1% edit/save, 10,000 text boxes | 399.320 ms | 361.915 ms | -9.37% p50 / -9.37% mean | Allocation calls -11.67%; peak heap flat; profiler RSS +1.28% |
| ODS no-op edit/save, 32,768 cells | 76.894 ms | 67.838 ms | -11.78% p50 / -12.08% mean | Peak heap flat; profiler RSS -0.13% |
| ODS one-cell edit/save, 32,768 cells | 384.150 ms | 376.237 ms | -2.06% p50 / -2.19% mean | Changed package rewrite/readback still dominates |
| ODS row-local one-cell edit/save, 32,768 cells | 359.011 ms | 324.774 ms | **-9.54% p50 / -9.32% mean** | Allocation calls -5.85%; peak heap -27.18%; uninstrumented RSS improved |
| ODS media-rich one-cell edit/save, 2,048 cells + 16 MiB media | 325.902 ms | 310.472 ms | **-4.73% p50 / -5.73% mean** | p95 -7.65%; peak heap -8.78%; existing no-media guard p50 -0.77% |
| ODS durable-patch sharing, 2,048 cells + 16 MiB media | 326.694 ms | 297.958 ms | **-8.80% p50 / -9.07% mean** | p95 -13.85%; redundant package SHA stack absent; 33.58 MB copy site removed; peak heap -1.92%; RSS flat |
| ODS checked row-splice raw publication, 2,048 cells + 16 MiB media | 287.766 ms | 74.365 ms | **-74.16% p50 / -74.17% mean** | p95 -74.11%; instructions -69.04%; unchanged-media rebuild/deflate subtree absent; peak heap/RSS flat |
| ODS shared worksheet archive handoff, 2,048 cells + 16 MiB media | 76.440 ms | 60.140 ms | **-21.32% p50 / -21.30% mean** | p95 -21.15%; peak heap -22.03%; uninstrumented RSS -20.57% |
| ODP media-rich text-box edit/save, 12 slides + 16 MiB media | 227.606 ms | 12.665 ms | **-94.44% p50 / -94.43% mean** | p95 -94.29%; allocation calls +0.52%; peak heap/RSS flat |
| ODS public cell sweep, 32,768 cells | 2.049 ms | 0.374 ms | **-81.74% p50 / -80.72% mean** | Lazy 3,216-byte dense index; peak heap/RSS flat; allocation calls +0.0004% process-wide |
| ODS full cell text, 32,768 cells | 3.047 ms | 1.443 ms | **-52.65% p50 / -52.30% mean** | Existing string clones/join remain; lookup work only is indexed |
| RTF full text, 10,000 paragraphs | 33.095 us | 24.134 us | -27.08% p50 / -25.37% mean | One fragment-vector allocation removed per first materialization |
| RTF one paragraph edit/save, 10,000 paragraphs | 12.408 ms | 9.208 ms | -25.79% p50 / -25.53% mean | Allocation calls -707 over 100 samples; peak heap flat; RSS +0.32% (flat) |
| RTF parser-state follow-up, one paragraph edit/save, 10,000 paragraphs | 8.630 ms | 7.634 ms | **-11.54% p50 / -11.71% mean** | `State::clone` profile frame removed; allocation calls, peak heap and RSS flat |
| RTF transport batching, open, 10,000 paragraphs | 3.159 ms | 2.316 ms | **-26.67% p50 / -26.56% mean** | Per-byte `SmallVec::extend` frame falls from 15.37% to 2.56%; allocations and peak heap flat |
| RTF transport batching, one paragraph edit/save, 10,000 paragraphs | 7.795 ms | 7.307 ms | **-6.26% p50 / -5.73% mean** | Instructions -18.40%; allocation count, peak heap and RSS flat |
| RTF byte-delimiter scan, open, 10,000 paragraphs | 2.479 ms | 2.052 ms | **-17.23% p50 / -17.99% mean** | `tokenize_with_spans` share 17.36% -> 11.06%; instructions -21.27%; peak heap/RSS flat |
| RTF byte-delimiter scan, one paragraph edit/save, 10,000 paragraphs | 7.554 ms | 6.447 ms | **-14.65% p50 / -14.84% mean** | p95 -16.34%; allocations effectively flat; complete edit/save readback unchanged |
| RTF retained body span, one paragraph edit/save, 10,000 paragraphs | 6.053 ms | 5.404 ms | **-10.72% p50 / -10.11% mean** | p95 -8.76%; 588 locator-subtree allocation calls over 20 edits removed; peak heap/RSS flat; candidate parse/readback unchanged |
| RTF bounded body-block reservation, open, 10,000 paragraphs | 2.073 ms | 1.634 ms | **-21.17% p50 / -21.00% mean** | p95 -21.04%; body-vector allocations 264 -> 22 over 22 parses; peak heap -29.73%; uninstrumented RSS flat |
| RTF bounded body-block reservation, one paragraph edit/save | 5.585 ms | 5.503 ms | **-1.46% p50 / -1.75% mean** | p95 -1.87%, p99 -4.11%; complete candidate parse/readback unchanged |
| RTF paragraph list, already-open 10,000-block story | 29.692 us | 25.225 us | **-15.04% p50 / -13.71% mean** | p95 -8.64%; reuses the parser-owned exact text length; allocations and peak heap/RSS flat |
| RTF middle paragraph, already-open 10,000-block story | 18.926 us | 13.780 us | **-27.19% p50 / -25.23% mean** | p95 -14.46%; paragraph boundaries, formatting, exact no-op and complete verification unchanged |
| RTF public paragraph count, already-open 10,000-block story | 28.898 us | 0.020 us | **-99.93% p50 / -99.91% mean** | p95 -99.86%; full parser validation retained; allocations and peak heap flat; collection p50 -1.61% |
| ODT no-op edit/save, 10,000 paragraphs | 3.950 us | 3.219 us | -18.51% p50 / -29.58% mean | Exactly two allocations and one 28.42 KiB archive copy removed per snapshot; peak heap/RSS flat |
| ODT full text, 10,000 blocks | 4.127 ms | 3.993 ms | **-3.25% p50 / -4.81% mean** | Allocation calls -15.48%, temporary allocations -45.52%; peak heap/RSS flat; open guard disclosed |
| ODT middle paragraph, 10,000 paragraphs | 3.202 ms | 1.647 ms | **-48.56% p50 / -48.33% mean** | Allocation calls -27.05%; peak heap -24.74%; uninstrumented RSS -10.93%; complete EOF validation retained |
| ODP middle slide, 100 slides | 1.019 ms | 0.977 ms | **-4.09% p50 / -4.20% mean** | p95 -5.18%; allocation calls -3.86%; peak heap/RSS flat; complete style/content EOF validation retained |
| ODP exact no-op transaction/save, 100 slides | 1.728 ms | 0.692 ms | **-59.96% p50 (2.50x) / -59.92% mean** | Large changed edit/save p50 -20.78%; allocations -20.13%; complete package/security and final readback retained; peak heap/RSS flat |
| ODP one-slide edit/save, 100 source slides | 3.573 ms | 2.417 ms | **-32.35% p50 / -32.92% mean** | p95 -35.95%; allocations -16.71%; final package reopen/audits/media checks retained; peak heap/RSS flat |
| ODT media-rich paragraph edit/save, 200 paragraphs + 16 MiB media | 249.177 ms | 11.001 ms | **-95.58% p50 / -95.63% mean** | p95 -95.43%; allocation calls -6.71%; peak heap flat; RSS -0.59% |
| ODT media-rich line-break edit/save, 200 paragraphs + 16 MiB media | 217.532 ms | 3.985 ms | **-98.17% p50 (54.59x) / -98.16% mean** | p95 -98.08%; instructions -78.34%; allocation calls -6.90%; peak heap/RSS flat |
| ODT media-rich append-run edit/save, 200 paragraphs + 16 MiB media | 225.431 ms | 3.635 ms | **-98.39% p50 (62.01x) / -98.38% mean** | p95 -98.27%; instructions -78.48%; allocation calls -7.00%; peak heap/RSS flat |
| ODT media-rich insert paragraph, 200 paragraphs + 16 MiB media | 220.507 ms | 3.969 ms | **-98.20% p50 (55.55x) / -98.19% mean** | p95 -98.10%; exact output/member identity; combined structural instructions -82.14% |
| ODT media-rich remove paragraph, 200 paragraphs + 16 MiB media | 219.315 ms | 3.791 ms | **-98.27% p50 (57.86x) / -98.25% mean** | p95 -98.13%; exact output/member identity; removal performs no resource GC |
| ODT direct snapshot sharing, 200 paragraphs + 16 MiB media | 32.270 ms | 7.798 ms | **-75.84% p50 / -73.84% mean** | Two archive-sized copies removed; p95 -75.41%; peak heap/RSS flat |
| ODT compact-audit package sharing, 200 paragraphs + 16 MiB media | 7.773 ms | 5.407 ms | **-30.44% p50 / -31.36% mean** | Three archive-sized audit copies removed; p95 -32.41%; allocations -0.57%; peak heap/RSS flat; exact no-op +39 ns disclosed |
| ODT envelope-classification sharing, 200 paragraphs + 16 MiB media | 5.555 ms | 4.921 ms | **-11.40% p50 / -11.95% mean** | One archive-sized envelope copy and two allocations/commit removed; p95 -12.19%; peak heap/RSS flat; large exact no-op +152 ns disclosed |
| ODT final changed-result byte handoff, 200 paragraphs + 16 MiB media | 5.216 ms | 4.030 ms | **-22.74% p50 / -22.56% mean** | One 16.79 MB result copy and redundant parse removed; p95 -21.48%; allocation calls -3.46%; independent final reopen and peak heap/RSS retained |
| ODT 1% paragraph edit/save, 10,000 paragraphs / 100 replacements | 906.439 ms | 15.615 ms | **-98.28% p50 (58.05x) / -98.27% mean** | One mutable candidate/publication/reopen/audit replaces 100; allocations -96.13%; peak heap and uninstrumented RSS flat; tool-inclusive RSS +9.93% disclosed |
| ODT mixed model-content publication, medium/large 80/320 operations | 25.640/25.052 ms scalar vs 0.803/0.785 ms batch; **31.93x–31.94x p50** | 2.759/2.756 s scalar vs 21.276/20.998 ms batch; **129.69x–131.24x p50** | One staged publication vs 49/193 scalar publications; per-shape output and logical hashes equal | Narrow repeated-publication comparison only; preparation, reopen/lifecycle/security/limits, I/O, serialization, allocation/RSS, and physical cold behavior excluded; see [change 0104](changes/0104-odt-mixed-model-publication-evidence.md) |
| Native XLS one-cell edit/save, 8,192 cells | 1.777 ms | 1.639 ms | **-7.72% p50 / -7.90% mean** | Allocation calls -1.19%; peak heap and uninstrumented RSS flat |
| Native XLS fixed-width numeric edit/save, 8,192 cells | 1.582 ms | 1.458 ms | **-7.83% p50 / -7.37% mean** | Complete public Workbook validation retained; peak heap -5.54%, RSS flat |
| Native DOC one-paragraph edit/save, 512 paragraphs | 1.506 ms | 1.348 ms | **-10.52% p50 / -10.48% mean** | Duplicate publication-site allocations nearly halved; peak heap and uninstrumented RSS flat |
| Native DOC open, 512 paragraphs | 790.727 us | 348.679 us | **-55.91% p50 / -55.78% mean** | Physical PieceTable scan self cycles 36.89% -> 4.17%; allocation calls +0.009%; peak heap and uninstrumented RSS flat |
| Native DOC one-paragraph edit/save after PieceTable index, 512 paragraphs | 1.379 ms | 0.950 ms | **-31.08% p50 / -31.68% mean** | Same private index accelerates mandatory candidate/public readbacks; patch/inverse and exact output checks unchanged |
| Native DOC open after PieceTable index, 512 paragraphs | 343.503 us | 304.199 us | **-11.44% p50 / -11.87% mean** | Paragraph-style validation 4.44% -> 0.83% self cycles; allocation calls -18.61%; peak heap and uninstrumented RSS flat |
| Native DOC one-paragraph edit/save after style cache, 512 paragraphs | 912.288 us | 875.736 us | **-4.01% p50 / -4.23% mean** | Same one-entry cache accelerates mandatory candidate/public readbacks; patch/inverse and exact output checks unchanged |
| Native DOC paragraph list after style cache, 512 paragraphs | 454.100 us | 358.414 us | **-21.07% p50 / -20.93% mean** | CHPX range query changes from a full scan per paragraph to binary start plus matching slice; p95 -20.00%; allocations and peak heap/RSS flat |
| Native DOC exact-source paragraph list after CHPX index, 512 paragraphs | 206.644 us | 168.142 us | **-18.63% p50 / -19.04% mean** | Ordered piece/PAPX containment uses predecessor binary search; instructions -26.13%; allocations and peak heap flat |
| Native DOC one-paragraph edit/save after PAPX containment index | 888.602 us | 817.424 us | **-8.01% p50 / -7.88% mean** | p95 -7.71%, p99 -8.37%; patch/inverse, candidate owner and independent public readback unchanged |
| Source-backed DOC Word97+ paragraph splice | — | — | Correctness-only; no matched release comparison | One ordinary main-story paragraph in one uncompressed Unicode piece with unchanged UTF-16 width; positional bounded chunk selector, exact no-op/source/fingerprint/stale checks, candidate reopen/readback, inverse and typed partial output; complete artifact fingerprints and CFB validation/publication scans remain; no end-to-end latency, I/O/range, allocation/RSS, cold/high-latency, real-producer, or broad DOC CRUD claim; see [change 0105](changes/0105-doc-source-backed-paragraph-splice.md) |
| Native PPT root snapshot open, 144 shapes | 37.522 us | 34.227 us | **-8.78% p50 / -10.58% mean** | Allocation calls -5.01%, temporary allocations -12.22%; peak heap and uninstrumented RSS flat |
| Native PPT direct text edit/save, 144 shapes | 206.209 us | 177.089 us | **-14.12% p50 / -15.39% mean** | Allocation calls -3.53%, temporary allocations -6.05%; peak heap/RSS flat; minor faults +315.43% with zero major faults |
| Native PPT root text edit/save, 144 shapes | 352.306 us | 286.805 us | **-18.59% p50 / -17.83% mean** | p95 -16.58%; allocation calls -6.54%; peak heap and uninstrumented RSS flat; custom limits retain full reopen |
| XLSX one-cell commit + first read, 4,096 cells | 4.431 ms | 3.402 ms | **-23.23% p50 / -23.15% mean** | Allocation calls -21.01%; peak heap +4.29%; unrestricted dense-wide retention rejected |
| Rejected XLSX 1% commit + save, 4,096 cells | 15.235 ms | 14.990 ms | -1.61% p50 / -1.26% mean | Fully reverted as immaterial; p99 +0.18%, peak heap flat |
| Rejected XLSX 1% commit + save, 131,072 cells | 514.926 ms | 511.407 ms | -0.68% p50 / -0.66% mean | Fully reverted as immaterial; process allocation calls -0.0623% |
| Source-backed XLSX calculation-metadata publication, 12 Parts + 16 MiB media | 215.457 ms | 1.612 ms | **-99.2519% p50 (133.67x) / -99.2507% mean** | Materializations 12 -> 1; allocation calls -10.81%; peak heap flat; uninstrumented RSS -1.20% |
| Source-backed XLSX defined-name publication, 12 Parts + 16 MiB media | 220.101 ms | 4.752 ms | **-97.84% p50 (46.32x) / -97.81% mean** | Materializations 12 -> 1; allocation calls -12.77%; peak heap and uninstrumented RSS flat |
| Source-backed XLSX page-break publication, 12 Parts + 16 MiB media | 216.789 ms | 4.647 ms | **-97.86% p50 (46.65x) / -97.86% mean** | Materializations 12 -> 2; allocation calls -15.95%; peak heap and uninstrumented RSS flat |
| Source-backed XLSX page-margin publication, 12 Parts + 16 MiB media | 216.799 ms | 4.492 ms | **-97.93% p50 (48.26x) / -97.93% mean** | Materializations 12 -> 2; allocation calls -12.10%; peak heap and uninstrumented RSS flat |
| Source-backed XLSX print-options publication, 12 Parts + 16 MiB media | 219.294 ms | 4.668 ms | **-97.87% p50 (46.98x) / -97.88% mean** | Materializations 12 -> 2; allocation calls -12.10%; peak heap and uninstrumented RSS flat |
| Source-backed XLSX page-setup publication, 12 Parts + 16 MiB media | 218.626 ms | 4.847 ms | **-97.78% p50 (45.10x) / -97.79% mean** | Materializations 12 -> 2; allocation calls -10.50%; peak heap and uninstrumented RSS flat |
| Source-backed XLSX sheet-protection publication, 12 Parts + 16 MiB media | 221.877 ms | 4.982 ms | **-97.75% p50 (44.54x) / -97.75% mean** | Materializations 12 -> 2; instructions -77.87%; allocation calls +2.73% within policy; peak heap and uninstrumented RSS flat |
| Source-backed XLSX data-validation publication, 12 Parts + 16 MiB media | 222.945 ms | 5.009 ms | **-97.75% p50 (44.51x) / -97.75% mean** | Materializations 12 -> 2; instructions -73.43%; allocation calls +4.92% within policy; peak heap flat and RSS -1.49% |
| Source-backed XLSX auto-filter publication, 12 Parts + 16 MiB media | 219.615 ms | 4.946 ms | **-97.75% p50 (44.40x) / -97.75% mean** | Materializations 12 -> 3; instructions -73.57%; allocation calls -1.94%; peak heap flat and RSS -1.35% |

The underlying records are:

- [`0001-opc-publication-plan.md`](changes/0001-opc-publication-plan.md)
- [`0002-cfb-lookup-and-sector-buffers.md`](changes/0002-cfb-lookup-and-sector-buffers.md)
- [`0003-legacy-owned-stream-handoff.md`](changes/0003-legacy-owned-stream-handoff.md)
- [`0004-opc-exact-owned-source.md`](changes/0004-opc-exact-owned-source.md)
- [`0008-targeted-opc-preservation.md`](changes/0008-targeted-opc-preservation.md)
- [`0009-range-source-and-scaling.md`](changes/0009-range-source-and-scaling.md)
- [`0010-docx-pptx-semantic-queries-and-edits.md`](changes/0010-docx-pptx-semantic-queries-and-edits.md)
- [`0011-odf-semantic-baseline-and-ods-snapshot.md`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md)
- [`0012-docx-coalesced-paragraph-edits.md`](changes/0012-docx-coalesced-paragraph-edits.md)
- [`0013-rtf-semantic-baseline-and-text-paths.md`](changes/0013-rtf-semantic-baseline-and-text-paths.md)
- [`0014-odt-shared-snapshot-bytes.md`](changes/0014-odt-shared-snapshot-bytes.md)
- [`0015-native-ole2-semantic-baseline.md`](changes/0015-native-ole2-semantic-baseline.md)
- [`0016-xls-commit-editor-reuse.md`](changes/0016-xls-commit-editor-reuse.md)
- [`0017-doc-batched-stream-publication.md`](changes/0017-doc-batched-stream-publication.md)
- [`0018-ods-row-local-publication.md`](changes/0018-ods-row-local-publication.md)
- [`0019-rtf-parser-state-specialization.md`](changes/0019-rtf-parser-state-specialization.md)
- [`0020-rtf-ascii-transport-batching.md`](changes/0020-rtf-ascii-transport-batching.md)
- [`0021-opc-shared-regenerated-payload.md`](changes/0021-opc-shared-regenerated-payload.md)
- [`0022-zip-generated-local-span-move.md`](changes/0022-zip-generated-local-span-move.md)
- [`0023-odt-full-text-owned-blocks.md`](changes/0023-odt-full-text-owned-blocks.md)
- [`0024-ppt-slide-order-open-reuse.md`](changes/0024-ppt-slide-order-open-reuse.md)
- [`0025-xlsx-validated-store-handoff.md`](changes/0025-xlsx-validated-store-handoff.md)
- [`0026-ppt-text-edit-resolver-reuse.md`](changes/0026-ppt-text-edit-resolver-reuse.md)
- [`0027-ods-adaptive-cell-locator.md`](changes/0027-ods-adaptive-cell-locator.md)
- [`0028-xls-terminal-render-handoff-rejected.md`](changes/0028-xls-terminal-render-handoff-rejected.md)
- [`0029-rtf-transport-and-producer-coverage.md`](changes/0029-rtf-transport-and-producer-coverage.md)
- [`0030-xlsx-action-plan-flattening-rejected.md`](changes/0030-xlsx-action-plan-flattening-rejected.md)
- [`0031-ods-unchanged-media-preservation.md`](changes/0031-ods-unchanged-media-preservation.md)
- [`0032-xlsx-no-extension-scan.md`](changes/0032-xlsx-no-extension-scan.md)
- [`0033-ole-common-publication-handoffs-rejected.md`](changes/0033-ole-common-publication-handoffs-rejected.md)
- [`0034-odp-unchanged-media-preservation.md`](changes/0034-odp-unchanged-media-preservation.md)
- [`0035-odt-content-only-paragraph-publication.md`](changes/0035-odt-content-only-paragraph-publication.md)
- [`0036-ole-common-stage-attribution.md`](changes/0036-ole-common-stage-attribution.md)
- [`0037-opc-source-backed-one-part-publication.md`](changes/0037-opc-source-backed-one-part-publication.md)
- [`0038-odt-direct-snapshot-sharing.md`](changes/0038-odt-direct-snapshot-sharing.md)
- [`0039-docx-source-backed-semantic-publication.md`](changes/0039-docx-source-backed-semantic-publication.md)
- [`0040-rtf-byte-delimiter-scanning.md`](changes/0040-rtf-byte-delimiter-scanning.md)
- [`0041-odt-compact-audit-package-sharing.md`](changes/0041-odt-compact-audit-package-sharing.md)
- [`0042-odt-envelope-package-sharing.md`](changes/0042-odt-envelope-package-sharing.md)
- [`0043-rtf-decoded-body-ownership-rejected.md`](changes/0043-rtf-decoded-body-ownership-rejected.md)
- [`0044-pptx-source-backed-semantic-publication.md`](changes/0044-pptx-source-backed-semantic-publication.md)
- [`0045-odt-coalesced-paragraph-publication.md`](changes/0045-odt-coalesced-paragraph-publication.md)
- [`0046-xlsx-source-backed-calculation-metadata-publication.md`](changes/0046-xlsx-source-backed-calculation-metadata-publication.md)
- [`0047-odt-indexed-paragraph-selector.md`](changes/0047-odt-indexed-paragraph-selector.md)
- [`0048-rtf-retained-body-source-span.md`](changes/0048-rtf-retained-body-source-span.md)
- [`0049-odp-indexed-slide-selector.md`](changes/0049-odp-indexed-slide-selector.md)
- [`0050-doc-piece-table-physical-index.md`](changes/0050-doc-piece-table-physical-index.md)
- [`0051-doc-adjacent-style-baseline-cache.md`](changes/0051-doc-adjacent-style-baseline-cache.md)
- [`0105-doc-source-backed-paragraph-splice.md`](changes/0105-doc-source-backed-paragraph-splice.md)
- [`0052-odt-final-result-byte-handoff.md`](changes/0052-odt-final-result-byte-handoff.md)
- [`0053-doc-chpx-range-index.md`](changes/0053-doc-chpx-range-index.md)
- [`0054-ods-shared-durable-patch-blobs.md`](changes/0054-ods-shared-durable-patch-blobs.md)
- [`0055-rtf-body-block-reservation.md`](changes/0055-rtf-body-block-reservation.md)
- [`0056-doc-papx-containment-index.md`](changes/0056-doc-papx-containment-index.md)
- [`0057-ods-row-splice-raw-publication.md`](changes/0057-ods-row-splice-raw-publication.md)
- [`0058-ods-exact-noop-handoff.md`](changes/0058-ods-exact-noop-handoff.md)
- [`0059-xls-fixed-numeric-inventory-carry.md`](changes/0059-xls-fixed-numeric-inventory-carry.md)
- [`0060-odp-snapshot-slide-projection-reuse.md`](changes/0060-odp-snapshot-slide-projection-reuse.md)
- [`0061-xlsx-source-backed-page-break-publication.md`](changes/0061-xlsx-source-backed-page-break-publication.md)
- [`0062-ppt-root-text-publication-adoption.md`](changes/0062-ppt-root-text-publication-adoption.md)
- [`0063-pptx-atomic-source-backed-shape-text-batch.md`](changes/0063-pptx-atomic-source-backed-shape-text-batch.md)
- [`0064-rtf-retained-story-length.md`](changes/0064-rtf-retained-story-length.md)
- [`0065-odp-final-snapshot-handoff.md`](changes/0065-odp-final-snapshot-handoff.md)
- [`0066-rtf-sparse-paragraph-nth.md`](changes/0066-rtf-sparse-paragraph-nth.md)
- [`0067-xlsx-source-backed-page-margin-publication.md`](changes/0067-xlsx-source-backed-page-margin-publication.md)
- [`0068-ods-shared-worksheet-archive-handoff.md`](changes/0068-ods-shared-worksheet-archive-handoff.md)
- [`0069-rtf-retained-paragraph-count.md`](changes/0069-rtf-retained-paragraph-count.md)
- [`0070-xlsx-source-backed-print-options-publication.md`](changes/0070-xlsx-source-backed-print-options-publication.md)
- [`0073-xlsx-source-backed-page-setup-publication.md`](changes/0073-xlsx-source-backed-page-setup-publication.md)
- [`0071-odt-content-only-line-break-publication.md`](changes/0071-odt-content-only-line-break-publication.md)
- [`0072-odt-content-only-run-publication.md`](changes/0072-odt-content-only-run-publication.md)
- [`0074-odt-content-only-hyperlink-publication.md`](changes/0074-odt-content-only-hyperlink-publication.md)
- [`0075-odt-structural-paragraph-publication.md`](changes/0075-odt-structural-paragraph-publication.md)
- [`0076-xlsx-source-backed-defined-names-publication.md`](changes/0076-xlsx-source-backed-defined-names-publication.md)
- [`0077-pptx-source-backed-multi-slide-batch-publication.md`](changes/0077-pptx-source-backed-multi-slide-batch-publication.md)
- [`0078-xlsx-source-backed-sheet-protection-publication.md`](changes/0078-xlsx-source-backed-sheet-protection-publication.md)
- [`0079-xlsx-source-backed-data-validation-publication.md`](changes/0079-xlsx-source-backed-data-validation-publication.md)
- [`0080-xlsx-source-backed-auto-filter-publication.md`](changes/0080-xlsx-source-backed-auto-filter-publication.md)

The DOC ownership-transfer variant was rejected and removed after a 58.42%
p50 regression. The earlier full-rewrite mutated-OPC guardrail was neutral on
incompressible data; targeted raw publication supersedes it only for the
strictly proved same-topology owned-source case. Fallback still uses that
validated full rewrite. Rejected, fallback and memory results are retained
rather than hidden in an aggregate.

An ODS target-package adoption candidate was likewise removed after large
one-cell edit/save improved only 0.44% p50 and p95 regressed 0.30%. The existing
package/readback boundary remains; no production or test code from that
candidate is retained.

An ODT final-document adoption candidate was also fully reverted. It improved
large one-edit/save p50 5.70%, but a dedicated medium one-paragraph read guard
regressed 6.33% mean and 17.64% p95. The accepted snapshot-byte sharing remains;
the rejected parsed-document retention contributes no production or test code.
Change 0052 is deliberately narrower: it shares only immutable final bytes and
retains a fresh independent final reopen; its same guard stays within 3% p50
and mean with a better p95.

The first ODT full-text ownership candidate also moved strings for structured
list and one-paragraph callers. Their large-corpus p50 regressed 5.71% and
5.30%, respectively, so that broad version was removed. The accepted private
full-text mode retains the original structured path; the rejected raw reports
remain linked from change 0023.

## Work removed

- Exact unchanged owned OPC publication no longer regenerates manifests,
  reconstructs ZIP records, or recompresses logical Parts. It copies the
  complete validated source to the caller's sequential sink in writes bounded
  to 64 KiB and verifies complete output in the benchmark.
- Targeted same-topology OPC publication no longer recompresses unchanged
  Parts. It audits the ordinary publication plan, regenerates only changed
  payload/relationship/content-type closures, and raw-copies unchanged local
  spans and central records, including unknown non-part members.
- The low-level source-backed one-Part publisher no longer converts the
  positional package into an eager owning package or recompresses every Part.
  It materializes and validates the selected original payload, regenerates that
  member, and raw-copies every other member while monitoring source version.
- The changed ordinary Part now shares its already-owned immutable logical
  payload with ZIP regeneration rather than allocating and copying it again.
  Generated XML and the required compressor/archive buffer stay owned.
- After the generated member has passed complete ZIP validation, its local span
  now moves into the prepared entry instead of being allocated and copied a
  second time. Central-directory framing remains separately retained.
- Rewritten OPC publication constructs and audits generated XML and stable
  Part order once before emission rather than once for validation and again
  for writing.
- CFB lookup follows the validated sibling-tree ordering with SID-aligned
  cached comparison keys rather than scanning the complete sibling tree.
- CFB FAT/DIFAT/MiniFAT parsing reuses a bounded sector buffer, MiniFAT decodes
  into its final table, and directory sectors read into their final buffer.
- `SharedOleFile::read_stream_range` follows only the logical sectors needed by
  a caller-owned bounded range and leaves the lazy MiniFAT root-stream cache
  untouched. Change 0094's MiniFAT ABBA removes the full-stream source request
  amplification while preserving exact payload hashes; FAT remains a one-call,
  one-4-MiB-request control. The release result is substrate evidence only and
  does not imply DOC/XLS/PPT semantic adoption.
- Fresh XLS and PPT writers transfer already-owned generated stream buffers to
  CFB without a second payload copy. DOC deliberately retains its measured
  faster exact-sized copy.
- Native XLS changed commit reuses the already rendered/reopened object editor
  instead of discarding one BIFF owner parse and reopening/capturing the CFB a
  second time before final validation.
- Native DOC applies ordinary WordDocument and table-stream replacements to
  one isolated object-editor candidate and renders/reopens the CFB once rather
  than once per replacement.
- Native DOC paragraph FKP parsing reuses one resolved initial-style baseline
  across repeated source runs instead of reconstructing and revalidating the
  same inheritance chain. Direct properties, piece modifiers and direct style
  switches still execute independently for every PAPX.
- DOCX one-paragraph selection no longer constructs the complete paragraph
  collection, and source-backed paragraph counts no longer construct any
  paragraph views. Complete XML validation and limits still run.
- Canonical multi-paragraph DOCX text edits now plan disjoint replacements and
  emit/reparse one candidate XML document instead of rebuilding and reparsing
  the complete main document once per paragraph. Durable patches remain
  ordinary source-checked paragraph operations with complete final readback.
- Repeated PPTX shape-text edits no longer parse the selected slide scene a
  second time solely to map the already selected shape to its raw XML span.
- DOCX plaintext package output exposes the underlying forward-only OPC sink
  instead of imposing an unused `Seek` bound.
- ODS unified snapshot construction reuses its one validated package for full
  facade readback instead of cloning package bytes and parsing the package a
  second time.
- Eligible same-topology ODS worksheet commits serialize only changed modeled
  rows instead of regenerating every worksheet row; untouched XML source spans
  are copied exactly and structural edits retain full-table fallback.
- Eligible compact ODS `content.xml` commits raw-copy every other validated ZIP
  member. Source/target effect checks use exact local and central member bytes
  to skip unchanged payload inflation only while the manifest is also exact;
  every unproved case retains logical comparison and established rebuild.
- Eligible same-topology ODS row edits now carry their already checked exact
  source ranges through that raw package publisher. They no longer fall back
  to recompressing unchanged media merely because the flattened result cannot
  be rediscovered as one conservative maximal diff.
- The adjacent unified ODS worksheet handoff now moves and shares its exact
  archive allocation across nested worksheet snapshots, package parsing,
  commit readback and candidate validation instead of repeatedly copying it.
  Failure paths restore the original bytes and allocation; durable patch and
  final validation boundaries are unchanged.
- RTF first full-text materialization retains only a byte count during parse,
  then allocates the final string once and copies blocks in one pass instead of
  allocating and joining a temporary fragment vector.
- RTF borrowed stories now receive that already validated byte count instead
  of rescanning every retained style block to establish paragraph and inline
  iterator endpoints.
- RTF canonical text emission writes ordinary ASCII spans in chunks instead of
  one formatted write per character. Text-only commits skip paragraph-property
  vectors/scans, and a successful paragraph selector stops at its target.
- Ordinary RTF body-text flushes no longer clone the complete parser state.
  They copy the effective encoding and block properties; insertion/deletion
  runs alone retain full state for revision author/date and exact range data.
- All-ASCII RTF source tokens now enter transport buffers in one extension
  rather than one generic `SmallVec::extend` call per character. The checked
  byte-valued non-ASCII and invalid-Unicode fallback is unchanged.
- ODT transaction snapshots created from an already validated `Document` clone
  its private immutable package handle instead of allocating and copying the
  complete archive. Direct snapshot byte ingress keeps independent validation.
- ODT full-text extraction moves each parser-created validated block string
  into the element and consumes it into final output instead of cloning the
  string at both private handoff boundaries. Structured block queries retain
  their original ownership behavior.
- ODT line-break, run, hyperlink, and plain paragraph insertion/removal
  transactions now use the accepted content-only publisher. They regenerate
  compact `content.xml` while raw-copying every eligible unchanged core/media
  member; other structural, mixed, oversized and security-sensitive cases
  retain the established path.
- Eligible XLSX changed sheets move their exact commit-validated semantic
  store into the published snapshot after byte and style/shared-string lineage
  checks. Retention is capped at 4,096 cells and 1 MiB of worksheet XML; larger
  sheets keep the cold-cache path.
- Direct PPT text-edit setup uses its full editor preflight to resolve the
  selected persisted record instead of opening and capturing the CFB a second
  time. Commit still opens a fresh editor and performs exact source comparison,
  publication, complete snapshot reopen and semantic readback.
- Repeated public ODS cell queries lazily build one private, sheet-aligned
  locator after 64 successful lookups. Direct runs retain compact row
  descriptors; repeated runs add cumulative endpoints under a 4 MiB cap and
  any build failure permanently falls back to linear lookup.

No unsafe code, ambient I/O, dependency edge, public archive type, or global
synchronization primitive was introduced. Exact-source authorization is
revoked conservatively on every mutable OPC entry point, including failed and
semantic no-op calls. Borrowed ingress, topology-changing edits, and unsupported
ZIP layouts use the fully validated owning rewrite path; the narrow
source-backed publisher instead returns a typed zero-output refusal.

## Evidence and verification

The current standalone harness provides 320 selectable cases. Change
0091 adds four committed opt-in XLS visibility selectors, change 0094 adds four
committed opt-in CFB selective-range selectors, and change 0099 adds one opt-in
ODF repair-plan selector. The visibility and repair selectors
are correctness/coverage evidence only. Change 0094 has a pinned 30-warmup,
500-sample release ABBA summary: MiniFAT exact-range source bytes fall from
261,184 to 36 and from 2,096,192 to 36, with stable read-stage p50/p95 gains
and only modest total-p50 movement; FAT retains one 4 MiB read request/call.
No p99, cold-filesystem, allocation, peak-RSS, or DOC/XLS/PPT semantic claim
is accepted by that record. Change 0144 separately accepts only the configured
harness-simulator result: the 36- and 4095-byte MiniFAT targets reduce to one
exact request and improve total p50/p95 in both 200-sample ABBA directions,
while the exact-work FAT control stays near neutral. This is not real
cold/network/device evidence. See the [exact-range summary](results/cfb-selective-range-abba-0106-summary.json)
and [simulated-range summary](results/cfb-simulated-range-0144-summary.json).

Change 0145 adds two opt-in PPTX cross-presentation slide-copy selectors with
separate plan, commit, sequential OPC publication, and reopen diagnostics.
Exact slide XML/media payloads, relationship topology, content types,
untouched destination members, output semantics, source immutability, durable
patches, and refusal paths are checked outside timing. This is correctness and
sink-counter evidence only; no speedup, allocation, RSS, release-ABBA, or
physical-I/O claim is made. See
[`0145`](changes/0145-pptx-cross-slide-copy-evidence.md).

Change 0146 adds twelve opt-in selectors that call the public generic-CFB
`SharedOleFile::open_stream` path for exact 36-byte and 4,095-byte MiniFAT
targets. One-shot, repeat-3, and sequential repeat-8 operations retain exact
per-invocation hashes, source events, root identity, version/refusal gates, and
matched deterministic-range-model evidence. This is correctness/counter
evidence only; release ABBA and all latency, allocation, RSS, physical-I/O,
native DOC/XLS/PPT, cross-format, and iWork claims remain open. See
[`0146`](changes/0146-cfb-open-stream-evidence.md).

Change 0147 supplies the clean CPU-2 release `A1/B1/B2/A2` comparison for all
24 target/shape records. Under the exact configured range simulator, all four
one-shot cells improve total p50/p95/p99/mean by about 62-64% in both
directions and reduce exact source work from the complete root Mini Stream to
one target range. Repeats retain the candidate's extra first request and show
small many-small modeled regressions, so no generic repeat improvement is
claimed. Local wall-clock, resource, physical-I/O, native-format, and
cross-format claims remain open. See
[`0147`](changes/0147-cfb-open-stream-release-abba.md) and the
[compact summary](results/cfb-open-stream-abba-0147-summary.json).

Change 0148 extends the current harness with six production-only CFB selectors
for different-SID A-B-A, public bulk A-B-A, and overlapping same-target calls
at 36-byte and 4,095-byte MiniFAT targets. They retain ordered workload names,
output hashes/lengths, exact source positional events, source-version checks,
and typed missing-stream refusal. This is correctness/source-event evidence
only; failure/retry, ineligible-root, FAT, native semantic, resource, and
performance acceptance for those extended selectors remain open. See
[`0148`](changes/0148-cfb-same-target-repeat-policy.md).

Change 0149 retains a clean release comparison for the target-aware repeat
policy. Four CPU-2 legs in strict `A1 control, B1 candidate, B2 candidate, A2
control` order cover 36 records each, 20 warmups, and 200 samples, for 28,800
retained samples. The identical-harness control restores only `shared.rs` and
`shared_bulk.rs` from the immediate pre-change production revision.

Under the configured range model, all eight repeat-3/repeat-8 aggregate-total
cells improve in both adjacent directions: roughly 60-64% for repeat-3 and
56-64% for repeat-8 at p50, with matching p95/p99/mean direction. Exact source
work changes from `[L,R,0...]` to `[L,L,...]`; configured-simulator one-shot
controls remain near neutral. Later per-invocation calls are target reads rather
than zero-source cache hits, and the noisy local bulk/concurrent distributions
contain explicit >5% review triggers. Consequently no local, per-invocation,
bulk, concurrent, allocation/RSS, physical-I/O, cold/network/device, native-
format, or generic claim is accepted. See the
[release record](changes/0149-cfb-same-target-repeat-release-abba.md) and
[machine-readable summary](results/cfb-repeat-abba-0149-summary.json).

Change 0152 supplies the final clean release ABBA for same-target MiniFAT
single-flight: control `e486e4b1` versus candidate `f46381c6` (introduced by
`c270c8f3b`) on CPU 2, with 20 warmups, 500 samples, 24 records per leg, and
48,000 retained samples. All correctness/source-event invariants pass, and
existing concurrent scenarios record 6,473 candidate versus 8,000 control
logical source calls (19.09% fewer). This accepts source-event/correctness
evidence only. At the 0152 revision the 291-name matrix was unchanged; change
0153 adds four RTF selectors measured at the pre-staged publication-call
interval, making that matrix 295. Change 0154 adds six ODF content-COW
publication selectors, making that matrix 301; change 0159 later made it 302,
change 0160 made it 303, change 0162 made it 305, change 0163 made it 309,
change 0164 made it 311, change 0166 made it 315, change 0174 made it 319, and
change 0175 makes the current matrix 320.
No runtime selector was added to 0152; only `cfg(test)` source-event
acceptance and tests changed. Root
MiniStream cache/resource-accounting boundaries and broader performance gaps
remain, while local/generic latency, allocation/RSS/peak memory, physical
I/O/syscalls, cold-cache/device/network, decompression, native semantic,
OOXML/ODF/RTF/iWork claims are withheld. See the
[0152 release record](changes/0152-cfb-same-target-singleflight-release-abba.md)
and [summary](results/cfb-singleflight-abba-0152-summary.json).

Change 0153 adds four matched RTF tail selectors. Their pre-staged
Commit/PublicationPlan `elapsed_ns` is exactly the publication-call interval
around the respective public write call to a fixed 16 KiB sink; the calls have
intentionally asymmetric validation and publication work. Planning, reopen,
lifecycle, durable patch, cancellation, sink failure/partial progress, limits,
source-version and exact semantic gates remain untimed. Retained
source/candidate/window bytes are explicit.
No end-to-end, rich-format, allocation/RSS, physical-I/O, or ABBA latency
claim is made. See
[0153](changes/0153-rtf-tail-publication-plan-evidence.md).

Change 0154 adds matched owned-rebuild/source-positional publication selectors
for ODT, ODS, and ODP. Clean CPU-2 A/B/B/A evidence accepts 96.35%-96.63% p50
improvement in both pair directions at the prepared publication-call boundary;
p95, p99, and mean agree. Exact semantic/content/raw-order/no-op/limit/
cancellation/source gates remain outside timing. No end-to-end,
allocation/RSS, physical-I/O, decompression, cold-cache, filesystem,
real-producer, broad ODF CRUD, or iWork claim is made. See
[0154](changes/0154-odf-content-cow-publication-evidence.md) and the
[summary](results/odf-content-cow-abba-0154-summary.json).

Change 0119 adds three opt-in native PPT selected-shape controls and preserves
the 36-case / 198-record default. The query-only and fresh-open-plus-query
eager/source-backed pairs use the same deterministic target; separate untimed
source replays retain exact logical counters and selected-text hashes. They are
correctness and fixture-scoped logical-read evidence only, with no accepted
latency, physical-I/O, allocation/RSS, cold-cache, or publication claim.

Change 0121 adds two opt-in native PPT repeated selected-shape selectors,
bringing the matrix to 229 names at that point (before changes 0122-0124) while
preserving the default 36-case / 198-record tranche. The matched eager/source-backed controls retain one
prepared owner for eight identical queries; source timing uses an
uninstrumented source and independent replays report exact logical calls,
bytes, prior-covered range bytes, and a canonical semantic digest. The frozen
production two-query regression binds 74 calls / 8,310 bytes for legacy CFB
reconstruction and 66 calls / 3,190 bytes with a retained parsed CFB index.
This is logical-I/O/correctness evidence, not a latency or resource claim.

Changes 0122, 0123, and 0124 add four ODP media-rich, four ODP unified-root
filesystem, and six ODS unified-root/source selectors respectively. They move
the selectable matrix from 229 to 233, 237, and finally 243 names while
preserving the default 36-case / 198-record tranche. Each change keeps corpus
and owner preparation outside the named timing boundary, verifies complete
semantic/metadata/member/hash parity, and reports only correctness and logical
compressed-range evidence; no latency, physical-I/O, decompression,
allocation, RSS, or release claim is accepted. See [`change 0122`](changes/0122-odp-media-source-read-evidence.md),
[`change 0123`](changes/0123-odp-unified-root-filesystem-evidence.md), and
[`change 0124`](changes/0124-ods-unified-root-filesystem-evidence.md).

Change 0125 adds two matched 4095-byte MiniFAT boundary selectors, preserving
the 36-case / 198-record default and bringing the selectable matrix to 245
names. The focused gate records separate open/read/total timing, exact source
calls/bytes/range sizes, and payload hashes; it expects legacy root-mini-stream
amplification and one exact 4095-byte positional request over the 64 logical
mini-sectors. This is correctness/request-amplification evidence only, with
no release latency, tail, cold/high-latency, allocation/RSS, physical-I/O, or
native semantic claim.

Change 0126 adds eight ordinary-root DOCX filesystem selectors and brings the
selectable matrix to 253 names without changing the default 36 cases / 198
records. The unchanged source-edit corpus contains 200 paragraphs and eight
deterministic incompressible 2 MiB media Parts. Eager open times `fs::read` plus
`Document::from_bytes`; source open times `Document::open(path)`; each query
root is prepared outside the timer and only its exact query is timed. An
independent typed `litchi_docx::source_backed::Package` replay records calls,
bytes, request sizes, compressed-range coverage, and materializations. It
requires zero main/media/unselected/core overlap at source open, complete
coverage of the compressed main range during query-selector preparation, and
zero such overlap after the query begins. Untimed eager/source parity covers
semantic projections and metadata; exact source SHA plus logical OPC
part/relationship/content-type/blob-hash gates cover package preservation,
including media hashes and source immutability. This is correctness/logical-
range evidence only; no latency, physical-I/O, decompression, allocation, RSS,
cold-cache, ABBA, broad-security, or Markdown-performance claim is accepted.
See [`0126`](changes/0126-docx-root-source-path-evidence.md).

Change 0127 adds two matched ODS repeated-cell sweep selectors, bringing the
selectable matrix to 255 names without changing the default 36 cases / 198
records. The fixed two-sheet 32 by 32 media-rich corpus is opened outside the
timer; four identical sweeps include the adaptive locator transition. An
independent instrumented replay per measured source sample resets counters
after preparation and requires zero source reads during the sweep. Semantic
digest/count plus source SHA, member topology, semantic grid, manifest-media,
and retained-media payload checks remain untimed. This is
correctness/logical-read evidence only; no latency, physical-I/O,
decompression, allocation, RSS, cold-cache, ABBA, or release claim is accepted.
See [`0127`](changes/0127-ods-source-cell-sweep-evidence.md).

Change 0134 adds matched eager/source-backed ODS ordered cell-batch sweep
selectors over the same two-sheet 32 by 32 media-rich corpus. Owners and
2,048 borrowed selectors are prepared before timing; each timed sample covers
four bounded `cell_batch` calls and 8,192 black-boxed result slots. An
independent source replay records exactly eight version observations and zero
post-preparation payload reads per four-call sweep. Ordered digest/count and
source/member/media identity checks remain untimed. At that stage the
selectable matrix was 257 while the default 36 cases / 198 records remained
unchanged. This is
correctness/logical-read evidence only, with no release speed or resource claim
without ABBA. See
[`0134`](changes/0134-ods-source-cell-batch-sweep-evidence.md).

Change 0135 adds four opt-in native XLS fixed-width numeric selectors and brings
the current selectable matrix to 261 names while preserving the default
36-case / 198-record tranche. The Number controls reuse the deterministic
comments corpus `Untouched!E21` (`42` -> `43`); the RK/MulRK controls use one
standalone RK and one two-cell MulRK record and edit all three values in one
transaction. Timed vectors separate transaction creation, `set_number` or
`set_numeric`, eager/source-backed commit, and complete publication to one
preallocated sink. Untimed gates cover source ingress, expected outputs,
complete Snapshot/Workbook reopen, storage-family/value readback, deterministic
digests, untouched CFB topology/member bytes, equal Workbook lengths on the
source-backed path, patch apply/inverse/stale, no-op/fingerprint,
signed/macro/protected/unsupported refusal, and the 54016.xls real-producer
reopen/inverse check. Source evidence reports complete target materialized
bytes on both paths and explicitly does not claim bounded artifact memory or
positional I/O; no speedup or broad-producer claim is accepted.

Change 0136 records the first pinned release baseline for those four selectors
at clean revision `9577cd16f` (CPU 2, 20 warmups, 200 samples). Number p50 is
31.492 ms eager and 146.410 ms source-backed; RK/MulRK p50 is 0.100 ms eager
and 1.627 ms source-backed. P95/p99 are 34.116/35.916 ms versus
149.108/150.693 ms for Number and 0.120/0.127 ms versus 1.659/1.690 ms for
RK/MulRK. Both implementations retain the complete target and produce the same
family output hash. The 4.65x/16.25x source-backed/eager ratios are descriptive
matched-implementation baselines, not a before/after regression result. The
run is single-process and warm, with no allocation, RSS, hardware-counter,
physical-I/O, or cold-cache evidence. See
[`0136`](changes/0136-xls-numeric-current-revision-baseline.md) and the linked
raw schema-1 artifact.

Change 0137 adds two opt-in plan-only native XLS numeric selectors, bringing
the selectable matrix to 263 while preserving the default 36 cases / 198
records. The plan-only Number and RK/MulRK paths include validated overlay-plan
construction and composed semantic validation in the commit vector, then time
complete `write_to` publication separately. They retain no complete target
snapshot or target byte vector at commit, report zero target materialization and
unsupported patch/inverse fields, and still publish complete sink bytes. Full
reopen, topology/member identity, no-op and exact source/target fingerprint
preflights, partial-sink, security/unsupported and 54016.xls forward producer
gates remain untimed. Composed semantic validation may allocate/read a
candidate Workbook model, so zero target-artifact bytes at commit is not a
bounded total-memory claim. This
is correctness/descriptive evidence only; no latency, allocation, RSS, I/O,
memory, or speedup claim is accepted before balanced release ABBA. See
[`0137`](changes/0137-xls-numeric-plan-only-publication.md).

Change 0138 is the balanced release acceptance record for the plan-only
candidate. A clean exact-revision release binary ran Number and RK/MulRK in
strict `A1, B1, B2, A2` order, one process at a time, with 20 warmups and 200
samples. The complete-operation p50/p95/p99/mean candidate/control directions
all agree: Number p50 improves 27.57% and 28.58%, and RK/MulRK 24.90% and
24.56%. Commit phase directions agree as well; publication is near-neutral and
is not accepted in isolation. A matched `/usr/bin/time -v` 3-warmup/30-sample
capture records process VmHWM reductions of 10.73% and 10.66% for Number in
both directions, but disagreeing RK/MulRK RSS. Valid heaptrack A/B profiles
show whole-process allocation reductions but identical 205.56/154.93 MiB
peak heaps for Number/RK, so no operation-only allocation or peak-heap claim
is accepted. Exact family output hashes and complete plan-only
zero-target-artifact/correctness gates hold on every leg. See
[`0138`](changes/0138-xls-numeric-plan-only-release-abba.md) and the linked
schema-1 raw results.

Change 0139 adds two opt-in source-backed ODP repeated-text selectors, taking
the selectable matrix from 263 to 265 while preserving the default 36 cases /
198 records. Both selectors prepare the same 12-slide, eight-picture owner and
four output slots before timing, then perform four full-text projections. The
control reproduces the historical uncached public sequence; the candidate
calls `SourceBackedPresentation::text()` and exercises the threshold-two
cache. Untimed source replays bind zero post-preparation reads and exact
freshness vectors `[3,3,3,3]` and `[3,5,2,2]`, with archive, text, and 16 MiB
media identity retained. This is correctness and logical replay evidence
only; no latency, physical-I/O, decompression, allocation, RSS, cold-cache,
ABBA, or release claim is accepted. See
[`0139`](changes/0139-odp-repeated-text-cache-evidence.md).

Change 0140 supplies the clean-revision release result for that exact matched
selector pair. Four fresh CPU-2 processes ran in `A1, B1, B2, A2` order with
20 warmups and 200 samples: p50 improves 45.80%/46.32%, p95 45.25%/45.83%,
p99 39.91%/45.41%, and mean 45.74%/46.33% in the paired directions. Four
Heaptrack profiles record identical-direction whole-process reductions of
14.31% in allocation calls and 17.25% in temporary allocations; peak heap is
unchanged at 89.22M. Process VmHWM is near-neutral (0.00%/0.16%), so no RSS or
peak-heap reduction is accepted. The result applies only to four repeated
full-text projections on the prepared source-backed corpus and makes no
single-call, open, physical-I/O, decompression, cold-cache, operation-local
allocated-byte, or generic ODF claim. See
[`0140`](changes/0140-odp-repeated-text-cache-release-abba.md).

Change 0144 adds six opt-in CFB simulated-range selectors, taking the current
selectable matrix from 265 to 271 while preserving the default 36 cases / 198
records. A clean release `A1, B1, B2, A2` run on CPU 2 uses 20 warmups and 200
samples for each target/shape. Both MiniFAT targets reduce the selected read to
one exact request and improve total p50/p95 in both directions; the matched
4 MiB FAT control keeps identical request, byte, and modeled-service work and
stays near neutral. The claim is limited to the named configured simulator;
real cold/network/device I/O, production scheduling, allocation/RSS, and
native DOC/XLS/PPT semantic adoption remain open. See
[`0144`](changes/0144-cfb-simulated-range-source-evidence.md).

Change 0146 adds twelve generic CFB `open_stream` selectors, taking the current
selectable matrix from 273 to 285 while preserving the default 36 cases / 198
records. The current tree's exact counter/range formulas and the parent
revision's root-materializing shape are both supported by the same runner.
This is correctness/counter evidence only, pending clean release ABBA. See
[`0146`](changes/0146-cfb-open-stream-evidence.md).

Change 0147 adds no selector. Its 19,200-sample clean release matrix accepts
only exact one-shot source-work reduction and the named configured simulator's
one-shot latency direction. The repeated-work cost remains explicit and no
generic local or resource claim is made. See
[`0147`](changes/0147-cfb-open-stream-release-abba.md).

Change 0103 adds a separate pinned release ABBA capture for
`cfb_file_same_length_overlay_atomic_save` on CPU 2 (five warm-ups and 30
fresh-child samples per leg). The atomic save's duplicate post-emission scan
was removed mechanically, taking its scan shape from `4N` to `3N`; direct
`write_to` retains its scan. Before legs report 2,084 logical calls and
101,751,908 requested/returned bytes; after legs report 1,825 calls and
84,838,500 bytes. The exact reductions are 259 calls (12.4280%) and
16,913,408 bytes (16.6222%), with the same 16,913,408-byte output SHA-256 on
all legs. The paired p50 directions disagree (+3.7963% and -10.0141%), so
this is logical source-work and correctness evidence only; no latency,
allocation, RSS, peak-memory, physical-cold, high-latency, or storage claim is
accepted. The four raw reports and compact summary are in
[`change 0103`](changes/0103-cfb-atomic-save-scan-evidence.md).

Change 0117 records two CPU-pinned balanced release attempts for eight native
PPT `Pictures` selectors. Timed source-backed samples use uninstrumented
`OwnedSource`, while separate untimed replays confirm zero `Pictures` overlap
at open, one complete stream read at a cold all-images query, and zero
additional reads at a cached query. The longer attempt used 20 warmups and 200
samples in each of eight fresh children per phase, including a directly timed
fresh open-plus-all-images pair. Every phase failed at least one fixed
same-implementation 5% p50 / 10% p95 drift gate, so the raw timing distributions
are retained but no latency result is accepted. Whole-process RSS was
91,136--91,584 KiB and is not per-operation memory attribution. No allocation,
cold-cache, hardware-counter, producer-breadth, or save-path claim is accepted.
See
[`change 0117`](changes/0117-ppt-pictures-release-evidence.md) and its
[raw report](results/ppt-pictures-release-0117.json).

Change 0120 adds eight opt-in PPTX ordinary-root filesystem selectors while
preserving the 36-case / 198-record default. They compare eager byte-root and
the unified source-path `litchi::Presentation::open(path)` root over the fixed
200-slide/eight-text-box/eight-2 MiB-media corpus for open, complete owned slide
listing, slide-count, and selector-first slide-100 queries. Source samples run
in fresh warm/cold-requested children and receive one separate untimed
`SourceBackedPresentation` replay. The replay must classify open/count as zero
slide/media payload overlap, selected as target-slide-only, and list as all
slide-payload/no-media overlap. Full source hashes, archive length, metadata,
slide size, names, text hashes, and eager/source parity are verified outside
timing; eager controls explicitly have no `ReadAt` replay. The counters are
logical compressed-range evidence for this generated corpus only. No latency,
tail, allocation, RSS, decompression, physical-I/O, cold-cache, or release-ABBA
claim is accepted. See
[`change 0120`](changes/0120-pptx-root-source-path-evidence.md).

The previously measured
default matrix remains 198 records across deterministic ZIP/OPC, positional
CFB/OPC, source-backed XLSX,
public DOC/XLS/PPT writer and semantic corpora, and DOCX/PPTX/RTF/ODT/ODS/ODP
semantic corpora. RTF includes deterministic raw CP-1252 and LZFu inputs plus
a content-addressed producer watermark; its separate native `relsize` chain is
an offline correctness gate rather than a timed paragraph case.
It records
p50/p95/p99, raw samples, mean, sample deviation, Student's-t 95% mean interval,
corpus/output hashes, environment, bounded sequential-write behavior,
deterministic logical/physical range distributions, and exact execution
tasks/bytes. The committed release cache artifact retains per-sample counters,
occupancy, Budget diagnostics, flights and waiters but intentionally does not
claim speedup; the repeated filesystem release artifact retains 300 fresh-child
tmpfs samples but does not prove physical cold-cache behavior. CI runs a
non-gating deterministic smoke check and a scheduled/manual release matrix.

The current local evidence includes consolidated changed-crate tests with
byte/hash checks, focused changed-crate warning-denied Clippy, formatter and
diff checks, YAML parsing, and JSON parsing. The umbrella all-feature `litchi`
attempt exhausted local disk, so it is not represented as a passing umbrella
gate. The historical stage-1 all-feature gate and its pre-existing Cargo
warning about DOCX/PPTX example output name `owner_native_smoke` remain scoped
to that earlier capture.

The repository-wide warning-denied rustdoc command remains blocked by existing
broken/private intra-doc links in unchanged OPC, DOC, XLS, and PPT files. The
dependency-direction checker unit suite passes, while the live policy check
reports existing unclassified edges (including `litchi-opc -> xml-minifier`
and several dev-only `-> soapberry-zip` edges); this tranche changes no Cargo
manifest or dependency edge. These pre-existing gate failures are not counted
as passing verification.

During the stage-1 capture, hardware counters were unavailable because that
host had `perf_event_paranoid=4`. The later targeted-OPC capture ran after the
environment reported `1`: matched process counters show cycles -69.21% and
instructions -69.85% for that one save cell. No counter claim is retroactively
made for stage 1 or generalized to other workloads. Later XLS, DOC and ODS
change records also retain matched process counters; the row-local ODS workload
reports cycles -5.47%, instructions -6.92%, and cache misses -6.58%.
The later RTF parser-state workload reports cycles -10.50%, instructions
-9.28%, and cache references -8.61%; its profiler removes the former 8.53%
exclusive state-clone frame. The subsequent RTF transport workload reports
cycles -11.22%, instructions -18.40%, and branches -14.04%; its per-byte
`SmallVec::extend` share falls from 15.37% to 2.56% on open. The OPC
shared-payload follow-up removes one 4.19 MiB allocation, cuts peak heap 3.42%,
task clock 21.08%, cycles 19.41% and cache misses 31.12% on its matched
few-large compressible process. The local-span follow-up removes the next 4.20
MiB allocation, cuts peak heap another 3.20% and task clock 2.11%; its other
major hardware counters stay within 5%. Uninstrumented RSS is flat for both.
The source-backed one-Part publisher removes three unnecessary Part
materializations and recompressions on its four-Part corpus: operation p50
falls 73.12%, instructions 65.42%, allocation calls 6.41%, peak heap 3.20% and
maximum observed uninstrumented RSS 3.26%. Physical source bytes remain flat
because every unchanged compressed span is still copied to the output.
The ODT full-text follow-up removes 420,019 allocation calls over ten samples,
cuts temporary allocations 45.52%, task clock 2.39%, instructions 2.51% and
cache misses 13.05%; peak heap and uninstrumented RSS remain flat.
The PPT root-snapshot follow-up removes 45 allocation calls per open, cuts
task clock 6.56%, instructions 9.57% and cycles 6.85%, and keeps peak heap and
uninstrumented RSS flat. Its 15.00% cache-miss increase is disclosed rather
than presented as a locality improvement.
The bounded XLSX changed-sheet handoff removes the duplicate public first-read
parse on eligible commits: task clock falls 24.25%, instructions 23.05%,
cycles 24.29% and allocation calls 21.01% on the medium attribution process.
Peak heap rises 4.29% under the bound and uninstrumented RSS is flat. The
unrestricted dense-wide prototype's 8.99% peak-heap increase triggered its
rejection and the retained cold-cache fallback.
The direct PPT text-edit follow-up removes one repeated editor open and cuts
task clock 3.60%, allocation calls 3.53% and temporary allocations 6.05%.
Peak heap and uninstrumented RSS remain flat. Its +315.43% minor-fault trigger
has zero major faults and is disclosed rather than presented as a
memory-locality improvement.
The PPT root text-publication handoff removes the immediate second root reopen
after the validated text owner: scoped p50 falls 18.59%, task clock 8.76%,
instructions 6.75%, and allocation calls 6.54%. Peak heap and uninstrumented
RSS remain flat; nondefault limits retain the prior complete reopen.
The ODS cell-locator follow-up reduces the large public sweep's p50 81.74% and
the full-text aggregate's p50 52.65%. Matched sweep-process task clock falls
10.28%, cycles 9.74%, instructions 6.72% and cache misses 7.90%; peak heap,
Heaptrack RSS and uninstrumented RSS remain flat. The 318 added allocation
calls across 105 snapshots (+0.0004%) and the retained 3,216-byte dense index
are disclosed.
The ODS unchanged-media follow-up reduces its 16 MiB media-rich edit/save p50
4.73%, mean 5.73% and p95 7.65%. Peak heap falls 8.78%, task clock 3.54% and
cache references 5.92%; allocation calls rise 0.11%, while branch/cache misses
move +0.42%/+1.29% and are disclosed. The existing medium no-media case remains
slightly better at -0.77% p50.
The ODS row-provenance follow-up reduces the same media-rich edit/save p50
74.16%, mean 74.17% and p95 74.11%. Instructions fall 69.04%, branches 72.96%
and branch misses 94.07%; allocation calls fall 1.26%, while peak heap and
uninstrumented RSS remain flat. Tiny/medium/large open, read and no-op p50
guards remain within the 3% gate or improve.
The ODS shared-worksheet ownership follow-up reduces the remaining media-rich
transaction p50/mean/p95 by 21.32%/21.30%/21.15%. Peak heap falls 22.03%,
uninstrumented RSS 20.57%, cache misses 23.50% and page faults 27.31%; large
ordinary open/no-op/one-edit guards remain within 1.6%.
Lock-wait evidence remains missing.

## Remaining highest-impact work

The largest remaining limitation is the incomplete migration from eager OPC to
source-backed CRUD: selective open, source versions, a finite pinned-aware
single-flight cache, managed charging across physical `InputBytes`, cumulative
declared cold-load `Work`, retained `Objects`, and `Memory`, and a low-level
consuming one-Part publisher now exist and are correctness-tested. Release
contention ABBA evidence is structural/distribution-only with no accepted
speedup; allocation/RSS/hardware/copied/decompressed-byte/CPU resource
instrumentation and broad semantic edit/patch coverage are incomplete. Raw
ZIP preservation is integrated for
eager owned same-topology mutation and this narrow source-backed case; format
facades, topology changes, signatures and real-producer/media matrices remain.
The changed-Part handoff and post-validation local-span copies are removed;
the required selected-Part/compressor buffer remains to be attributed and
reduced independently.

CFB now has a public bounded exact-range read with release ABBA evidence for a
MiniFAT target and exact logical source-work evidence for the atomic overlay
save scan reduction. Current-revision phase attribution assigns 135,680 bytes
to open, 33,962,596 to plan/validation and 50,740,224 to atomic publication in
each of 400 samples. Change 0143 accepts the bounded follow-up without weakening
the two planning brackets or three save-time scans: fingerprint-only requests
are capped at 1 MiB while comparison/publication remain at 64 KiB. Clean
CPU-2 `A1, B1, B2, A2` evidence reduces logical calls from 1,825 to 857 with
unchanged 84,838,500 bytes and exact output. Warm p50 improves 3.3327%/1.3163%
and advisory-cold p50 10.7679%/9.4641%; p95 and mean agree in both directions.
The code-local fingerprint window grows by at most 983,040 bytes and a matched
whole-process RSS boundary shows no candidate increase, but operation-only
allocation/peak memory remains unmeasured. These results remain substrate-only:
no DOC/XLS/PPT semantic owner consumes the exact-range seam, FAT tail behavior
is withheld, and high-latency range, physical-cold and storage-device evidence
remain open. The older Change 0103 `4N -> 3N` scan-removal ABBA still has
disagreeing latency directions; only its exact logical-work reduction was
accepted independently.

Other high-priority gaps are physical cold-filesystem and real range-source
matrices beyond the debug smoke and repeated tmpfs release capture, threshold
tuning and cache-acceptance work beyond the committed explicit scaling curves,
operation-local allocation and total/peak-memory attribution for bounded
forward-only XLSX creation, richer XLSX authoring plus physical/cold-I/O and
producer evidence, allocation/peak-memory evidence for the now latency-measured
RTF creation path, and broad format-semantic CRUD coverage beyond the generated
text/grid slices
(bulk action distinctions, dependency-copy, merge/split, patch timing, repair,
security, malformed and real-producer corpora, plus broader ODF and RTF
coverage). Native DOC/XLS/PPT semantic baselines now have accepted XLS
editor-reuse, DOC batched-publication, PPT root-open reuse, direct PPT
text-edit resolver reuse, and checked PPT root text-publication adoption.
Native DOC also indexes physical
PieceTable intervals and reuses one resolved PAPX initial-style baseline after
distinct profiles attributed 36.89% of large-open self cycles to scalar FKP
range mapping and 6.94% to repeated style resolution/validation; full
validation remains. The highest-priority follow-up from that DOC attribution,
the eager before/after fingerprint work in the retained patch/validation
workflow, is accepted in change 0165: the lazy cache and same-lineage identity
fast path preserve both independent validation layers and exact-byte
authorization while making the deferred diagnostic demand explicit.
Remaining native work requires new attribution inside the retained final
owner/public-reader validation layers beyond that exact opportunity. The
rejected XLS terminal-render
handoff is not a reusable shortcut for those checks. The new opaque-heavy
common OLE2 case rejects direct shared writer payloads (+32.02% p50), an
editor-wide validated-render cache, and inline recapture allocation reuse. The
last improves isolated publication 6.49% p50 but only 2.61% end to end, so it
too was reverted. ODT full-text block
ownership is accepted, and repeated ODS facade cell lookup now has a bounded
lazy index. ODP one-slide lookup now retains only the selected semantic
projection, and its editing snapshot reuses its validated complete slide
projection during transaction staging. Compact ODS and content-only ODP/ODT
edits preserve unchanged ZIP members, and eligible ODS row-range provenance
now survives through raw package emission. Broader ODF source-backed reads, repeated independent
ODT/ODP scans, resource-adding/structural publications, package-parse reuse and
structural-edit profiles remain open.
XLSX changed-sheet validation can now seed a bounded first-read cache. Direct
writer-local action regrouping was immaterial and reverted; distinct bulk
actions, any larger planning/emission coalescing, large-sheet retention,
source-backed editable publication, structural changes and broad preservation
matrices remain independent work.
The rejected direct ODS target-package and parsed ODT final-document adoptions
are not evidence that those broader paths are complete or that validation
should be weakened. Change 0052 shares final bytes only and retains the
independent parse boundary.
iWork work is deliberately deferred while the `iwa-*` crates are changing
independently.
The scenario-by-scenario gap map and next case queue are in
[`CRUD_COVERAGE.md`](CRUD_COVERAGE.md).
The ranked source-level queue and path maps are maintained in
[`HOTSPOTS.md`](HOTSPOTS.md), and architectural gates are in
[`ADR_COMPLIANCE.md`](ADR_COMPLIANCE.md).

## Latest accepted and rejected tranche

Change 0175 adds one opt-in owned-CFB atomic-save filesystem selector and keeps
the historical default matrix fixed. Sealed immutable ownership removes two
complete source fingerprint scans per save: 33,826,816 logical bytes and 34
large reads on the retained 16.9 MiB corpus. Generic sources keep both fences;
owned publication retains full emission source/target hashes and atomic
durability. Both measured candidate directions are lower, but control drift
exceeds 5%, so only the deterministic work reduction is accepted.

Change 0176 records and reverts two non-useful micro-optimizations. ODS
authenticated source-content reuse regresses source-backed p50 in both paired
directions, and XLSX conditional-formatting readback reuse disagrees across
directions. No production behavior or performance claim remains from either
experiment.

Change 0177 accepts the existing source-backed ODS one-cell edit/save path at
the full open/stage/commit/sequential-publication boundary. On a clean CPU-2
release A/B/B/A run with 500 samples per workload and leg, p50 is
75.03%/74.27% lower than eager ownership and p50/mean/p95/p99 stability gates
all pass. The deterministic 21-cell 1% path remains correctness and phase
evidence only because same-implementation mean/tail drift exceeds policy.
Logical source replay counters are explicitly not physical-I/O evidence.

Change 0178 specializes only CFB plans rooted in sealed `Arc<[u8]>` ownership.
After the initial complete fingerprint, candidate reopen, and optional native
owner readback, it omits the redundant final complete scan. Generic positional
sources keep that final mutation fence. Each effective XLS comments/Number plan
removes 16,995,840 logical bytes and 17 one-MiB reads; RK/MulRK removes 202,752
bytes and one read. A clean four-case CPU-2 A/B/B/A observes lower candidate
p50 in every paired direction, but every workload fails at least one stability
threshold, so only deterministic logical work is accepted. Publication hashes,
atomic durability, physical-I/O/resource, cold-cache, producer, and broader CRUD
claims remain unchanged or withheld.
