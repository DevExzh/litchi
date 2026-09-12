# 0518: reuse a retained DOCX source snapshot after current-source proof

`performance_claim: scoped managed DOCX publication improvement`
`claim_authorized: true`

This batch reuses the caller's retained source snapshot during managed DOCX
publication after the current source has been read and proved unchanged. The frozen base is `afb62ab7a70859dad4a4b9c8eea91402d1ef4052`.
The [frozen plan](../results/change-0518/plan.json),
[candidate plan](../results/change-0518/candidate-plan.json), and
[source review](../results/change-0518/final-source-review.md) define the
candidate and its evidence boundaries. OLE2/OOXML remains the active scope;
ODF is deferred and iWork is excluded. Historical 0499/0500 limitations are
not closed by this batch.

## Candidate contract

`publish_document_commit_to_stream` offers `Patch::source()` only for a
managed, source-authorized changed commit whose retained `before` snapshot
has the current source identity. The OPC owner first performs the ordinary
current-Part read and retains its read-ahead, freshness, security,
classification, limits, cancellation, and source/context fences. A hint is
eligible only when its lineage, version, equivalent Part URI, content type,
`ReadLimits`, and original-payload ownership agree. The original bytes then
must match exactly: pointer-and-length identity is sufficient, while distinct
allocations use an interruptible bounded byte comparison. That comparison
charges `Resource::Work` conservatively, including the full original length
for a pointer-identical proof.

On an exact hit, OPC returns a clone of the retained original proof and DOCX
returns a clone of `before`; the candidate avoids rebuilding the same source
snapshot. A foreign or derived hint, identity or limit mismatch, or byte
mismatch uses full validation and reconstruction from the current `PartData`,
without a second read. Security, freshness, cancellation, and malformed-XML
failures retain their typed refusals; DOCX retains its existing signed-source
fallback and separate no-op path. `Patch::apply` still decides whether the
commit source matches. An eligible hint comparison consumes its conservative
Work charge even on a byte mismatch, so Work exhaustion can precede a
subsequent malformed-XML error. The
candidate adds no global cache, mutable byte access, unsafe code, archive
handle, or public snapshot API. OPC owns the source proof; DOCX owns the
source identity and snapshot match. The required hit, fallback, security,
freshness, cancellation, limit, derived-hint, exact-output, readback, and
release tests are catalogued in the
[focused-test review](../results/change-0518/focused-test-review.md).

## Native end-to-end result

The unchanged managed paragraph harness covers 24 shape/replacement,
provider, and route cases. Baseline campaigns are `r1` and `r2`; candidate
campaigns are `after-r1` and `after-r2`, with the second campaign reversing
case order. Each case has 30 measured samples, three warmups, and two
internal repetitions in fresh CPU-2 child processes. The lifecycle `elapsed`
clock includes the existing open/edit/commit/publication path; `publish_ns`
includes destruction of the returned Snapshot. The table gives
baseline → candidate nearest-rank p50 values in nanoseconds; RSS is the
whole-child maximum in KiB. The [native comparison](../results/change-0518/candidate-comparison.md)
and [machine-readable comparison](../results/change-0518/candidate-comparison.json)
retain all means, p95/p99 values, bootstrap intervals, and guard counters.

The compact flag column lists matched-pair p95/p99 and RSS threshold flags.
The reports retain p50/mean flags on the short open/commit/drop phases: `r1`
has 31 flags across 14 records and `r2` has 38 across 18 records. No matched
E2E or publication p50 regression is flagged, and no RSS value exceeds the
5% threshold; the largest RSS increase is 4.63%.

| Workload | r1 p50: E2E / publish / RSS | r2 p50: E2E / publish / RSS | p95/p99/RSS flags (r1; r2) |
| --- | --- | --- | --- |
| p128-k1-file-batch | 578333→433162 (-25.10%) / 285642→135021 (-52.73%) / 6608→6516 KiB | 578192→426162 (-26.29%) / 284511→134870 (-52.60%) / 6564→6520 KiB | open.p99 +18.5%; open.p99 +35.1% |
| p128-k1-file-repeated | 574922→425412 (-26.01%) / 286381→135100 (-52.83%) / 6524→6524 KiB | 580093→426421 (-26.49%) / 287391→135230 (-52.95%) / 6512→6568 KiB | —; commit.p99 +979.5% |
| p128-k1-owned-batch | 572473→375171 (-34.46%) / 286411→94141 (-67.13%) / 7204→7328 KiB | 523692→418932 (-20.00%) / 241651→135931 (-43.75%) / 7152→7212 KiB | open.p99 +37.5%; open.p99 +16.2%, drop.p99 +865.1% |
| p128-k1-owned-repeated | 570932→373442 (-34.59%) / 285291→92951 (-67.42%) / 7080→7328 KiB | 530532→424232 (-20.04%) / 244571→134770 (-44.90%) / 7188→7176 KiB | —; open.p99 +64.9%, commit.p99 +32.4% |
| p128-k32-file-batch | 791273→643143 (-18.72%) / 332212→181571 (-45.34%) / 6604→6600 KiB | 788963→633853 (-19.66%) / 332561→180721 (-45.66%) / 6604→6704 KiB | —; open.p99 +19.1% |
| p128-k32-file-repeated | 5850597→5727155 (-2.11%) / 336451→184391 (-45.20%) / 6648→6680 KiB | 5829054→5746046 (-1.42%) / 334002→183781 (-44.98%) / 6708→6616 KiB | open.p99 +46.8%; drop.p99 +167.3% |
| p128-k32-owned-batch | 719203→566942 (-21.17%) / 288151→138421 (-51.96%) / 7104→7164 KiB | 718763→568943 (-20.84%) / 288711→138851 (-51.91%) / 7216→7184 KiB | —; open.p99 +21.7%, commit.p99 +286.3% |
| p128-k32-owned-repeated | 5414485→5263474 (-2.79%) / 291091→141431 (-51.41%) / 7164→7192 KiB | 5412172→5259004 (-2.83%) / 294001→141851 (-51.75%) / 7152→7192 KiB | open.p95 +54.1%; drop.p99 +184.6% |
| p128-k8-file-batch | 617523→462472 (-25.11%) / 288861→135191 (-53.20%) / 6820→6624 KiB | 616722→465142 (-24.58%) / 286641→135630 (-52.68%) / 6868→6624 KiB | drop.p99 +491.6%; drop.p99 +531.0% |
| p128-k8-file-repeated | 1518926→1373926 (-9.55%) / 329861→180111 (-45.40%) / 6584→6624 KiB | 1531466→1376416 (-10.12%) / 330692→179241 (-45.80%) / 6628→6684 KiB | —; drop.p99 +19.3% |
| p128-k8-owned-batch | 560362→412412 (-26.40%) / 244191→94290 (-61.39%) / 7332→7296 KiB | 559302→412272 (-26.29%) / 243831→93931 (-61.48%) / 7300→7344 KiB | drop.p99 +537.3%; — |
| p128-k8-owned-repeated | 1455066→1296166 (-10.92%) / 288181→137300 (-52.36%) / 7108→7348 KiB | 1452606→1295246 (-10.83%) / 288951→138470 (-52.08%) / 7092→7096 KiB | open.p95 +99.6%, open.p99 +129.5%; open.p95 +92.0%, open.p99 +92.4% |
| p512-k1-file-batch | 2026208→1434816 (-29.19%) / 966174→370461 (-61.66%) / 6448→6448 KiB | 2017258→1429467 (-29.14%) / 962364→371771 (-61.37%) / 6448→6448 KiB | —; — |
| p512-k1-file-repeated | 2010368→1430487 (-28.84%) / 956184→373442 (-60.94%) / 6608→6452 KiB | 2032368→1429507 (-29.66%) / 965224→370752 (-61.59%) / 6436→6484 KiB | —; — |
| p512-k1-owned-batch | 2010778→1382406 (-31.25%) / 916204→326841 (-64.33%) / 6996→7020 KiB | 1969998→1380876 (-29.90%) / 919814→327831 (-64.36%) / 7128→6960 KiB | open.p95 +21.9%; — |
| p512-k1-owned-repeated | 1965628→1375846 (-30.00%) / 915424→328331 (-64.13%) / 7144→6996 KiB | 1957978→1369076 (-30.08%) / 913544→326522 (-64.26%) / 6992→6960 KiB | —; — |
| p512-k32-file-batch | 2168071→1583528 (-26.96%) / 952725→363452 (-61.85%) / 6992→6952 KiB | 2153009→1563317 (-27.39%) / 949854→351842 (-62.96%) / 6996→6992 KiB | commit.p95 +12.5%, commit.p99 +361.1%; — |
| p512-k32-file-repeated | 16906138→16304593 (-3.56%) / 940705→365352 (-61.16%) / 7024→6936 KiB | 16962342→16248513 (-4.21%) / 952274→355182 (-62.70%) / 7008→6952 KiB | —; open.p99 +38.4% |
| p512-k32-owned-batch | 2109971→1532517 (-27.37%) / 922775→334492 (-63.75%) / 7344→7444 KiB | 2113859→1529357 (-27.65%) / 925074→330332 (-64.29%) / 7380→7384 KiB | open.p99 +43.8%, drop.p99 +45.8%; — |
| p512-k32-owned-repeated | 16417726→15818511 (-3.65%) / 921165→334702 (-63.67%) / 7396→7428 KiB | 16738450→15813630 (-5.53%) / 941864→334572 (-64.48%) / 7464→7320 KiB | open.p99 +26.1%; commit.p99 +289.8% |
| p512-k8-file-batch | 2000830→1440027 (-28.03%) / 920535→348272 (-62.17%) / 6736→7048 KiB | 2037349→1397036 (-31.43%) / 960444→325002 (-66.16%) / 6716→6952 KiB | commit.p95 +17.4%; drop.p99 +222.6% |
| p512-k8-file-repeated | 5206317→4511910 (-13.34%) / 979395→372971 (-61.92%) / 6792→6684 KiB | 5084392→4521470 (-11.07%) / 957374→373372 (-61.00%) / 6596→6704 KiB | —; — |
| p512-k8-owned-batch | 2065051→1415256 (-31.47%) / 936125→328031 (-64.96%) / 7292→7520 KiB | 2007879→1402686 (-30.14%) / 919344→326512 (-64.48%) / 7484→7276 KiB | —; — |
| p512-k8-owned-repeated | 5036417→4441070 (-11.82%) / 925905→330211 (-64.34%) / 7284→7224 KiB | 4996131→4423960 (-11.45%) / 911093→329871 (-63.79%) / 7588→7240 KiB | —; open.p99 +20.4% |

Across the two campaigns, lifecycle p50 reductions range from 2.11% to
34.59% in `r1` and 1.42% to 31.43% in `r2`; publication p50 reductions range
from 45.20% to 67.42% and 43.75% to 66.16%, respectively. These are
synthetic managed-DOCX observations. They do not establish a host-general
speedup or erase the short-phase tail flags.

## Scoped instruction and allocation evidence

The six owned-source Callgrind arms have two fresh baseline and two fresh
candidate profiles. Collection is toggled only around
`Package::publish_document_commit_to_stream`, so it excludes the caller's
returned-Snapshot drop. The [profile comparison](../results/change-0518/profile-comparison.md)
records 12 paired rows and passing raw/annotation checks: one positive
publication call, matching incoming and direct-edge totals, the unchanged
`SourceBackedPackage::write_topology_to_stream` owner, exactly one positive
XML-validator call, and no positive candidate cost in the fresh DOCX scan or
fresh Snapshot builder.

| Owned profile arm | r1 publication inclusive Ir (baseline→candidate) | r2 publication inclusive Ir (baseline→candidate) |
| --- | ---: | ---: |
| p128-k1-batch | 5,467,227→2,520,473 (-53.90%) | 5,470,253→2,521,228 (-53.91%) |
| p128-k1-repeated | 5,468,556→2,519,698 (-53.92%) | 5,468,080→2,520,826 (-53.90%) |
| p512-k1-batch | 17,747,024→6,097,734 (-65.64%) | 17,746,149→6,098,491 (-65.63%) |
| p512-k1-repeated | 17,743,288→6,097,136 (-65.64%) | 17,743,774→6,096,605 (-65.64%) |
| p512-k32-batch | 17,752,919→6,098,980 (-65.65%) | 17,751,271→6,099,762 (-65.64%) |
| p512-k32-repeated | 17,755,856→6,099,746 (-65.65%) | 17,752,841→6,101,321 (-65.63%) |

The separate allocator probe has 96 validated captures: 24 cases in each of
two baseline and two candidate child repeats, with one measured sample, no
warmups, and one internal repeat. All samples are measured with zero failed
allocations and verified cleanup. Values below are per case; absolute region
peaks vary with the retained entry state, so the incremental peak is the
comparable publication-region demand.

| Shape | Allocation calls (baseline→candidate) | Allocated bytes (baseline→candidate) | Incremental region peak bytes (baseline→candidate) | Absolute region peak range, baseline→candidate |
| --- | ---: | ---: | ---: | ---: |
| p128 | 3,252→109 | 300,922→75,001 | 74,547→72,535 | 3,284,754–3,829,495→3,282,747–3,827,488 |
| p512 | 12,470→109 | 947,578→75,001 | 77,619→72,535 | 3,400,338–3,966,583→3,395,259–3,961,504 |

The allocator region ends when the publication method returns, before the
caller drops the returned Snapshot. Its absolute callback-ordered live peak
is therefore separate from whole-child RSS; the native captures have no RSS
increase over the 5% threshold (maximum +4.63%). The raw lanes and probe
custody are retained in [candidate repeat 1](../results/change-0518/alloc-candidate-r1/),
[candidate repeat 2](../results/change-0518/alloc-candidate-r2/), and the
corresponding baseline directories.

The allocator probe is a separate Cargo workspace with its own frozen lock
file and default release profile (LTO off and panic-unwind), while the normal
workspace uses LTO and panic-abort. The matched probe builds make the
source-change comparison valid, but absolute allocator counts and peaks are
not measurements of the normal executable's generated code or allocator
behavior.

## Work, identity, and disposition

The [counter comparison](../results/change-0518/counter-comparison.json)
checks 3,168 paired rows. Source read calls, requested and returned source
bytes, output bytes and hashes, release state, live memory, live objects, and
the after-input/output gauges are unchanged. Output identities match 24/24 in
each same-API campaign pair, and no release/output/input/source-read guard
counter changes are present. The allowed Work reduction is
6,604,771 bytes of charged Work for every p128 case and 103,485,667 bytes for
every p512 case. Work is a resource-budget charge, not an instruction or
allocator-call count; the exact comparison and fallback Work charge remains
part of the candidate contract.

The baseline and candidate native binaries and source manifests are bound in
the reports (`637df10f…` / `409fa85f…` and `75a26797…` / `7aa4c9c0…`,
respectively), with the candidate patch replay bound to the frozen base. The
implementation check and the recorded focused runs pass: DOCX managed tests
40, OPC public hint tests 11, and corrected private hint-unit tests 4. These
receipts establish the reviewed path and guards. The final all-features
OPC, DOCX, XLSX, PPTX, and XLSB suite passed 5,003 tests, and the independently generated
ZIP64 preservation case passed separately: 5,004 executed tests in total.
Formatting, workspace checking, warning-denied lint and documentation, crate
boundaries, and strict claims checks also passed.

The implementation is retained for the scoped end-to-end and publication
improvements above. The evidence audit and cleanup receipts accompany the
raw results. Two separate debug fallback profiles confirm that foreign and
derived hints still reach full XML validation; these are safety-path evidence
and are excluded from performance comparisons. The
whole-child hardware counters are retained as diagnostic evidence and are not
publication-local attribution. This change makes no cold-cache, native-Office
producer, broad-provider, parallel-scaling, or ODF claim. The full
OLE2/OOXML optimization goal remains open, and historical 0499/0500 flags
remain explicit limitations.
