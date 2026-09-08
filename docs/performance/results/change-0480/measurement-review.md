# Independent measurement review for change 0480

This review recomputes the measurements directly from the 48 retained raw
capture reports and their GNU `time -v` resource files.  There are two arms
(control and candidate), two reverse-order repeats, two binaries (normal and
counting allocator), three corpus sizes, and total/phase modes.  Each report
has 30 measured samples after three warmups, for 1,440 measured samples.  The
review was derived without relying on a generated summary file; a later
`summary.json`, if produced, must be compared against these raw-derived
values before the bundle is sealed.

Means are arithmetic means.  Percentiles are nearest-rank values over the 30
raw samples (`p50`, `p95`, and `p99`).  Elapsed values below are milliseconds.
The pair direction is candidate versus control, including repeat 2 even
though the process order is reversed.  A positive pair percentage is a
candidate cost; a negative value means the candidate's observed value was
lower.  Repeat percentages are repeat 2 versus repeat 1 within the same arm.

## Measurement verdict

Keep the two-line candidate for this scoped ownership change.  The direct
allocator evidence is exact: it removes the publication duplicate's request
from the measured operation, equal to the candidate XML payload plus 40 bytes
at all three sizes, and it removes two allocation callbacks and two matching
deallocation callbacks.  The operation ends with zero net retained bytes, all
48 processes exit successfully, and the output, source reads, sink accounting,
and inverse/corpus oracles are unchanged.

Normal total-lifecycle latency has no positive candidate-versus-control pair
change above 5%; repeat 1 mean changes are +0.085%, +0.091%, and -0.237% at
64, 8,192, and 131,072 paragraphs.  Those observations do not establish a
speedup.  The apparent large-case whole-process RSS reduction is useful
supporting evidence, but GNU `time` covers setup, corpus/oracle storage,
warmups, measured lifecycles, report serialization, and teardown.  The direct
operation allocator counters are the stronger heap result.

This remains a materialized one-paragraph transaction.  It does not establish
an explicit bounded-window append, repeated append scaling, or the broader
non-iWork goal in `docs/GOAL.md`.

## Corpus, output, and I/O identity

The candidate copies exactly one 49-byte paragraph fragment.  The candidate
main XML is therefore source XML plus 49 bytes.  The compressed ZIP size may
change by a different amount because the changed XML member is deflated.

| paragraphs | source XML | candidate XML | source archive | candidate archive | candidate output SHA-256 |
| ---: | ---: | ---: | ---: | ---: | :--- |
| 64 | 3,287 | 3,336 (+49) | 2,108 | 2,112 (+4) | `297f40d52e7494cd3f4601ef75a7f93e53bd5999dfa90cf4ecf5501aa7a61ca8` |
| 8,192 | 401,559 | 401,608 (+49) | 23,357 | 23,362 (+5) | `c29151eb2672bd2910fdbdf55f922dfdd01aa39d1cdfa2ffb6136d6579c4fdc9` |
| 131,072 | 6,422,679 | 6,422,728 (+49) | 343,945 | 343,949 (+4) | `2152458e09392b4aa836b8f22c4d27fd11d58d3af4dbabc261f202ce09714e4e` |

For every size, the corpus record, source/candidate XML oracle, semantic
paragraph/text/order oracle, physical member order, untouched member bytes,
and output archive identity were the same in all control/candidate arms, both
repeats, and both total/phase executions.  The output sink hash was the
candidate archive hash above in every report.

The lifecycle source and sink observations were also identical in all arms,
repeats, binaries, and modes.  The source histogram columns are
`0 / 1–512 / 513–4096 / 4097–16384 / 16385–65536 / >65536` bytes.

| paragraphs | source `ReadAt` calls | requested = returned bytes | source request histogram | sink accepted bytes | sink writes | largest write | sink histogram |
| ---: | ---: | ---: | :--- | ---: | ---: | ---: | :--- |
| 64 | 42 | 7,969 | 3 / 36 / 3 / 0 / 0 / 0 | 2,112 | 19 | 877 | 0 / 18 / 1 / 0 / 0 / 0 |
| 8,192 | 42 | 71,716 | 3 / 35 / 1 / 0 / 3 / 0 | 23,362 | 20 | 16,384 | 0 / 17 / 1 / 2 / 0 / 0 |
| 131,072 | 62 | 1,033,480 | 3 / 35 / 1 / 3 / 20 / 0 | 343,949 | 39 | 16,384 | 0 / 17 / 1 / 21 / 0 / 0 |

The phase `ReadAt` counters sum exactly to the lifecycle counters for every
sample.  The total and phase reports are separate executions, so their timing
and process RSS are compared as separate runs; phase measurements are not
added to a total-operation peak.

## Exact allocation removal and zero exit

For allocator reports, the operation incremental peak is

`region_peak_live_bytes - live_bytes_before`.

The fields in this table were constant for all 30 samples in each repeat.  The
same values were observed in repeats 1 and 2.  The allocation reduction is
exactly `candidate_main_xml_bytes + 40` in every row, and the incremental peak
and request-byte reductions agree exactly.

| paragraphs | incremental peak control | incremental peak candidate | delta | allocated bytes control | allocated bytes candidate | delta | allocation/deallocation callbacks control → candidate | realloc callbacks |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | :--- | ---: |
| 64 | 509,974 | 506,598 | -3,376 = -(3,336 + 40) | 1,036,969 | 1,033,593 | -3,376 | 1,888 / 1,821 → 1,886 / 1,819 | 67 |
| 8,192 | 3,071,310 | 2,669,662 | -401,648 = -(401,608 + 40) | 1,615,057,345 | 1,614,655,697 | -401,648 | 213,346 / 196,893 → 213,344 / 196,891 | 16,453 |
| 131,072 | 41,793,870 | 35,371,102 | -6,422,768 = -(6,422,728 + 40) | 549,272,923,585 | 549,266,500,817 | -6,422,768 | 3,653,986 / 3,146,013 → 3,653,984 / 3,146,011 | 507,973 |

The phase attribution places the same reduction entirely in `publish`; the
`open`, `snapshot`, `stage`, `commit`, and `drop` allocation byte/callback
counters are unchanged apart from the two-byte live-owner baseline shift.
For the publish phase, the control-to-candidate allocated bytes are
`569,770 → 566,394`, `807,803,986 → 807,402,338`, and
`274,640,123,986 → 274,633,701,218` for the three sizes.  The corresponding
allocation/deallocation callback pairs are `890/898 → 888/896`,
`106,619/98,434 → 106,617/98,432`, and
`1,826,939/1,572,994 → 1,826,937/1,572,992`; realloc callback counts do not
change.

All 48 capture metadata records have exit code zero.  Across allocator total
samples, `live_bytes_after - live_bytes_before` is exactly zero.  Across every
allocator phase sample, the accounting equation

`live_before + allocated_bytes - deallocated_bytes = live_after`

holds; adjacent phase boundaries are continuous, and `drop.live_after` equals
`open.live_before`.  Normal reports intentionally expose unavailable
allocation metrics as unavailable rather than zero.  `allocated_bytes`
includes the full requested `new_size` on realloc callbacks; there is no
physical-copy-byte or separate reallocated-byte counter.

## Normal total latency and whole-process RSS

RSS in these tables is the one GNU `/usr/bin/time -v` maximum resident set
size per process, in KiB.  It is not a 30-sample operation-local statistic.

| paragraphs | repeat | control mean / p50 / p95 / p99 | candidate mean / p50 / p95 / p99 | candidate mean delta | GNU RSS control → candidate | RSS delta |
| ---: | ---: | :--- | :--- | ---: | ---: | ---: |
| 64 | 1 | 0.168799 / 0.167000 / 0.177131 / 0.182921 | 0.168942 / 0.167671 / 0.175420 / 0.180341 | +0.085% | 4,744 → 4,824 | +1.69% |
| 64 | 2 | 0.168659 / 0.166301 / 0.179751 / 0.188360 | 0.170074 / 0.168110 / 0.183391 / 0.191751 | +0.839% | 4,744 → 4,828 | +1.77% |
| 8,192 | 1 | 16.555023 / 16.543697 / 16.632267 / 16.632917 | 16.570107 / 16.564220 / 16.662310 / 16.698900 | +0.091% | 11,184 → 10,728 | -4.08% |
| 8,192 | 2 | 16.873467 / 16.857759 / 17.187361 / 17.237431 | 16.903092 / 16.889139 / 17.097400 / 17.153589 | +0.176% | 10,924 → 10,732 | -1.76% |
| 131,072 | 1 | 270.012341 / 269.310218 / 274.485198 / 274.662879 | 269.373280 / 269.179225 / 271.426324 / 273.144572 | -0.237% | 113,388 → 107,200 | -5.46% |
| 131,072 | 2 | 272.926225 / 272.548574 / 277.595654 / 278.069437 | 265.883534 / 265.787657 / 271.286070 / 271.389856 | -2.580% | 113,648 → 107,200 | -5.67% |

There is no positive candidate-versus-control total-latency pair flag above
5%.  Repeat 1 normal total means remain within 0.3% at every size; the
negative values are retained observations, not speedup claims.

The phase runs are separate processes.  The following is the sum of six phase
elapsed values within each sample, included for attribution context only.

| paragraphs | repeat | control phase-sum mean / p50 / p95 / p99 | candidate phase-sum mean / p50 / p95 / p99 | candidate mean delta | GNU RSS control → candidate | RSS delta |
| ---: | ---: | :--- | :--- | ---: | ---: | ---: |
| 64 | 1 | 0.168296 / 0.166461 / 0.177801 / 0.178550 | 0.171243 / 0.168990 / 0.182011 / 0.186411 | +1.751% | 4,792 → 4,800 | +0.17% |
| 64 | 2 | 0.173089 / 0.170421 / 0.183221 / 0.204120 | 0.170921 / 0.169412 / 0.179831 / 0.190421 | -1.252% | 4,744 → 4,584 | -3.37% |
| 8,192 | 1 | 16.603923 / 16.540038 / 16.975279 / 16.975840 | 16.650560 / 16.598940 / 16.972251 / 16.992112 | +0.281% | 11,200 → 10,732 | -4.18% |
| 8,192 | 2 | 16.355452 / 16.305747 / 16.536269 / 16.554189 | 16.404770 / 16.387986 / 16.525127 / 16.544968 | +0.302% | 10,928 → 10,732 | -1.79% |
| 131,072 | 1 | 272.435212 / 271.959547 / 277.133599 / 277.442419 | 271.885150 / 271.717026 / 273.442374 / 273.706924 | -0.202% | 113,288 → 107,456 | -5.15% |
| 131,072 | 2 | 270.684164 / 269.221166 / 278.707774 / 279.791898 | 264.807100 / 263.243847 / 268.590138 / 271.850832 | -2.171% | 113,388 → 107,052 | -5.59% |

## Exhaustive pair flags above 5%

The table includes every latency/RSS statistic whose absolute pair change is
strictly greater than 5%.  A `+` is a candidate cost and a `-` is a lower
candidate observation.  `RSS` is one scalar process value; its four summary
positions are therefore one flag, not four independent samples.  The phase
flags are ratio-sensitive microsecond-scale quantiles, especially `commit`
and `drop`; they are retained for review rather than promoted to a speed claim.

| repeat | binary | paragraphs | mode | flags (candidate versus control) |
| ---: | :--- | ---: | :--- | :--- |
| 1 | normal | 64 | phases | `snapshot.p95` +9.92%; `snapshot.p99` +9.37%; `stage.p95` +7.03%; `stage.p99` +16.12%; `commit.p95` -14.29%; `commit.p99` -14.29%; `publish.p99` +7.79%; `drop.mean` -7.31% |
| 1 | normal | 8,192 | phases | `open.p99` -16.69%; `commit.mean` -15.84%; `commit.p50` -28.57%; `commit.p95` -33.33%; `commit.p99` +177.78%; `drop.mean` -26.01%; `drop.p50` -26.03%; `drop.p95` -26.04%; `drop.p99` -18.14% |
| 1 | normal | 131,072 | total | `RSS` -5.46% |
| 1 | normal | 131,072 | phases | `RSS` -5.15%; `open.mean` -15.98%; `open.p50` -16.46%; `open.p95` -14.83%; `open.p99` -11.25%; `commit.mean` -16.23%; `commit.p50` +5.26%; `commit.p95` -9.09%; `commit.p99` -10.42%; `drop.mean` -99.89%; `drop.p50` -99.88%; `drop.p95` -99.89%; `drop.p99` -99.88% |
| 1 | allocator | 64 | phases | `snapshot.p95` -14.86%; `stage.p95` -17.78%; `stage.p99` +39.98%; `commit.mean` +55.70%; `commit.p95` +200.00%; `commit.p99` +200.00%; `publish.p99` -14.48% |
| 1 | allocator | 8,192 | phases | `commit.mean` -18.37%; `commit.p50` -16.67%; `commit.p95` -12.50%; `commit.p99` -22.22%; `drop.mean` -27.86%; `drop.p50` -26.74%; `drop.p95` -41.49%; `drop.p99` -40.85% |
| 1 | allocator | 131,072 | total | `RSS` -5.39% |
| 1 | allocator | 131,072 | phases | `RSS` -5.43%; `open.mean` -16.74%; `open.p50` -16.33%; `open.p95` -23.25%; `open.p99` -21.05%; `commit.mean` -27.68%; `commit.p50` -6.25%; `commit.p95` -17.14%; `commit.p99` -13.16%; `drop.mean` -99.85%; `drop.p50` -99.84%; `drop.p95` -99.86%; `drop.p99` -99.85% |
| 2 | normal | 64 | phases | `open.p99` -19.68%; `snapshot.p99` +15.33%; `stage.p95` -15.21%; `stage.p99` -9.49%; `commit.mean` +23.08%; `commit.p50` +20.00%; `commit.p95` +16.67%; `commit.p99` +112.50%; `publish.p99` -9.49% |
| 2 | normal | 8,192 | phases | `open.p95` -6.08%; `open.p99` -21.26%; `commit.mean` +52.17%; `commit.p50` +100.00%; `commit.p95` -29.63%; `commit.p99` -38.89%; `drop.mean` -27.20%; `drop.p50` -27.08%; `drop.p95` -31.47%; `drop.p99` -23.92% |
| 2 | normal | 131,072 | total | `RSS` -5.67% |
| 2 | normal | 131,072 | phases | `RSS` -5.59%; `open.mean` -24.26%; `open.p50` -24.67%; `open.p95` -23.03%; `open.p99` -17.92%; `commit.mean` +23.62%; `commit.p50` +7.69%; `commit.p95` +51.85%; `commit.p99` +55.17%; `drop.mean` -99.83%; `drop.p50` -99.82%; `drop.p95` -99.83%; `drop.p99` -99.86% |
| 2 | allocator | 64 | total | `latency.p99` -7.01% |
| 2 | allocator | 64 | phases | `open.p99` -53.36%; `snapshot.p95` +11.50%; `snapshot.p99` +13.48%; `stage.p95` -9.47%; `stage.p99` -24.38%; `commit.p95` -28.57%; `commit.p99` -12.50%; `publish.p95` +7.60%; `publish.p99` +12.17% |
| 2 | allocator | 8,192 | phases | `open.mean` -29.32%; `open.p50` -9.53%; `open.p95` -61.92%; `open.p99` -58.74%; `commit.mean` -43.75%; `commit.p50` -16.67%; `commit.p95` -72.00%; `commit.p99` -46.15%; `drop.mean` -38.67%; `drop.p50` -27.59%; `drop.p95` -67.30%; `drop.p99` -62.21% |
| 2 | allocator | 131,072 | total | `RSS` -5.41% |
| 2 | allocator | 131,072 | phases | `RSS` -5.61%; `open.mean` -14.62%; `open.p50` -14.12%; `open.p95` -20.55%; `open.p99` -20.90%; `commit.mean` +28.94%; `commit.p50` +31.25%; `commit.p95` +16.00%; `drop.mean` -99.85%; `drop.p50` -99.85%; `drop.p95` -99.83%; `drop.p99` -99.83% |

The positive candidate-cost pair flags are therefore limited to phase
statistics (including a few means): normal r1 at 64 (`snapshot`, `stage`, `publish`), normal r1 at
8,192 (`commit.p99`), normal r1 at 131,072 (`commit.p50`), allocator r1 at 64
(`stage.p99`, `commit` mean/p95/p99), normal r2 at all three sizes' listed
`snapshot`/`commit` fields, allocator r2 at 64 (`snapshot`/`publish`), and
allocator r2 at 131,072 (`commit`).  There is no positive total-lifecycle
latency or RSS cost above 5%.

## Exhaustive repeat-drift flags

These are all fields with absolute repeat-2 versus repeat-1 drift strictly
greater than 5%.  They are process-order and measurement-stability signals,
not candidate regressions.

| arm | binary | paragraphs | mode | absolute drift flags |
| :--- | :--- | ---: | :--- | :--- |
| control | normal | 64 | phases | `open.p95` +7.72%; `open.p99` +35.03%; `snapshot.p99` -8.85%; `stage.p95` +23.16%; `stage.p99` +17.72%; `commit.p95` -14.29%; `commit.p99` +14.29%; `publish.p99` +23.18%; `drop.p95` -6.25% |
| control | normal | 8,192 | phases | `open.p99` +8.93%; `commit.mean` +25.25%; `commit.p95` +200.00%; `commit.p99` +300.00%; `drop.p95` +5.90% |
| control | normal | 131,072 | phases | `commit.mean` -34.35%; `commit.p50` -31.58%; `commit.p95` -38.64%; `commit.p99` -39.58%; `drop.p99` +20.31% |
| control | allocator | 64 | phases | `RSS` +5.29%; `open.p99` +128.82%; `snapshot.p95` -14.38%; `stage.p95` -8.08%; `stage.p99` +25.13%; `commit.p95` +40.00%; `commit.p99` +60.00%; `publish.p99` -11.97%; `drop.p50` +5.26%; `drop.p95` +7.50%; `drop.p99` +7.50% |
| control | allocator | 8,192 | phases | `open.mean` +38.72%; `open.p50` +7.88%; `open.p95` +153.40%; `open.p99` +159.34%; `commit.mean` +55.10%; `commit.p95` +212.50%; `commit.p99` +188.89%; `drop.mean` +19.45%; `drop.p95` +77.77%; `drop.p99` +101.91% |
| control | allocator | 131,072 | phases | `commit.p95` +42.86%; `commit.p99` +76.32% |
| candidate | normal | 64 | total | `p99` +6.33% |
| candidate | normal | 64 | phases | `open.p99` +7.33%; `snapshot.p95` -6.99%; `stage.p99` -8.23%; `commit.mean` +28.86%; `commit.p50` +20.00%; `commit.p95` +16.67%; `commit.p99` +183.33%; `drop.p50` +8.33%; `drop.p95` -9.38%; `drop.p99` -6.06% |
| candidate | normal | 8,192 | phases | `commit.mean` +126.47%; `commit.p50` +180.00%; `commit.p95` +216.67%; `commit.p99` -12.00%; `drop.p99` -6.16% |
| candidate | normal | 131,072 | phases | `open.mean` -11.98%; `open.p50` -12.47%; `open.p95` -11.70%; `open.p99` -11.93%; `commit.p50` -30.00%; `drop.mean` +48.02%; `drop.p50` +51.06%; `drop.p95` +53.39%; `drop.p99` +37.31% |
| candidate | allocator | 64 | phases | `open.p99` +6.54%; `snapshot.p95` +12.13%; `snapshot.p99` +11.18%; `stage.p99` -32.40%; `commit.mean` -35.34%; `commit.p95` -66.67%; `commit.p99` -53.33%; `publish.p95` +11.82%; `publish.p99` +15.47% |
| candidate | allocator | 8,192 | phases | `open.p95` -6.67%; `commit.mean` +6.87%; `commit.p99` +100.00%; `drop.p99` +29.01% |
| candidate | allocator | 131,072 | phases | `commit.mean` +83.70%; `commit.p50` +40.00%; `commit.p95` +100.00%; `commit.p99` +93.94%; `drop.p95` +16.67%; `drop.p99` +15.19% |

The repeat flags explain why phase-level ratios are not suitable for a speed
claim.  In particular, the large control allocator phase drift and the
candidate large-case `drop` drift occur in short attribution regions, while
the direct allocation counters and output/I/O identities remain exact.

## Scope and limitations

The normal binary does not install the counting allocator, so normal
allocation fields are unavailable rather than zero.  The allocator's
`allocated_bytes` counter includes realloc `new_size`; it does not report
physical bytes copied by reallocations.  The per-sample process observer's
`rss_bytes` is a saturating RSS delta and `peak_rss_bytes` is an absolute VmHWM
endpoint; the RSS tables here use the broader GNU `time` maximum to make the
scope explicit.

The total lifecycle includes package construction, snapshot, staging, commit,
publication, sink digest finalization, and owner destruction.  Corpus and
independent oracle storage are prepared outside that measured lifecycle while
remaining inside the whole-process RSS envelope.  Phase peaks are separate,
non-nested attribution regions and are never summed into a total-operation
peak.

The measured edit remains one plain-paragraph copy at the tail.  It still owns
the complete materialized XML and paragraph index and does not cover repeated
reopen/append cycles, an explicit bounded-window source, section properties,
native/cold/range sources, or broader CRUD/scaling.  The two-line ownership
candidate can therefore be retained for its measured heap/request reduction
without changing the broader goal's bounded-memory status.
