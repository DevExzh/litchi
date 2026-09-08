# Independent measurement review for change 0481

This review recomputes the measurements directly from the 48 retained raw
capture reports and their GNU `time -v` resource files.  The matrix has two
arms (control and candidate), two reverse-order repeats, normal and counting
allocator binaries, three corpus sizes, and total/phase modes.  Each report
has 30 measured samples after three warmups, for 1,440 measured samples.  No
generated summary file was used to derive these values.  After finalization,
`summary.json` matched the raw rows, all 261 pair flags, and all 89 repeat
flags below.

Means are arithmetic means.  Percentiles are nearest-rank values over the 30
raw samples (`p50`, `p95`, and `p99`).  Elapsed values below are milliseconds.
The pair direction is candidate versus control, including repeat 2 even
though the process order is reversed.  A positive pair percentage is a
candidate cost; a negative value means the candidate's observed value was
lower.  Repeat percentages are repeat 2 versus repeat 1 within the same arm.

## Measurement verdict

Keep the seven-line candidate for this scoped scanner ownership change.  The
raw allocator evidence shows the predicted removal exactly: 24*N+28
allocation callbacks and 24*N+108 requested bytes at each corpus size.  The
same reductions appear in matching allocation and deallocation callbacks;
reallocation counts do not change.  The incremental operation peak is
unchanged after subtracting each arm's full entry live bytes; both the
entry-live and absolute region-peak values differ by two bytes, so
the direct result is lower allocation traffic and lower latency rather than
a lower operation high-water mark.

Normal total-lifecycle means fall by 7.844% to 12.000% across the two repeats
and three sizes.  Allocator total means fall by 23.013% to 28.352% in repeat
1 and by 23.383% to 28.101% in repeat 2.  These are observations from this
single pinned host and do not authorize a general speedup claim.  There is no
positive candidate-versus-control total-latency or GNU process-RSS pair flag
above 5%; the phase flags below are retained because the phase regions are
short and repeat-sensitive.

All 48 capture processes exited with status zero.  Every allocator total
sample returned to its entry live-byte count, and every allocator phase
sample satisfied live-byte conservation and continuous phase boundaries.  The
source reads, sink records, corpus identities, and output oracles were
unchanged between arms, repeats, binaries, and modes.

This remains a materialized one-paragraph transaction.  It does not establish
an explicit bounded-window append, repeated append scaling, section-property
support, or the broader non-iWork goal in `docs/GOAL.md`.

## Scanner accounting and source-derived prediction

The source diff was checked against the current candidate and the retained
`before-source.txt`.  It only replaces the three fallible local-name byte
clones in `Start`, `Empty`, and `End` event handling with borrowed
`LocalName` wrappers.  The call graph still performs four scanner passes per
total lifecycle:

1. the initial snapshot scans the source XML of N paragraphs;
2. `copy_fragment` constructs the projected snapshot and scans candidate XML
   of N+1 paragraphs;
3. publication captures a fresh source snapshot and scans source XML again;
4. `Patch::apply` clones the candidate bytes and scans candidate XML again.

The phase harness places these same passes in `snapshot` (source), `stage`
(candidate), and `publish` (source plus candidate).  The source call graph,
not an inferred timing pattern, was used to verify that no scanner pass was
removed.  The reports do not expose a separate parser-event counter, so the
four-pass statement is source-level accounting while the allocation and
timing measurements are raw runtime evidence.

The source-derived predictions are therefore two source-N scans and two
candidate-(N+1) scans, with the three local-name clone sites exercised within
those unchanged passes.  The observed allocator deltas are:

| paragraphs | removed allocation callbacks | formula | removed requested bytes | formula |
| ---: | ---: | :--- | ---: | :--- |
| 64 | 1,564 | `24*64 + 28` | 1,644 | `24*64 + 108` |
| 8,192 | 196,636 | `24*8192 + 28` | 196,716 | `24*8192 + 108` |
| 131,072 | 3,145,756 | `24*131072 + 28` | 3,145,836 | `24*131072 + 108` |

The borrowed `LocalName` does not escape its parser-event match arm.  The
scanner still owns the namespace bytes and layout, and the exact source
fragments copied by the edit are unchanged.

## Corpus, output, source reads, and sink identity

The candidate copies one 49-byte paragraph fragment.  The source and
candidate XML and archive identities below are the frozen corpus values used
by every formal report.

| paragraphs | source XML bytes | candidate XML bytes | source archive bytes | candidate archive bytes | source archive SHA-256 | candidate archive SHA-256 |
| ---: | ---: | ---: | ---: | ---: | :--- | :--- |
| 64 | 3,287 | 3,336 (+49) | 2,108 | 2,112 (+4) | `b8756cc3ca2cf55572d989ec95629d71d1e6f515b08d42253db78d0a493539a9` | `297f40d52e7494cd3f4601ef75a7f93e53bd5999dfa90cf4ecf5501aa7a61ca8` |
| 8,192 | 401,559 | 401,608 (+49) | 23,357 | 23,362 (+5) | `aca7dc0072331b91f04fd2bf7a4262575dd9329f75db1453501bb890d967881d` | `c29151eb2672bd2910fdbdf55f922dfdd01aa39d1cdfa2ffb6136d6579c4fdc9` |
| 131,072 | 6,422,679 | 6,422,728 (+49) | 343,945 | 343,949 (+4) | `b545060378379ffcd21a72d6534c5128875dbb368a588c68a2a07f73d4edbd42` | `2152458e09392b4aa836b8f22c4d27fd11d58d3af4dbabc261f202ce09714e4e` |

The corresponding main-XML SHA-256 pairs are, in size order,
`6c6bf4974aa8ac4cc18ae89ebeb75cdbbe0255173fbbc78418a937a647db37fa9` →
`65fea96781565784866917b629db330827813dae70dbf726497d6c6c4eb73921`,
`4268bed02a76ff093b425c87e61c691df76518b4bee447039336d951a77da68a` →
`5db002ddff00bd44e51fc3e53b7f0454fbdb5c5b2647dbdcebc2a5917d1ff697`, and
`24d43839e52d4aa836c7c4e3a574b722baab3f366e17e354aa0e96d28d6a962a` →
`2283078ddbdf199b0eba8b5891c146d02994495d60789af768d7f21631c93d8f`.

Every report retained the same semantic paragraph/text/order oracles, exact
untouched ZIP members, physical member order, source-unchanged checks, and
inverse/replay checks.  The source and candidate archive identities above
were constant across all 48 reports; each sink hash was the candidate archive
hash for its size.

The source-read histogram columns are `0 / 1–512 / 513–4096 /
4097–16384 / 16385–65536 / >65536` bytes.  These lifecycle observations were
constant in all arms, repeats, instruments, and modes.

| paragraphs | source `ReadAt` calls | requested = returned bytes | source request histogram | sink accepted bytes | sink writes | largest write | sink histogram |
| ---: | ---: | ---: | :--- | ---: | ---: | ---: | :--- |
| 64 | 42 | 7,969 | 3 / 36 / 3 / 0 / 0 / 0 | 2,112 | 19 | 877 | 0 / 18 / 1 / 0 / 0 / 0 |
| 8,192 | 42 | 71,716 | 3 / 35 / 1 / 0 / 3 / 0 | 23,362 | 20 | 16,384 | 0 / 17 / 1 / 2 / 0 / 0 |
| 131,072 | 62 | 1,033,480 | 3 / 35 / 1 / 3 / 20 / 0 | 343,949 | 39 | 16,384 | 0 / 17 / 1 / 21 / 0 / 0 |

The phase source-read columns below are `calls / requested bytes`.  They sum
to the lifecycle values for every sample, and the same phase allocation and
I/O identities held for both repeats and both arms.

| paragraphs | `open` | `snapshot` | `stage` | `commit` | `publish` | `drop` |
| ---: | :--- | :--- | :--- | :--- | :--- | :--- |
| 64 | 13 / 893 | 6 / 2,493 | 0 / 0 | 0 / 0 | 23 / 4,583 | 0 / 0 |
| 8,192 | 13 / 893 | 6 / 44,991 | 0 / 0 | 0 / 0 | 23 / 25,832 | 0 / 0 |
| 131,072 | 13 / 893 | 21 / 686,167 | 0 / 0 | 0 / 0 | 28 / 346,420 | 0 / 0 |

## Allocator totals and zero-exit checks

For allocator reports, the incremental peak is
`region_peak_live_bytes - live_bytes_before`.  The values below were
constant across all 30 samples in both repeats.  `allocated_bytes` is the
requested-byte counter and includes the full `new_size` on realloc callbacks;
it is not physical copy traffic.

| paragraphs | incremental peak control → candidate | delta | allocation calls control → candidate | deallocation calls control → candidate | requested bytes control → candidate | delta | realloc calls |
| ---: | :--- | ---: | :--- | :--- | :--- | ---: | :--- |
| 64 | 506,598 → 506,598 | 0 | 1,886 → 322 | 1,819 → 255 | 1,033,593 → 1,031,949 | -1,644 | 67 → 67 |
| 8,192 | 2,669,662 → 2,669,662 | 0 | 213,344 → 16,708 | 196,891 → 255 | 1,614,655,697 → 1,614,458,981 | -196,716 | 16,453 → 16,453 |
| 131,072 | 35,371,102 → 35,371,102 | 0 | 3,653,984 → 508,228 | 3,146,011 → 255 | 549,266,500,817 → 549,263,354,981 | -3,145,836 | 507,973 → 507,973 |

The candidate entry and exit live-byte values are two bytes higher than the
control values because of the surrounding executable's baseline ownership;
that shift is removed by the incremental-peak calculation.  The absolute
region peaks are correspondingly `530,831 → 530,833`, `2,736,401 →
2,736,403`, and `36,079,019 → 36,079,021`; none is a candidate reduction.

All 360 allocator total samples have `live_bytes_after == live_bytes_before`.
All 360 allocator phase samples have zero failed allocations, satisfy
`live_before + allocated_bytes - deallocated_bytes == live_after`, preserve
adjacent phase boundaries, and return to the entry live count at `drop`.
Normal reports intentionally expose allocation metrics as unavailable rather
than zero.  All 48 capture envelopes have exit code zero.

## Phase allocation attribution

The table gives allocation calls, deallocation calls, requested bytes,
retained live-byte delta, and the absolute phase region peak as
control → candidate, followed by the candidate-minus-control delta in each
cell.  Reallocation callbacks were unchanged in every phase at every size.
The phase peaks are attribution endpoints and are not summed into an
operation peak.

| paragraphs | phase | allocation calls | deallocation calls | requested bytes | retained delta | region peak |
| ---: | :--- | :--- | :--- | :--- | :--- | :--- |
| 64 | open | 151 → 151 (+0) | 95 → 95 (+0) | 306,689 → 306,689 (+0) | 3,408 → 3,408 (+0) | 219,371 → 219,373 (+2) |
| 64 | snapshot | 433 → 45 (-388) | 412 → 24 (-388) | 153,780 → 153,372 (-408) | 4,748 → 4,748 (+0) | 173,698 → 173,700 (+2) |
| 64 | stage | 414 → 20 (-394) | 403 → 9 (-394) | 6,730 → 6,316 (-414) | 4,432 → 4,432 (+0) | 100,529 → 100,531 (+2) |
| 64 | commit | 0 → 0 (+0) | 0 → 0 (+0) | 0 → 0 (+0) | 0 → 0 (+0) | 98,982 → 98,984 (+2) |
| 64 | publish | 888 → 106 (-782) | 896 → 114 (-782) | 566,394 → 565,572 (-822) | 1,844 → 1,844 (+0) | 592,992 → 592,994 (+2) |
| 64 | drop | 0 → 0 (+0) | 13 → 13 (+0) | 0 → 0 (+0) | -14,432 → -14,432 (+0) | 100,826 → 100,828 (+2) |
| 8,192 | open | 151 → 151 (+0) | 95 → 95 (+0) | 306,689 → 306,689 (+0) | 3,408 → 3,408 (+0) | 261,877 → 261,879 (+2) |
| 8,192 | snapshot | 53,297 → 4,141 (-49,156) | 49,180 → 24 (-49,156) | 403,481,316 → 403,432,140 (-49,176) | 533,068 → 533,068 (+0) | 796,995 → 796,997 (+2) |
| 8,192 | stage | 53,279 → 4,117 (-49,162) | 49,171 → 9 (-49,162) | 403,465,354 → 403,416,172 (-49,182) | 532,752 → 532,752 (+0) | 1,329,723 → 1,329,725 (+2) |
| 8,192 | commit | 0 → 0 (+0) | 0 → 0 (+0) | 0 → 0 (+0) | 0 → 0 (+0) | 1,198,128 → 1,198,130 (+2) |
| 8,192 | publish | 106,617 → 8,299 (-98,318) | 98,432 → 114 (-98,318) | 807,402,338 → 807,303,980 (-98,358) | 660,212 → 660,212 (+0) | 2,798,562 → 2,798,564 (+2) |
| 8,192 | drop | 0 → 0 (+0) | 13 → 13 (+0) | 0 → 0 (+0) | -1,729,440 → -1,729,440 (+0) | 1,858,340 → 1,858,342 (+2) |
| 131,072 | open | 151 → 151 (+0) | 95 → 95 (+0) | 306,689 → 306,689 (+0) | 3,408 → 3,408 (+0) | 903,055 → 903,057 (+2) |
| 131,072 | snapshot | 913,457 → 127,021 (-786,436) | 786,460 → 24 (-786,436) | 137,315,271,396 → 137,314,484,940 (-786,456) | 8,520,268 → 8,520,268 (+0) | 11,391,453 → 11,391,455 (+2) |
| 131,072 | stage | 913,439 → 126,997 (-786,442) | 786,451 → 9 (-786,442) | 137,317,221,514 → 137,316,435,052 (-786,462) | 8,519,952 → 8,519,952 (+0) | 19,911,381 → 19,911,383 (+2) |
| 131,072 | commit | 0 → 0 (+0) | 0 → 0 (+0) | 0 → 0 (+0) | 0 → 0 (+0) | 17,813,706 → 17,813,708 (+2) |
| 131,072 | publish | 1,826,937 → 254,059 (-1,572,878) | 1,572,992 → 114 (-1,572,878) | 274,633,701,218 → 274,632,128,300 (-1,572,918) | 10,613,492 → 10,613,492 (+0) | 36,141,180 → 36,141,182 (+2) |
| 131,072 | drop | 0 → 0 (+0) | 13 → 13 (+0) | 0 → 0 (+0) | -27,657,120 → -27,657,120 (+0) | 28,427,198 → 28,427,200 (+2) |

## Normal total latency and GNU whole-process RSS

RSS in these tables is the one GNU `/usr/bin/time -v` maximum resident set
size per process, in KiB.  It is not a 30-sample operation-local statistic;
the same process value is shown for mean, p50, p95, and p99 when the raw
summary is derived.  Total and phase runs are separate processes.

### Total mode

| paragraphs | repeat | control mean / p50 / p95 / p99 | candidate mean / p50 / p95 / p99 | candidate mean delta | GNU RSS control → candidate | RSS delta |
| ---: | ---: | :--- | :--- | ---: | ---: | ---: |
| 64 | 1 | 0.168671 / 0.166951 / 0.180891 / 0.183601 | 0.155033 / 0.153471 / 0.164351 / 0.174640 | -8.086% | 4,808 → 4,800 | -0.166% |
| 8,192 | 1 | 16.597530 / 16.557378 / 17.027420 / 17.132290 | 14.786540 / 14.767391 / 14.921771 / 14.924561 | -10.911% | 10,728 → 10,980 | +2.349% |
| 131,072 | 1 | 264.665767 / 264.071112 / 267.857768 / 273.712351 | 236.834936 / 236.738972 / 237.896047 / 237.990327 | -10.515% | 107,200 → 107,348 | +0.138% |
| 64 | 2 | 0.169057 / 0.168831 / 0.172771 / 0.177640 | 0.155796 / 0.154071 / 0.168821 / 0.170571 | -7.844% | 4,804 → 4,588 | -4.496% |
| 8,192 | 2 | 16.545505 / 16.571138 / 16.734568 / 16.948739 | 14.624062 / 14.609741 / 14.727752 / 15.023713 | -11.613% | 10,720 → 11,020 | +2.799% |
| 131,072 | 2 | 270.711302 / 270.448847 / 272.990398 / 273.869812 | 238.225314 / 238.031517 / 239.282742 / 239.493653 | -12.000% | 107,052 → 107,200 | +0.138% |

### Phase mode

The phase timing columns are the statistics of the per-sample sum of the six
phase elapsed values, included as attribution context.  The individual phase
statistics remain separate in the exhaustive flag table below.

| paragraphs | repeat | control phase-sum mean / p50 / p95 / p99 | candidate phase-sum mean / p50 / p95 / p99 | candidate mean delta | GNU RSS control → candidate | RSS delta |
| ---: | ---: | :--- | :--- | ---: | ---: | ---: |
| 64 | 1 | 0.171305 / 0.169710 / 0.181431 / 0.183831 | 0.157258 / 0.156241 / 0.163911 / 0.164852 | -8.200% | 4,608 → 4,800 | +4.167% |
| 8,192 | 1 | 16.587737 / 16.580618 / 16.784568 / 16.837709 | 14.637393 / 14.637129 / 14.724390 / 14.781801 | -11.758% | 10,728 → 10,984 | +2.386% |
| 131,072 | 1 | 272.559360 / 272.051366 / 275.139768 / 275.374848 | 236.766001 / 236.683904 / 238.446879 / 240.047037 | -13.132% | 107,048 → 107,200 | +0.142% |
| 64 | 2 | 0.172345 / 0.170491 / 0.181631 / 0.182801 | 0.157124 / 0.155340 / 0.165851 / 0.173580 | -8.832% | 4,592 → 4,592 | +0.000% |
| 8,192 | 2 | 16.555417 / 16.542058 / 16.666248 / 16.697759 | 14.611155 / 14.591350 / 14.768642 / 14.920413 | -11.744% | 10,732 → 10,980 | +2.311% |
| 131,072 | 2 | 268.341205 / 268.244116 / 272.356513 / 272.995625 | 236.489232 / 236.144429 / 238.889239 / 239.071886 | -11.870% | 107,300 → 107,484 | +0.171% |

Allocator total mean timings, for context, were 0.206564 → 0.159027 ms
(-23.013%), 20.719287 → 14.844871 ms (-28.352%), and 333.497739 →
240.207486 ms (-27.973%) in repeat 1 for 64, 8,192, and 131,072 paragraphs.
Repeat 2 was 0.207790 → 0.159202 ms (-23.383%), 20.593167 → 14.806256 ms
(-28.101%), and 333.609728 → 240.099394 ms (-28.030%).

## Exhaustive pair flags above 5%

All 24 candidate-versus-control pairs were checked for every latency mean,
p50, p95, and p99 in total mode and for every individual phase in phase mode.
The GNU RSS check uses the scalar process maximum for the corresponding total
or phase process.  A `+` is a candidate cost and a `-` is a lower candidate
observation.  The table includes every latency flag with absolute change
strictly greater than 5%; `RSS` would identify a GNU process-RSS flag.

No GNU RSS flag occurred in either total or phase mode: all 24 total/phase
process-RSS pair changes were within ±5%.  Thus the phase-mode `RSS` values in
the timing table are retained process-scope observations, not omitted phase
RSS evidence.  The short internal phase `peak_rss_bytes` and `rss_bytes`
fields are separate observer metrics and are not substituted for GNU RSS.

| repeat | binary | paragraphs | mode | flags (candidate versus control) |
| ---: | :--- | ---: | :--- | :--- |
| 1 | normal | 64 | total | `latency.mean` -8.09%; `latency.p50` -8.07%; `latency.p95` -9.14% |
| 1 | normal | 64 | phases | `open.mean` +5.61%; `open.p95` +16.66%; `open.p99` +48.87%; `snapshot.mean` -12.00%; `snapshot.p50` -11.15%; `snapshot.p95` -16.37%; `snapshot.p99` -26.34%; `stage.mean` -14.72%; `stage.p50` -14.39%; `stage.p95` -13.92%; `stage.p99` -23.15%; `commit.mean` -16.85%; `commit.p50` -16.67%; `commit.p95` -14.29%; `commit.p99` -12.50%; `publish.mean` -7.15%; `publish.p50` -6.66%; `publish.p95` -10.75%; `publish.p99` -10.60%; `drop.p99` -8.82% |
| 1 | normal | 8,192 | total | `latency.mean` -10.91%; `latency.p50` -10.81%; `latency.p95` -12.37%; `latency.p99` -12.89% |
| 1 | normal | 8,192 | phases | `snapshot.mean` -14.97%; `snapshot.p50` -14.97%; `snapshot.p95` -15.20%; `snapshot.p99` -15.39%; `stage.mean` -13.81%; `stage.p50` -13.94%; `stage.p95` -12.94%; `stage.p99` -12.81%; `commit.mean` -8.06%; `commit.p95` -22.22%; `commit.p99` -22.22%; `publish.mean` -9.95%; `publish.p50` -9.73%; `publish.p95` -11.44%; `publish.p99` -11.62%; `drop.mean` -5.64%; `drop.p95` -15.47%; `drop.p99` -30.25% |
| 1 | normal | 131,072 | total | `latency.mean` -10.52%; `latency.p50` -10.35%; `latency.p95` -11.19%; `latency.p99` -13.05% |
| 1 | normal | 131,072 | phases | `open.mean` +7.04%; `open.p50` +5.66%; `open.p95` +11.36%; `open.p99` +12.17%; `snapshot.mean` -16.43%; `snapshot.p50` -16.56%; `snapshot.p95` -16.61%; `snapshot.p99` -15.82%; `stage.mean` -15.84%; `stage.p50` -15.63%; `stage.p95` -16.83%; `stage.p99` -17.79%; `commit.mean` -37.40%; `commit.p50` -55.00%; `commit.p95` -6.45%; `publish.mean` -10.96%; `publish.p50` -11.01%; `publish.p95` -10.79%; `publish.p99` -10.78%; `drop.p95` +8.59%; `drop.p99` +28.40% |
| 1 | allocator | 64 | total | `latency.mean` -23.01%; `latency.p50` -22.99%; `latency.p95` -22.34%; `latency.p99` -19.40% |
| 1 | allocator | 64 | phases | `open.p95` -24.19%; `open.p99` -7.46%; `snapshot.mean` -27.35%; `snapshot.p50` -26.59%; `snapshot.p95` -27.05%; `snapshot.p99` -40.22%; `stage.mean` -33.70%; `stage.p50` -33.93%; `stage.p95` -33.75%; `stage.p99` -25.98%; `commit.p95` +16.67%; `commit.p99` -41.67%; `publish.mean` -20.04%; `publish.p50` -20.05%; `publish.p95` -17.69%; `publish.p99` -21.65% |
| 1 | allocator | 8,192 | total | `latency.mean` -28.35%; `latency.p50` -28.33%; `latency.p95` -28.44%; `latency.p99` -28.40% |
| 1 | allocator | 8,192 | phases | `open.p99` -6.51%; `snapshot.mean` -32.83%; `snapshot.p50` -32.89%; `snapshot.p95` -32.83%; `snapshot.p99` -34.22%; `stage.mean` -33.57%; `stage.p50` -33.87%; `stage.p95` -32.07%; `stage.p99` -31.95%; `commit.mean` -7.10%; `commit.p50` -16.67%; `commit.p95` +16.67%; `commit.p99` +14.29%; `publish.mean` -24.11%; `publish.p50` -24.17%; `publish.p95` -23.30%; `publish.p99` -23.57%; `drop.p95` -19.45%; `drop.p99` -18.88% |
| 1 | allocator | 131,072 | total | `latency.mean` -27.97%; `latency.p50` -28.01%; `latency.p95` -28.05%; `latency.p99` -27.98% |
| 1 | allocator | 131,072 | phases | `snapshot.mean` -31.87%; `snapshot.p50` -31.79%; `snapshot.p95` -31.93%; `snapshot.p99` -31.82%; `stage.mean` -31.24%; `stage.p50` -31.95%; `stage.p95` -29.08%; `stage.p99` -28.85%; `commit.mean` -23.40%; `commit.p50` -30.00%; `commit.p95` +24.24%; `commit.p99` +33.33%; `publish.mean` -24.33%; `publish.p50` -24.34%; `publish.p95` -24.10%; `publish.p99` -23.94%; `drop.p50` +6.72%; `drop.p95` -9.14%; `drop.p99` -5.97% |
| 2 | normal | 64 | total | `latency.mean` -7.84%; `latency.p50` -8.74% |
| 2 | normal | 64 | phases | `open.p99` -33.89%; `snapshot.mean` -10.67%; `snapshot.p50` -11.05%; `snapshot.p95` -12.31%; `stage.mean` -12.58%; `stage.p50` -13.33%; `stage.p95` +8.91%; `stage.p99` -11.20%; `commit.mean` +6.85%; `commit.p95` +40.00%; `commit.p99` +16.67%; `publish.mean` -8.43%; `publish.p50` -8.48%; `publish.p95` -6.27%; `publish.p99` -8.98%; `drop.p95` +6.67%; `drop.p99` -20.00% |
| 2 | normal | 8,192 | total | `latency.mean` -11.61%; `latency.p50` -11.84%; `latency.p95` -11.99%; `latency.p99` -11.36% |
| 2 | normal | 8,192 | phases | `open.p95` -6.91%; `snapshot.mean` -14.42%; `snapshot.p50` -15.17%; `snapshot.p95` -12.50%; `snapshot.p99` -7.71%; `stage.mean` -13.59%; `stage.p50` -13.80%; `stage.p95` -12.12%; `stage.p99` -12.77%; `commit.mean` +8.33%; `commit.p50` +20.00%; `commit.p95` +42.86%; `commit.p99` +25.00%; `publish.mean` -10.19%; `publish.p50` -10.06%; `publish.p95` -10.00%; `publish.p99` -10.31%; `drop.p99` -26.33% |
| 2 | normal | 131,072 | total | `latency.mean` -12.00%; `latency.p50` -11.99%; `latency.p95` -12.35%; `latency.p99` -12.55% |
| 2 | normal | 131,072 | phases | `open.p95` +8.43%; `snapshot.mean` -13.72%; `snapshot.p50` -13.62%; `snapshot.p95` -15.06%; `snapshot.p99` -14.70%; `stage.mean` -13.71%; `stage.p50` -13.88%; `stage.p95` -14.75%; `stage.p99` -14.64%; `commit.mean` -33.42%; `commit.p50` -33.33%; `commit.p95` -37.21%; `commit.p99` -33.33%; `publish.mean` -10.58%; `publish.p50` -10.50%; `publish.p95` -10.32%; `publish.p99` -10.78%; `drop.p95` +14.96%; `drop.p99` +23.08% |
| 2 | allocator | 64 | total | `latency.mean` -23.38%; `latency.p50` -23.52%; `latency.p95` -22.81%; `latency.p99` -23.05% |
| 2 | allocator | 64 | phases | `open.p95` -7.53%; `open.p99` +31.20%; `snapshot.mean` -29.71%; `snapshot.p50` -28.22%; `snapshot.p95` -39.62%; `snapshot.p99` -45.46%; `stage.mean` -32.29%; `stage.p50` -33.45%; `stage.p95` -33.32%; `commit.p99` +40.00%; `publish.mean` -20.24%; `publish.p50` -20.97%; `publish.p95` -17.05%; `publish.p99` -16.56% |
| 2 | allocator | 8,192 | total | `latency.mean` -28.10%; `latency.p50` -28.05%; `latency.p95` -27.65%; `latency.p99` -27.93% |
| 2 | allocator | 8,192 | phases | `open.p95` -5.22%; `snapshot.mean` -32.44%; `snapshot.p50` -32.38%; `snapshot.p95` -32.31%; `snapshot.p99` -31.91%; `stage.mean` -33.24%; `stage.p50` -33.32%; `stage.p95` -33.29%; `stage.p99` -33.57%; `commit.mean` +29.49%; `commit.p95` +43.75%; `commit.p99` +17.39%; `publish.mean` -24.50%; `publish.p50` -24.51%; `publish.p95` -24.16%; `publish.p99` -23.96%; `drop.p99` +29.02% |
| 2 | allocator | 131,072 | total | `latency.mean` -28.03%; `latency.p50` -28.10%; `latency.p95` -28.02%; `latency.p99` -27.87% |
| 2 | allocator | 131,072 | phases | `open.p99` -6.12%; `snapshot.mean` -31.43%; `snapshot.p50` -31.45%; `snapshot.p95` -30.68%; `snapshot.p99` -30.67%; `stage.mean` -32.31%; `stage.p50` -32.49%; `stage.p95` -31.80%; `stage.p99` -31.17%; `commit.mean` +20.95%; `commit.p50` +44.44%; `commit.p95` +7.14%; `commit.p99` -11.54%; `publish.mean` -24.62%; `publish.p50` -24.76%; `publish.p95` -24.58%; `publish.p99` -23.80%; `drop.p95` +11.45% |

The positive candidate-cost pair flags are confined to short phase regions;
the negative total-lifecycle flags are the consistent scanner-cost reduction
seen in the full operation.  None of the phase ratios is promoted to a
general performance claim.

## Exhaustive repeat-drift flags

All 24 arm/binary/size/mode combinations were checked for absolute repeat-2
versus repeat-1 changes in total latency, every phase latency statistic, and
GNU process RSS.  Twelve phase rows contain 89 latency flags; the other 12
rows have no flag.  No GNU RSS repeat drift exceeded 5%.

| arm | binary | paragraphs | mode | absolute drift flags (repeat 2 versus repeat 1) |
| :--- | :--- | ---: | :--- | :--- |
| control | normal | 64 | phases | `open.mean` +5.64%; `open.p99` +60.82%; `snapshot.p95` -5.27%; `snapshot.p99` +5.22%; `stage.p95` +5.90%; `stage.p99` -10.49%; `commit.mean` -20.65%; `commit.p50` -16.67%; `commit.p95` -28.57%; `commit.p99` -25.00%; `drop.p99` +17.65% |
| control | normal | 8,192 | phases | `open.p95` +13.52%; `open.p99` +7.34%; `commit.mean` -9.68%; `commit.p50` -16.67%; `commit.p95` -22.22%; `commit.p99` -11.11%; `drop.p95` -13.41%; `drop.p99` -6.30% |
| control | normal | 131,072 | phases | `open.p95` -7.70%; `commit.mean` +28.67%; `commit.p50` +35.00%; `commit.p95` +38.71%; `commit.p99` +64.52%; `drop.mean` -16.37%; `drop.p50` -14.96%; `drop.p95` -22.09%; `drop.p99` -23.08% |
| control | allocator | 64 | phases | `open.p95` -21.35%; `snapshot.p95` +21.19%; `snapshot.p99` +40.40%; `commit.mean` -7.55%; `commit.p95` -16.67%; `commit.p99` -58.33%; `publish.p99` -17.16%; `drop.mean` -5.58%; `drop.p95` -9.09% |
| control | allocator | 8,192 | phases | `open.p95` +5.51%; `commit.mean` +28.40%; `commit.p50` -16.67%; `commit.p95` +166.67%; `commit.p99` +228.57%; `drop.p95` -18.46%; `drop.p99` -15.12% |
| control | allocator | 131,072 | phases | `open.p95` +6.34%; `open.p99` +10.84%; `commit.p50` -10.00%; `commit.p95` +27.27%; `commit.p99` +57.58%; `drop.p95` -10.75% |
| candidate | normal | 64 | phases | `open.p95` -10.54%; `open.p99` -28.57%; `snapshot.p99` +46.88%; `stage.p95` +33.98%; `commit.p95` +16.67%; `drop.p95` +6.67% |
| candidate | normal | 8,192 | phases | `snapshot.p99` +9.40%; `commit.mean` +6.43%; `commit.p95` +42.86%; `commit.p99` +42.86% |
| candidate | normal | 131,072 | phases | `open.mean` -7.46%; `open.p50` -6.45%; `open.p95` -10.14%; `open.p99` -11.33%; `commit.mean` +36.84%; `commit.p50` +100.00%; `commit.p95` -6.90%; `commit.p99` +9.68%; `drop.mean` -12.33%; `drop.p50` -14.06%; `drop.p95` -17.51%; `drop.p99` -26.27% |
| candidate | allocator | 64 | phases | `open.p99` +36.71%; `snapshot.p99` +28.09%; `stage.p99` +30.40%; `commit.mean` -9.62%; `commit.p95` -28.57%; `publish.p99` -11.78%; `drop.p95` -6.82%; `drop.p99` -6.67% |
| candidate | allocator | 8,192 | phases | `commit.mean` +78.98%; `commit.p95` +228.57%; `commit.p99` +237.50%; `drop.p99` +35.00% |
| candidate | allocator | 131,072 | phases | `open.p50` +5.08%; `commit.mean` +54.60%; `commit.p50` +85.71%; `commit.p95` +9.76%; `drop.p95` +9.47% |

The drift flags are stability signals for short attribution regions, not
candidate regressions.  In particular, none changes the exact allocator
callback/request-byte identity or the zero-net-retention result.

## Scope and limitations

The normal binary does not install the counting allocator, so normal
allocation fields are unavailable rather than zero.  The allocator's
`allocated_bytes` includes realloc `new_size`; it does not report physical
bytes copied by reallocations.  The per-sample process observer's
`rss_bytes` is a saturating RSS delta and `peak_rss_bytes` is an absolute
VmHWM endpoint.  The RSS tables and RSS flag checks here deliberately use the
broader GNU `time` maximum for the complete process.

The measured total lifecycle includes package construction, source scanning,
snapshot, staging, commit, publication, sink digest finalization, and owner
destruction.  Corpus and independent oracle storage are prepared outside the
timed operation while remaining inside the GNU whole-process RSS envelope.
Phase peaks are separate attribution regions and are never summed into a
total-operation peak.

The source manifest covers 7,048 files; control is bound to
`87acdfee1a43725688a0c4ffee0a623463b45a87168faaf683829fe99bb95893` and the
candidate to `bcec8208a7ce40c82eb59f20868afc090a0c314fb6431909f17a757132ee88b8`.
The source-level scanner accounting and all runtime evidence concern this
single borrowed-name change.  The broader explicit-window append capability
and the remaining non-iWork goal stay open.
