# 0689 — retain the validated SST chain position for indexed XLS replay

Status: retained after correctness, performance and evidence review. `performance_claim: none`.
OLE2/OOXML remain active; iWork is excluded.

## Measured hypothesis

After [0688](0688-cfb-checked-chain-hot-path.md), a fresh two-million-query
54016 owned-source profile still assigns 57.42% self samples to checked CFB
chain walking. Separate temporary tracing distinguishes its callers: a warm
first-cell query walks 14 links from the worksheet checkpoint and then 478
links from the workbook stream's first sector to its selected shared string.
Plan1 repeats six worksheet links and 97 SST links. Numeric late targets have
only worksheet traversal. These counts identify repeated work that the
previous call-inlining change could not remove.

The trace is diagnostic only: its stderr writes invalidate its timing output.
Its instrumentation patch, build, source/binary bindings and route output are
retained separately from the uninstrumented latency and allocation probes.
The original sources are restored after every trace run, including failures.

## Implementation and invariants

The private worksheet index gains one immutable SST chain checkpoint, alongside
its existing worksheet-start checkpoint. A building selected-cell scan captures
it only after a successful `LabelSst` target result, only once, and only while
the candidate remains active. The first selected result may be a typed error
value with no SST read; its checkpoint then has no useful position. Failed
reads/decodes do not capture a new position. Numeric/formula/missing targets do
not create an SST checkpoint.

Publication still requires the completed worksheet scan and existing source
and execution fences. Replay seeds a fresh local resolver hint from the opaque
CFB checkpoint. Existing owner, stream, allocation-table, first-sector and
backward-position checks govern reuse. A request before the retained SST
position restarts normally. Every replay still reads and decodes its own SST
payload; no source bytes, decoded values or errors are cached. Duplicate
occurrences remain ordered and errors cannot be hidden by later duplicates.

The checkpoint retains a weak index identity, not a source or allocation table.
Initialization, publication, abandonment and drop follow the existing index
lifecycle. Fixed logical index overhead rises 192 → 224 bytes, charged before
collection through the existing local and hierarchical memory reservations.
The extra field occupies every admitted index, including indexes built for
numeric/missing targets with no useful SST point. There is no CFB API change,
new dependency, unsafe code, mutable shared hint or extra source read.

## Evidence scope

The [packet](results/change-0689/README.md) binds baseline `9ba2c0b6b`, unchanged
probes and fixtures, final sources, binaries, commands and raw captures.
Native A/A then A/B/B/A cover 24 case/source groups, 14,400 fresh owners and
115,200 queries. Eighth-query latency and open-plus-eight timer sums remain
separate. The long-loop controls average 50,000 queries in each of nine
processes per leg; all use the default 2 MiB ceiling, even the missing route
whose native case uses 1 MiB. These are distributions of loop means, not
individual query latencies.

Allocation and counted-I/O probes remain separate from native timing. Counter
and native-child RSS diagnostics compare N=10 with N=100,010; counters include
whole-process setup and the wrapper. RSS, allocator live gauges and logical
reservation charges are distinct. CPU 12, shared-host conditions and warm OS
file caches are recorded. Physical cold-device, remote, concurrency,
cross-platform and native Office performance remain outside this experiment.

## Main paired results

| Eighth query | Owned median change | File median change, warm OS cache |
|---|---:|---:|
| 54016 first, default index limit | −49.09% / −49.55% | −20.00% / −19.78% |
| Plan1 first | −18.03% / −16.39% | −4.09% / −5.80% |
| 45365-2 first | −7.27% / −8.93% | −1.38% / −1.38% |
| Simple first | −4.15% / −2.13% | +0.95% / +0.24% |
| 54016 numeric late | −0.68% / −1.34% | +0.75% / +0.37% |

54016 first owned falls from 1,100/1,110 ns to 560 ns in both pairs.
Its longer same-owner controls improve 58.03–58.04% owned and
21.38–22.18% file. Plan1 owned loop means improve 20.11–20.43%, and
45365-2 first owned improves 6.68–6.75%. The numeric late 54016 controls
remain approximately flat (owned −0.95%/−0.27%, file −0.22%/−0.58%),
consistent with the absence of an SST resolve on that route.

Open-plus-eight does not inherit the large warm-query gain: default-index
54016 first improves 0.79–1.47% owned and 0.70–1.06% file. Plan1 windows
move about −0.5–1.1%. Simple/file windows rise 1.18–1.44%; formula-refusal
owned windows rise 1.13–2.32%, with q8 up 1.30–1.57%. These smaller costs
remain in the complete matrix. Disabled-index owned 54016 q8 improves
5.75–5.92%, even though no checkpoint is retained; no checkpoint mechanism
is attributed to that unrelated movement.

There is no paired native p50 regression above 5% in the main matrix, but
[mean/tail triggers](results/change-0689/tail-regressions.md) remain. They
include 45365-2/file q8 p99 rises around 53–54%, generated warm-query p99,
and several first/open/workflow tails. One hundred samples do not establish
a population p99. The original 54016 missing/1 MiB owned A2 first-query
median also doubles versus A1 (+101.83%), while q2 and q8 remain stable;
its open-plus-eight control drifts +30.93%. The resulting apparent −22.89%
second-pair workflow change is not accepted as a performance gain. Additional
controls below investigate these specific uncertainties without overwriting
the original data.

## Verified mechanism and memory tradeoffs

Matched diagnostic tracing shows the selected repeated SST walks becoming
zero links: 54016 478 → 0, Plan1 97 → 0, 45365-2 24 → 0, and
Simple/generated first-cell 2 → 0. Worksheet walks and source offsets remain
exactly unchanged; numeric late targets retain 1,044/72/258 worksheet links
for 54016/Plan1/45365-2. Missing targets have no indexed replay chain walk, and the formula
refusal still occurs before indexed replay. Both traces are excluded from
latency evidence.

The uninstrumented two-million-query profile falls 1.908 → 0.805 seconds
(about 57.8%). Chain-walk self share falls 57.42% → 6.85%; remaining leading
shares are indexed query work (28.75%), cursor reads (20.38%) and directory-name
construction (17.16%). These percentages describe the smaller candidate total,
not absolute growth of those functions.

Whole-process extra-query estimates support the removed work:

| Owned target | Instructions before → after | Cycles before → after |
|---|---:|---:|
| 54016 first | 15,788 → 6,743 | 4,202 → 1,751 |
| Plan1 first | 9,559 → 7,757 | 2,262 → 1,819 |
| 54016 numeric late | 24,593 → 24,591 | 6,478 → 6,452 |

First-target 54016 instructions/cycles fall 57.29%/58.33%, and Plan1 falls
18.85%/19.57%. Branch counts fall on the SST-heavy routes. Differenced miss
and page-fault counts can be near zero or negative and are not used to claim
locality improvements. Native timers and longer controls remain the latency
proof. Whole-binary `.text` changes 637,523 → 637,379 bytes (−144); the
inspected `query_cell` symbol changes 34,500 → 34,104 bytes. CFB cursor/walk
code sizes are unchanged. The `query_cell` stack reservation grows 1,992 → 2,040
bytes (+48), an explicit cost despite the code shrinkage. This is not a peak
stack bound. No clean-build claim is made.

All 12 counted routes preserve opening and eight queries' read counts, bytes,
freshness observations and outcomes. All 72 non-build allocation groups match
exactly, and every group has three identical repeats per binary. Successful
ordinary-capacity index construction adds exactly 32 requested and retained bytes,
with no additional allocation call and unchanged peak live bytes. This applies
to numeric and missing-target indexes too.

Two near-limit builds also change the existing bounded-growth capacity.
For 54016 missing at 1 MiB, q2 requested bytes change 2,751,590 → 2,751,574,
retained delta 1,048,336 → 1,048,320, and peak 1,113,780 → 1,113,732.
The 32-byte larger header is offset by 48 fewer bytes of slot capacity.
For the generated 70,001-cell fixture, requested/retained deltas rise only
8 bytes while peak falls 24: one fewer 24-byte slot offsets the larger header.
Both still admit an index and keep the same post-build allocations and I/O.
These are allocator observations, not logical-budget headroom calculations;
the independent weighted reservation rules continue to enforce the limit.

Index-disabled and refused builds remain unchanged. Native-child RSS has no
median increase above 5% in the measured diagnostic matrix. The largest are
54016 first/file long, 4,008 → 4,200 KiB (+4.79%), and late/file short,
4,012 → 4,204 KiB (+4.79%); Simple owned short rises 4.69%. These process
changes are separate from the additional 32-byte field and do not establish
a memory or RSS bound.

## Larger follow-up controls

The four flagged groups receive a separate A/A and A/B/B/A capture with 1,000
fresh owners per leg and ten warmups: 24,000 owners and 192,000 queries.
Frozen binaries and probes are unchanged. Original captures remain intact;
absolute timings from these different shared-host windows are not pooled.
Raw records are losslessly compressed and independently audited.

The 54016 missing-target q1 doubling does not recur (baseline A2/A1 −1.06%).
Its open-plus-eight medians change −0.45%/+0.10%, while opening costs
+4.41%/+4.21% and q2 costs +1.34%/+1.72%. Eighth-query medians stay 120 ns.
The original apparent −22.89% workflow improvement is not accepted as a gain.
For 45365 first/file, q8 p99 changes −0.88%/+0.88%, and open p95 changes
+1.09%/+0.43%; the earlier large paired tail flags do not recur. Workflow
medians are −0.26%/+0.32%. For 45365 late/file, warm-mean p99 changes
−10.54%/−5.74%, with baseline A/A +8.91% and ABBA baseline drift +22.77%.

The synthetic owned case retains costs: opening medians rise +4.27%/+3.91%
and q2 medians +1.40%/+0.99%. A new open p99 flag is +41.36%/+8.44%
(12,960 → 18,320 ns and 13,611 → 14,760 ns), against A/A +11.78%.
Warm-mean p99 changes +15.61%/−5.33%, so its earlier large paired flag is
not reproduced in both legs. Workflow medians are −0.75%/−0.12%, with
workflow p99 −0.92%/−0.22%. These observations support no universal tail
improvement. See the packet's `followup-summary.md` and complete comparisons.

## Validation and disposition

Six owner/facade quality gates and DOC/PPT consumer tests pass: **4,400 tests
passed, zero failed, 27 existing ignored**. Six new tests cover nonzero SST
checkpoints, backward offsets, clones/reopened owners, concurrent cloned
handles, malformed tails, error precedence and checkpoint lifetime/accounting.
The initial lifecycle test retained an abandoned candidate's cache owner until
the final accounting assertion; explicitly dropping that test variable fixes
the assertion. Its failed run and differing test source remain archived in
`initial-checks/`; production behavior did not change for that correction.

All 126 real XLS fixtures on both source types and the full 70,001-cell
visitor/digest differential match. Repository boundary, claim, report,
coverage and non-iWork gates pass. Separate audits bind builds, source hashes,
raw captures, assembly, profiles, diagnostic routes and supplementary controls.
Independent source, test, performance and evidence review is recorded in
[the review record](results/change-0689/review.md).

Retain the narrowly scoped SST replay optimization. The substantial measured
SST-heavy query savings justify the logical 32-byte cost on every admitted
index and the measured 48-byte query stack increase. Opening/build costs,
near-limit slot capacity changes and tail uncertainty remain explicit.
Numeric/missing-built indexes have no SST checkpoint; backward SST requests
may restart traversal. Directory-name construction, remaining worksheet walks
and index construction are candidates for later measurement. No registered
performance claim, CRUD coverage promotion, device-cold, remote, cross-platform
or concurrency performance claim follows. The broader non-iWork GOAL remains
active.
