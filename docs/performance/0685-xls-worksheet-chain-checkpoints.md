# 0685 — immutable worksheet-start chain checkpoints

Status: final validation passed; scoped retention reviewed below. `performance_claim: none`;
no registry or CRUD coverage promotion. OLE2/OOXML remain active; iWork is excluded.

## Mechanism and identity proof

The retained occurrence index from [0684](0684-xls-occurrence-query-index.md)
eliminates worksheet scanning on warm selected queries, but every replay still
walked the CFB allocation chain from the beginning of the Workbook stream.
A fresh baseline profile on 54016 attributes 67.64% self samples to
`litchi_cfb::shared::next_chain_sector`. The next change removes only the
already-validated prefix before the selected worksheet.

CFB exposes an opaque `StreamChainCheckpoint`, captured from a borrowed
`StreamChainHint`. It contains a weak reference to the immutable parsed-index
allocation and the existing private validated chain position. Restoring it
compares the weak and current strong allocation pointers without dereferencing
either; a foreign or expired identity
produces a cold hint. The weak control block prevents allocation-identity reuse
while a checkpoint exists, without keeping the reader, source bytes, or parsed
allocation tables alive. Capture and restore allocate nothing and take no lock.

Every current SharedOleFile constructor creates its own private parsed index;
there is no reader clone or public index-adoption path. Thus index identity
identifies the reader today. A future API that shares an index between readers
must preserve this boundary explicitly, for example with a private per-reader
identity token. Equality of source version or bytes is not a substitute.

The existing stream SID, allocation-table, first-sector and backward-offset
checks remain in the ordinary hinted read methods. Links after the checkpoint
still pass through the same chain walker and checks. Source reads retain their
existing freshness fences and error identities. No physical sector identifier
is exposed through ordinary document APIs.

XLS captures the worksheet-start checkpoint immediately after its existing
successful seek, while building an optional occurrence index. It publishes the
checkpoint only when the complete worksheet scan and final fences succeed.
Failed/cancelled candidates discard it. Every indexed query restores a fresh
local hint, so there is no shared mutable cursor, I/O lock, query-order
assumption, or last-query cache. All indexed occurrences lie at or beyond the
fixed worksheet start and still replay in source order.

Only admitted worksheet indexes retain checkpoints. Zero-budget and ordinary
first-query paths, visitors, and text extraction keep their previous behavior.
The SST resolver remains independent; no string value, SST window or last-read
position is retained across queries. Within-sheet chain traversal remains.

## Resource boundary

The checkpoint adds no allocation and does not pin CFB metadata. Its inline
fields live inside the existing managed, pinned worksheet-index value.
`INDEX_OVERHEAD` rises from 128 to 192 logical bytes; the layout assertion
covers the index, Arc/cache metadata and LRU linkage. The fixed cache-table
overhead remains 128. Slot/vector capacity accounting, hierarchical candidate
reservations, concurrent builder admission, losing publication, eviction and
last-pin release remain unchanged. This conservative weight is not exact
allocator or RSS accounting.

## Evidence

The [packet](results/change-0685/README.md) binds baseline `c207f2c43`, both
CFB and XLS sources, unchanged goal/ADRs, and all 126 XLS fixtures. It reuses
0684's already-committed native, counting, allocation and repeat probes.
Twelve cases add late targets in 54016, 45365-2 and Plan1 to the prior first,
missing, tiny and refusal cases, so a worksheet-prefix improvement is not
presented as constant-time access to every cell.

### Final paired observations

Native A/A and A/B/B/A cover 48 route groups and 8,640 records. Twelve
cases include first, late, missing, tiny and formula-refusal queries on owned
and file sources. Warm query p50 changes against the 0684 cache baseline:

| Stored target | Owned | File, warm OS cache |
|---|---:|---:|
| 54016 first | −22.55% / −26.70% | −15.28% / −15.55% |
| 54016 late | −19.72% / −19.01% | −12.07% / −12.47% |
| Plan1 first | −21.76% / −22.35% | −9.90% / −8.50% |
| Plan1 late | −31.87% / −32.72% | −12.01% / −10.40% |
| 45365-2 first | −7.41% / −5.69% | −3.31% / −3.13% |
| 45365-2 late | −1.09% / −2.63% | +1.13% / +1.61% |

The separate 200-sample-per-leg owned control confirms 54016 warm query
2,740/2,735 → 2,060/2,040 ns (−24.82%/−25.41%), while open plus three
queries improves only 1.20%/1.08%. Plan1 warm improvement is 20.16%/19.38%;
its open-plus-three window remains approximately flat. Neither the first-query
54016 improvement nor open-time variation is attributed to checkpoint reuse:
those routes do not restore a checkpoint.

### Regressions and limits

[Every paired >5% p50 trigger](results/change-0685/regressions.md) is retained,
including phases and three-query totals. Tiny owned queries regress roughly
8–16%; the longer control reproduces missing-cell first query 440 → 510 ns,
build 825/830 → 920/900 ns and warm missing lookup 80 → 90 ns. Tiny missing
open plus three queries rises 7,610/7,620 → 8,040/8,010 ns (+5.65%/+5.12%);
tiny stored rises 4.74%/4.56%. Formula refusals regress 5.41–10.55% per query
across file/owned routes, and the owned refusal visitor regresses 5.26%/6.16%.
These are accepted costs of retaining this scoped checkpoint optimization;
they are not noise exclusions or claims that ordinary first queries improved.
The cause of cold/refusal shifts has not been isolated. Removing two temporary
weak-count atomics from the first candidate did not remove these regressions;
that candidate and its full evidence remain archived.

A longer visitor control has +9.00% in one leg and −0.88% in the other, with
8.91% candidate-leg drift; it does not establish a stable visitor improvement.
All means, p95/p99 and control spreads remain in the machine-readable
comparisons. Thirty-sample p99 is a sample maximum, and outliers prevent a
uniform tail-latency claim (54016 late/file candidate B2 maximum is 18,570 ns
versus A2 13,360 ns).

### Work, allocation and memory

The separate repeat-process diagnostic subtracts N=10 from N=1,010, repeated
three times. Median extra-query instructions decrease 46.28%/35.90% for
54016 owned/file and 38.29%/22.81% for Plan1; tiny instructions stay within
0.33% of baseline. These loop diagnostics exclude the fresh-owner query shape
and do not cancel its tiny regressions. The same two-million-query profile
reduces chain-sector self share from 67.64% to 53.07%. Sampling covers the
whole process; user-space symbols are available, kernel symbols are restricted.
Within-sheet and separate SST chain walks remain material.

All 88 allocation groups agree across their three repeats. Building an index
adds exactly 32 requested and retained bytes, with unchanged allocation count
and measured peak; other measured phases are unchanged. For example, 54016
retains 1,572,948 → 1,572,980 bytes after its build. This differs deliberately
from the conservative +64 logical cache charge. Process peak RSS has no stable
increase across the larger cases: the Plan1 file N=10 median rises
2,772 → 2,936 KiB (+5.92%), while N=1,010 is 2,796 → 2,792 KiB. This flag
is disclosed without attributing process noise to 32 bytes or claiming an RSS
bound. Full process measurements retain both lengths and all repeats.

### Validation and disposition

All six final quality gates pass: formatting, owner check, warning-denied
Clippy, 1,870 CFB/XLS tests (two existing ignored), 61 facade tests and
warning-denied rustdoc. Ten new CFB tests cover identity/lifetime, chain and
source errors, FAT/MiniFAT and concurrency; two XLS integration tests exercise
nonmonotonic coordinates and concurrent handles. The 126-fixture differential
matches for owned/file sources; all 24 counted routes preserve source reads,
bytes, freshness checks and outcomes exactly. Dependency, strict/structural
claim, report classification, coverage and non-iWork gates pass.

Disposition: retain after independent source and performance review. Repeated
selected-cell queries on worksheets with a meaningful CFB prefix show useful,
reproduced reductions in latency and instructions. The tiny/refusal costs are
real candidate overhead and remain a follow-up priority; no gain is claimed
for opening, first/build queries, visitors, refusals, tiny sheets, arbitrary
random access or XLS generally. This is a scoped measured optimization, with
`performance_claim: none`, not a universal tradeoff-free improvement. The
[review](results/change-0685/review.md) records this decision.
Physical cold-cache, remote latency, parallel throughput and cross-platform
claims remain outside this batch; no broad Office-performance completion claim
follows these checks. Tiny-query/refusal costs, admission retries, SST and
within-sheet walks remain explicit follow-up work.
