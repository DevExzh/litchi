# 0958 — XLSX edit/commit allocation attribution

Six current-source heaptrack captures confirm that cloning omitted numeric
cells allocates once per retained numeric cell. This accounts for **8.28–8.43%**
of selected edit/commit allocation calls, but only **0.28–0.38%** of requested
bytes. This batch records source attribution, not a production optimization,
native speedup, allocation reduction, or completion of the performance program.

## Capture and custody

The retained, uninstrumented 0957 native benchmark was run in six fresh serial
processes: medium one-percent, dense-sparse one-percent, medium one-edit, then
the reverse order. Each used one worker, zero warmups, three measured samples,
CPU 12, a 900-second deadline, no retries and stop on failure. All six passed:
18 measured samples with the existing exact source/output/semantic hashes,
inverse restoration, finite resources and sink/task accounting. Setup,
preflight and verification also appear in the whole-process profiles.

The binary is 43,259,328 bytes, SHA-256
`5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3`.
The [bounded packet](results/change-0958/README.md) binds its retained source,
build commands, fixture oracles, capture plan and analysis. Relevant production
source is unchanged from 0957. No Rust source was edited or rebuilt here.
Heaptrack 1.5.0 records requested allocation bytes, not copied bytes, live heap
or RSS. Profiled timings and GNU time RSS are diagnostic only.

The capture manifest was sealed before offline decoding: 57 files totaling
92,466,528 bytes, SHA-256
`bd30035f2c7926ec579fea58ecd1e5c5fe34da6b2ae129473bee033630e5c465`.
Large raw, decoded and folded artifacts remain outside the checkout under
`/home/zhuhe/litchi-performance-0958`; committed manifests bind names, sizes
and SHA-256 values. Source/build reconstruction and replay commands are in the
packet. No executable or file of 1 MiB or more is committed.

## Attribution boundary and conservation

Selected rows require a resolved `benchmark::timed_iteration` owner in the
exact benchmark module and either `benchmark::stage_timed_commit` or the
nested `MultiSourceEdit::commit` callsite. The latter is a **commit-core subset
surrogate**, not independent evidence for the entire timed phase. No unknown
frame is admitted. Both disjoint resolution buckets remain explicit below.
Their union exactly matches the separate 0957 allocator phase counters in every
capture; that correspondence corroborates coverage but does not turn a
surrogate stack into a directly resolved stage stack.

All counts below are per iteration, dividing each three-sample capture by
three. Both repeats agree exactly for selected costs.

| Case | Exact stage calls / requested bytes | Commit-core subset calls / requested bytes | Combined selected calls / requested bytes |
| --- | ---: | ---: | ---: |
| Medium one-percent | 68,212 / 14,041,541 | 41,912 / 3,125,139 | 110,124 / 17,166,680 |
| Dense-sparse one-percent | 130,114 / 18,769,492 | 79,304 / 5,834,581 | 209,418 / 24,604,073 |
| Medium one-edit | 17,152 / 3,224,904 | 10,175 / 724,411 | 27,327 / 3,949,315 |

Across all six processes, whole-process raw totals are 13,032,599 calls and
3,406,101,409 requested bytes; selected totals are 2,081,214 calls and
274,320,408 bytes. Selected plus unqualified rows equals raw totals in every
run. Decoder size/trace aggregates and every heaptrack histogram bucket agree.
Folded stacks agree on calls; they contain no requested-byte column, so byte
conservation against folded output is explicitly unavailable. Each profile
retains one unknown, unqualified allocation of 73,728 bytes. These allocations
and all setup/verification stacks remain in whole-process accounting.

## Numeric ownership and competing allocation mechanisms

The resolved path is `Number::clone` through `Value`, `Cell`, `Stored` and
`Store::merge_omitted_cells`, reached from rewritten-source snapshot creation.
The source census predicts every observed omitted-cell clone:

| Case | Omitted numeric clones | Requested lexical bytes | Selected calls share | Selected requested-byte share |
| --- | ---: | ---: | ---: | ---: |
| Medium one-percent | 9,123 | 58,723 | 8.28% | 0.34% |
| Dense-sparse one-percent | 17,614 | 92,907 | 8.41% | 0.38% |
| Medium one-edit | 2,303 | 10,933 | 8.43% | 0.28% |

These are the `stored_clone` intersection, not every numeric construction or
clone in the process. Source numeric owners number 9,216, 17,792 and 2,304;
updates replace 93, 178 and 1 respectively. Existing snapshots share their
`Arc<Store>`; ordinary snapshot cloning is not the cell-clone source.

The additive leaf-mechanism partition prevents overlapping stack tags from
being added together. For medium one-percent, per iteration:

| Exclusive mechanism | Calls | Requested bytes |
| --- | ---: | ---: |
| Reallocation/growth | 42,561 | 12,324,253 |
| Other selected owner | 947 | 2,922,651 |
| Merge leaf | 4 | 1,695,744 |
| Copy leaf | 47,598 | 90,483 |
| Clone leaf | 9,705 | 73,636 |
| Numeric construction leaf | 9,309 | 59,913 |
| Total | 110,124 | 17,166,680 |

These labels classify allocating stack mechanisms; they are not exclusive
semantic subsystems or a timing ranking. Requested reallocation sizes are not
bytes copied. The complete three-case partitions and source callsites remain
in the machine-readable analysis. Number/clone/merge/copy/reallocation source
tags and their intersections overlap and must never be summed. In particular,
the omitted numeric clone cost is already included in the clone leaf row.

## Decision and next gate

1. Numeric lexical ownership is now a measured, bounded candidate for removing
   allocation calls. Compare shared and short inline storage only after
   measuring `Number`, `Value`, `Cell` and `Stored` layouts. Preserve exact
   lexical forms, finite-number validation, allocation failure handling and
   the public `Number::new(impl Into<Box<str>>)` contract. Shared storage may
   add construction/control-block costs; inline storage may enlarge objects.
   Neither tradeoff is established by this capture.
2. Growth and merge backing storage dominate requested bytes. Examine their
   concrete capacity/lifetime owners before changing layout or indexes. A
   broad store overlay remains an unqualified 0587 proposal. The 0556 linear
   merge pilot stopped at unstable baseline qualification in 0557; it is not
   a measured rejected candidate or an implemented overlay.
3. Attribute duplicate-check allocations remain a historical trap. The source
   audit connects a singleton shortcut to the rejected 0797 strategy; the
   later 0805 microbenchmark improvement did not yield eligible 0806 public
   workflow benefit. Allocation counts alone do not justify reviving it.

Any candidate needs before/after ordinary native latency and tail measurements,
allocator/peak evidence, one-edit and dense controls, and preservation,
source-freshness, inverse, malformed-input and resource checks. Complete
rewritten XML validation and canonical fallback must remain. No ambient pool,
parallelism policy, public API, compression policy or preservation rule changes.
The 0956 width-four RSS, 0957 observer-peak and earlier ordinary one-edit tail
flags remain open. iWork remains excluded.

## Validation and analysis history

Prelaunch runner review corrected executing-file identity, source/build
binding, exact report admission, terminal-before-parse ordering, interruption
cleanup and fail-stop behavior. Plans v2–v4 were never launched; v4 failed its
own revision identity check. Only v5 ran the six children. Offline controls
cover plan admission, wrong runner identity, fail-stop at each child index,
interruption cleanup and zero-return launch errors. Timeout handling exists
in the runner but is not claimed as a separately exercised control.

The first decoder launch refused an existing empty output directory before
processing; its failure receipt is retained. The subsequent decode passed
all six profiles. Attribution v1 failed on space-delimited folded rows; v2
corrected the parser, and v3 corrected source classification. Earlier outputs
are retained as superseded history, not accepted evidence. The final analyzer,
controls and manifests are bound by the packet. Independent review confirms
conservation, owner boundaries, source census and counter correspondence.
No Rust check suite was rerun because this batch changes evidence and offline
tooling only; capture oracles and analyzer controls are the applicable checks.
