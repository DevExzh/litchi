# 0895 — selected-stream allocation profile

The rejected 0894 candidate moves the producer full-cell scan into an
allocation-heavy MCE stream path. Four fresh Heaptrack processes reproduce
identical counts within each role and locate the largest candidate allocation
groups in name expansion, owned name/value copies, and temporary attribute
collections. Production is unchanged; the 0894 rejection stands. This is
allocation attribution, with no new latency, RSS, or adoption claim.

## Protocol and qualification

The frozen [plan](results/change-0895/plan.json) binds the exact retained 0894
native baseline and rejected candidate binaries, their build/source/quality/
corpus evidence, the 0893 worksheet guard contract, seven 0895 analysis helpers,
and Heaptrack/heaptrack_print/taskset identities. Heaptrack 1.5.0 records four
fresh sequential processes in baseline/candidate/candidate/baseline order on
CPU 12, using the medium producer full-cell case, warm filesystem, zero
warmups, and 20 measured scans per process (80 total). The shared host was not
exclusive. Other batch agents stayed idle during capture; no new build ran.

A separate one-scan candidate preflight qualified symbols and export syntax;
it is excluded from the formal population. Independent
[methodology review](results/change-0895/methodology-review.json) approved the
frozen protocol before capture. All four children pass the strict existing
producer oracle and forced per-worksheet tolerance vectors. Offline
[verification](results/change-0895/verification.json) replays source/binary/
helper custody, exact commands, all semantic vectors, and reduction equality.

Heaptrack records the entire process, including corpus generation, setup,
scan work, post-timing worksheet guards, oracle checks, and destruction.
Instrumented report durations are not native timing evidence. Allocation-call
costs in collapsed stacks, the allocation-size histogram, printed totals,
and capture stderr agree exactly for every process. One allocation event per
process contains an unresolved frame; those events remain in the totals.
The separate temporary-allocation summaries are not used for reconciliation
(the preflight capture and printer differ by one temporary allocation).

## Whole-process allocation attribution

Each cell below is the allocation count in one 20-scan process. Both repeats
of a role give exactly these values. Categories assign each complete stack to
its rightmost recognized owner frame; they are disjoint call-path categories,
not native self time. `other` retains stacks outside the named owners.

| Closest recognized owner | Baseline | Candidate |
| --- | ---: | ---: |
| Legacy codec | 535,697 | 529,082 |
| MCE stream | 14,418 | 2,450,201 |
| Selected worksheet | 1,361 | 84,193 |
| x14ac | 9,792 | 17,694 |
| Other | 112,395 | 112,955 |
| **Total** | **673,663** | **3,194,125** |

MCE stream accounts for **76.71%** of candidate whole-process allocation calls
under this assignment, versus **2.14%** for baseline. Separate inclusive
presence counts overlap and must not be summed with these categories. Raw
profiles, complete collapsed stacks, histograms, and calls-ranked exports are
retained in [the evidence packet](results/change-0895/).

## Scan closure and post-scan guards

[Attribution](results/change-0895/attribution.json) further partitions stacks
by the benchmark scan closure, forced worksheet guard, remaining runner, and
remaining process. This is symbol-presence attribution, not profiler timestamp
segmentation. The closure is the source interval in
`tools/perf-baseline/src/producer_shape.rs:2520–2569`; endpoint observations
precede the forced worksheet guard at lines 2585 onward.

| Disjoint stack region | Baseline | Candidate |
| --- | ---: | ---: |
| Scan closure | 518,660 | 2,523,420 |
| Forced worksheet guard | 0 | 488,980 |
| Remaining runner | 847 | 847 |
| Remaining process | 154,156 | 180,878 |
| **Total** | **673,663** | **3,194,125** |

The scan-closure totals divided by 20 are **25,933** and **126,171**, exactly
matching the earlier 0894 per-scan allocator counters. This observed agreement
does not make whole-process Heaptrack totals equivalent to timed counters.
Zero baseline guard allocations means no allocation stack matched that guard;
it does not mean the guard was skipped. All report vectors prove it ran.

In the candidate scan closure, 2,508,360 calls (**99.40%**) have an MCE stream
frame. The nearest MCE frame divides that population into the following main
groups, including descendant work outside the MCE module:

| Nearest MCE frame | Calls across 20 scans | Share of all candidate scan-closure calls |
| --- | ---: | ---: |
| `clone_bounded_name_part` | 711,520 | 28.20% |
| `expand` | 546,800 | 21.67% |
| `Processor::parse_element` | 388,660 | 15.40% |
| `clone_bounded_string` | 338,840 | 13.43% |
| `clone_bounded_bytes` | 169,800 | 6.73% |
| `invoke_active` | 86,840 | 3.44% |
| `clone_raw_attributes` | 85,080 | 3.37% |
| `semantic_attrs` | 85,080 | 3.37% |
| `validate_duplicate_attributes` | 85,080 | 3.37% |

These are allocation sites/call paths, not estimates of removable work or of
the share of the latency regression caused by each function. Inlining and
symbol availability affect the nearest-frame boundary.

## Next bounded hypothesis and limits

The [source audit](results/change-0895/source-audit.md) maps these groups to
unchanged MCE/XML-name source and distinguishes intentional selected output,
x14ac matching state, raw/semantic frame lifetimes, and temporary event data.
The highest-priority small hypothesis is allocation-free QName validation and
prefix/local resolution in stream `expand`, avoiding construction of an owned
`QualifiedName` solely to inspect its parts. Owned expanded local names remain
a separate cost. Broader borrowed event views and attribute scratch reuse are
higher-risk follow-ups, not part of this batch.

Any candidate must preserve QName grammar, namespace/default-prefix semantics,
raw-before-active callbacks, exact typed errors and recovery, resource limits,
and first-refusal fallback. It must pass the existing MCE differential and
format/quality gates before a fresh native matched trial. A stream-only change
must be evaluated against restored current production first; the rejected
metadata route cannot be reinstated on allocation evidence alone. No savings
or adoption result is predicted here.

No source edits, new worktrees, build targets, or disposable executable copies
were created. Retained packet files are evidence, not temporary build scratch.
The broader OLE2/OOXML performance program remains active; iWork is excluded.
