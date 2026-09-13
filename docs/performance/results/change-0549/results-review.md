# 0549 candidate results review

The 0549 `CheckedBitSet::test_and_mark` candidate is rejected for production
adoption and retained as measurement-only evidence. Every one of the eight
required primary native p50 comparisons regresses. The smallest regression is
`+2.079931%`, which is already short of the required `3%` improvement, and the
largest is `+10.876280%`. Allocation and public guard gates pass, and the
required Callgrind Ir directions pass, but those results cannot offset the
primary native latency failure.

The candidate quality stage separately passed all 15 checks with 4,388 test
executions. That count belongs to the candidate stage; the final quality stage
is a separate restoration check and must be taken from its final summary. The
coordinator has restored production to a source manifest byte-identical to the
measured baseline.

## Evidence bindings

| Artifact | SHA-256 |
| --- | --- |
| `comparison.json` | `d64357af42deea745a0074a16aa0683a6c98734152643dbaeda41bbc79f49efe` |
| `guard-analysis.json` | `a2a32d36148fe07798f4e6db0b5ac71a517a167ae90b69a60a0cd75d4987c765` |
| `profile-comparison.json` | `b4560164ecbe02e98acbfb47946c9229f8bea73d67e9ea643f1eb48051762956` |
| `instruction-analysis-comparison.json` | `59e85a64dc4f3cb1d314b83d142a1108211d1a9db96e47fa9c1488eab617c9dc` |
| `adverse-review.json` | `e81a0ed9549367b337aa41ba8ba9d902b71fdfdccc4030535118d60cd9833421` |
| candidate source `candidate-sources/file.rs` | `711572fd5779aff9efdfd3e452040daa378e23aad22bb6ecb83841dbbe903377` |
| `candidate.patch` | `0602a50efe65c800ea36163d22af19d4f44dac9a77f4d87861230956a23fb40d` |
| `candidate-sources/source-note.md` | `135628353a35fd53cbb72b91914530454bac279796bdda0ee152727489929660` |
| `source-review.md` | `133d09e6a19ebf3ea9d904eb72631de3dc1383719b7375ce4772d1e32c9ab756` |

The candidate is the one-file source snapshot based on repository revision
`6d9fbb7401729aacf2afc5d6b6c681a9e7384056`. The source review binds the
unchanged bounds and error precedence, duplicate write behavior, allocation
order, reset and ownership lifetimes, and unchanged callers of
`CheckedBitSet::insert`. This results review did not build, test, capture, or
edit Rust source.

## Primary native gate

The values below are p50 elapsed nanoseconds for the fixed
`256-comments-opaque-heavy` XLS shape. A positive change is slower. The gate
requires at least a 3% improvement in both repeats for every listed workflow.

| Workflow | Baseline R1 | Candidate R1 | Change | Baseline R2 | Candidate R2 | Change |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xls_source_backed_open` | 101,786 | 109,010 | +7.097243% | 104,090 | 106,255 | +2.079931% |
| `xls_source_backed_open_one_cell` | 103,665 | 112,221 | +8.253509% | 105,161 | 110,480 | +5.057959% |
| `xls_owned_source_open` | 92,265 | 102,300 | +10.876280% | 93,715 | 100,325 | +7.053300% |
| `xls_owned_source_open_one_cell` | 94,980 | 103,630 | +9.107180% | 95,675 | 101,440 | +6.025608% |

All eight rows therefore fail the native primary requirement. The JSON review
retains and individually explains all 87 matched >5% latency/RSS rows. The
other 46 main rows above the absolute 5% threshold are same-build repeat
variations: 21 baseline rows and 25 candidate rows. They are retained as drift
diagnostics and do not receive candidate credit or blame.

## Other admission evidence

Allocation calls, allocated bytes, and incremental peak live bytes are equal
between stages across the retained allocation rows, so the allocation gate
passes. The malformed-input guard analysis also passes its aggregate envelope:
the maximum same-invalid ratio is `1.311728`, below `4.0`, and the maximum
baseline-valid ratio is `1.071459`, below `2.0`.

The guard review covers all 77 adverse rows and all 30 same-build drift rows.
Those timings cover only `OleFile::open`; oracle checks and returned-value
destruction are outside the clock. They do not support an allocation or RSS claim. Each row remains bound to its
exact case, size, repeat, metric, and values in `adverse-review.json`.

The required profile rows all move in the required direction in both repeats:

| Profile metric | R1 baseline → candidate | R2 baseline → candidate |
| --- | ---: | ---: |
| XLS-owned constructor inclusive Ir | 11,317,748 → 11,152,433 (-1.460670%) | 11,315,393 → 11,155,279 (-1.415011%) |
| XLS-owned `collect_exact` exclusive self Ir | 5,601,140 → 5,436,525 (-2.938955%) | 5,601,140 → 5,436,525 (-2.938955%) |
| CFB few-large `collect_exact` exclusive self Ir | 5,571,945 → 5,408,115 (-2.940266%) | 5,571,945 → 5,408,115 (-2.940266%) |

All profile `delta_percent` values were checked for absolute changes above 5%:
there are zero such profile diagnostic rows and zero same-build profile drift
rows. Callgrind Ir is a mechanism diagnostic; its decrease does not establish
a native wall-time gain.

The hardware lane contains measured, identity-matching grouped counters, but
the command measures the whole child, including setup, copies, queries,
oracles, drop, and report construction. One matched diagnostic crosses the
absolute 5% threshold: context switches fall from 82 to 71 in repeat 1
(`-13.414634%`). One same-build diagnostic crosses it: baseline repeat 1 to 2
falls from 82 to 70 (`-14.634146%`). No other event, IPC, or branch-miss
diagnostic crosses 5%. These observations are scheduling/workload diagnostics
and provide no operation-local speedup or regression-cause claim.

## What the measured assembly shows

The two `assembly-2` receipts map the same
`SectorChainScratch::collect_exact` symbol at function base `0x2f223c0`.
The baseline loop loads the visited word, uses register `bt` at function
offset `+0x1ee` (`0x2f225ae`), and then uses memory `or` at `+0x1fd`
(`0x2f225bd`). The candidate loads the word at `+0x1f3`, copies it, emits
register `bts %r15,%r10` at `+0x1fa` (`0x2f225ba`), stores the register at
`+0x1fe` (`0x2f225be`), and emits a separate register `bt %r15,%r9` at
`+0x202` (`0x2f225c2`), followed by the duplicate branch at `+0x206`.

The candidate instruction row count is 309 versus 313 for the baseline, and
the XLS-owned function self Ir is 1,087,305 versus 1,120,228. The candidate
`bts` is a register operation after an explicit load; this receipt contains no
memory BTS and no single collapsed carry-returning test-and-set path. Fewer
instruction rows or lower Ir do not prove faster wall time. No opcode latency,
uop, cache, or other speculative assembly-causality explanation is asserted;
the native p50 measurements are the deciding evidence.

The assembly artifacts are bound in the JSON review, including both
`assembly-index.json` files, both `assembly-2.stdout` dumps, and both
`assembly-2.receipt.json` files. The unchanged `insert` symbol remains present
in both instruction reports, while the candidate's combined helper is inlined
into the measured collector symbol.

## Disposition and direction

The final disposition is **rejected / measurement-only**. Keep the exact
restored baseline for this batch. OLE2 and OOXML remain the active optimization
priority; ODF work can wait. A future candidate should isolate the collector
mechanism and repeat the full native, profile, allocation, guard, correctness,
and quality gates before any assembly or Ir reduction is treated as a
performance improvement.
