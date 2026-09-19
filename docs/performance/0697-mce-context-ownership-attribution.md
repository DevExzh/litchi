# 0697 — attribute remaining MCE context ownership

Status: completed attribution experiment; production unchanged.
`performance_claim: none`. Production revision: `1bf58ace2`.

This experiment measures the remaining MCE work after the empty namespace
installation guard in 0696. It combines a fresh native isolated-input baseline,
repeated hardware counters and instruction-level sample attribution with a
source review of temporary namespace ownership. It does not introduce a
production optimization or claim a speedup against earlier sessions.

The [packet](results/change-0697/README.md) retains the standalone locked probe,
raw measurements, source and corpus identities, native disassembly, annotated
instruction samples, independent audits and cleanup receipts. The
[next investigation](results/change-0697/next-investigation.md) defines the
smallest candidate and its semantic proof obligations.

## Scope and reproducibility

The corpus is the exact 18-call sequence from the retained 0693 real-deck
capture trace: presentation XML three times, slides 1–13 once each, then
presentation XML twice. It uses fourteen distinct byte owners from the
LibreOffice `slide-section-test.pptx` fixture. Complete PPTX source-tree
identity binds the historical call topology to the measured production;
the MCE implementation itself is current. Historical timings are not reused.

Sixteen cases comprise the complete sequence, presentation and slide groups,
and thirteen individual slides. Four native process legs each record 200
four-execution batches after ten warmup batches, reversing case order on odd
legs. Timers include default capabilities, MCE processing and output drop;
input loading, identity hashes, sample-buffer allocation and printing are
outside. These are warm measurements on CPU 12 of a shared host. The recorded
12,800 samples represent 51,200 sequence executions. Tail statistics describe
four-execution averages, not individual-call tails.

The host is Linux x86-64 on an AMD EPYC 9R45, using Rust/Cargo 1.95.0 and
perf 7.0.14. The probe is a release build with debug information and the default
system allocator. Exact environment, lockfiles, commands and binary/source
hashes are retained; no allocator instrumentation is inside the native runs.

Hardware counters use separate processes, three repetitions of 10 and 210
samples at ten executions per sample. Subtraction divided by 2,000 estimates
per-sequence work while removing fixed startup. Sampling and annotation use
another process; no perf-instrumented timing is used as native latency.

All group/member ratios are kernel attribution ratios, not additive capture fractions.
Package open, non-MCE consumers, edits, publication, saving and I/O are absent.
Intervening consumer allocation lifetimes are absent too. No full-workflow,
cold-cache, allocation, RSS, concurrency or interoperability claim follows.
Instruction samples may skid; source and inline attribution are imperfect.
An atomic instruction in native code does not by itself establish contention,
and a sample percentage is not an achievable speedup estimate.


## Fresh native results

The complete isolated sequence takes **2.370–2.392 ms** across four leg medians.
Presentation work remains 3.32% of the median-of-leg-medians; slide 11 contributes
42.99%. The maximum case-level inter-leg median spread is 2.48%. These are
observed ranges, not population confidence intervals. No A/B comparison with
0695 is made: production changed between the sessions, and these measurements
were designed to attribute the current implementation, not estimate that change.

| Case | Calls per sequence | Input bytes | Median range (µs) |
| --- | ---: | ---: | ---: |
| all | 18 | 280,963 | 2370.097–2392.310 |
| presentation | 5 | 11,785 | 78.695–79.583 |
| slides | 13 | 269,178 | 2268.491–2308.692 |
| slide1 | 1 | 2,589 | 20.514–20.934 |
| slide2 | 1 | 17,006 | 140.506–141.728 |
| slide3 | 1 | 3,038 | 24.203–24.471 |
| slide4 | 1 | 31,153 | 258.786–260.752 |
| slide5 | 1 | 17,833 | 149.787–150.643 |
| slide6 | 1 | 3,037 | 24.168–24.360 |
| slide7 | 1 | 28,449 | 237.596–238.956 |
| slide8 | 1 | 2,674 | 21.191–21.357 |
| slide9 | 1 | 2,676 | 21.190–21.715 |
| slide10 | 1 | 3,037 | 24.180–24.313 |
| slide11 | 1 | 121,498 | 1017.067–1027.331 |
| slide12 | 1 | 29,160 | 249.635–251.204 |
| slide13 | 1 | 7,028 | 57.028–57.552 |

The three full-sequence counter slopes are 10.634, 10.894 and 10.728 million
cycles, and 49.656, 49.654 and 49.598 million instructions. IPC is 4.56–4.67.
Presentation/full ratios are 3.30–3.38% of cycles and 3.51–3.52% of instructions.
All seven requested events ran at least 99% of their enabled duration; each
raw event and repetition is retained in `counter-summary.json` and its inputs.
The independently parsed source sequence contains 10,072 start events and
23 starts with namespace declarations. This is a source-syntax census, not
production branch instrumentation.

## Instruction attribution and decision

Fresh self shares are **26.52% for `start`, 7.16% for `Inherited` destruction,
and 4.86% for `Ctx` destruction**. This isolated MCE denominator differs from
0696's open-plus-capture profile and its percentages must not be compared as
whole-workflow improvements. There were no lost samples reported.

The frozen binary separates the temporary clones from the owners retained by
`Ctx` and the child frame. The table records samples at the branch immediately
after each listed atomic instruction. The sample locations support a source
hypothesis; they do not measure the exact cost of the preceding atomic.

| Operation | Atomic address | Following branch address | Samples at branch |
| --- | --- | --- | ---: |
| Parent `Ctx` namespace clone | `0x4311a` | `0x4311e` | 205 |
| Parent `Ctx` directive clone | `0x43131` | `0x43135` | 0 |
| Temporary `Inherited.ns` clone | `0x43168` | `0x4316c` | 62 |
| Temporary `Inherited.emitted` clone | `0x432fb` | `0x432ff` | 59 |
| Ordinary-path child emitted-boundary owner | `0x45b2e` | `0x45b32` | 213 |
| Temporary `Inherited.ns` release | `0x4767f` | `0x47683` | 132 |
| Temporary `Inherited.emitted` release | `0x47697` | `0x4769b` | 43 |
| `Ctx` namespace release | `0x4746f` | `0x47473` | 119 |

`start` has 649 instruction samples; the two temporary-clone following branches
account for 121 of them. All 175 `Inherited` destructor samples occur at its
two following branches. All 119 `Ctx` destructor samples occur at its namespace
release's following branch. These are counts from one sampling run, not dynamic
instruction execution counts. Zero samples do not prove a path never executes.
The machine-readable observations and the raw assembly/annotations are retained.
The `start` stack reservation remains `0x598` in this build; this is not a
whole-parser stack bound.

The next candidate is a stack-scoped borrowed `Inherited` view. Its two values
are observations of owners already held by the parent frame. Borrowing should
remove those two temporary clone/drop pairs when present, while retaining the
child frame's necessary owner clone. The review requires finishing parent
`AlternateContent` mutation before taking the borrowed views, then constructing
an owned child frame before `close` can mutate or reallocate the frame vector.
Borrowing from the child context is incorrect because local declarations can
replace that context's namespace head.

This is a recommendation for a separate source-bound A/B experiment, not an
implemented optimization or a promised gain. The larger `Ctx`/`Frame` redesign
remains separate: those owners survive to later child/sibling/end events, and
removing them needs broader invariants. Required evidence includes exact output,
whole Report, ownership and refusal parity across the shared oracle and existing
controls, declaration-heavy controls, native edit latency, tails, allocations,
code/stack size and applicable integration gates.

## Validation and limitations

Only diagnostic probe/driver sources and performance documentation change.
Production sources remain byte-identical to the bound revision. All accepted
ADR and goal hashes were checked against the previously read 33-file receipt.
The independent audit checks the historical corpus/topology, current production
and probe bindings, all raw native samples, repeated counters, deterministic
summaries and instruction commands/bytes. No new production behavior or public
API requires testing in this batch; probe formatting, warning-denied Clippy and
rustdoc accompany the six applicable repository evidence gates.

The first locked build refused the copied probe's stale package name; the
corrected own-package name leaves dependency versions unchanged. The initial
symbol-name-based objdump command exited successfully but produced no selected
instructions. The instruction audit caught the empty output. The final driver
uses the symbol table's exact address and size; both initial and final receipts
are retained. No timing run or production source changed for that correction.
Raw perf data and the compiled target are owned scratch; cleanup preserves their
hashes, textual evidence and the root workspace lock.

All nine fresh validation gates and the pre-cleanup audit passed. The final
seal rechecks the audit and four documentation gates after exact scratch cleanup.
The independent [instruction review](results/change-0697/instruction-review.md)
supports the proposed narrow experiment without making a speedup claim.
