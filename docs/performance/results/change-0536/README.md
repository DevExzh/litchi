# 0536: matched CFB collector cold-error-layout experiment

The previous turn made progress: 0535 mapped the existing collector loop,
committed its verified evidence, and removed its owned scratch. This batch
uses that attribution to test private cold error-formatting helpers without
changing successful chain validation or any reset/allocation/error order.

The frozen plan retains nine XLS scenarios and CFB tiny, many-small and
few-large. Native order is baseline r1, candidate r1, candidate r2, baseline
r2 with 20 warmups and 1,000 samples per case. Normal and allocator binaries
are separate. The allocation lane uses 3 warmups and 30 samples per case,
with all four guarded allocation vectors retained. Exact-constructor
Callgrind captures retain instruction positions, direct-callee costs and
collection-off call/jump metadata separately; setup dumps are excluded from
timed constructor totals. Native clocks are the wall-time evidence. Hardware
counters remain whole-child diagnostics and may be unavailable.

All four primary XLS workflows must improve p50 by at least 3% in both
paired repeats. XLS-owned constructor inclusive Ir and collector exclusive
Ir in XLS-owned/CFB few-large must decrease in both repeats. Allocation,
correctness and quality gates must pass; all matched adverse and same-build
variation flags over 5% remain individually reviewed. Any rejection restores
production and runs final quality. No speedup is claimed before admission.

Baseline source exactly matches 0535, allowing sealed reuse of final 0534's
14 quality gates and 4,382 Rust test executions after custody validation.
Candidate and any restored final state receive the applicable 14 checks.
The coordinator runs all Rust builds, tests and captures serially. Owned
storage is /home/zhuhe/litchi-goal-0536-target with a /tmp/litchi-goal-0536
alias. TMPDIR is owned from the first child and profiles disable vgdb.
Captures and precleanup verification passed; both owned build paths were
removed. Final postcleanup verification is recorded in verification.json.

OLE2 and OOXML remain first. ODF is deferred until their optimization goal
completes, and iWork is excluded. This batch does not complete the broader
CRUD, producer, cold/range-provider or scaling program.

## Measured disposition

The candidate is rejected and the runtime is restored exactly to the baseline.
Four of eight primary p50 comparisons miss the frozen 3% threshold, although
all eight improve. XLS-owned constructor inclusive Ir also increases slightly
in both repeats. Collector self Ir falls by only 0.004642% in XLS and
0.001795% in CFB few-large. The collector body shrinks from 1,436 to 1,238
bytes, while the normal executable grows by 2,608 bytes. These observations
do not support retaining this cold-helper implementation.

All 24 allocation pairs have identical allocation-call, reallocation-call,
allocated-byte and incremental-peak vectors. The comparison retains 32 matched
adverse flags and 60 same-build variation flags. Candidate quality passed all
14 gates with 4,382 test executions; restored final quality passed all 14 gates with 4,382 executions. Raw
profiles and instruction analysis remain bound to each measured stage.

All 96 receipt intervals passed serial and ABBA verification. Both owned
build paths and the coder draft directory were removed. Five tamper probes
passed. The final postcleanup receipt and SHA256SUMS bind the retained bundle.

Full postcleanup replay passed after correcting external report-output handling
in the two analyzers. Sealed internal writes and non-identical overwrites
remain refused; the focused replay regression probe verifies those boundaries.
Raw measurements and derived report contents remain unchanged.
