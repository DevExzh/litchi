# 0715 DOCX section publication collection

Status: **rejected**, with 46 of 48 hard gates passing. The first NumberedList
counting-publication pair regresses 42.30% in median and 41.32% in mean.
All production changes and added tests have been restored to the exact baseline.
The candidate patch and all failed measurements remain archived.

The initial four current-source Callgrind profiles select `write_plain` by
its positive incoming `ordinary_save::run_case` edge. Generated documents
retain five setup dumps and one measured dump; NumberedList retains four
setup dumps and one measured dump. Every numbered dump and the zero-Ir
termination dump remain in this packet.

The paired pilot tests fusion of section-part and explicit-relationship
collection. `pilot-plan.json` and `pilot-freeze.json` were frozen before A1;
A1 runs on baseline source before the implementation. Later stages run with
candidate source checked out, using exact frozen baseline/candidate binaries.
Each stage has native and allocator children for generated/NumberedList
counting publication and lifecycle: 32 children in ABBA order. Native uses
100 samples and ten warmups; allocation instrumentation uses three samples
and no warmup. Source, fixture, argv, environment and output hashes bind each
child. CPU 12 and the owned filesystem root are explicit.

The primary threshold is at least 10% generated counting-publication p50 and
mean improvement in both pairs. NumberedList counting publication and both
lifecycle corpora may regress by at most 3% in p50/mean. Allocation requests
and requested bytes have the same 3% nonregression limit. Every over-5% tail
or repeat spread remains a review flag. Instrumented elapsed values are not
native latency. Independently sampled phase quantiles are not additive.

The separately frozen post-acceptance mechanism plan is deferred because the
pilot failed. No candidate profiles or RSS children were captured. The four
initial profiles remain baseline evidence only. None of these tools supplies
hardware counters, cold-cache, scaling or native Office compatibility evidence.

The archived candidate changes three private DOCX implementation files.
Eight differential tests use the original two-pass logic as a test-only
reference. Candidate checks pass 1,461 all-feature/all-target tests,
warning-denied all-target Clippy, 75 doctests and rustdoc; 31 doctests are
ignored. Six repository evidence gates passed on that candidate. Exact commands
and outcomes are in `quality.json` and `evidence/`.

The restored baseline reuses exact-source 0713 verification, bound in
`reused-verification.json`: 4,995 tests, 92 passed/46 ignored doctests and seven
baseline-proven PPTX/XLSB test Clippy violations. Candidate checks do not
replace that final-source scope.

The final audit replays baseline profiles, all 32 pilot children and the
sample-order diagnostic after removal of the three owned scratch roots,
using exact binary cleanup witnesses. The diagnostic establishes a sustained
B1 slowdown without identifying its cause. Profile and pilot corruption
checks verify refusal of altered evidence. The manifest covers every packet
file except itself. The unrelated API design file is outside this batch.

```sh
python3 -B docs/performance/results/change-0715/audit.py
python3 -B docs/performance/results/change-0715/artifact-seal.py --check
```
