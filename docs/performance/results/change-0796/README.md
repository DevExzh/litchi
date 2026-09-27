# 0796 — direct XML attribute construction and consumption

This is a diagnostic follow-up to the rejected 0794 iterator and the 0795
instruction profiles. Production remains unchanged. No helper timing or guest
counter result can adopt that candidate or replace its failed public-workflow
and cross-format gates.

One independent evidence binary compiles exact local copies of the baseline
and rejected OPC helper. Both use the same dependency resolution and release
configuration. The 31 fixed inputs cover distinct counts 0/1/4/5/8/9/16/17/32/33/64;
valid duplicates at 1/4/5/32/33; quoted, unterminated, and unquoted duplicates
at 1/33; and three syntax-error tails after 0/4/33 names. Long values have 4,096
bytes. Construction creates, exposes through black_box, and drops an iterator
without advancing it. Consumption creates, consumes to first error/end, and
drops it. Construction’s opaque observation deliberately forces iterator state
to remain observable; it does not model every optimization available to a real
inlined caller. Iterator size is reported separately.

The native lane uses six alternating source-leg blocks, 30 measured samples
after three warmups, and 4,096 identical-input iterations per sample. Each
case/mode/leg/block uses a fresh process on CPU 12. Elapsed values describe a
whole repeated batch, not a single isolated call; division by 4,096 is a derived
average. These are hot repeated-input microbenchmarks, not Office-workflow
measurements. The nearest-rank process p50 feeds six paired ratios. The median
ratio bootstrap uses seed 796079, 10,000 resamples, and endpoints 250/9749.
A ratio above 1.05 with interval low above one is a diagnostic flag only.

The Callgrind lane uses two alternating source-leg repeats, one sample without
warmup, and one iteration. Exact non-inlined source-leg/mode owners delimit the
counted region. Collection starts off, is zeroed/toggled at entry, and dumps at
exit. Ir/Bc/Bcm/Bi/Bim counts must conserve, with one positive dump and an empty
termination dump. Simulated prediction counts are not hardware predictors or
CPU-cycle shares. Work before collection may affect simulator state; this is
not a cold-predictor claim. Function attribution depends on compiler inlining;
nested inclusive costs overlap.

The matrix has 744 native reports/22,320 samples and 248 profile reports/samples:
992 reports and 22,568 samples total. Input generation, exact item/error oracles,
clone/fusion checks, and JSON serialization remain outside measured owners.
Semantic preflight compares both implementations with quick-xml up to its first
error, including byte positions and values. Checksums validate equal consumed
work without allocating or formatting errors in the measured region.

All build, quality, native, and profiler work is root-only and serial. Source
review and offline analysis are delegated independently. Fresh probe checks are
separate from inherited exact-source production correctness evidence. Full
recapture needs a new packet and owned target; drivers refuse to overwrite
retained output. The owned target is removed only after all captures finish and
its executable identity is recorded. Unrelated working-tree files remain intact.
OLE2/OOXML work remains active, ODF is deferred, and iWork excluded.

Replay retained evidence with Python 3 from the repository root:

```sh
python3 -B docs/performance/results/change-0796/validate.py --require-final-seal
python3 -B docs/performance/results/change-0796/seal_packet.py --check-head
```

The second command checks this batch's committing HEAD; after later commits,
use the first command to replay the sealed packet. Probe unit-test modules are
excluded (`cfg(test)`); production helper tests are inherited from exact-source
0794 evidence. Error-followed-valid-tail recovery coverage is not claimed.
