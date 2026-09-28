# 0803 — isolate exact-empty check placement

This is a diagnostic control, not a production-candidate qualification.
Before is the exact rejected 0802 candidate; after moves its exact-empty
conditional from construction to the first next call. The linear array,
ordered backend and duplicate-preflight implementation stay unchanged.
Production remains untouched and 0802 remains rejected, regardless of the
observed ratio. No result here permits workflow advancement or adoption.

The same 39 literal cases and construct/consume owners run in one fresh
release binary. Six paired native blocks use 30 samples, three warmups and
4,096 iterations per process on CPU 12. A fresh seed 803080 drives the same
10,000-resample paired-median bootstrap. Two separate Callgrind repeats retain
guest instruction and branch diagnostics. Historical timings are not pooled.

The hypothesis is that deciding empty state during construction contributes
to 0802's opaque construction/drop cost and early syntax regressions. Moving
the check is one source intervention; associated compiler/code-layout effects
remain inseparable from that intervention in this experiment. This does not
isolate backend design, iterator size or a particular instruction as the cause.
No before leg in this packet is the current production helper.

The full source/architecture witnesses separately preserve the production
baseline. Candidate-before copies are bound to the sealed 0802 candidate,
and the five-copy mirror tests plus probe oracles check behavioral parity.
The patch is archive-relative, checked against candidate/before; it must not
be applied to production. Root runs all builds and captures serially, then
independent offline replay and review verify the retained results. Owned-target
cleanup retains executable identity before removal.

## Final result

All 39 construction rows improve against the rejected control, with ratios
0.296307–0.372513. Empty/one/two consumption ratios are
0.519075/0.761308/0.868048. No diagnostic regression flag occurs; all 73
spread flags remain visible. Neither leg is production and no workflow
advancement or adoption is authorized.

Both mirror legs pass 100 tests and Clippy. The initial after-only private-state
test failure and its correction are retained. All 1,248 reports and 28,392
samples pass independent replay. The owned build target is removed with exact
executable identity retained. See `summary.md`, `decision.json`, source/results
reviews and the main 0803 report for limits and individual results.

Offline verification:

```sh
python3 -B docs/performance/results/change-0803/validate.py --require-final-seal
python3 -B docs/performance/results/change-0803/seal_packet.py --check-head
```
