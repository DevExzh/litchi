# Minimum-service policy inside the unchanged managed PPTX lifecycle

Both policies wrap explicit caller-owned OwnedSource with ReadAt. The fixed
200 us delay still occurs before every nonempty delegated read; the 64 KiB cap,
returned bytes, source versions, caller budgets and cache behavior are unchanged.
The separate-sleeps policy then requests the full nominal transfer sleep.
The minimum-service policy samples elapsed time from just before the fixed wait,
and sleeps only for the remainder of fixed-plus-nominal-transfer service time.
Fixed-wait overshoot, adapter bookkeeping and the wrapped-source read contribute
to this minimum. This models a minimum service duration, not additive transport
latency or a physical/shared-link bandwidth constraint.

Both expose the same nominal transfer-delay counters. These are target components,
not actual sleep durations or sleep call counts. Whole serial API duration must
cover the sum of fixed and nominal transfer targets for all reads in each phase.
No nominal transfer target accrues for an empty, zero-cap, EOF or failed read.

Source/destination open, cross-slide-copy planning and consuming publication
keep their exact existing API clocks. Fixture construction, corpus/refusal gates,
source wrapper creation, diagnostics, full-output comparison, readback and final
owner drops remain outside. The CountingSink retains full output. Operation
allocation attribution is unavailable; managed boundary gauges and process RSS
are distinct and do not prove bounded total memory. CPU profiles include untimed
work and omit blocked sleep. Both configurations are one-worker simulations;
no production speedup, physical bandwidth, cold I/O or native/scaling claim follows.
