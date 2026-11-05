# 0362yda — return matched ODP lifecycle regressions

Reviewed evidence: `9965f37d8d`, measuring candidate `12385bc045` against the
original changed-route baseline `e4a0614f5d`. The tracked common probe,
adapters, fixed executable identity and external artifact manifest improve
reproducibility. They do not close the performance return.

The candidate reduces peak heap on open and reader routes, and matches the
append peak. It increases cumulative append/replay allocation by 2.03%.
All ten recorded whole-process instruction counts increase: the displayed
repeated `write_text` ratio is 1.024823, with append/replay ratios 1.002322
and 1.002360. These increases are returned for repair under the coordinator
rule; neither a small percentage nor the heap improvement is an exception.

The whole-process Callgrind counts also contain fixture setup, unlike the
allocator lifecycle window. Record 0362ye must identify inclusive route
costs or use matched instrumentation boundaries, retaining owner open,
operation, output verification and teardown in the relevant lifecycle.
Keep whole-process measurements as context. Verify semantic reader outputs
as well as serialized outputs; a zero-byte sink hash alone cannot prove
equal slide or text observations.

Repair the measured costs while retaining bounded streaming, encryption and
XML fidelity, content-based replay and retained limits. Reuse the original
baseline and tracked driver, preserve rejected evidence, and rerun matched
allocation/instruction checks. Native ABBA qualification remains required;
no native timing was claimed by 0362yd.
