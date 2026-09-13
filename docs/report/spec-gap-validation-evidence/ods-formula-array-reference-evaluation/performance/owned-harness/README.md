# Owned evaluation profile harness

This candidate-only harness measures the public `Evaluated::to_owned` copy
boundary.  The `own` phase parses and evaluates one expression before the
timed loop, then times only conversion and result inspection.  The
`parse-evaluate-own` phase is an explicitly broader end-to-end control; its
numbers must not be compared with `own` as if they covered the same work.
`setup` measures repeated preparation separately so callers can see that
preparation is excluded from the isolated conversion lane.

The corpus contains scalar number and Unicode text controls, square inline
arrays with 1/16/256/4096 cells, duplicate local reference lists at those
sizes, and 3-D reference lists spanning Main through Archive.  Four bounded
refusal controls exercise text, aggregate-memory, work, and cancellation
failures.  The executable reports allocator request/release counters, live
and allocator-peak bytes, execution-budget usage, exact owned retained
reservation, typed failure labels, and a structural checksum.  A retained
reservation is the `OwnedEvaluated` budget charge while the result is live;
it is not a process RSS peak or a transient allocator peak.

`run.py` is a serial CPU-6 runner.  It requires a new output directory and
records source binding, compiler/build metadata supplied by the caller,
binary and all five harness input hashes before and after the run, child stdout/stderr, GNU
`time -v` RSS, commands, and `raw.csv`.  It does not build the harness.  A
capture is diagnostic evidence for this API boundary and makes no speedup or
zero-copy claim.

The executable runs an independent untimed semantic preflight before sampling:
expected scalar/text values, array cells and shapes, reference form, duplicate
counts and lexical metadata must match the fixture. Borrowed and owned checksums
include reference lexical fields and must agree. The runner retains the supplied
source manifest and rejects changes to it during capture; this binds the supplied
manifest, while the caller remains responsible for the binary build provenance.
Warmups are capped at 100, iterations at 1,000, and executable repeats at 4,096.

Current debug validation is diagnostic: 48 of 54 lanes pass with the independent
oracle. The six 4,096-entry reference-list lanes exceed the production evaluator's
default preparation work budget. These failures remain visible; no limits or
corpus entries were changed to conceal them. See
[the retained preflight receipt](../diagnostics/owned-preflight-03.json).
The default 15 samples provide coarse p95/p99 estimates and do not establish
stable tail latency.
