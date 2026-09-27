# 0785 — exact known namespace URI candidate

The 0784 capture profiles identify `notes::resolved` as a repeatedly observed
nested function. This candidate uses exact byte equality with the six existing
valid UTF-8 namespace constants to avoid validating those same URI bytes again.
Unknown bound values retain `from_utf8`; unbound and undeclared-prefix behavior
remain unchanged. No namespace cache, global state, unsafe code, dependency,
public API or relaxed XML check is introduced.

The frozen matrix covers capture, staged commit and the full public lifecycle
across the inherited tiny/medium/large fixtures and two new medium vendor
fixtures. Each vendor text element carries six unknown namespace attributes,
with URIs sharing the six known byte lengths. One variant uses ASCII near
misses, the other valid Unicode URIs. Fixture construction is outside clocks.
The text/source/output oracles remain mandatory. Ordinary fixture identities
are checked against the retained 0780 corpus.

Native capture uses six alternating process blocks with thirty samples and
three warmups. Separate allocator processes use two blocks of three samples
without warmup. No profiler timing is pooled into native measurements.
The adoption policy is frozen before builds: any public-case p50 slowdown over
5% with a 95% bootstrap lower bound above one rejects the candidate, as does
increased net-live or peak-above-entry allocation. At least one capture or
lifecycle case must improve at least 3% with an upper bound below one. All
flags remain visible. Cumulative allocation reductions alone are insufficient.

Build and capture reproduction needs a fresh checkout, output packet and target
at the recorded base, copied workspace and probe locks, and the recorded
reference links. Run baseline `build.py before`, `probe_tests.py`, `capture.py qualification`,
apply the reviewed candidate patch, then `quality.py`, `build.py after`,
`capture.py native`, and `capture.py allocation`, serially. Retain every failed
attempt. Regenerate analysis, tables, decision and seal only in that new packet;
do not overwrite this evidence or copy the old decision.

This batch does not certify physical cold-cache, range-source, concurrency,
scaling, complete CRUD or native Office interoperability. OLE2/OOXML remain
active, ODF deferred and iWork excluded.


Offline replay uses `python3 -B validate.py` and `python3 -B tables.py --check`
from this directory. `raw_audit.py --check` independently reconstructs paired
p50 and memory checks. Neither replay invokes benchmarks. The full report is
[0785](../../0785-pptx-known-namespace-uris.md).

`failed-build-before-0` and `failed-build-before-1` preserve the pre-measurement
lock and type-inference failures. Their original receipts use `build-before/`
paths; correction receipts record each relocation. `probe-checks` retains the
initial incompatible synthetic/real-allocator test invocation; `probe-tests`
contains the passing established separate protocol. No primary measurement
was retried. `revision-transition.json` records the two documentation-only
agent commits preceding successful baseline builds; production bytes match
origin exactly.
