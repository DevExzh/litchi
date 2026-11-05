# 0362ul57 — role-aware DrawingML raw preparation repair

Date: 2026-10-25. This is a source-only repair for the failed tooling row at
integration checkpoint `319690ffc68ad9fac5e28ce2b0756bc289bc5c82`. It does not
run Cargo, a native probe, or a timing workload. The previous failed receipt
and its source freeze remain historical inputs; a fresh preparation receipt is
required after this repair.

## Failed row retained

The isolated tooling row ran the existing full tooling step against the clean
integration source and failed one test after 1,688 tests and 108.883 seconds.
The failure was the stale DrawingML adapter expectation:

~~~text
expected tools/perf-baseline/src/lib.rs:
  280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a
observed:
  ff0ca3a924852cfc5d8497f812fbf412c4eb71b5f289f1c3b3136dea92c7ed9f
~~~

The row exited 1 with no target growth, an empty post-exit trace, a clean
source handoff, and no owned processes remaining. The retained artifacts are:

~~~text
state:   (deleted-coord-artifact)/state.json
wrapper: (deleted-coord-artifact)/tooling-row-319690ffc6-20261025T-wrapper.log
wrapper bytes/SHA-256: 3999 / d69f9b504d5eeaba2b86c77ed737fcb96ea5ee95be1cdb2e240af872cd6c9fb6
child:   (deleted-coord-artifact)/48-tooling-unit-attempt-1792890578893939273-1202833.log
child bytes/SHA-256: 636346 / 1f52db64e95036ff40de1ef51bc5793fd5b085f0b9a8c86991f2b79323d1bf59
source freeze: (deleted-coord-artifact)/tooling-row-319690ffc6-20261025T-source-freeze.txt
candidate lib.rs SHA-256: ff0ca3a924852cfc5d8497f812fbf412c4eb71b5f289f1c3b3136dea92c7ed9f
~~~

## Cause and bounded repair

RTF commit `1d273acbf2` changed exactly two consumer spellings in
`tools/perf-baseline/src/lib.rs`:

~~~text
litchi_rtf::write::StreamingRtfLimits::new  ->  litchi_rtf::streaming::Limits::new
litchi_rtf::write::StreamingRtfWriter::new  ->  litchi_rtf::streaming::Writer::new
~~~

The current candidate must use the new RTF API. The historical DrawingML
baseline still uses the old RTF API, so copying the whole current `lib.rs`
unchanged into that baseline would make its historical consumer fail to
compile. Refreshing the expected hash alone would hide this incompatibility.

`prepare_drawingml_raw_inputs.py` now keeps the candidate adapter byte-exact
and applies a role-aware compatibility transform only while materializing the
baseline copy of `tools/perf-baseline/src/lib.rs`. It requires exactly one
occurrence of each current spelling, rejects an already-transformed source,
and records both role hashes and the exact two-replacement recipe:

~~~text
candidate lib.rs: ff0ca3a924852cfc5d8497f812fbf412c4eb71b5f289f1c3b3136dea92c7ed9f
baseline lib.rs:  280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a
replacement count: 2
changed bytes: 16
~~~

All other adapter paths remain byte-identical between roles. The receipt now
binds the role-specific `lib.rs` hash and the complete transformation metadata;
old receipts are not rewritten or promoted. The DrawingML probe, allocator,
fixture, and timed operation code are not transformed. The focused fidelity
test constructs the baseline bytes from the candidate bytes and proves that
the only changes are those two API spellings.

## Source-only checks and fresh handoff

The targeted check is:

~~~sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_drawingml_raw_qualification
~~~

It must pass before a new preparation receipt is created. The coordinator must
then run the existing preparer with a fresh external output path, validate the
new receipt, and only afterward build or measure the five raw lanes. No prior
receipt can be reused because its role hash describes the pre-repair policy.
