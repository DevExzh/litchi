# 0826 — allocation-vector schema preflight

The reusable `tools/perf_allocation_schema.py` validates the report shape that
caused the frozen 0825 checker to fail: allocation counters are per-sample
metric objects, not scalar integers. `tools/test_perf_allocation_schema.py`
contains 30 regression tests, including malformed vectors, signed net-live
changes, region/lifetime peaks, unavailable native counters, and CLI behavior.
Both normal and optimized Python runs pass.

The positive preflight pins 325 existing reports / 6,733 sample envelopes from
0819, 0821, and the failed 0825 qualification by path, size, SHA-256, and their
historical seals. It checks all twelve real-file format/phase selectors and
reproduces the old checker's failure on the same report the new helper accepts.
These are historical schema fixtures, not new measurements. No comparative
statistics are calculated and no optimization or speedup decision is made.

Nonzero failed-allocation counters are valid schema; trial policy decides
whether to admit them. The first passing draft preflight, which conflated that
policy with schema validation, is preserved with its helper source under
`preflight-attempt-0`. The final preflight uses the corrected policy separation.
There were no failed test, preflight, or validator commands in this batch.

Replay from the repository root:

```sh
python3 -B -m unittest tools.test_perf_allocation_schema -v
python3 -B -O -m unittest tools.test_perf_allocation_schema -v
python3 -B docs/performance/results/change-0826/preflight.py --check
python3 -B docs/performance/results/change-0826/validate.py
python3 -B docs/performance/results/change-0826/seal.py --check-head
```

The source review and origin census prove production Rust, benchmark runtime,
locks, normative inputs, and unrelated files remain unchanged. The test command
receipts bind all tested source snapshots and logs. No Cargo/build/scratch root
was created; temporary CLI fixtures use automatically removed directories.
The seal binds the two Python tools, this packet, and the report/index updates.

For the next matched trial, call `validate_report` with expectations from the
frozen protocol, freeze the helper's hash as an execution input, and retain
independent source/binary/corpus/artifact checks. Also correct the prepared
reader's accepted-attempt path and canonical JSON replay before freezing it.
0825 remains an aborted attempt; this repair does not retroactively admit it.
