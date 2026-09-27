# Observer-lane review

`observe.py` is the pure offline hook for the final 0780 analyzer.  Import
`analyze(packet)` and consume its returned mapping; it reads receipts and
retained files, hashes those files to check the recorded size/SHA, and does
not write `analysis.json`, start a command, inspect a profiler, or hash its
own source.  The default packet is the directory containing this file, and a
packet path may be passed explicitly for replay after relocation.

The frozen `observer-plan.json` fixes CPU 12, three `perf stat` blocks, the
`capabilities/tiny` and `commit/large` cases, 30 samples plus three warmups,
and `instructions:u,cycles:u`.  This is exactly 3 blocks × 2 cases × 2 legs =
12 process receipts.  The hook checks every command token, including
`taskset`, event list, report/stats paths, sample arguments, and the native
binary receipt.  It checks each build manifest's source census, probe inventory,
lock, build logs, environment, and native/allocation binary identity.  A
missing executable is accepted only when `cleanup.json` is verified and
contains the exact raw path, byte count, and SHA tuple from the binary receipt.
Absolute paths captured in the owned worktree are mapped through
`origin.json`; packet artifacts remain packet-bound and cannot resolve outside
the packet.

For every perf process the hook verifies the report schema, mode, shape,
dimensions, marker, sample count, source digest on every sample, semantic
verification, and deterministic publication output.  It compares source and
output identities across all before/after processes in each case.  It parses
both counters from the semicolon perf file and retains the running percentage
for each counter.  The returned `pairs` list contains all six block/case
pairs, with before/after values and percentage changes.  These are whole
process counters, including probe setup and output/readback verification; they
are not operation-region counters and carry no isolated timing claim.

Each heaptrack lane must contain one `capabilities/tiny` process for its leg.
The hook checks the exact `heaptrack` command, report, log, source parity, and
every raw trace receipt.  The trace is retained as the primary observer
artifact.  Root's decode receipt should use the shape already produced by the
baseline capture:

```json
{
  "command": ["heaptrack_print", "-f", ".../trace.zst", "-H", ".../histogram", "-n", "15", "-s", "3"],
  "exit_code": 0,
  "trace": {"path": "...", "bytes": 0, "sha256": "..."},
  "log": {"path": ".../print.log", "bytes": 0, "sha256": "..."},
  "histogram": {"path": ".../histogram", "bytes": 0, "sha256": "..."}
}
```

`observe.py` accepts `log` as the decoded print output and then streams the
retained `.zst` interpreted Heaptrack v3 records through offline `zstd`.  It
resolves string/instruction/trace ancestry, attributes `run_capabilities` and
`Capabilities::ooxml_baseline` allocation events, and requires the complete
trace event count and requested-byte sum to match the histogram, plus the
allocation-event count to match the print summary.  The demangled text is retained as a stack witness as well.
It returns the before/after attribution under `heaptrack.diagnostic`.  The
diagnostic is complete only when both decode receipts are present and bound to
their raw traces.  Until the after trace is decoded it reports
`pending_decode` (or `unbound` when text exists without decode metadata).  The
attribution is allocation-site evidence only: it makes no timing, RSS,
physical-copy, or causal-cost claim.

Current baseline evidence has 24,951 whole-process allocation events and
3,434,778 requested bytes.  The exact interpreted ancestry gives 21,003
`run_capabilities` events / 2,398,413 bytes and 21,000 direct
`Capabilities::ooxml_baseline` events / 2,398,000 bytes.  The print log also
shows the constructor site through `HashSet::insert`; its SHA and the raw
trace/histogram SHAs are bound by `heaptrack-before/decode.json`.  The
candidate after lane must use the same shape and command before the two
diagnostics are compared.

## Completed observer replay

Both final lanes now pass. Exact constructor attribution is 21,000 events / 2,398,000 requested bytes before and zero after; enclosing diagnostic reporting retains three events / 413 bytes in both legs. Twelve perf-stat processes and two Heaptrack processes match their bound commands, source, executable, output and decode identities. The before Heaptrack census is checked against the before build, and the after census against the after build. Returned analysis contains stable identities while retained-file checks or exact cleanup witnesses remain mandatory. Root replay after target removal agrees byte-for-byte with observer-analysis.json and analysis.json.
