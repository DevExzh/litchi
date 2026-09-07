# 0464: source-backed PPTX pair lifecycle baseline

This batch adds a manifest-driven, harness-only PPTX slide-copy command with
separate source and destination owners, bytes and capped-range providers,
four API timing regions, allocator diagnostics, sequential output accounting,
and an independent ZIP/XML closure oracle. It does not change production code
or claim a performance improvement. The full non-iWork goal remains open.

The [pair manifest](pair.json) binds a public LibreOffice QA source and a
two-slide Litchi-derived destination. This is a **same-source-derived positive
control**, not an independently authored native pair. See
[fixture provenance](fixture-provenance.md) and the
[qualified corpus screening](native-pair-audit/README.md).

## Measurements

[protocol.json](protocol.json) freezes eight serialized lanes: bytes/range ×
normal/allocator × R1/R2, with 30 samples and three warmups per lane, CPU 2 and
one explicit worker. R1 runs forward; R2 reverses the lane order. The range
adapter caps returned bytes at 256 and has zero configured delay. It measures
logical short reads, without physical/network or cold-cache claims.

The [summary](summary.json) derives nearest-rank p50/p95/p99, means and
fixed-seed bootstrap mean intervals. Each `api_sum_ns` is the sum of source
open, destination open, plan and publication clocks. It excludes input loading,
adapter construction, sink reservation, between-phase diagnostics, independent
oracles, artifact writes and drops; it is not a contiguous end-to-end clock.
GNU time RSS covers the whole process. Phase allocator peaks are not additive.
Configured memory budgets apply to individual execution owners and read
profiles, not to the entire harness process or its outside-timing oracles.

All 240 retained samples reproduce the same 55,891-byte, three-slide output.
Each retained output passes the independent oracle. The Python oracle derives
the copied closure and insertion order from the inputs, verifies relationship
targets/content types and exact payloads, and checks untouched destination ZIP
records with only central-directory relocation offsets masked.

## Reproduction

The original measurement binary build commands and source epochs are recorded
in [binding.json](binding.json). The five changed source files at that epoch
are archived in [source-code.json](source-code.json). Build on its recorded
base revision with those exact archived sources to reproduce the measured
harness. The supplemental inventory diagnostic has a separate
[build binding](inventory-binding.json); [source-compatibility.json](source-compatibility.json)
proves that only that independent binary source changed after measurement.

```sh
python3 -B docs/performance/results/change-0464/summarize.py --check
python3 -B docs/performance/results/change-0464/verify.py
```

Portable verification uses retained inputs, outputs, manifests and receipts;
it does not require the temporary executables. To make new captures, use a
fresh result directory and bind the new binaries and protocol. Existing
capture/check artifacts are immutable. The exact capture commands are in
`protocol.json`; the generic harness selector is `pptx-pair-lifecycle`.

## Validation and limits

The retained history includes the initial regenerated-relationship oracle
failure, unnecessary-mut lint failures, and smoke assertions corrected to
match the actual typed budget errors. Latest applicable gates, rather than
superseded failures, determine the batch status. The broad format command
finds an existing Keynote formatting difference; the changed harness passes
its package-scoped formatting gate and iWork remains untouched.

Native application execution and semantic readback are separate from the
timing matrix. LibreOffice saves the output, but its markup-compatibility
markup causes the source-backed picture inventory to refuse that surface.
The supplemental diagnostic preserves those typed refusals while collecting
the available slide/text and eager observations. Image equivalence and
rendering remain unproved. Retained native receipts state the exact checks
that succeeded and the unavailable observations.

See the [source review](source-review.md), [per-change record](../../changes/0464-pptx-derived-pair-lifecycle.md)
and [next work](next-work.md). This supplementary command does not promote a
checked default-baseline coverage row: 439 selectors and 36 defaults remain.

Final evidence passes [precleanup verification](precleanup.json),
[fresh-copy portable replay](portable-verification.json), and a
[resealed one-nanosecond summary mutation rejection](tamper-test.json).
[Cleanup](cleanup.json) removed 12 owned temporary files totaling 252,885,687
bytes and two audit temporaries whose archived hashes matched. The complete
bundle remains independently verifiable after executable removal.

The retained [verification attempts](verification-attempts/) document helper
integration corrections: historical source/binary binding shapes, JSON
normalization of native inventory tuples, and portable repository discovery.
These corrections changed no measured captures or required semantic checks.
