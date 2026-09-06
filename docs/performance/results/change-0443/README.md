# 0443: Compact ODP source-fragment scanner frames

The private source-fragment scanner now records an exact element kind and byte
offset in each open frame. It no longer copies namespace and local-name bytes
that were used only for fixed-name comparisons. Root bindings, retained source
XML, opaque markup, byte spans, style names, limits and publication readback
remain unchanged. The entire original scanner is retained byte-for-byte as a
test-only differential reference.

The frozen allocation-call gate passes: medium/large use 14.886%/14.943%
fewer calls across both repeats. Requested bytes fall 2.557%/2.663%; peak
and retained live bytes are unchanged. The normal latency gate fails, with
medium p50 changes -0.606%/+2.460% and large +0.267%/+1.402% (R1/R2).
The [decision](decision.json) retains the change for fewer allocation calls,
accepting those sub-5% normal p50 costs. No latency speedup is claimed.
Two instrumented medium R1 tail regressions and all eight repeat flags remain
visible; no measurement is excluded or replaced.

The [frozen protocol](protocol.json) uses A1/B1/B2/A2, 24 reports/720 samples,
64/4,096/8,192 slides, normal/allocator executables, CPU 2, one worker,
30 samples and three warmups. Four fresh whole-process profiles follow.
The inherited freeze timestamp is a disclosed metadata error; see
[freeze-review.json](freeze-review.json). The original file, thresholds and
measurement lanes remain unchanged. Capture timestamps record actual chronology.

The independent report oracle is byte-identical to 0439–0442. Rust fixture
gates inspect actual archives for slide semantics, one exact tail append,
opaque preservation, exact no-op sharing, reversible patches and stale-source
refusal. Python independently regenerates expectations and verifies report
contracts; it does not reopen archives embedded in reports.

Portable verification needs Python 3 without Cargo, perf, original binaries,
temporary directories or the original checkout:

```sh
python3 -B verify.py --sealed --cleanup
python3 -B derive.py --check
python3 -B profile-summary.py --check
python3 -B measurements.py --check
```

The verifier binds protocol/oracle/drivers, exact source/build identities,
ordered captures and commands, all selected report oracles, profile artifacts,
chronology, check receipts, reference-scanner identity, cleanup and the sealed
inventory. Independent corrupted copies refresh inventories and report hashes
before rejection checks. Completed probe receipts are resealed.

See [measurements](measurements.md), [decision](decision.json), and
[validation notes](validation-notes.md) for results, flags, gates and limitations.
Capture/evidence tooling is adapted from 0442. The optional prior-attribution.py
reads the sealed 0442 profile and is not needed for portable verification.

The registry stays at 436 selectors, default matrix at 36 cases, and coverage
at 15 categories/33 mappings (10 measured, 23 correctness-only). No promotion.
Bounded existing append, Part addition, repackaging, native breadth, cold/range,
scaling and other non-iWork goal requirements remain open.
