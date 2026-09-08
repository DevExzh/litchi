# Change 0478: PPTX generated-name metadata spool

This bundle documents the explicit-scratch generated-name route for fresh,
text-only PPTX streaming creation. The final result is a scoped
operation-heap observation and same-source correctness/storage-policy
comparison; it makes no latency or general-memory claim.

The public route is exposed by
`StreamingPresentationWriter::with_options_and_metadata_spool` and its
default-policy convenience `with_metadata_spool`, together with
`StreamingPresentationScratchLimits { max_bytes, buffer_bytes }`. The caller
supplies a `Read + Write + Seek + Send + Sync + 'static` provider. The provider
stores finalized ZIP central-directory records under the caller's byte quota
and replays them through the configured fixed buffer. Production does not open
an ambient scratch file or expose ZIP plan types through the PPTX API. The
provider must preserve the admitted range and exclusive coherent access,
including backing aliases; provider-owned storage is outside the library's
working-memory bound.

## Name-plan scope

PPTX preflights the exact current serialization topology before constructing
the physical writer or accepting output. The checked plan contains 37 fixed
members and two members per declared slide, in the existing emission order:

```text
9 leading literal members
11 slide-layout XML/relationship pairs (22 members)
6 trailing literal members
N slide XML/relationship pairs (2N members)
```

Thus a deck with `N` slides emits exactly `37 + 2N` physical members. The ZIP
plan stores only bounded descriptors and a scalar cursor. Its symbolic proof
checks canonical production names, decimal-range overflow, exact and
ASCII-folded collisions, and whole-component ancestor/descendant conflicts
without expanding the slide range. The OPC owner validates absolute PackURI
literals and indexed endpoints, then enforces the checked sequence while
retaining no generated-mode `PartNameSet` or completed-name history. A name is
advanced only after the corresponding member is successfully finalized.

The default constructors remain on the ordinary storage path. They retain
their existing arbitrary-name and duplicate/ancestor validation policy. The
new route preserves the existing semantic limits, output byte accounting,
slide/relationship identifiers, serializers, deterministic bytes, and typed
forward-only failure behavior.

## Final measured results

The final matrix contains 24 isolated reports and 720 measured operations.
Elapsed values below are milliseconds; p50/p95/p99 are nearest-rank
percentiles, and RSS is the contextual process maximum from the matching
`/usr/bin/time` observation. RSS includes setup and teardown and is not a
library heap measurement. Control has no scratch extent.

| Repeat | Slides | Policy | Mean ms | p50 ms | p95 ms | p99 ms | RSS KiB | Scratch bytes |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 8 | control | 0.905319000 | 0.904344000 | 0.915683000 | 0.919723000 | 4,044 | — |
| 1 | 8 | spool | 0.933504600 | 0.933944000 | 0.943084000 | 0.947064000 | 4,044 | 4,074 |
| 2 | 8 | control | 0.901271433 | 0.900583000 | 0.915664000 | 0.917114000 | 4,064 | — |
| 2 | 8 | spool | 0.931108567 | 0.929294000 | 0.945394000 | 0.961084000 | 4,052 | 4,074 |
| 1 | 256 | control | 6.344201900 | 6.339666000 | 6.455467000 | 6.455507000 | 6,808 | — |
| 1 | 256 | spool | 6.320574300 | 6.321027000 | 6.341257000 | 6.345406000 | 6,832 | 40,842 |
| 2 | 256 | control | 6.471067100 | 6.480767000 | 6.492937000 | 6.523597000 | 6,844 | — |
| 2 | 256 | spool | 6.295404733 | 6.293517000 | 6.316657000 | 6.323947000 | 6,596 | 40,842 |
| 1 | 8,192 | control | 185.536250633 | 184.195604000 | 188.825695000 | 189.065036000 | 102,860 | — |
| 1 | 8,192 | spool | 182.077139700 | 182.301259000 | 182.950661000 | 183.892064000 | 99,536 | 1,237,692 |
| 2 | 8,192 | control | 185.602817200 | 185.164615000 | 189.965465000 | 190.741678000 | 101,472 | — |
| 2 | 8,192 | spool | 181.198961100 | 182.400040000 | 182.932063000 | 183.125533000 | 99,416 | 1,237,692 |

The eight-slide spool mean is 3.113% above control in repeat 1 and 3.311%
above control in repeat 2, the measured small-operation setup cost. The
analyzer reports no positive policy flag above 5% across the 12 policy pairs (six normal and six allocator) and no repeat-drift flag above 5%.

Allocator values are identical in both external repeats; each arrow is
control → spool and each RSS pair is repeat 1 / repeat 2:

| Slides | Allocation calls | Requested bytes | Incremental peak live bytes | RSS KiB | Scratch bytes |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 8 | 850 → 617 | 499,462 → 475,312 | 435,701 → 432,436 | 4,028 / 4,060 → 4,036 / 4,028 | 4,074 |
| 256 | 9,049 → 5,825 | 1,415,242 → 861,904 | 681,819 → 432,436 | 6,812 / 6,800 → 6,844 / 6,600 | 40,842 |
| 8,192 | 278,153 → 179,677 | 31,428,173 → 13,379,518 | 8,875,252 → 432,436 | 99,524 / 101,316 → 102,880 / 99,412 | 1,237,692 |

The scoped memory gate covers 180 spool allocator samples. Every incremental
peak is exactly 432,436 bytes, with zero failed allocations and zero live-byte
exit delta; the global range is 0% against a 1% threshold. This is an
operation-scoped generated-route allocator observation, not a bound on total
RSS, provider storage, native allocations outside the instrumented allocator,
or the ordinary control route.

The configured scratch quota is 67,108,864 bytes and the replay window is
16,384 bytes. Observed scratch is 4,074 / 40,842 / 1,237,692 bytes for 8 /
256 / 8,192 slides, matching the central-directory identity over 37 fixed
names and two names per slide. The provider is on tmpfs in this environment;
the measurement does not describe durable-storage latency.

The benchmark compares the same source and matching normal/allocator builds
across the control and explicit-file-spool policies at the retained small,
medium, and large slide counts, with reversed external repeats. Each formal
report contains 30 samples after 3 warmups on CPU 2 of the recorded AMD EPYC
9R45 environment, using Rust 1.98.1 and the frozen four-job, no-incremental,
frame-pointer/unwind-table build policy. Mean intervals use t(29) = 2.045;
normal and allocator timings are separate. Both lanes must produce
byte-identical output and pass physical and semantic reopen checks.
Preliminary captures are superseded and excluded: the two archived
formal matrices were captured before the OPC percent-prefix guard and before
the cleanup fix. The corrected PPTX and ZIP harnesses propagate errors before
cleanup and retain scratch on failure, including pre-existing files refused by
`create_new`.

## Final custody, validation, and cleanup

The final production source is commit `69765160b`, and the diagnostic harness
is `b16fb3962`. The final copied binaries bind the unchanged 7,046-file source
manifest. `verify.py --data-only` passes all 24 captures, 720 samples, output
identities, scratch arithmetic, physical members, semantic PPTX checks, and
summary invariants.

The final evidence set includes `rust-validation.json` for the complete gate
ledger, `live-state.json` for source/template/binary custody,
`cleanup.json` for authenticated runtime removal and shared-cache retention,
and `portable.json` plus `SHA256SUMS` for fresh-copy replay and negative
probes. The final gate totals are 1,887 library tests with 5 ignored over 103
targets, 387 harness tests with 1 ignored, 180 shared streaming tests, DOCX 4,
ODF 14, ODT 14, ODS 16, ODP 12, and 21 Python evidence tests. These ledger,
live, cleanup, and portable receipts are required completion evidence and are
retained before the batch is committed.

The allocator binary copy first hit errno 122 after a 319,131,648-byte partial
destination. The authenticated partial copy was removed, and
`build-copy-recovery.json` records the successful 414,504,696-byte copy
matching its origin at SHA-256
`6279bccb61337e572d215b7f788a670e95ea839a6d6dfe9a83e5f61fae52b632`, with
the source unchanged.

For portable replay after the final seal, run:

```text
python3 -B docs/performance/results/change-0478/verify.py
```

The verifier authenticates the exact argv, environment, frozen protocol,
normal/allocator binary hashes, and embedded XML source bindings. Capture and
build helpers use exclusive-create outputs and must be run against a fresh
evidence root; they do not overwrite sealed reports. A new observation requires
a new explicit scratch path and a new bundle, while the retained final bundle
is replayed read-only.

The bundle's retained source-review and protocol files describe the proof and
measurement obligations. It must not be read as evidence of constant total
memory, durable storage latency, arbitrary-name streaming, logical append,
Part addition, repackaging, native/source-variant breadth, or completion of
the wider non-iWork goal.

The final source also rejects an OPC indexed prefix ending `%` before any
endpoint allocation: valid endpoint URIs can otherwise hide an invalid
percent-encoded name inside the range. The regression checks that refusal is
atomic and that fixed percent escapes remain accepted. Both earlier matrices
are retained separately; the final rerun binds this guard as well as the
benchmark cleanup fixes.
