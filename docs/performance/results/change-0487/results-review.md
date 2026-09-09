# 0487 results review

Retaining consumed replay prefixes inside the existing adapter allocation is
accepted for this scoped change. Both normal repeats show authored-heavy file
p50 falling from 239.982/238.837 ms to 153.609/153.778 ms (35.99/35.61% lower).
Owned input falls 9.02/8.53%; short-read input 7.91/9.68%; injected-latency input
9.62/7.88%. Source-heavy normal medians stay within 1% across all six routes.
Memory-store authored-heavy medians change less than 1%; file-store medians
fall 0.93/8.57%, an inconsistent magnitude that is not a general speedup claim.

[All comparison rows](comparison-summary.md) retain 144 formal children and
4,320 measured samples, with two separate process repeats per phase and arm.
The [plot](latency-comparison.png) shows normal p50 only; tails, allocator
measurements, and adverse rows remain in the complete tables. There is no
confidence interval, reserved-host guarantee, or cold-cache inference.

## Adverse observations

Every latency increase above 5% is listed here. All concern the 64-source,
64-authored deterministic route. Values are signed percent changes.

| Input | Binary | Repeat | p50 | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: |
| owned | normal | 1 | +7.74% | +2.74% | -3.20% |
| owned | allocator | 1 | +8.73% | +8.76% | +8.23% |
| short-read | allocator | 1 | +6.97% | +7.18% | +6.66% |
| short-read | allocator | 2 | +5.74% | +5.70% | +5.26% |

Normal owned repeat 2 improves 6.42%; normal short-read repeats change
-6.64/+4.55%. These do not eliminate the allocator short-read regression,
which appears in both repeats. Small-workload overhead remains a tradeoff.
No other latency row exceeds +5% at any retained quantile.

Three distinct whole-child RSS pairs increase above 5%, all small-workload
repeat 1: normal short-read 6,443,008 to 6,836,224 bytes (+6.10%); allocator
latency input 6,565,888 to 6,905,856 (+5.18%); allocator memory-store
6,836,224 to 7,397,376 (+8.21%). Each is one observation per process; the
summary's three identical quantile labels are not independent RSS samples.

Operation heap quantiles are unchanged for deterministic and memory-store
arms. File-store p50 decreases by two bytes, attributable to measured path
storage rather than a meaningful buffer reduction. No operation-heap row
increases above 5%. This change adds no allocation or larger window.

## Mechanism and identity

The [implementation review](implementation-review.md) covers tail appends,
consumed-byte accounting, fragment boundaries, and failure ordering. Sink
freshness fences now follow full windows instead of each short replay read;
provider callbacks retain their own pre/post freshness checks. The exact
version-counter test requires `4 + 2*payload_len + 3*window_flushes` for the
one-byte replay fixture. Standalone source XML validation remains intact.

[Whole-child diagnostics](profiles/profiles1-summary.md) include setup and
oracles, with one measured sample after one warmup per profile child.
Authored-heavy file statx calls fall 1,475,055 to 591,477 (59.90%); pread64
stays 114. Source-heavy file statx falls 25,219 to 21,781 (13.63%); pread64
stays 264. Owned controls retain 12 statx and 6 pread64 calls. Authored-heavy
file instructions fall 19.92% and cycles 25.96%; source-heavy instruction
counts change less than 0.2%. These diagnostic counts support the mechanism
but are not operation-only CPU measurements or a universal causal claim.

Every formal source, authored, candidate XML, semantic, and preservation
oracle passes. Candidate ZIP length and SHA-256 also match across phases:
there are zero candidate archive identity changes. The after source manifest
is `cb69be7b5a1f443191a27000fdfab285b707ba28517cccc3ee764d833cde0ba2`;
only the OPC adapter and its focused public replay test differ from baseline.

## Validation and retained failures

The final source passes formatting, OPC/DOCX all-feature tests (1,996 passed,
32 ignored), no-default-feature tests (1,955 passed, 32 ignored), Clippy with
warnings denied, rustdoc with warnings denied, and crate boundaries. Six new
adapter cases bring its focused suite to 16; a public Stored/Deflate test
checks accepted output after a publication reader error. Twelve Python helper
tests pass. ASan fuzz validation passes 27 required-positive cases and two
10,000-run campaigns with seeds 484 and 485 and a 65,536-byte input cap.

The first verification lane passed all five allocator harness tests. The
final lane's parallel harness run failed the live-byte delta assertion in
`global_allocator_records_successful_alloc_and_dealloc_with_live_peak`.
The counter is process-global while the test mutex excludes only other test
bodies; concurrent test-runner deallocations can affect this delta. This is
the likely cause, not a proven reproduction. The retained serial retry
`bench-tests-serial1` passes all five tests without source changes. Both the
failed output and retry remain in the inventory; the harness race is open.

The first before capture failed before any benchmark launched because it
required the future after build. Historical helpers and protocol remain in
`development/protocol1`; selected-phase loading and its regression test fix
the new capture helper. The verifier independently binds the actual protocol
hash and validates before/after gates against their respective source builds.

## Remaining work

The non-iWork GOAL remains open, including broader CRUD/provider intersections,
cold-cache and concurrent behavior, and native Office validation. A read-only
lifecycle review identifies repeated candidate XML audits as a possible next
experiment. A standalone source proof cannot replace candidate validation.
Any future private candidate-audit capability must bind immutable candidate
bytes and frozen XML limits while retaining source/replay hashes, authenticated
EOF, freshness, cancellation, Work and memory accounting, sink counts, and
final DOCX reopen. No such audit reuse is implemented in this batch.
