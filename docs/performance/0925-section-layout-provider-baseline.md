# 0925 — DOCX section-layout provider baseline

This batch extends the existing section-layout benchmark with warm file and
controlled range providers. It changes the benchmark harness, not production
library behavior. The comparison is a provider baseline at base revision
`aa077c030819c3019288e350f34959666ab7204f`; it is not a before/after optimization
or a universal provider ranking.

**The sole 42-child run failed while saving its aggregate terminal receipt.**
All 42 raw reports and process logs remain retained, and no child was rerun or
replaced. The missing original terminal receipt invalidates the run's protocol
qualification. Offline analysis of those existing reports is descriptive only;
it cannot reconstruct the lost in-memory receipt as original evidence.

## Fixed measurement protocol

The protocol fixes seven routes across six alternating forward/reverse blocks:
42 fresh processes, each with five warmups and 50 measured samples. This yields
300 measured samples per route and 2,100 overall. Children run sequentially on
CPU 12 of the AMD EPYC 9R45 Linux host. Each child has a 900-second deadline;
observation polling does not authorize restarting a live process. Any terminal
failure invalidates the cohort, and failed or adverse rows cannot be replaced.

Pooled medians and nearest-rank p95/p99 use raw samples, not averages of child
quantiles. Per-child distributions remain visible. Phase shares divide summed
phase nanoseconds by summed total nanoseconds. GNU time peak RSS covers the
whole child, including corpus construction and correctness checks; it is not
operation-scoped allocation or retained memory. This is a descriptive baseline,
with no paired speedup test or adoption threshold.

## Descriptive results from the retained raw reports

The [descriptive recovery receipt](results/change-0925/comparison-recovery-final-v3.json)
summarizes the existing raw reports after the terminal receipt failure. These
results do not constitute a successful protocol qualification,
a provider ranking, or an optimization result. Each route contributes six
processes and 300 raw samples. RSS is a whole-process minimum–maximum range.

| Route | p50 ms | p95 ms | p99 ms | Child p50 range ms | Peak RSS range MiB |
| --- | ---: | ---: | ---: | ---: | ---: |
| Generic ReadAt | 4.488 | 4.528 | 4.580 | 4.485–4.495 | 35.11–35.67 |
| Owned vector | 3.712 | 3.741 | 3.770 | 3.701–3.726 | 35.38–35.96 |
| Owned reader | 3.848 | 4.077 | 4.102 | 3.831–4.072 | 35.61–38.38 |
| Warm tmpfs FileSource | 4.600 | 4.647 | 4.672 | 4.592–4.613 | 33.64–33.95 |
| Short range | 4.537 | 4.572 | 4.600 | 4.527–4.547 | 35.37–35.95 |
| Zero-fixed-delay paced range | 73.456 | 73.786 | 75.518 | 73.400–73.528 | 35.11–35.96 |
| 1 ms fixed-delay paced range | 238.517 | 239.068 | 240.447 | 238.497–238.546 | 35.64–35.95 |

Publication accounts for 60.27% of summed warm-file elapsed time and 59.57%
for the short-range route; query adds 29.79% and 29.70%, respectively. For the
paced range models, publication shares are 66.81% and 69.62%. These are shares
of measured elapsed time, including simulated service, rather than CPU profiles.
They do not identify a safe production optimization by themselves.

At the operation boundary, every new provider returns 6,308,061 bytes. The
file and paced range routes record 165 nonempty calls. The 4 KiB short-range
adapter records 1,607 calls and 1,506 short returns: callers request 84,096,471
bytes across repeated suffix buffers, while delegated requests and returned
bytes total 6,308,061. Requested-buffer volume is not actual transferred data.
The generic observer records 169 calls for the same returned-byte volume,
including empty calls; its historical checked-iteration count remains 202
calls and 8,409,488 bytes.

The paced transfer target is 60,158,393 ns per operation. The delayed route adds
165 configured 1 ms targets, yielding nominal service of 225,158,393 ns.
Actual elapsed time also includes work and scheduling; no difference between
these targets and the latency table is attributed to a specific CPU or I/O cost.

## Workload and provider boundaries

All routes use `litchi-docx-section-layout-source-v1`: a 2,101,427-byte DOCX
with 257 main-story paragraphs, three sections, and four 512 KiB media parts.
The selected paragraph at position 64 changes from two columns to three.
The source SHA-256 is
`183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8`;
expected output is 2,101,446 bytes with SHA-256
`5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`.

The six timed phases remain preparation, open, query, edit, commit, and
publication. Exact output, untouched ZIP members and ordering, section
semantics, inverse restoration, cache observations, and source checks run
outside those intervals. The shared stale/foreign/signed/no-op/limit/partial-sink
control flag describes the generic provider control suite, not separate
negative-case qualification of each new provider.

| Route | Provider setup |
| --- | --- |
| `generic_read_at` | Existing instrumented in-memory positional source. |
| `owned_vec` | Existing moved-vector ingress. |
| `owned_reader` | Existing instrumented sequential-reader ingress. |
| `file_warm` | Recently staged and synced `/tmp` tmpfs file; fresh `FileSource` opens during preparation. |
| `range_short` | Immutable memory, 4 KiB maximum delegated return, no pacing. |
| `range_zero` | Immutable memory, 64 KiB cap, 100 MiB/s nominal transfer, zero fixed delay. |
| `range_delayed` | Same transfer model and cap, plus 1 ms fixed delay per nonempty request. |

The range routes reuse `PptxRangeSource` with `MinimumService`: request work
and fixed-delay overshoot count toward the fixed-plus-transfer target.
Reported nominal service is configured fixed delay times delayed calls plus
per-request transfer targets. It does not measure actual sleeping or network
bandwidth. The filesystem route uses warm, recently written tmpfs storage on
this host; no cache flush or physical cold-read qualification is implied.

New provider observations stop immediately after publication, before reopen,
inverse, digest, and preservation checks. Logical caller requests and capped
delegated requests are separate boundaries. The direct file route has no
second independent boundary and emits null underlying observations.
These counters measure wrapper calls, not syscalls, disk traffic, or networking.
The generic observer includes empty calls; new provider counters exclude them.
The generic route also gains operation-end call/byte vectors while preserving
its historical whole-checked-iteration vectors. Owned routes do not fabricate
positional observations; new provider routes leave owned-ingress vectors empty.

File staging and file length/digest/byte validation are outside the timers.
Provider version checks use the same handle before the operation, at publication,
and after verification; version values are not serialized as report vectors.
The range backing is retained for a direct immutability check, and every route
must restore the exact corpus through inverse publication.

## Review and validation

Independent review verifies source SHA-256
`7904a6161e7b40a675d61acd9867b31b81a47f049b1ebb5a9979778bf14a373a`
and the [full harness patch](results/change-0925/section-layout-provider-final.patch).
The [quality receipt](results/change-0925/reviewer-quality-receipt-v5.json)
records two DOCX harness tests and thirteen existing range-adapter tests.
They cover the section/output/inverse/refusal controls and range short reads,
empty/EOF distinctions, version forwarding, pacing arithmetic, and counter
overflow. These tests ran on the semantically identical preformat version;
root independently confirmed exact formatter-output equality with final source.

The initial formatter check used an incomplete configuration and missed newly
introduced formatting differences. Independent checking with the actual
repository configuration passed the base and failed the candidate. This was
corrected before any performance capture. The first release build succeeded
but is superseded; its [receipt](results/change-0925/setup-v3-build.json),
[log](results/change-0925/setup-v3-build.log), and
[format failure](results/change-0925/setup-v3-format-failure.log) remain visible.
The final source passes Rust 1.95.0 formatting with the repository configuration.
The clean final build used `cargo +1.95.0 build --offline -j2 --locked --release`
in the standalone benchmark workspace with no profile environment overrides.
Its 70,187,352-byte binary has SHA-256
`7138c6c458dc13f1677deba3f413120ffc1af155d07848a243a43ebcf335feb7`.

All seven final-binary release preflights pass with one warmup and three measured
samples per route. They verify report shape, sample/phase alignment, output,
provider observations, and binary identity; they are separate correctness
preflights, not samples in the formal cohort. See the
[preflight validator receipt](results/change-0925/comparison-preflight-final.json)
and [build receipt](results/change-0925/measurement-build-v4-final.json).

## Terminal receipt failure and offline recovery

The frozen runner embedded each validator's raw sample/provider vectors in its
child receipt, then tried to write the aggregate `capture-state.json` under a
1 MiB writer limit. The final write failed; no original state file survives.
The [root failure receipt](results/change-0925/capture-v4/capture-failure.json)
records the terminal error and hashes all 42 existing reports. All GNU time
logs report child exit status zero and provide RSS, but the missing runner
return-code records, deadline flags, per-child timestamps, and post-capture
custody status cannot be restored as original evidence.

The post-capture tooling correction stores a compact scalar validation receipt
while retaining samples in their separately hashed raw reports. The
[actual-helper control](results/change-0925/comparison-postcapture-fix-controls-final-v2.json)
uses all 42 existing reports: its test state occupies 99,640 bytes, versus
1,992,796 bytes with the old full-vector shape. These are control fixtures,
not a reconstruction of the missing original receipt. No benchmark is rerun. The original
executed scripts remain unchanged in `capture-v4/scripts/`. This correction
supports future runs and does not repair this run's qualification.

Offline recovery validates the frozen scripts, source/build/protocol receipts,
report identities, phase sums, provider contracts, output hashes, and GNU time
logs. It explicitly retains invalid terminal status. The missing state remains
absent; recovery must never be presented as the original successful receipt.

## Reproduction and retained artifacts

Replay the descriptive analysis from the committed packet without building or
starting a benchmark process:

```sh
python3 -B docs/performance/results/change-0925/comparison-recover.py \
  --packet docs/performance/results/change-0925/capture-v4 \
  --failure-receipt docs/performance/results/change-0925/capture-v4/capture-failure.json \
  --authorized-manifest docs/performance/results/change-0925/provider-capture-manifest-v4-authorized.json \
  --output /tmp/0925-descriptive-recovery.json
```

The ordinary successful-capture replay rejects this packet because its original
terminal state is missing. Use the explicitly labeled recovery path above.
It validates existing evidence and is not permission to recapture the cohort.
The [root moved-packet replay](results/change-0925/root-moved-recovery-check.json)
exactly matches the canonical per-route timing, phase, provider and RSS results.
The [independent final approval](results/change-0925/comparison-final-approval.json)
confirms those results and the actual runner-helper control; it explicitly
retains the failed protocol status. Post-cleanup root replay also passes.

All individual committed artifacts are smaller than 1 MiB. The normalized
115,220,480-byte final source archive remains external:
`measurement-source-v4.tar`, SHA-256
`41ebda126104fe45e384176857419d38f92c42e4ac9bea238f309f1038a4022c`.
Its [descriptor and regeneration command](results/change-0925/measurement-source-v4-archive.json)
bind the source listing and reviewed patch. The release binary remains external;
its size, SHA-256, locks, compiler and build command are in the
[build receipt](results/change-0925/measurement-build-v4-final.json).

Cleanup removed 16,750,887,370 apparent bytes of obsolete mutable source,
build dependencies and formatting scratch. Both recorded binaries, source
archives, final source, raw reports, preflight evidence and receipts remain
available; see the [cleanup receipt](results/change-0925/cleanup-receipt-v1.json).

## Review and limits

Source review caught a redundant range backing-vector clone before release
measurement. The corrected setup shares the actual immutable backing through
`OwnedSource::from_arc`, so its retained reference checks the bytes used by the
provider. This was a premeasurement harness correction, not a measured production
optimization.

Routes include their declared setup and instrumentation costs. Cross-route
latencies do not isolate syscall overhead or scheduler delay.

This workload does not qualify real networking, cold storage, physical I/O,
operation-scoped allocation, concurrent scaling, or other document families.
The broader non-iWork performance goal remains active.
