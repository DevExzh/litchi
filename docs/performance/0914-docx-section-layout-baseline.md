# 0914 — DOCX existing-section layout baseline

This batch measures `docx_source_backed_existing_section_layout_edit_save`,
the unmeasured DOCX route identified by the [0906 opportunity
audit](results/change-0906/ooxml-opportunity.md). It establishes a current
baseline for an existing implementation. There is no candidate, matched eager
control, production change, or speedup claim.

## Measurements

Four fresh CPU-12 processes each execute 30 warmups and 200 measured iterations
on an AMD EPYC 9R45 Linux host, using Rust 1.95.0 and a locked release build
from `a4ca85e28cdea151176534b7f8dc912253a1ca1e`. All 800 measured samples and
the setup/preservation controls pass. The [reduction](results/change-0914/reduced.json)
retains individual child distributions and source/build/report identities.

| Metric | Result |
| --- | ---: |
| Pooled total p50 | 4.810184 ms |
| Pooled total p95 | 4.850650 ms |
| Pooled total p99 | 4.876280 ms |
| Range of four child total medians | 4.804199–4.816565 ms |
| Logical source reads per measured iteration | 202 calls / 8,409,488 bytes |
| Pre-publication successful main-Part loads | 1 |
| Output size / largest sink write | 2,101,446 / 65,536 bytes |
| Whole-child maximum RSS range | 36,672–38,056 KiB |

The pooled quantiles are descriptive. Four child medians provide a repeat
range, not a confidence interval for other machines or host conditions. RSS
includes fixture generation, warmups, verification, report serialization, and
teardown; it is not operation peak live memory. This is a shared host.

| Phase | Pooled median, ms | Share of summed timed nanoseconds |
| --- | ---: | ---: |
| Preparation | 0.034 | 0.75% |
| Open | 0.045 | 0.94% |
| Query | 1.331 | 27.68% |
| Edit | 0.053 | 1.12% |
| Commit | 0.345 | 7.22% |
| Publication | 2.996 | 62.29% |

Phase medians are rounded and need not sum to the total median. Shares use
the sum of each phase over all samples divided by the sum of all sample totals.
Query and publication together account for 89.98% of this measured interval;
their internal work needs attribution before choosing an optimization.
The 8,409,488 logical source bytes are about four times the 2,101,427-byte
archive size, but do not by themselves identify how many bytes were hashed,
decompressed, or copied in memory.

## Workload and timing boundary

The deterministic `litchi-docx-section-layout-source-v1` fixture has three
sections, 257 main-story paragraphs, a table-cell section sentinel, opaque
markup, header/footer relationships, and four 512 KiB media members. The edit
changes the columns of the section at public paragraph position 64 from two to
three. The other sections retain two columns.

Elapsed time is the sum of six separately measured intervals: source/sink
preparation, package open, section query, edit, commit, and sequential
publication. Preparation includes cloning the archive and reserving the sink.
The total excludes inter-phase gaps, correctness checks, and iteration teardown;
it is not whole-process elapsed time. Source counters describe the original
instrumented in-memory source, not physical storage traffic.
Each iteration constructs a fresh source and package inside a warmed process;
the post-timer checks also touch document data before the next iteration. This
is not a cold-cache experiment.

The existing harness checks semantic readback, table-cell sentinel exclusion,
raw untouched ZIP members, ordering and comments, header/footer relationships,
source immutability, forward/inverse patch behavior, and exact source restoration
through inverse publication. For untouched ZIP members, local spans must match
exactly; central records must match except for the relocated local-header offset.
Setup controls check exact no-op behavior and typed
foreign, stale, signed-change, section-limit, and partial-sink refusals. These
checks are outside the six timed intervals. See the [source-scope
audit](results/change-0914/source-scope.md).

## Build and capture corrections

The initial clean export omitted two compile-time inputs used by other harness
cases: `test-data/poi/test-data/spreadsheet/54016.xls` and
`test-data/rtf/watermark.rtf`. The failure is recorded in
[`custody-rejections.json`](results/change-0914/custody-rejections.json). Both inputs were
restored from the same pinned commit, and the next locked release build passed.
Neither is the measured DOCX input.

The first four measured children are rejected because their source/build
metadata was subsequently rewritten. Their original receipts and measurements
remain separate; they are not rebound to the newer metadata. The accepted four
children ran at 14:45:05–14:45:20 UTC on 2026-10-05, after the final metadata
was frozen, and their recorded protocol, build, and report hashes match.
Only this second cohort contributes to the tables above.

The retained `litchi-perf-baseline` executable is 69,940,744 bytes, SHA-256
`fe56e5192c1198e831583e870beb5e1ba4731fb2077faaffdfc46044bb0ed2db`.
The `source-HEAD-code.tar` archive is 171,120,640 bytes, SHA-256
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`.
Both remain outside the repository. To regenerate into a new external directory,
use the [capture helper](results/change-0914/capture.py) with
`LITCHI_PERF_0914_ROOT` set to that directory and
`LITCHI_PERF_0914_REVISION=a4ca85e28cdea151176534b7f8dc912253a1ca1e`.
Run its `prepare`, `freeze`, `build`, and `run` actions in that order. The build
is `cargo build --release --locked -j2` against the extracted
`tools/perf-baseline/Cargo.toml`, with an external `CARGO_TARGET_DIR`.

The [independent review](results/change-0914/review.md) checks all 9,357
extracted and archived source files against the pinned Git tree, including
contents and modes, with no differences. It also rejects deliberately corrupted
real reports covering identities, phase sums, sample order, cache/source counters,
and sink evidence. This verifies the measurement packet; no whole-workspace
test, lint, or producer-wide certification is claimed by this documentation
and tooling batch.

The committed reports and receipts support an external-artifact-free replay:

```sh
python3 -B docs/performance/results/change-0914/replay.py
```

This checks receipt hashes and ordering, report gates, source/output identities,
process metrics, phase sums, and independently recomputed quantiles. It does not
inspect the external executable or source archive. `capture.py verify` performs
the additional external custody checks. Replay was also tested with an absent
external root; corrupted pooled statistics and raw report bytes were rejected.

## Architecture and scope

| Constraint | Treatment |
| --- | --- |
| ADR 0003: source-bound edits and reversible patches | Existing source checks, semantic commit readback, and inverse restoration remain exercised. |
| ADR 0005: explicit I/O and measured performance | Only the external harness reads files, invokes processes, and measures time; library behavior is unchanged. Logical source counters and process metrics remain distinct. |
| ADR 0006: preservation and typed refusals | Existing byte-preservation and refusal controls remain required for every admitted child. |
| ADR 0010/0011/0024: format and physical ownership | No library, dependency, public API, or ownership changes. |
| ADR 0008: verification scope | This is one generated-fixture baseline, not workspace or producer-wide certification. |

The source shows full-artifact fingerprinting in the query and again during
publication, along with output fingerprinting. This is a candidate for
operation-local profiling, not a measured hashing cost share or permission to
remove source validation. Any optimization must preserve exact-source,
freshness, budget, and conflict behavior.

Real-producer diversity, filesystem/cold/range input, allocation and peak-live
metrics, hardware counters, copy/decompression accounting, and parallel scaling
remain outside this capture. The broader non-iWork goal remains active.
