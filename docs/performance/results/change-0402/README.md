# Change 0402 OPC source-overlay publication evidence

This directory retains the audited evidence for Change 0402, the
`litchi-opc` source-backed no-op overlay validation change. The candidate
reuses one indexed decoder session while validating an unmanaged multi-part
overlay, while the harness retains the publication-only allocation envelope.
The control is revision
`46ef44966d5be16f153b1f3375ac14401b7139ac`; the candidate is revision
`51964019db3f6b0787645e3a56c2ecb83bdca65c`.

The result is intentionally scoped. The normal ABBA summary accepts only the
cell/statistic combinations listed below, and the allocator run contributes
exact operation-scoped call/byte vectors only. No overall matrix speedup is
claimed.

## Fixed corpus and operation

The harness used the deterministic generator
`litchi-opc-source-overlay-multi-part-v1`. Each source is an OPC/ZIP archive
with 32 ordinary entries plus the two OPC metadata members. The target entry
is `benchmark/parts/00016.bin`; each replacement is non-empty and has the same
payload as the source, so the semantic operation is a verified no-op.

| Shape | Payload | Entry bytes | Uncompressed payload | Archive bytes | Archive SHA-256 | Overlay counts |
| --- | --- | ---: | ---: | ---: | --- | --- |
| `overlay-small` | compressible | 1,024 | 32,768 | 7,451 | `4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509` | 2, 8, 32 |
| `overlay-large` | incompressible | 65,536 | 2,097,152 | 2,103,195 | `8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73` | 2, 8, 32 |
| `overlay-media-incompressible` | incompressible | 262,144 | 8,388,608 | 8,396,580 | `bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883` | 2, 8, 32 |

The corpus, source/output, sink, semantic no-op, raw-member ordering, and
phase-sum identities are repeated in every retained report and checked by the
custom validators. `0402-opc-source-overlay-abba-manifest.json` records the
compressed and raw identity of all eight report frames;
`evidence-manifest.json` additionally records every projection and text file.

## Normal ABBA result

The normal release binary was measured in `A1_control`, `B1_candidate`,
`B2_candidate`, `A2_control` order. Each leg has 20 warmups and 500 retained
in-process samples on CPU 2 with one execution worker. The selected cache
state is warm. The reports retain the harness's global configuration list
`["warm", "cold-requested"]`, but that list does not prove a cold run. The
normal claim therefore does not assert fresh-child or process-isolated
semantics.

The timed metric is only
`source.opc_source_overlay.publication_ns`, the publication call that writes
the overlay output. For each sample, the top-level `elapsed_ns` value is
checked as the exact sum of `preparation_ns`, `open_ns`, `planning_ns`, and
`publication_ns`. It is retained as a phase-sum oracle and is never
summarized or claimed. Reopen, digest, preservation, cache-probe, configured
sink/cache ceilings, and semantic checks are also outside the publication
timer.

The fail-closed policy accepts a statistic only when both paired candidate
reductions are positive and both same-implementation drifts are within the
5%/5%/10%/15% (`p50`/`mean`/`p95`/`p99`) ceilings. The accepted matrix is:

| Shape | 2 overlays | 8 overlays | 32 overlays |
| --- | --- | --- | --- |
| `overlay-small` | none | `p50`, `mean`, `p95`, `p99` | `p50`, `mean`, `p95`, `p99` |
| `overlay-large` | none | none | `p50` only |
| `overlay-media-incompressible` | `p50`, `mean`, `p95`, `p99` | none | none |

`summary.json` is the byte-for-byte output of
`tools/perf_opc_overlay_abba_summary.py` for the four retained normal
reports. It preserves all four statistics for every cell, including rejected
cells and adverse directions. `latency-metrics.tsv` is a deterministic,
publication-only projection of those values, reductions, drifts, and
per-cell authorization decisions. The complete reasons are in
`adjudication.json`.

In particular, the summary must not be read as an overall 0402 latency
result. The small/2 and large/2 cells are rejected, media-incompressible/2
accepts all four statistics, large/32 authorizes only p50, and every
cell/statistic not listed in the accepted matrix is rejected. The rejected
cells include same-implementation drift failures as well as non-positive paired
reductions.

## Allocator observation

The allocator release binary uses
`CountingSystemAllocator(std::alloc::System)`. It has three warmups and 30
retained samples per ABBA leg. These reports are retained for operation-vector
alignment and exact allocation observations. The allocator elapsed vectors,
absolute live-byte snapshots, and absolute high-water snapshots are
non-claimable because the instrumentation changes elapsed behavior and those
snapshots are not operation memory measurements. The allocator reports also
contain no per-sample PID/envelope proof; no independent-process claim is
made, even though the global harness envelope names fresh-child and
process-isolated configuration fields.

The six claimable metrics are ordered as follows:

`allocation_calls`, `deallocation_calls`, `reallocation_calls`,
`failed_allocation_calls`, `allocated_bytes`, `deallocated_bytes`.

Every shape has the same exact vector for a given overlay count. The validator
still checks all nine shape/count cells independently. The control vector,
candidate vector, and candidate-minus-control delta are:

| Overlay count | Control | Candidate | Candidate minus control |
| ---: | --- | --- | --- |
| 2 | `[52, 338, 0, 0, 229884, 259086]` | `[50, 336, 0, 0, 149564, 178766]` | `[-2, -2, 0, 0, -80320, -80320]` |
| 8 | `[90, 388, 0, 0, 720314, 756104]` | `[76, 374, 0, 0, 158074, 193864]` | `[-14, -14, 0, 0, -562240, -562240]` |
| 32 | `[236, 582, 0, 0, 2681938, 2744080]` | `[174, 520, 0, 0, 192018, 254160]` | `[-62, -62, 0, 0, -2489920, -2489920]` |

These are exact observations for this fixed operation and corpus. They are
not a law per overlay, a memory reduction, or a prediction for other
corpora. `allocation-metrics.json` is the compact projection produced by
`tools/validate_opc_overlay_allocator_abba.py` from the four raw allocator
reports and `allocator-contract.json`.

The allocator binary identity is `reported+contract-bound` and
`file_rehashed` is `false`: the binary files may have been removed after
collection. The contract nevertheless binds each report to the recorded
revision, SHA-256, byte count, mode bits, and release profile. The normal
report envelope carries the corresponding normal binary identities.

## Retained artifacts

The eight report files are deterministic zstd frames made with `/usr/bin/zstd`
v1.5.7, compression level 3, one thread, XXH64 frame checksum, and content
size enabled:

- `a1-normal.json.zst`, `b1-normal.json.zst`, `b2-normal.json.zst`,
  `a2-normal.json.zst`
- `a1-allocator.json.zst`, `b1-allocator.json.zst`, `b2-allocator.json.zst`,
  `a2-allocator.json.zst`

`0402-opc-source-overlay-abba-manifest.json` is the package-level manifest
for those eight frames and `summary.json`. `evidence-manifest.json` is the
final self-excluding evidence manifest. It binds every retained byte stream,
its SHA-256, and (for JSON or compressed frames) its canonical or decompressed
identity. It also binds the measured source revisions and validator files.

The normal publication validator is bound to revision
`8720c9f103243903fcb1047eaaf3d384e964bfa5` and the exact source hash recorded
in both manifests. The allocator validator is bound to the `db29f7705`
lineage and its current source hash; that lineage includes the cache-envelope
acceptance fix used by the retained raw reports.

## Reproduction and integrity checks

Run from the repository root after checkout:

```sh
set -eu
result_dir=docs/performance/results/change-0402
check_dir=$(mktemp -d /tmp/litchi-0402-bundle-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT

for profile in normal allocator; do
  for leg in a1 b1 b2 a2; do
    compressed="$result_dir/$leg-$profile.json.zst"
    raw="$check_dir/$leg-$profile.json"
    /usr/bin/zstd -q -t "$compressed"
    /usr/bin/zstd -q -d -c "$compressed" > "$raw"
  done
done

python3 tools/perf_opc_overlay_abba_summary.py \
  --a1 "$check_dir/a1-normal.json" \
  --b1 "$check_dir/b1-normal.json" \
  --b2 "$check_dir/b2-normal.json" \
  --a2 "$check_dir/a2-normal.json" \
  --json-out "$check_dir/recomputed-summary.json"
cmp "$result_dir/summary.json" "$check_dir/recomputed-summary.json"

python3 tools/validate_opc_overlay_allocator_abba.py \
  --contract "$result_dir/allocator-contract.json" \
  --a1 "$check_dir/a1-allocator.json" \
  --b1 "$check_dir/b1-allocator.json" \
  --b2 "$check_dir/b2-allocator.json" \
  --a2 "$check_dir/a2-allocator.json" \
  --output "$check_dir/recomputed-allocation.json"
cmp "$result_dir/allocation-metrics.json" "$check_dir/recomputed-allocation.json"

python3 -m json.tool "$result_dir/summary.json" >/dev/null
python3 -m json.tool "$result_dir/allocator-contract.json" >/dev/null
python3 -m json.tool "$result_dir/allocation-metrics.json" >/dev/null
python3 -m json.tool "$result_dir/adjudication.json" >/dev/null
python3 -m json.tool "$result_dir/0402-opc-source-overlay-abba-manifest.json" >/dev/null
python3 -m unittest tools.test_perf_opc_overlay_abba_summary
python3 -m unittest tools.test_validate_opc_overlay_allocator_abba
```

The manifest verification should additionally confirm that
`evidence-manifest.json` does not list itself, all listed byte counts and
SHA-256 values match, every zstd frame tests successfully, and each frame
decompresses to the declared raw and canonical identities. The two validator
test suites are standard-library-only and do not require the discarded binary
or corpus staging roots.

## Scope and withheld claims

This evidence authorizes only the explicitly listed warm normal
`publication_ns` cell/statistic entries and the exact six-metric allocator
vectors. It does not authorize:

- an overall 0402 matrix latency reduction, or any rejected cell/statistic;
- top-level `elapsed_ns`, allocator-instrumented elapsed time, or a phase-sum
  latency claim;
- RSS, live bytes, peak live bytes, total-memory behavior, or a memory-saving
  claim;
- fresh-child, process-isolated, cold-cache, cold-verified, or cache-transition
  behavior;
- physical I/O, filesystem read volume, locality, throughput, concurrency, or
  worker-count scaling;
- other OPC shapes, overlay counts, payloads, formats, platforms, or broad
  unified-facade behavior; or
- results from non-retained, stale, wrong-toolchain, or contaminated captures.
