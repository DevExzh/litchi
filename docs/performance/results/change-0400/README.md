# Change 0400 XLSX selected-cell streaming evidence

This directory retains the evidence for the scoped XLSX selected-cell
optimization described in the [production change record](../../changes/0400-xlsx-selected-dimension-streaming.md).
The measured candidate is cumulative: control is revision
`2e47ccebf449ef88943c0abcecd32bd9141eb520`, and candidate is revision
`f159c0aed603672aacee8e5923586ce4aa8753f7`. The candidate combines
dimension-bearing selected streaming with reusable numeric value scratch. The
ABBA results therefore do not isolate the contribution of either mechanism.

The selected path keeps a valid worksheet `<dimension>` in the streaming
fast path and validates its reference without using it to bound the query.
For numeric and untyped cells it reuses a private, bounded scratch buffer for
the cell value. The public setup is `litchi::Workbook::open`; workbook
preparation is outside the timed operation. The timed operation is only the
case-insensitive sheet selection and exact `M29` cell query.

## Fixed corpus and oracle

All eight retained reports use the same deterministic XLSX corpus:

- generator: `litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`
- four sheets, each 48 rows by 48 columns: 9,216 cells total
- 17 ZIP members, 4,226,429 compressed bytes, deflate compression
- archive SHA-256: `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`
- canonical sheet `Bench01`, selected with `bEnCh01`, cell `M29`
- selected Number lexical value: `1028012`
- semantic corpus SHA-256: `020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e`
- selected-cell evidence digest: `36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`

The reports use schema version 1. Their strict corpus identity and selected-cell
oracle are embedded in each report and are checked again by the summary. No
schema-2 corpus-catalog sidecar was requested for this package; there are no
catalog files to validate or claim.

## Normal ABBA result

The normal release binary was measured in strict `A1_control`,
`B1_candidate`, `B2_candidate`, `A2_control` order. Each leg used 20
warmups and 500 retained samples, a fresh process per sample, process-isolated
filesystem evidence, the selected filesystem root, a warm cache, one
execution worker, CPU affinity `2`, and Rust 1.98.1 on the AMD EPYC 9R45
host. Positive values mean that the candidate was faster.

| Statistic | A1 to B1 reduction | A2 to B2 reduction |
| --- | ---: | ---: |
| p50 | +27.775881% | +27.711459% |
| mean | +27.728990% | +27.691341% |
| p95 | +27.657228% | +27.705070% |
| p99 | +28.150563% | +27.835790% |

The four statistics were accepted by the fail-closed summary. Same-
implementation drift stayed below the configured ceilings of 5% for p50 and
mean, 10% for p95, and 15% for p99. Candidate drift was 0.133415%/0.130111%/
0.151022%/0.715051% for p50/mean/p95/p99; control drift was
0.044177%/0.077977%/0.217298%/0.275743%.

`summary.json` contains the normal-leg metadata, sample counts, statistics,
oracle bindings, and recomputed ABBA math; the raw frames retain every sample.
The standard package manifest
`0400-xlsx-selected-cell-streaming-abba-manifest.json` binds that summary and
the four normal report frames:
`a1-normal.json.zst`, `b1-normal.json.zst`, `b2-normal.json.zst`, and
`a2-normal.json.zst`.

## Allocator ABBA evidence

The allocator run used separate release binaries instrumented with the Rust
system allocator counter. It used the same four-leg order, warm cache, one
worker, CPU affinity `2`, fresh child per sample, and process isolation, with
three warmups and 30 retained samples per leg. Allocator elapsed time is
observational only; the instrumentation changes elapsed time and is not a
second latency claim.

Every retained sample had the same operation-scoped vector within each
implementation and across its two legs:

| Metric | Control A1/A2 | Candidate B1/B2 | Candidate minus control |
| --- | ---: | ---: | ---: |
| allocation calls | 100,992 | 84,221 | -16,771 (-16.6063%) |
| deallocation calls | 98,642 | 84,206 | -14,436 (-14.6347%) |
| reallocation calls | 38 | 12 | -26 (-68.4211%) |
| failed allocation calls | 0 | 0 | 0 |
| allocated bytes | 13,925,077 | 10,706,565 | -3,218,512 (-23.1131%) |
| deallocated bytes | 13,412,029 | 10,705,182 | -2,706,847 (-20.1822%) |

The paired setup baselines differ by four live-before bytes and peak-live-
before bytes. Live-after is lower by 511,661 bytes for the candidate, while
peak-live-after is higher by 371 bytes; those absolute process live and
high-water snapshots are reported for transparency and are not memory claims.
The four `*-allocator.json.zst` reports retain the full per-sample vectors;
`allocation-metrics.json` is the compact projection of those vectors.

## Integrity checks

Check every retained zstd frame, then verify the normal package manifest's
compressed and decompressed identities. The package manifest is deliberately
self-excluding and records both compressed and raw byte counts and SHA-256
digests for each normal frame.

```sh
set -eu
result_dir=docs/performance/results/change-0400
check_dir=$(mktemp -d /tmp/litchi-0400-evidence-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT

for profile in normal allocator; do
  for leg in a1 b1 b2 a2; do
    zstd -q -t "$result_dir/$leg-$profile.json.zst"
  done
done

zstd -q -d -c "$result_dir/a1-normal.json.zst" > "$check_dir/a1-normal.json"
zstd -q -d -c "$result_dir/b1-normal.json.zst" > "$check_dir/b1-normal.json"
zstd -q -d -c "$result_dir/b2-normal.json.zst" > "$check_dir/b2-normal.json"
zstd -q -d -c "$result_dir/a2-normal.json.zst" > "$check_dir/a2-normal.json"
python3 tools/perf_abba_summary.py \
  --case xlsx_file_selected_cell \
  "$check_dir/a1-normal.json" "$check_dir/b1-normal.json" \
  "$check_dir/b2-normal.json" "$check_dir/a2-normal.json" \
  --json-out "$check_dir/recomputed-summary.json"
python3 -m json.tool "$result_dir/summary.json" >/dev/null
python3 -m json.tool \
  "$result_dir/0400-xlsx-selected-cell-streaming-abba-manifest.json" >/dev/null
python3 - "$result_dir" "$check_dir/recomputed-summary.json" <<'PY'
import hashlib
import json
import pathlib
import subprocess
import sys

root = pathlib.Path(sys.argv[1])
recomputed = json.loads(pathlib.Path(sys.argv[2]).read_text())
summary_path = root / "summary.json"
summary_bytes = summary_path.read_bytes()
summary = json.loads(summary_bytes)
manifest = json.loads(
    (root / "0400-xlsx-selected-cell-streaming-abba-manifest.json").read_text()
)

def digest(data):
    return hashlib.sha256(data).hexdigest()

def canonical(value):
    return json.dumps(
        value, sort_keys=True, separators=(",", ":"), ensure_ascii=True
    ).encode()

assert summary == recomputed
summary_identity = manifest["summary"]
assert len(summary_bytes) == summary_identity["bytes"]
assert digest(summary_bytes) == summary_identity["sha256"]
summary_canonical = canonical(summary)
assert len(summary_canonical) == summary_identity["canonical_bytes"]
assert digest(summary_canonical) == summary_identity["canonical_sha256"]

for artifact in manifest["artifacts"]:
    path = root / artifact["path"]
    compressed = path.read_bytes()
    assert len(compressed) == artifact["bytes"]
    assert digest(compressed) == artifact["sha256"]
    raw = subprocess.run(
        ["zstd", "-q", "-d", "-c", str(path)],
        check=True,
        stdout=subprocess.PIPE,
    ).stdout
    assert len(raw) == artifact["uncompressed_bytes"]
    assert digest(raw) == artifact["uncompressed_sha256"]

print("verified summary and four normal package artifacts")
PY
```

The final `evidence-manifest.json` additionally binds the four allocator
frames, `allocation-metrics.json`, `adjudication.json`, this README, and all
other textual projections. For each compressed artifact in that manifest,
check its recorded compressed and raw values with the same frame test and
decompression pattern:

```sh
set -eu
result_dir=docs/performance/results/change-0400
check_dir=$(mktemp -d /tmp/litchi-0400-raw-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT
for profile in normal allocator; do
  for leg in a1 b1 b2 a2; do
    compressed="$result_dir/$leg-$profile.json.zst"
    raw="$check_dir/$leg-$profile.json"
    zstd -q -t "$compressed"
    zstd -q -d -c "$compressed" > "$raw"
    stat -c '%n %s bytes' "$compressed" "$raw"
    sha256sum "$compressed" "$raw"
  done
done
python3 tools/validate_perf_allocator_abba.py \
  --a1 "$check_dir/a1-allocator.json" \
  --b1 "$check_dir/b1-allocator.json" \
  --b2 "$check_dir/b2-allocator.json" \
  --a2 "$check_dir/a2-allocator.json" \
  --projection "$result_dir/allocation-metrics.json"
python3 -m json.tool "$result_dir/allocation-metrics.json" >/dev/null
python3 -m json.tool "$result_dir/adjudication.json" >/dev/null
python3 -m json.tool "$result_dir/evidence-manifest.json" >/dev/null
```

Compare those `stat` and `sha256sum` lines with the corresponding
`compressed_bytes`, `raw_bytes`, `compressed_sha256`, and `raw_sha256`
fields in `evidence-manifest.json`; parse `allocation-metrics.json` and
`adjudication.json` with `python3 -m json.tool` as well. The allocator
validator recomputes every operation vector and projection delta but does not
evaluate allocator elapsed statistics. No Cargo invocation is required to
inspect this evidence package.

## Scope and exclusions

The authorized result is limited to the warm normal ABBA timing of this exact
filesystem-backed selected-cell operation and the exact allocator operation
vectors above. It does not claim:

- allocator-enabled elapsed time, cold-cache behavior, throughput, or physical
  I/O;
- RSS, peak operation memory, or a reduction in the peak-live snapshots;
- whole-workbook/open timing, ranges, other selectors, other XLSX shapes, or
  general XLSX workloads;
- other workbook facades, formats, execution-worker counts, or hosts; or
- an independently attributable speedup for numeric scratch versus dimension
  streaming.

Preliminary captures made before dimension-bearing streaming was enabled
fell back to eager parsing and showed no scratch-path effect. Captures made
with the wrong toolchain or concurrent host load are rejected and are not
part of this package.
