# Change 0401 XLSX selected numeric ownership-elision evidence

This directory retains the audited evidence for Change 0401, the narrow XLSX
selected-cell scanner optimization that validates an unselected numeric lexical
value by borrow and avoids constructing an owned `Number`. The control is
Change 0400's final revision `0859063be5a67bd2aafb3531f2126020b2b5000d`; the
candidate is `87f26d5ee02a1903e668bf7f60fa3ef954a0c3fb`.

The candidate keeps ownership for selected values, formula caches, inline
strings, shared strings, and all other paths that need semantic output. The
borrowed validation shares the lexical finite-number checks with `Number::new`.
The focused regression cases cover valid selected and unselected values,
malformed and non-finite values, formula-cache ownership, and inline validation
(including malformed, overlong, and unpaired-surrogate input).

The [Change 0400 baseline evidence](../change-0400/) records the control's
earlier dimension-bearing streaming and numeric-scratch work. The [normal ABBA
summary helper](../../../../tools/perf_abba_summary.py) and [Change 0401 allocator
validator](../../../../tools/validate_perf_allocator_abba_0401.py) are bound by
hash in `evidence-manifest.json`.

## Fixed corpus and oracle

All retained reports use one deterministic filesystem-backed XLSX corpus:

- generator: `litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`
- shape: `medium`; four sheets, 48 rows by 48 columns
- 9,216 workbook cells and 17 ZIP members
- archive bytes: `4,226,429`; archive SHA-256:
  `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`
- uncompressed payload bytes: `4,231,168`; deflate compression
- canonical sheet: `Bench01` at zero-based position 1
- prepared case-insensitive selector: `bEnCh01`
- exact cell: `M29` (zero-based row 28, column 12)
- selected view and value kind: stored Number
- selected lexical value: `1028012`
- selected-cell oracle digest:
  `36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`
- semantic workbook SHA-256:
  `020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e`

The normal and allocator raw samples independently carry this corpus and
selected-cell identity. The normal reports carry the semantic oracle too; all
2,000 normal samples and all 120 allocator samples retained the expected
selected value and oracle identities.

## Normal ABBA result

The normal release binary was measured in strict `A1_control`, `B1_candidate`,
`B2_candidate`, `A2_control` order. Each leg used 20 warmups and 500 retained
samples, a fresh child process per sample, process-isolated filesystem
measurement, a warm cache, one execution worker, CPU affinity `2`, and Rust
1.98.1 on the AMD EPYC 9R45 host. The timed operation is only case-insensitive
worksheet selection plus the exact `M29` read; workbook open, selector
preparation, and post-operation oracles are outside the timer.

Positive values mean that the candidate was faster. The fail-closed ABBA
adjudication accepts only statistics that are lower in both paired directions:

| Statistic | A1 to B1 reduction | A2 to B2 reduction |
| --- | ---: | ---: |
| mean | +0.099577940251% | +0.026562239637% |
| p95 | +0.625379111895% | +0.198122423529% |
| p99 | +1.170167332729% | +0.045344544337% |

The p50 result is explicitly rejected: the paired reductions are
`-0.012690677428%` (A1 to B1) and `-0.035254218167%` (A2 to B2), so the
candidate is not lower in both directions. Same-implementation drift remained
inside the configured ceilings (5% for mean and p50, 10% for p95, and 15% for
p99). `latency-metrics.tsv` is a deterministic projection of the four raw
normal frames and includes both accepted and rejected statistics.

`summary.json` is byte-identical to the supplied `normal-abba-summary.json`.
The normal package manifest is retained as
`0401-xlsx-selected-numeric-elision-abba-manifest.json`; its two summary path
fields are normalized from `normal-abba-summary.json` to the retained
`summary.json` alias, with the summary bytes and digest unchanged.

## Allocator observation

The allocator run used a separate release binary with
`CountingSystemAllocator(std::alloc::System)`, the same warm cache, fresh-child
and process-isolated setup, CPU affinity `2`, and the same ABBA order. It used
three warmups and 30 retained samples per leg. The allocator timer is
observational only and is never used for latency acceptance.

Every retained sample has a constant vector within each implementation and the
control and candidate vectors are equal across their two legs. The six
operation-scoped call/byte metrics are ordered as follows:

| Metric | Control | Candidate | Candidate minus control |
| --- | ---: | ---: | ---: |
| allocation calls | 84,221 | 81,918 | -2,303 |
| deallocation calls | 84,206 | 81,903 | -2,303 |
| reallocation calls | 12 | 12 | 0 |
| failed allocation calls | 0 | 0 | 0 |
| allocated bytes | 10,706,565 | 10,690,444 | -16,121 |
| deallocated bytes | 10,705,182 | 10,689,061 | -16,121 |

These are exact operation-scoped observations for this selector and corpus.
The scan has one selected cell and 2,303 unselected numeric cells, but the
observed 16,121-byte delta is not generalized as a 7-byte-per-cell rule or a
proportional saving. The retained process live-before/live-after and
peak-live-before/peak-live-after snapshots, allocator elapsed time, RSS, and
any other process-wide memory behavior are non-claimable.

## Retained artifacts

The eight compressed frames are copied mechanically from the audited package:
`a1-normal.json.zst`, `b1-normal.json.zst`, `b2-normal.json.zst`,
`a2-normal.json.zst`, and the corresponding four allocator frames. The package
also retains `summary.json`, `allocation-metrics.json`,
`latency-metrics.tsv`, `adjudication.json`, and the normalized normal package
manifest. `evidence-manifest.json` is the final self-excluding manifest; it
binds every retained file with byte counts and SHA-256 values, plus the
compressed-frame raw identities and the relevant validator, test, and source
file hashes.

The source binding covers the three changed candidate/control files:
`crates/litchi-xlsx/src/cell.rs`,
`crates/litchi-xlsx/src/raw/worksheet/selected.rs`, and
`crates/litchi-xlsx/src/raw/worksheet/tests.rs`. The validator binding covers
`tools/perf_abba_summary.py`, `tools/test_perf_abba_summary.py`,
`tools/validate_perf_allocator_abba_0401.py`, and
`tools/test_validate_perf_allocator_abba_0401.py`.

## Deterministic integrity and reproduction checks

Run from the repository root after checkout:

```sh
set -eu
result_dir=docs/performance/results/change-0401
check_dir=$(mktemp -d /tmp/litchi-0401-evidence-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT

for profile in normal allocator; do
  for leg in a1 b1 b2 a2; do
    compressed="$result_dir/$leg-$profile.json.zst"
    raw="$check_dir/$leg-$profile.json"
    zstd -q -t "$compressed"
    zstd -q -d -c "$compressed" > "$raw"
  done
done

python3 tools/perf_abba_summary.py \
  --case xlsx_file_selected_cell \
  "$check_dir/a1-normal.json" "$check_dir/b1-normal.json" \
  "$check_dir/b2-normal.json" "$check_dir/a2-normal.json" \
  --json-out "$check_dir/recomputed-summary.json"

python3 - "$result_dir" "$check_dir/recomputed-summary.json" <<'PY'
import hashlib
import json
import pathlib
import subprocess
import sys

root = pathlib.Path(sys.argv[1])
recomputed = json.loads(pathlib.Path(sys.argv[2]).read_text())
summary = json.loads((root / "summary.json").read_text())
assert summary == recomputed
package = json.loads((root / "0401-xlsx-selected-numeric-elision-abba-manifest.json").read_text())
evidence = json.loads((root / "evidence-manifest.json").read_text())
assert evidence["self_excluded"] is True
assert all(item["path"] != "evidence-manifest.json" for item in evidence["artifacts"])

def digest(data):
    return hashlib.sha256(data).hexdigest()

def canonical(value):
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True).encode()

summary_bytes = (root / "summary.json").read_bytes()
assert len(summary_bytes) == package["summary"]["bytes"]
assert digest(summary_bytes) == package["summary"]["sha256"]
assert len(canonical(summary)) == package["summary"]["canonical_bytes"]
assert digest(canonical(summary)) == package["summary"]["canonical_sha256"]

for item in package["artifacts"]:
    path = root / item["path"]
    compressed = path.read_bytes()
    assert len(compressed) == item["bytes"]
    assert digest(compressed) == item["sha256"]
    raw = subprocess.check_output(["zstd", "-q", "-d", "-c", str(path)])
    assert len(raw) == item["uncompressed_bytes"]
    assert digest(raw) == item["uncompressed_sha256"]

for item in evidence["artifacts"]:
    path = root / item["path"]
    data = path.read_bytes()
    assert len(data) == item["bytes"]
    assert digest(data) == item["sha256"]
    if item.get("encoding") == "zstd":
        raw = subprocess.check_output(["zstd", "-q", "-d", "-c", str(path)])
        assert len(raw) == item["raw_bytes"]
        assert digest(raw) == item["raw_sha256"]

for name in ("summary.json", "allocation-metrics.json", "adjudication.json", "0401-xlsx-selected-numeric-elision-abba-manifest.json"):
    json.loads((root / name).read_text())
print("verified summary, package manifest, eight zstd frames, and self-excluding evidence manifest")
PY

python3 tools/validate_perf_allocator_abba_0401.py \
  --a1 "$check_dir/a1-allocator.json" \
  --b1 "$check_dir/b1-allocator.json" \
  --b2 "$check_dir/b2-allocator.json" \
  --a2 "$check_dir/a2-allocator.json" \
  --projection "$result_dir/allocation-metrics.json"

LITCHI_0401_EVIDENCE="$check_dir" \
  python3 -m unittest tools.test_validate_perf_allocator_abba_0401
```

The deterministic unit checks are:

```sh
python3 -m unittest tools.test_perf_abba_summary
```

The allocator validator must report four reports, 120 samples, 120 unique
fresh child processes, and 40 exact allocator vectors. It rejects allocator
latency statistics and non-claimable live/peak projections.

## Scope and exclusions

The authorized result is limited to the accepted warm normal mean/p95/p99
statistics and the exact operation-scoped allocator call/byte vectors for this
single selected-cell operation. It does not claim:

- normal p50 latency reduction;
- allocator-enabled elapsed latency;
- RSS, peak operation memory, total-memory behavior, or a memory reduction;
- physical I/O, filesystem read volume, locality, throughput, or scaling;
- cold-cache, cold-verified, or cache-transition behavior;
- workbook-open or selector-preparation timing;
- other cells, worksheet selectors, corpus shapes, payloads, or XLSX workloads;
- broad XLSX, unified-facade, other-format, cross-platform, or worker-count
  generalization;
- a 7-byte-per-cell law or any proportional saving derived from this corpus; or
- any result from non-retained, wrong-toolchain, preflight, or contaminated
  captures.

The result is an incremental comparison against the Change 0400 control and
makes no claim about the independent effects of earlier dimension-bearing
streaming or numeric scratch work.
