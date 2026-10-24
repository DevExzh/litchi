# 0362umd — OPC managed read-ahead endpoint ABBA lane

- Status: protocol prepared; no build, capture, or timing has been run for this
  lane.
- Owner: the OPC read-ahead retention repair (0362umc) and its coordinator.
- Scope: compare the retained-source-window repair at the exact source pair
  below. The forward-start arm measures the configured `Weak` upgrade path; the
  exact arm is a control and must not configure or upgrade a read-ahead window.
- Evidence: all build receipts, reports, and summaries stay in an external
  scratch directory. This record contains the reproducible driver contract and
  the small manifest schema only.

## 1. Frozen source pair and shared inputs

The baseline is the parent of the repair. The candidate is the frozen reviewed
merge, not the repair branch tip; the repair is an ancestor of that merge.

| endpoint | revision | role |
|---|---|---|
| A | `410bd99f00fdfb4161a94215710340f16ea536bb` | baseline |
| B | `49c25ef67f62a8bb943e7c6a0bf464e49ae87d07` | candidate, frozen merge |

The selector and its build inputs are byte-identical at both endpoints:

| path | SHA-256 |
|---|---|
| `tools/perf-baseline/Cargo.toml` | `c944031ad9afbfef854cac48d400d30fdf6135e3f2c83d105bbf014f1d3dab40` |
| `tools/perf-baseline/Cargo.lock` | `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184` |
| `tools/perf-baseline/src/docx_managed_read_ahead.rs` | `6592815ab69ea362d0163b4548d141c4e7e07cdd1f387f7075aaa18f207fe75f` |
| `tools/perf-baseline/src/main.rs` | `b51dd8de3030ee9fa9be7676b05d36a1203172b91d8b28c4bee8a07dfb65447e` |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |

The Rust toolchain file pins Rust `1.95.0`. The binary is built once per
endpoint and reused for both of its ABBA legs; the manifest records its byte
count and SHA-256 after the build. A dirty tree, a source revision mismatch, a
different lock/toolchain/input digest, or a changed build environment aborts
the run before any report is accepted.

## 2. Existing driver and fixed policy

`tools/perf-baseline/src/docx_managed_read_ahead.rs`, selected by
`docx-managed-read-ahead`, is the driver. No second executable is needed. It
generates and checks the pinned `source-edit-media-v1` archive (16,793,036
bytes, 20 members, archive SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`) and the
10,000-byte extracted-text oracle (SHA-256
`ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af`). Its
timing scope includes managed package open, the production window
construction, document materialization, text extraction, diagnostics, and
package/document drop. Hashing, oracle checks, source-version checks and
post-drop budget checks remain outside the clock.

Both arms use the same finite limits, `--max-range 65536`, `--delay-us 0`,
`--samples 30`, and `--warmup 3`. No transfer pacing flags may be supplied:
zero artificial delay means both the fixed range delay and transfer delay are
absent. The two policy invocations are:

```text
forward-start: --policy forward-start --window-bytes 4096
exact:         --policy exact                         (no window argument)
```

The forward report must publish enabled source-read diagnostics configured for
4,096 bytes. The exact report must publish no source-read diagnostics. Every
report must also pass the existing harness checks: one payload load, unchanged
source version, exact text oracle, physical-returned/input-byte agreement, and
zero memory/object usage after package drop.

This is the ordinary `litchi-perf-baseline` binary only. Allocator-instrumented
reports are a different build and are not mixed into this latency lane. If an
allocator qualification is later requested, it gets a separate protocol and
binary identity rather than contaminating these ABBA legs.

## 3. Reproducible build and capture commands

The coordinator chooses two clean endpoint checkouts and one idle CPU. The
following variables are placeholders, not paths used while preparing this
record:

```sh
BASE_TREE=/external/checkout-410bd
CAND_TREE=/external/checkout-49c
BASE_TARGET=/external/target-opc-ra-410bd
CAND_TARGET=/external/target-opc-ra-49c
EVIDENCE=/external/evidence-opc-ra-<unique-run>
CPU=<one-idle-cpu>
CPU_LOCK=/external/locks/opc-read-ahead.cpu.lock
```

Build each endpoint separately, with the same pinned toolchain, dependency
lock, profile and flags. Separate target directories prevent a cross-endpoint
incremental or executable collision. `RUSTFLAGS` is deliberately empty and is
recorded as such; changing it requires a new protocol revision.

```sh
env -u RUSTFLAGS \
  RUSTUP_TOOLCHAIN=1.95.0 CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_PROFILE_RELEASE_LTO=false CARGO_TARGET_DIR="$BASE_TARGET" \
  cargo build --release --locked --offline \
    --manifest-path "$BASE_TREE/tools/perf-baseline/Cargo.toml" \
    --bin litchi-perf-baseline

env -u RUSTFLAGS \
  RUSTUP_TOOLCHAIN=1.95.0 CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_PROFILE_RELEASE_LTO=false CARGO_TARGET_DIR="$CAND_TARGET" \
  cargo build --release --locked --offline \
    --manifest-path "$CAND_TREE/tools/perf-baseline/Cargo.toml" \
    --bin litchi-perf-baseline

sha256sum "$BASE_TARGET/release/litchi-perf-baseline" \
          "$CAND_TARGET/release/litchi-perf-baseline"
stat -c '%n %s' "$BASE_TARGET/release/litchi-perf-baseline" \
                 "$CAND_TARGET/release/litchi-perf-baseline"
```

Before capture, record `git status --porcelain=v1`, `git rev-parse HEAD`,
`rustc -Vv`, `cargo -V`, the complete build argv, the five input hashes in
§1, the selected environment variables, and the binary path/size/SHA-256 in
the manifest. Do not use `cargo run`: the captured executable must be the
hashed file above.

The eight reports are captured serially under one advisory CPU lock. This is
the complete order for each policy: baseline, candidate, candidate, baseline.
The output directory must be new, and each selector invocation must receive a
new output basename because the harness uses exclusive creation.

```sh
set -euo pipefail
mkdir -p "$EVIDENCE/forward" "$EVIDENCE/exact" "$(dirname "$CPU_LOCK")"

(
  flock -n 9

  taskset -c "$CPU" "$BASE_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy forward-start --window-bytes 4096 \
    --max-range 65536 --delay-us 0 --samples 30 --warmup 3 \
    --source-revision 410bd99f00fdfb4161a94215710340f16ea536bb \
    --output "$EVIDENCE/forward/a1.json"
  taskset -c "$CPU" "$CAND_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy forward-start --window-bytes 4096 \
    --max-range 65536 --delay-us 0 --samples 30 --warmup 3 \
    --source-revision 49c25ef67f62a8bb943e7c6a0bf464e49ae87d07 \
    --output "$EVIDENCE/forward/b1.json"
  taskset -c "$CPU" "$CAND_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy forward-start --window-bytes 4096 \
    --max-range 65536 --delay-us 0 --samples 30 --warmup 3 \
    --source-revision 49c25ef67f62a8bb943e7c6a0bf464e49ae87d07 \
    --output "$EVIDENCE/forward/b2.json"
  taskset -c "$CPU" "$BASE_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy forward-start --window-bytes 4096 \
    --max-range 65536 --delay-us 0 --samples 30 --warmup 3 \
    --source-revision 410bd99f00fdfb4161a94215710340f16ea536bb \
    --output "$EVIDENCE/forward/a2.json"

  taskset -c "$CPU" "$BASE_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy exact --max-range 65536 --delay-us 0 \
    --samples 30 --warmup 3 \
    --source-revision 410bd99f00fdfb4161a94215710340f16ea536bb \
    --output "$EVIDENCE/exact/a1.json"
  taskset -c "$CPU" "$CAND_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy exact --max-range 65536 --delay-us 0 \
    --samples 30 --warmup 3 \
    --source-revision 49c25ef67f62a8bb943e7c6a0bf464e49ae87d07 \
    --output "$EVIDENCE/exact/b1.json"
  taskset -c "$CPU" "$CAND_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy exact --max-range 65536 --delay-us 0 \
    --samples 30 --warmup 3 \
    --source-revision 49c25ef67f62a8bb943e7c6a0bf464e49ae87d07 \
    --output "$EVIDENCE/exact/b2.json"
  taskset -c "$CPU" "$BASE_TARGET/release/litchi-perf-baseline" \
    docx-managed-read-ahead --policy exact --max-range 65536 --delay-us 0 \
    --samples 30 --warmup 3 \
    --source-revision 410bd99f00fdfb4161a94215710340f16ea536bb \
    --output "$EVIDENCE/exact/a2.json"
) 9>"$CPU_LOCK"
```

The selector itself has no process-level output; the JSON reports are the
timed evidence. Capture command lines, environment, host load, CPU id and
lock path in the manifest. A fresh run uses a fresh `EVIDENCE` directory and
does not overwrite an earlier report.

## 4. External evidence manifest

The capture driver records one small standard-library JSON manifest beside the
reports. Its schema is `docx-managed-read-ahead-abba-v1`:

```json
{
  "schema": "docx-managed-read-ahead-abba-v1",
  "status": "captured",
  "sources": {
    "baseline": {
      "revision": "410bd99f00fdfb4161a94215710340f16ea536bb",
      "tree": "...",
      "binary": {"path": "...", "bytes": 0, "sha256": "..."},
      "inputs": {
        "tools/perf-baseline/Cargo.toml": "...",
        "tools/perf-baseline/Cargo.lock": "...",
        "tools/perf-baseline/src/docx_managed_read_ahead.rs": "...",
        "tools/perf-baseline/src/main.rs": "...",
        "rust-toolchain.toml": "..."
      },
      "toolchain": {"rustc_vv": "...", "cargo_v": "..."},
      "build": {"argv": [], "environment": {}, "target_dir": "..."}
    },
    "candidate": {
      "revision": "49c25ef67f62a8bb943e7c6a0bf464e49ae87d07",
      "tree": "...",
      "binary": {"path": "...", "bytes": 0, "sha256": "..."},
      "inputs": {}, "toolchain": {}, "build": {}
    }
  },
  "corpus": {
    "version": "source-edit-media-v1",
    "archive_sha256": "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4",
    "text_sha256": "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af"
  },
  "protocol": {
    "selector": "docx-managed-read-ahead",
    "samples": 30, "warmup": 3, "max_range_bytes": 65536,
    "delay_us": 0, "transfer_bytes_per_second": null,
    "abba": ["a1", "b1", "b2", "a2"],
    "cpu": 0, "cpu_lock": "...",
    "arms": {
      "forward": {"policy": "forward_start", "window_bytes": 4096},
      "exact": {"policy": "exact", "window_bytes": null}
    }
  },
  "legs": {
    "forward/a1": {"endpoint": "baseline", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "forward/b1": {"endpoint": "candidate", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "forward/b2": {"endpoint": "candidate", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "forward/a2": {"endpoint": "baseline", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "exact/a1": {"endpoint": "baseline", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "exact/b1": {"endpoint": "candidate", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "exact/b2": {"endpoint": "candidate", "report": {"path": "...", "bytes": 0, "sha256": "..."}},
    "exact/a2": {"endpoint": "baseline", "report": {"path": "...", "bytes": 0, "sha256": "..."}}
  }
}
```

The real manifest fills each report binding with the report's `schema`,
`requested_source_revision`, `policy`, `samples` and `warmup` in addition to
the file digest. A leg is invalid if any of those disagree with its arm,
endpoint or this protocol. It also records the complete command argv, process
environment, host/load snapshot and the report's per-row summary values; no
timing value is copied into this record.

## 5. Acceptance and analysis

For every report, the verifier must check the selector's complete report
oracle, not just elapsed time. In particular:

1. `schema == "docx_managed_read_ahead_v1"`, `version == 1`, `samples == 30`,
   `warmup == 3`, `source_bytes`, corpus and text digests match §2, and
   `requested_source_revision` is exactly the endpoint revision.
2. `provider.max_range_bytes == 65536`, `provider.delay_us == 0`, and
   `provider.transfer_bytes_per_second == null` for all eight reports.
3. Forward-start reports have `policy.name == "forward_start"`, configured
   window 4096, enabled diagnostics and the expected physical trace. Exact
   reports have `policy.name == "exact"`, no configured window, no source-read
   diagnostics and the expected exact physical trace.
4. Every sample passes text, source-version, cache-load, physical-returned
   bytes, and post-drop budget/object checks already enforced by the Rust
   driver. A failed row rejects its entire leg; no row may be dropped.
5. The two baseline legs use the same baseline binary digest, and the two
   candidate legs use the same candidate digest. Build input hashes, toolchain,
   profile/flags, CPU lock and environment must match across endpoints except
   for the explicitly recorded source revision, tree, target directory and
   binary path/digest.

The summary reports p50, p95 and p99 elapsed nanoseconds for each leg, then
compares `median(b1,b2)` with `median(a1,a2)` separately for forward-start and
exact. The forward-start ratio is the repair qualification; the exact ratio is
the regression control. It is descriptive until the coordinator records an
acceptance budget. Do not pass these custom reports to the generic
`tools/perf_abba_summary.py`: that tool requires the standard
`litchi.perf-baseline` report schema and would not validate this selector's
managed-source or physical-range oracle.

No build or capture was run while preparing this record. Remaining blockers
are the coordinator's build/timing slot, two clean endpoint checkouts, the
actual binary SHA-256 receipts, the external eight reports, and the manifest
verification/summary. The candidate and baseline source pair, matched-input
hashes, fixed command arguments, zero-delay policy and quiet ABBA order are
ready for that run.
