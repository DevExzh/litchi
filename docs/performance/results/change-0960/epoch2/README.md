# 0960 epoch2 bounded screening packet

This is a separate, fresh 0960 screening epoch. It leaves the historical
parent packet (`docs/performance/results/change-0960/`) untouched; the first
incomplete epoch is excluded completely. Epoch2 has completed 36 native and 18
allocator children. The copied completed-packet descriptors bind every child
terminal and validation receipt, while raw reports remain external.

The bounded reduction is `analysis/screen-v1.json` (104,664 bytes); the copied
screening selector output is `analysis/selection-v1.json` (10,752 bytes) and
reports `reject_screen` with no adoption allowance. The root report owns the final interpretation. The independent raw audit
passed all 54 children and reported zero native/allocator metric mismatches;
its bounded summary, full audit output, and script are retained under
`review/`. The selector recomputation remains `reject_screen` with adoption false. The
independent audit files are `review/independent-raw-audit-v1.json` (119,253
bytes, SHA-256 `7329adea6e19745f925a18cdb07c1d4fb2079def4e87407afe420f00f1914512`),
`review/independent-raw-audit-summary-v1.json` (27,044 bytes, SHA-256
`57783d25d8b7a4d429c6d44255803cc603c4100cb8a45aa066996ea061e67d1a`), and
`review/audit_epoch2_raw_v1.py` (32,031 bytes, SHA-256
`5870717d3c0186ea3764f4df44f1ef12673ff59e016ec036e7eb05f8445ca661`). This
packet makes no production-adoption claim.

The external raw artifact manifest is
`review/external-artifact-manifest-v1.json` (SHA-256
`2fc2ca4141014e8149cddcb8299e255a248e629445d6bc09289a481b4def57bf`). It
inventories all regular files and symlink entries under the external epoch
root, including the native and allocator raw packets. Raw reports, logs,
source trees, and binaries remain outside this repository. The external
manifest records the native packet as 327 files / 2,054,441 bytes and the
allocator packet as 165 files / 1,491,661 bytes; the completed-packet files in
this subtree provide the lane-level terminal and validation descriptors.

The copied protocol files bind the final plan, prospective plan, host receipt,
source revalidation, preflight, root GO, and the independent rerun-feasibility
review. `history/live-state-v1.json` is retained only as a historical launch
state receipt; it must not be read as a terminal receipt. The two small Python
helpers preserve the epoch preparation and authorization commands. No raw
capture directory or compiled executable is copied here. `review/write-limits-cleanup-v1.json`
records the read-only Git staging failure and that no cleanup was attempted.

The v2 raw-audit helper adds a create-only `--output` path without changing the
arithmetic. `review/independent-raw-audit-v2-replay-receipt.json` records that
the v2 replay was byte-identical to the v1 audit, launched no workload, and did
not mutate the packet. The v2 helper is 32,535 bytes (SHA-256
`0e12714f378cb031af026c0618bbc168f43abc503f9d4403adccfbed060ebf12`) and its
replay receipt is 1,426 bytes (SHA-256
`b804195e5b75b15d122884ef5a8543f8a0fe9254e228659d0e5725cdda2f00a2`). The v1
helper and result remain the historical primary audit artifacts.

To regenerate the external manifest into a new path, using a destination that
does not already exist:

```sh
python3 -B docs/performance/results/change-0960/epoch2/tools/build_external_manifest_epoch2_v1.py \
  --root /tmp/litchi-performance-0960-epoch2 \
  --output /tmp/litchi-performance-0960-epoch2-manifest-replay.json
```

The following are the exact historical capture invocations, including the
outer CPU affinity and approval variable. They are provenance only and must
not be rerun against the retained packet roots:

```sh
LITCHI_PERF_0960_CAPTURE_APPROVED=1 /usr/bin/taskset --cpu-list 12-15 python3 -B /home/zhuhe/litchi-performance-0960/tools/screen_0960_v7.py \
  --capture --lane native --go \
  --plan /tmp/litchi-performance-0960-epoch2/plan-final-v1.json \
  --packet /tmp/litchi-performance-0960-epoch2/capture-native-v1
LITCHI_PERF_0960_CAPTURE_APPROVED=1 /usr/bin/taskset --cpu-list 12-15 python3 -B /home/zhuhe/litchi-performance-0960/tools/screen_0960_v7.py \
  --capture --lane allocator --go \
  --plan /tmp/litchi-performance-0960-epoch2/plan-final-v1.json \
  --packet /tmp/litchi-performance-0960-epoch2/capture-allocator-v1
```

Any future capture needs a newly reviewed plan, fresh source/build/binary
bindings, a new root GO receipt, and packet/output paths that do not exist.
The frozen plan points at retained evidence and is not a future recapture plan.

To reproduce the independent raw audit without changing the retained packets,
use the v2 helper with a fresh output path:

```sh
python3 -B docs/performance/results/change-0960/epoch2/review/audit_epoch2_raw_v2.py \
  /tmp/litchi-performance-0960-epoch2 \
  --output /tmp/litchi-performance-0960-epoch2-independent-raw-audit-v2-replay.json
# The retained root must contain plan-final-v1.json, both capture packet
# directories, analysis-v1/screen-v1.json, and analysis-v1/selection-v1.json.
# The output path must not already exist; no child workload is launched.
```

To regenerate the bounded reduction and selector from the retained external
packets, use fresh output paths. The reducer and rule/selector helpers are
bound by the historical parent packet and by the descriptors in the copied
plan and completed-packet files:

```sh
python3 -B /home/zhuhe/litchi-performance-0960/tools/reduce_0960_v3.py \
  --plan /tmp/litchi-performance-0960-epoch2/plan-final-v1.json \
  --native-packet /tmp/litchi-performance-0960-epoch2/capture-native-v1 \
  --allocator-packet /tmp/litchi-performance-0960-epoch2/capture-allocator-v1 \
  --output /tmp/litchi-performance-0960-epoch2-reduction-replay.json
python3 -B /home/zhuhe/litchi-performance-0960/tools/select_screen_0960_v1.py \
  --report /tmp/litchi-performance-0960-epoch2-reduction-replay.json \
  --rule /home/zhuhe/litchi-performance-0960/review/root-screening-rule-v2.json \
  --output /tmp/litchi-performance-0960-epoch2-selection-replay.json
```

All copied files are bounded below 1 MiB and no compiled executable is present.


The later writable continuation replayed and reviewed this packet successfully.
See `../review/closure-audit-v1.json`, `closure-review-v1.json`, and
`closure-cleanup-v1.json` for the completed verification and removal of the two
obsolete external build caches. Historical read-only receipts remain unchanged.
The retained source, executables and raw measurement packets were not removed.
