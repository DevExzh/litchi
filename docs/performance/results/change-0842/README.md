# 0842 evidence packet

Adopted runtime-detection configuration; see
[report](../../0842-opc-runtime-compression-dispatch.md) and
[decision](decision.json). This packet tests generic OPC operations, not semantic
Office documents. Base commit is `45ca1ca9fe9ff04f1ba14ded41cfe620629ffaa1`.

## Retained populations

- Original qualification/native: 168 reports, 4,344 samples; 144 native
  processes each run once, six alternating blocks and 30 samples/three warmups.
- Separate tiny-tail confirmation: 24 reports, 24,000 samples; 1,000 samples/ten
  warmups per child, still only six independent paired blocks.
- Auxiliary codec qualification: two processes, eighteen archives, 72 members.
  Nine named outputs per leg include a duplicate borrowed/reusable control.
- One initial failed qualification, with no report, is retained. All failed
  preflight builds/readers/gates and lock-integration scripts remain available.
  No formal native or confirmation run was selectively replaced.

`admission.json` freezes 506 files before native timing. `tail-admission.json`
separately freezes the confirmation plan, code, original results and binaries'
build descriptors. Original data is neither replaced nor pooled with the later
confirmation. Both feature-stage locks are retained; binary/source/command,
artifact, log and result identities are hash-bound.

## Replaying retained evidence

From the original repository path, with Python 3:

```sh
python3 -B docs/performance/results/change-0842/analyze.py
python3 -B docs/performance/results/change-0842/crosscheck.py
python3 -B docs/performance/results/change-0842/codec_check_v2.py
python3 -B docs/performance/results/change-0842/tail_analyze.py
python3 -B docs/performance/results/change-0842/tail_crosscheck.py
python3 -B docs/performance/results/change-0842/close_v2.py verify
```

The scripts use original absolute-path custody records. For another checkout
location, use a separate copy and an explicit path-remapping verifier rather
than editing the sealed packet. Reductions are repeat-safe and compare existing
outputs exactly. Closure also requires the recorded source state; after future
production changes, use a checkout of this batch. Final post-commit verification
checks the exact parent and changed-path seal.

For a fresh measurement, create a new packet and marker-owned target/scratch
roots. Build both recorded dependency declarations with their corresponding
probe lock snapshots, use the same probe source and release environment, retain
binary hashes, pass qualification/oracles first, then freeze and execute the
plan. See build.py, driver.py, plan.json, capture.py and tail-plan.json for exact
commands/settings. Existing capture destinations are exclusive; do not rerun a
capture into this packet. Temporary binaries and build trees are deliberately
removed at closure; qualified ZIP artifacts, source inputs and metadata remain.

## Interpretation and failures

`native.md`/`native.csv` retain all 84 metric rows. `analysis.json` is validity,
not adoption. Original tiny-write/read p99 flags are disclosed; separately
confirmed p99 estimates did not reproduce those increases. The read-tail
interval remains wide. RSS covers the full child, including validation/results,
and no allocation or operation heap claim follows.

`quality-after.json` scopes successful gates. Full workspace all-target and
documentation gates failed on baseline-confirmed iWork defects; see the two
workspace limitation records. User-excluded iWork source is unchanged. The
five affected crates have 6,931 passing tests, zero failures and 38 ignored.

`preflight-repairs.json` records the earlier probe compile/member-count/reader/
feature-lock corrections. `coordination-repair.json` retains the local premature
agent-commit correction. `lock-integration-repairs.json` retains the two failed
lock-orchestration versions and final successful repair. Native-resave's stale
local dependency graph is reproduced with baseline declarations in
`native-resave-lock-drift.json`; source-final.json distinguishes these two final
nested locks from measured source-after.json. All three live tools pass locked
metadata and compile checks.

`close_v2.py` requires the separate tail evidence and arithmetic check before
running the original close.py. Cleanup verifies and removes only the two owned
roots. `cleanup.json`, `closure.json` and `seal.json` record completion. The
unrelated FORMAT_IMPLEMENTATION_REVIEW, UNIFIED_OPS_API_DESIGN and empty
matrix-analysis files are preserved. The broad performance goal remains active.
