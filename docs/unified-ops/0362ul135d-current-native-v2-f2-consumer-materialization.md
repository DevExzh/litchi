# 0362ul135d — f2 consumer-only adapter materialization

Date: 2026-10-25. This is a source/Python-only repair to 0362ul135c. No
Cargo, rustdoc, native probe, timing, allocator, or performance workload was
run.

## Frozen baseline and the compatibility gap

The original baseline remains the immutable f2 source
`f2bbbc8d07ab5da26c180f1ff74813d39f632d54`, Git tree
`2e0d0cdf7f6a345961ceca361d53202200d00970`, and canonical tree-list SHA-256
`ce92652c439d810849b74b80e650821bc21de9029ea861cd6bbc13e38d809589`.
Its tracked `tools/perf-baseline/Cargo.toml` is SHA-256
`9e9d64678dafa5e96034f9f6b5cdb1fca28a65cc7a30adb83c42d6d4ad8b34cd` (3,478
bytes) and has no `rtf-legacy-api` feature. The f2 tree also has neither
`tools/perf-baseline/examples/refusal_contract_bench.rs` nor
`crates/litchi-ppt/examples/scoped_patch_correctness.rs`; it has only the
historical `scoped_patch_probe.rs` (SHA-256
`0350fc1696688e576613bc8b2c2ccb3d39c12b4b8ad890c5dcd14e3e7b4138ca`).

The original source archive remains the previously reviewed external artifact
`source-archive.tar.zst`, SHA-256
`de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280`,
6,614,528,000 bytes. The adapter receipt never replaces or rehashes that
archive as an overlay.

The baseline is therefore not made buildable by relabeling a later tree. A
coordinator creates a separate detached f2 worktree, copies the already
approved ignored root lock
`1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c` (139,441
bytes), and runs the materializer below. The pristine f2 worktree and its
source archive remain unchanged.

## Reviewed adapter overlay

The policy now binds exactly three endpoint files. The adapter source is
tracked in this protocol checkout under `tools/native_qualification/adapters/`
and is checked by SHA-256 before copying:

| endpoint path | reviewed adapter | bytes | SHA-256 |
| --- | --- | ---: | --- |
| `tools/perf-baseline/Cargo.toml` | `rtf-baseline-Cargo.toml` | 3,638 | `af876fbf715edb643d2c7abb8643efb252413c640824890e90b1849207889676` |
| `tools/perf-baseline/examples/refusal_contract_bench.rs` | `rtf-refusal_contract_bench.rs` | 12,782 | `0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81` |
| `crates/litchi-ppt/examples/scoped_patch_correctness.rs` | `ppt-scoped_patch_correctness-f2.rs` | 3,168 | `032e74e081d2b442df49ce5bd7fb2edc66b89a4a1d5368872dd7bfeb94c8ee7a` |

The RTF manifest overlay adds only the reviewed consumer features and retains
the f2 dependency graph. The RTF adapter selects `rtf-legacy-api`; the
candidate selects `rtf-current-api` in its own source endpoint. The PPT
adapter is the f2-compatible `slide_order::Snapshot` consumer from review
87051f9548. It emits the exact 0362g correctness fields: fixture, media ID 1,
`clip-0000.avi`, source/target extents and hashes, and inverse-source hash and
boolean. Forward-replay fields are accepted when a newer consumer emits them,
but are not required by the reviewed f2 consumer. A partial replay pair is
rejected.

The original manifest blob is separately checked from `git show HEAD:path`
before its working-tree replacement. The source archive is still verified
against the original f2 Git tree, not the adapter files. Any dirty, ignored,
sparse, symlinked, missing, or extra endpoint path outside these three fixed
overlays is refused.

## Materialization and guarded build packet

The source-only materializer is:

```text
PYTHONDONTWRITEBYTECODE=1 PYTHONPATH=tools \
  python3 tools/native_qualification/materialize_current_native_adapters.py \
  --source /absolute/external/f2-adapter-worktree \
  --receipt /absolute/external/f2-adapter-receipt.json
```

It requires the exact f2 commit/tree, the approved root lock, a complete
tracked materialization, and a separate source/protocol checkout. It writes
only the three fixed adapter files and records their source and endpoint
hashes, original Git identity, status rows, and the fact that no production
source or lock was regenerated. A receipt is refused if an existing overlay
is not either the untouched f2 blob or the exact reviewed adapter.

After materialization and source-archive preparation, the coordinator may
verify the build shape with the existing quiet-host guard. The baseline
commands are the following (commands are recorded here, not run in this
repair):

```text
cargo build --offline --locked --target x86_64-unknown-linux-gnu --release \
  --manifest-path tools/perf-baseline/Cargo.toml --no-default-features \
  --features rtf-legacy-api --example refusal_contract_bench

cargo build --offline --locked --target x86_64-unknown-linux-gnu --release \
  -p litchi-ppt --example scoped_patch_probe --example scoped_patch_correctness \
  --all-features
```

The correctness companion is outside the measured probe clock. Its output is
consumed by `docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py`
with the selected fixture; the validator checks the exact fields and rejects
malformed or fabricated media/path/hash relations. No timing or acceptance
claim follows until the coordinator independently runs both source endpoints,
the five-block RTF ABBA contract, and the existing evidence validators.

## Focused source-only verification

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2
```

Result: 29 tests passed. They cover exact adapter identities, non-f2
materialization refusal, reviewed f2 correctness receipts with optional
replay fields, malformed media/path/replay negatives, and the existing paired
source/lock/archive/route/ABBA boundary. No native or performance claim is
made.
