# 0362ul135a — reconcile current-native v2 f2 baseline locks

Date: 2026-10-25. This is a source/Python-only correction to 0362ul135.
The original 135 commit remains unchanged. No Cargo, rustdoc, native probe,
timing, allocator, or performance workload was run.

Independent review identified that the first 135 policy used the later
current-source perf/native lock identities for the historical f2 endpoint.
That would make `_tracked_materialization` and `_pin_file` reject the actual
f2 tree, so it could never validate its own fixed baseline. The policy now
uses the exact tracked blobs from the frozen f2 commit
`f2bbbc8d07ab5da26c180f1ff74813d39f632d54`:

| role | relative path | SHA-256 | bytes |
| --- | --- | --- | ---: |
| workspace | `Cargo.lock` (approved ignored artifact) | `1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c` | 139,441 |
| perf | `tools/perf-baseline/Cargo.lock` | `57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f` | 57,028 |
| native | `tools/native-resave/Cargo.lock` | `6aaa0e3800a4db72afbe8783aab382accdb4a6355d06bb97fb1c12dd78839408` | 56,929 |

The remaining baseline bindings were checked directly from f2 and remain
unchanged: Git tree `2e0d0cdf7f6a345961ceca361d53202200d00970`, canonical
tree-list SHA-256
`ce92652c439d810849b74b80e650821bc21de9029ea861cd6bbc13e38d809589`, the
source manifests `Cargo.toml` (`66a8f4d2b0ffb6b9ba4eb87d84fee8a3049e8ef7e4380b6ee48acaaf372dd7b5` /
10,608 bytes), `tools/perf-baseline/Cargo.toml`
(`9e9d64678dafa5e96034f9f6b5cdb1fca28a65cc7a30adb83c42d6d4ad8b34cd` /
3,478 bytes), and `tools/native-resave/Cargo.toml`
(`14a3b0fbf2da037af795d4dc820ed5d91ccbc2d4f5f009f7bb2182c25849fa8b` /
1,425 bytes), toolchain/configuration (`rust-toolchain.toml`
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` /
86 bytes; `.cargo/config.toml`
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` /
470 bytes), PPT probe
`0350fc1696688e576613bc8b2c2ccb3d39c12b4b8ad890c5dcd14e3e7b4138ca` /
5,158 bytes, and fixture
`f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a` /
43,520 bytes. The separately reviewed external f2 archive remains bound to
its existing identity and size.

The two current-v2 manifests were regenerated only for their derived trusted
baseline-lock projection. The candidate source/tree/archive/probe/fixture
binding remains pending final source freeze; no candidate digest was inferred
from this correction.

Focused verification:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2
```

Result: 21 tests passed, including copied-lock and protocol-checkout
boundary refusals. No native or performance claim is made.
