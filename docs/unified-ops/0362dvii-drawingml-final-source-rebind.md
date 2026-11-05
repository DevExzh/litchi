# 0362dvii — DrawingML ten-profile final-source rebind

This is the final-source rebind of the 0362dvii DrawingML packet. It binds
the packet to the clean 472b production source while preserving the validated
2f7, 996, 11728, and 945 baseline worktrees from the earlier c796 packet.
The earlier packet at `/tmp/dml-final-paired-baselines.Wz0XIZ` was not
modified, and no adapter files were recopied into those baselines.

This is source-only preparation. Cargo, native probes, ABBA, allocation
capture, timing, and performance qualification were not run.

## Final source identities

```text
candidate revision: 472b6705475b073e7fe08f6e11ee17fb1e29f838
candidate tree:     ebc1cee05f124c0c63977aff07a23288c19e7d00
candidate source tree identity: 9e713180f74f0515b807811a8c30ae166f43c05b4292ea9fc3824758055bb08e
workspace Cargo.lock: 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1 (139606 bytes)
historical Cargo.lock: 52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6 (139591 bytes)
perf-baseline Cargo.lock: eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c (57227 bytes)
plan: docs/unified-ops/results/0362dvie/qualification-plan.json
plan SHA-256: e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505
manifest SHA-256: a94e262d2fd2d9055de1036f2c87a31cc198142002bc27fec92f586eda2283e8
```

The final candidate root is
`/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source`.
The selected final-source adapter identities are explicitly bound in the
external receipt, including `tools/perf-baseline/src/lib.rs` (3,088,216 bytes,
SHA-256 `280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a`),
the native probe (247,616 bytes, SHA-256
`539e4641b1eecda7ef97f1bfca2e526348f0a0d4941cdec781cadb3c99a18477`), the
validator, preparer, paired runner, manifest, and fixture.

## Baseline profiles and validation

The four existing immutable baseline roots remain:

```text
945 owner routes: /tmp/dml-final-paired-baselines.Wz0XIZ/source-945
                  945eb1d6b9ef6c56fa26c82e12758d38c1b9015d
2f7 raw routes:   /tmp/dml-final-paired-baselines.Wz0XIZ/source-2f7
                  2f7db7ee7fa01893e9099217fc4243687d196f63
996 theme bulk:   /tmp/dml-final-paired-baselines.Wz0XIZ/source-996
                  996651b6b19717a0192dfd4144bccd367fb2ea44
11728 theme fam:  /tmp/dml-final-paired-baselines.Wz0XIZ/source-11728
                  11728db894181ffbcdbcea3ae0008d51fdd2d8af
```

The final rebind receipt is external at
`/tmp/dml-final-472b-rebind-20261024T1830Z/final-source-rebind-receipt.json`
(13,372 bytes, SHA-256
`f35c4794da122a699cf9e181bc996f4af31c452b43b4384ac1fb1eaa3fc75dca`). The
expanded ten-profile manifest is
`/tmp/dml-final-472b-rebind-20261024T1830Z/expanded-manifest.json` (12,318
bytes, SHA-256
`e8f36f68e0aef21cebfee42983458b611b2fe576a8b043890beb688585eef66f`). The
actual final-source runner accepted the manifest and independently validated
all ten preparation receipts:

```text
drawingml-chart-raw                         2f7
drawingml-diagram-raw                       2f7
drawingml-text-media-raw                    2f7
drawingml-theme-ink                         996
drawingml-geometry-color-formatcode2-transform 2f7
drawingml-theme-family-raw                  11728
drawingml-chart-owner                       945
drawingml-theme-family-owner                945
drawingml-ink-actions-owner                 945
drawingml-transform-owner                   945
```

The final-source static protocol tests (13 tests) and Python tool compilation
passed without Cargo or native execution. The generated diagnostic lock
`93878f7a5925cb0484a770bdfb7132cb124dd43c7d37a31ad08fe19b0d8dd741` was not
used.

## Handoff commands

The receipt records the exact future commands, with fresh external output and
target roots, for the paired build and ABBA phases:

```text
env -u LITCHI_CAPPED LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  python3 /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/tools/native_qualification/drawingml_raw_paired_runner.py \
  --manifest /tmp/dml-final-472b-rebind-20261024T1830Z/expanded-manifest.json --phase build

env -u LITCHI_CAPPED LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  python3 /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/tools/native_qualification/drawingml_raw_paired_runner.py \
  --manifest /tmp/dml-final-472b-rebind-20261024T1830Z/expanded-manifest.json --phase abba
```

These commands are recorded as a handoff only; no qualification result is
inferred from them.
