# 0362ul10h — OLE owner per-endpoint lock repair and compile diagnostics

- Date: 2026-10-24.
- Parent packet: 0362ul10g.
- Status: locked compile diagnostics passed; native execution, timing and
  paired qualification remain pending coordinator review.

The first compile packet exposed that one shared `bc5a133f...` lock could not
represent both source graphs.  The preparer now binds a reviewed base lock
(`bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`) and
explicit role locks:

| driver | SHA-256 | exact permitted local dependency additions |
| --- | --- | --- |
| baseline | `450f43c6b74536ef552bd5e9a19e321a3c0dea745f645fecc0aa9abe37868e14` | `litchi-perf-baseline` → `sha2 0.11.0` |
| candidate | `5a07adc1d62dea06dbac60986faa020007882506523dde49e4605708e1276aaf` | `litchi-perf-baseline` → `sha2 0.11.0`; `litchi-odraw` → `litchi-core` |

Both role locks are external files under
`/home/zhuhe/code/litchi-worktrees/.coord/ole-owner-driver-locks-e786/` and
are byte-for-byte the base lock plus those listed dependency edges.  The
preparer and the independent validator parse both locks, require identical
package identity sets, and compare every registry package's name, version,
source and checksum.  Focused mutations cover a registry checksum and an
unreviewed local edge.  No registry package, production source tree or
baseline overlay was changed.

Preparation was made from candidate commit
`69944be4a2e27063b46583b48150c8c83fa3a377`, with the clean baseline at
`0ce4bb8b61f03a1e5497633f51b954e76f98aa8c`.  The external receipt is
`/home/zhuhe/.cache/litchi-api-contract/ole-owner-preparation-69944.json`
(58,993 bytes,
`8a174e4c4f00ea30ec5eb5b5b7fb22e00ef1fbe055f8b935fb01d955ba0f90d4`), and
`validate_preparation_receipt` passed against the selected candidate.

With one Cargo job, offline mode, `--release --locked`, fresh external target
directories and the repository memory-capped Cargo dispatch, all four
compile-only checks passed:

```text
baseline/normal:    --features legacy-owner-api,legacy-property-api
                    --example ole_w5_probe --example ole_package_clone_probe
                    --example ole_owner_oracle
baseline/allocator: --features legacy-owner-api,legacy-property-api,w5-counting-allocator
                    --example ole_owner_oracle
candidate/normal:   --features current-api
                    --example ole_w5_probe --example ole_package_clone_probe
                    --example ole_owner_oracle
candidate/allocator:--features current-api,w5-counting-allocator
                    --example ole_owner_oracle
```

Each command used the corresponding prepared driver `Cargo.toml`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, `CARGO_NET_OFFLINE=true`,
`LITCHI_MEM_MAX=20G`, and a distinct target under
`/home/zhuhe/.cache/litchi-api-contract/ole-owner-compile-69944/`.  The exact
external logs are:

```text
locked-baseline-normal.log     5,243 bytes  c38a0393c62145f8d1bad5b623b995fba3d3eae80fe155572ccd50c9549d41e3
locked-baseline-allocator.log  5,993 bytes  0bd8a2517bb35a99ee255db81f8b7d731170e120b43035891407df8e1ba3dd94
locked-candidate-normal.log    5,088 bytes  675aeee66ad60f81c80bf08b1065ff0722d834a452fbb9f7dd9b1208f60c0223
locked-candidate-allocator.log 5,838 bytes  f3d9e08eb7f10bd883c90af73e630313cd59cdf5dadb48a01caf8bbb95d28512
```

Target directories were 140,028,164, 140,031,818, 140,233,291 and
140,236,899 bytes respectively.  These are compile diagnostics only; no
native binary was executed, no timing/calibration was performed, and no
performance or qualification acceptance is claimed.

Source-only checks and the focused protocol suite passed:

```text
python3 -m unittest tools.test_prepare_ole_owner_inputs \
  tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run -q
python3 -m py_compile tools/native_qualification/prepare_ole_owner_inputs.py \
  tools/native_qualification/ole_owner_paired_runner.py \
  tools/native_qualification/validate_ole_owner_run.py \
  tools/test_prepare_ole_owner_inputs.py tools/test_validate_ole_owner_run.py
```
