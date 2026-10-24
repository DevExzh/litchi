# 0360adca — XLSX raw paired-native qualification protocol

Status: preparation packet from staging commit
`76c21c49cfc7fedb995fb1b2c11351098eddef7b`. This closes the runner gap
identified by the `0360adc` historical and candidate diagnostics. It adds no
production-format change and records no timing or acceptance result.

## Existing evidence and gap

The historical raw identity lane remains the original production source
`0befb112d96d2409e6f40c9abfeb4813665905ad`, with the reviewed consumer-only
probe/manifest adapter copied from the final source into a disposable overlay.
The candidate diagnostic remains parameterized until the final assembly is
frozen; the previous `ff86f4a1c07f05a98a8dddf272daef56942656ea` result is
diagnostic only. The existing standalone probe already covers the four
corpora, owned open, first selected read, raw slicer/timeline insertion,
typed replacement, commit/replay/inverse, stale-patch refusal, no-op,
untouched worksheet bytes, native relationship identity, and deferred decode
counters. Its allocator and native binaries remain the two routes.

The catalog found no manifest-bound paired-native ABBA lane for that probe.
The generic `tools/native_qualification_runner.py` already supplies the
approved source/lock/probe/binary hashing, fresh external targets, 20 GB
target and address-space ceilings, process quietness guard, exact
`before, after, after, before` order, event journal, and build-to-ABBA
identity handoff. This packet registers the XLSX raw lane with those guards
instead of creating a second launcher.

## Protocol

`tools/native_qualification/manifests/w4-xlsx-raw.json` requires two programs
per leg:

- `xlsx-raw-identity-alloc`, built with `allocator-metrics`, reports measured
  allocation samples;
- `xlsx-raw-identity-native`, built without the allocator wrapper, reports the
  same lifecycle with allocation status unavailable.

The historical leg uses `--no-default-features --features
allocator-metrics,legacy-xlsx-raw-opc` and invokes the probe as `original`.
The candidate leg uses `allocator-metrics,current-xlsx-raw-opc` and invokes it
as `candidate`. Both run five measured iterations after two warmups; the
runner's four ABBA legs exercise both routes. `XLSX_BEFORE_REVISION` must be
the fixed original `0bef...` revision, while `XLSX_AFTER_REVISION` remains an
environment parameter until the final candidate freeze.

The validator checks every eight report files, both binary paths and SHA-256
identities, all four corpus rows, all four operation rows, all fidelity
booleans, native relationship identities, and the allocator/native status
split. It emits only a reduced correctness summary with
`performance_claim: false`; it does not interpret elapsed values.

Example final execution, after the coordinator supplies clean external source
trees and grants the native/build slot:

```sh
export NATIVE_QUAL_OUTPUT=/external/xlsx-raw-native/abba-20261024-01
export LITCHI_COORD_BIN=/home/zhuhe/code/litchi-worktrees/.coord/bin
export LITCHI_NATIVE_QUAL_ROOT=/path/to/qualification-checkout
export XLSX_BEFORE_SOURCE=/external/sources/xlsx-original-0bef-overlay
export XLSX_AFTER_SOURCE=/external/sources/xlsx-final-candidate
export XLSX_TARGET_ROOT=/external/xlsx-raw-native/targets/abba-20261024-01
export XLSX_BEFORE_REVISION=0befb112d96d2409e6f40c9abfeb4813665905ad
export XLSX_AFTER_REVISION=<final-assembled-candidate-sha>

python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4-xlsx-raw.json \
  --phase build
python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4-xlsx-raw.json \
  --phase abba --cpu <isolated-cpu>
```

The source overlay must be made from the actual original 0bef production tree
and contain only the reviewed consumer-side probe adapter. It must not use
the 5726 or ff86 source as a substitute baseline. Binaries, target trees,
reports, event journals, and Callgrind captures remain outside the checkout;
the runner refuses stale outputs, source mutation, target growth, unrelated
Litchi builders, and binary/source/lock identity drift.

## Static validation

No Cargo, release build, native binary, or timing workload was run in this
preparation packet. Python-only checks passed:

```text
python3 -m py_compile \
  tools/native_qualification_runner.py \
  tools/native_qualification/xlsx_raw_identity_validator.py \
  tools/test_xlsx_raw_native_qualification.py \
  tools/test_native_qualification_runner.py
PASS

python3 tools/test_xlsx_raw_native_qualification.py
PASS — 3 tests

python3 tools/test_native_qualification_runner.py
PASS — 22 tests
```

The full native qualification remains pending final source freeze and an
isolated coordinator build slot.
