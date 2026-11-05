# 0362ugjaa — PPT lazy native qualification protocol return

This return reviews the source-only qualification packet in
[`0362ugja-ppt-lazy-native-qualification.md`](0362ugja-ppt-lazy-native-qualification.md)
after the route-fixture and execution-binding corrections. It is a protocol
return only: no release build, native timing, allocator capture, or performance
acceptance claim is made here.

## Review result

The packet now binds the 28 routes to the reviewed four-fixture matrix. The
default fixture remains the generated one-slide package; the objects,
external-media, and picture-bullets families use the tracked representatives
with the required `ExObjList`/`Pictures`, movie, and `BlipCollection9`/`Pictures`
records. Each fixture has a pinned path, byte count, SHA-256, and generator
identity. Every route proof repeats its fixture key and canonical output digest,
and correctness, calibration, and timed reports reject a route or fixture
identity mismatch. Serialization and full fidelity checks remain outside the
timed operation interval.

The protocol tests also exercise real temporary Git/source roots, lock files,
probe source, build binaries, target directories, and a manifest. They verify
that mutating the manifest, lock, probe source, and each normal/allocator
binary for both baseline and candidate after the initial check is rejected by
the corresponding at-use guard. Source-root cleanliness and target path/size
guards are checked on every recheck, but fixture-copy, external build-receipt,
toolchain, and cargo-configuration mutation are outside this source-only
packet. The runner remains fail-closed for strict JSON, finite numeric limits,
source and resource overlap, exact source/lock/probe/binary identities, and
the fixed ABBA route set.

## Source-only checks

The reviewed packet passed:

```text
PYTHONDONTWRITEBYTECODE=1 PYTHONPATH=tools python3 -m unittest tools/test_ppt_lazy_native_qualification.py
python3 -m py_compile tools/ppt_lazy_native_qualification.py
python3 tools/ppt_lazy_native_qualification.py --manifest tools/native_qualification/manifests/ppt-lazy-edit.json --out-dir <external-plan-dir> --plan
rustfmt --edition 2024 --check tools/perf-baseline/src/ppt_lazy_native.rs
git diff --check
```

The manifest remains parameterized until the final candidate and retained
build receipt are frozen. Native execution, timing, and performance disposition
remain pending the authorized qualification slot.
