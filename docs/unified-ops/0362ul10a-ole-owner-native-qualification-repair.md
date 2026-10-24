# 0362ul10a — OLE owner qualification handoff and fidelity repair

- Date: 2026-10-24.
- Parent packet: 0362ul10.
- Status: source-only protocol repair; native build and acceptance remain pending.

The first 0362ul10 packet bound a build and ABBA runner but did not provide an
independent post-run closure check, and its timed probe's target-only checks
were not sufficient to establish complete OLE/property/DOC/XLS fidelity.  This
follow-up adds `tools/native_qualification/validate_ole_owner_run.py` and the
separate `ole_owner_oracle` example.

Preparation now records the historical external provenance labels and actual
hashes for the immutable `ad4c52c0…` probe, `0a23c99e…` clone probe and
`bc5a133f…` lock.  It reads the selected CFB source and rejects a baseline that
does not expose `open(reader)` plus `open_with_limits`, or a candidate that
does not expose only `open(reader, OleFileLimits)`.  It also checks the probe
call sites and the external driver file set, so a source revision or a marker
string cannot substitute for the reviewed adapter.  The baseline production
tree remains overlay-free.

`ole_owner_oracle` runs before calibration and never inside a timed operation.
It requires the complete no-op OLE bytes to equal the source, checks all 64
untouched sibling stream payloads and logical directory metadata after the
target replacement, and checks the replacement target bytes explicitly.  It
reads the exact link-base, hyperlink tuple and opaque property; checks the DOC
insertion kind/range/author and the XLS source/output cell values against
literal expected values; and emits all source/output artifacts.  The runner
and validator hash every artifact independently and require identical
non-allocator output digests across baseline/current and normal/instrumented
endpoints.

The instrumented oracle consumer records allocation/deallocation/reallocation
counts and bytes, including realloc old/new sizes, base/live/peak state and
actual retained history cardinality.  The normal oracle is explicitly marked
uninstrumented.  Both runner and independent validator enforce live-state
conservation, peak ≥ live, nonnegative counters, and 129 retained packages /
128 patches only for the two history routes.  Allocator elapsed values remain
non-claimable.

The independent validator re-reads the build and ABBA receipts, verifies the
actual runner/toolchain/features/artifact identities, all journal and log
hashes, claim scopes, immutable build-receipt hash, source/driver identities,
oracle digests and measurements.  It does not invoke Cargo or native code.
Its build-free round trip creates a real preparation receipt, synthetic
receipts/logs/artifacts and oracle files, then validates the packet; focused
mutations cover the build-receipt hash and allocator conservation.

Source-only evidence run on this packet:

```text
python3 -m unittest tools.test_prepare_ole_owner_inputs tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run
python3 -m py_compile tools/native_qualification/*.py tools/test_ole_owner_paired_runner.py tools/test_prepare_ole_owner_inputs.py tools/test_validate_ole_owner_run.py
rustfmt --edition 2024 --check tools/perf-baseline/examples/ole_owner_oracle.rs
```

No Cargo build, native binary, timing result, or performance acceptance claim
is made by this record.  The later execution must retain the preparation
receipt, four endpoint build artifacts, oracle outputs, raw journals/logs,
measurements, independent validator result, and reduced coordinator summary
outside the repository.
