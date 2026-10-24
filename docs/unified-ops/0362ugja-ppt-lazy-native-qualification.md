# 0362ugja — PPT lazy-entry native qualification protocol

This record prepares the missing native qualification lane for the PPT lazy
edit entry correction in 0362ugj. It is source and runner infrastructure only:
no Cargo command, release build, native timing, allocator capture, or
performance decision was made by this commit.

## Why a new lane is required

`0362ugj-ppt-lazy-edit-entry.md` is a consumer-validation receipt. It
explicitly disclaims native timing, ABBA, allocation, and performance
qualification. The earlier `0362f-ppt-replay-performance-review.md` measures
replay and does not exercise lazy edit construction or drop. Neither receipt
therefore covers the affected leaf entry boundary.

The first PPT lazy-entry implementation is commit
`120e7824a6c6b8a8ac0fe3928e45e4978630b116` (`feat(ppt): defer lazy edit
materialization`). Its direct parent
`8779d059add6c90c74552801372c44f770db9f0f` is the required baseline. The
baseline tree is `27e0b56e2f8e3ab11b00c119604b5636f358dfe2`, and its
57,158-byte `tools/perf-baseline/Cargo.lock` SHA-256 is
`4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184`.
The baseline is not a later consumer-validation tip or a replay comparison.

## Probe contract

`tools/perf-baseline/src/ppt_lazy_native.rs` is shared by the normal and
allocator binaries. It constructs the smallest valid PPT package before each
operation boundary, then measures only the operation. Its four routes are:

| route | operation inside the boundary | exact proof |
| --- | --- | --- |
| `primary_entry_drop` | `Package::edit()` followed by drop | no publication and no source change |
| `primary_noop_commit` | `Package::edit().commit()` | `Diagnostics::changed() == false`, patch is a no-op, published bytes equal source |
| `metadata_entry_drop` | `Package::edit_metadata()` followed by drop | no publication and no source change |
| `metadata_noop_commit` | `Package::edit_metadata().commit()` | shared diagnostics unchanged, patch is a no-op, published bytes equal source |

`--mode verify` runs all four routes once and emits the synthetic package
generator identity plus byte count/SHA-256. `--mode calibrate` doubles a
route-local operation batch to a fixed target. `--mode timed` runs warmups
outside the clock and returns one elapsed batch sample per request. The
allocator binary uses the existing `allocation_metrics` system-allocator
observer and reports allocation/deallocation/reallocation calls, bytes, live
bytes, and monotonic high-water bytes. The normal binary reports allocator
metrics as unavailable; an uninstrumented run is never interpreted as zero.
The probe rejects counter overflow, live-byte conservation failures, peak
regressions, and reallocation counts that exceed allocation counts; the runner
rechecks those same bounds for every timed sample.

The probe performs no source reopen or output serialization in the timed
loop. The verify phase is the semantic oracle; timed samples repeat only the
already-proved exact drop/no-op route. This keeps edit-entry cost separate from
package construction, fixture preparation, and report formatting.

## Runner contract

`tools/ppt_lazy_native_qualification.py` reuses the existing native
qualification process, target, load, and memory guards from
`odg_joined_qualification_builder.py`. It does not build a source revision.
For each endpoint it requires:

- a clean, non-sparse worktree whose `HEAD`, full tree SHA, and separate
  `tools/perf-baseline/Cargo.lock` identity match the manifest;
- independently bound normal and allocator release executables, with exact
  path, size, and SHA-256 identities below distinct external target
  directories;
- external probe source files with exact size/SHA-256 identities; and
- source, binary, lock, target, manifest, and probe identities rechecked
  before correctness, after correctness, after calibration, and after ABBA.

The checked-in manifest is intentionally parameterized. It pins the true
baseline ref/tree and lock path but leaves external roots, binary identities,
and the final candidate unset. `--plan` validates that shape without touching
those paths. At final freeze, fill every placeholder with the coordinator's
selected full candidate ref, source tree, lock, probe, target, and binary
identities; the runner rejects a still-parameterized candidate.

The correctness preflight runs both normal endpoints and requires the same
synthetic source digest and all four route proofs. The timed phase then uses
two `A1, B1, B2, A2` blocks for every route, with the same calibrated
iteration count for the normal and allocator legs. Each child is pinned to
the manifest CPU, receives the fixed 20,000,000,000-byte address-space cap,
and is monitored for output growth, RSS, target growth, descendants, host
load, and unrelated Cargo/rustc/rustdoc/rustfmt/clippy/profiler work. The
runner keeps logs and refuses stale output directories; it never deletes a
target or source artifact.

The fixed target preflight ceiling is 19,200,000,000 bytes and the hard target
ceiling is 20,000,000,000 bytes. Child output is capped at 4 MiB. A failed
guard, changed source/lock/binary/probe identity, mismatched correctness
proof, missing allocator observation, or incomplete ABBA order fails closed.
No threshold is weakened here; the eventual performance decision remains
subject to the coordinator's existing regression policy after a quiet-host
execution.

## Baseline adapter and build boundary

The baseline parent predates the probe module. At an authorized build slot,
copy the three checked-in probe sources into a temporary clean baseline
checkout at the exact paths below, and copy the one-line `lib.rs` module
registration if the old checkout does not already expose it:

```text
tools/perf-baseline/src/ppt_lazy_native.rs
tools/perf-baseline/src/bin/ppt_lazy_native.rs
tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs
tools/perf-baseline/src/lib.rs  (add `pub mod ppt_lazy_native;` only for the temporary build overlay)
```

Record the source-file hashes, baseline ref/tree, lock hash, exact commands,
external target paths, and both resulting executable hashes in the external
build receipt. Remove only the named temporary overlay and require the
baseline checkout to be clean before handing it to the runner. Build normal
and allocator targets separately, with one job, incremental compilation off,
offline/locked mode, and the repository's 20-GB memory-capped cargo shim.
The candidate uses the same sources and commands from its final clean source
tip. The build and receipt remain outside the repository; this record carries
no executable or raw evidence.

## Source-only validation

The following checks are intentionally build-free and are the only checks run
while preparing this lane:

```sh
python3 -m py_compile tools/ppt_lazy_native_qualification.py
PYTHONPATH=tools python3 -m unittest tools/test_ppt_lazy_native_qualification.py
rustfmt --edition 2024 --check \
  tools/perf-baseline/src/ppt_lazy_native.rs \
  tools/perf-baseline/src/bin/ppt_lazy_native.rs \
  tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs
git diff --check
python3 tools/ppt_lazy_native_qualification.py \
  --manifest tools/native_qualification/manifests/ppt-lazy-edit.json \
  --out-dir /external/ppt-lazy-native/plan \
  --plan
```

The plan proves the fixed route set, baseline parent, ABBA order, and resource
caps. It does not claim that either binary has been built or that any native
measurement has passed.
