# 0362ugja — PPT lazy-entry native qualification protocol

This record prepares the missing native qualification lane for the PPT lazy
edit entry correction in 0362ugj. It is source and runner infrastructure only:
no release build, native timing, allocator capture, or performance decision was
made by this commit. One authorized compile-only type check of the probe is
recorded below; it did not produce a qualification binary or execute a route.

The bounded compile follow-up and corrected historical-baseline fixture
materialization are recorded in
[`0362ugjab-ppt-lazy-native-preparation-return.md`](0362ugjab-ppt-lazy-native-preparation-return.md).
That return supersedes the source-only build-boundary paragraph below for
future endpoint preparation; this record remains the original protocol
contract and preserves its parameterized manifest.

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
allocator binaries. It opens the reviewed non-empty fixture before each
operation boundary, then measures only the operation. Every affected lazy
constructor has both an entry/drop route and an exact no-op commit route:

```text
primary, metadata, animations, broadcast, client_data, diagrams,
document_comparison, objects, external_media, fonts, master_metadata,
master_layout, picture_bullets, slide_sync
```

The default fixture has one slide with two text shapes, client data, a PPT9
broadcast record, a PPT10 diagram build list, and a valid main master. It is
tracked at `test-data/unified-ops/ppt-lazy-native-fixture.ppt`; its SHA-256 is
`02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c` and its
size is 14,848 bytes. Three owner-specific routes use additional tracked,
non-empty representatives: `objects` uses Apache POI's
`test-data/poi/test-data/slideshow/ppt_with_embeded.ppt` (180,736 bytes,
`afa1dd5adc36765e6e40725a133b0330e5ad1febd1506ca93ceae8c5c5400b53`) with a
real `ExObjListContainer` and `Pictures` stream; `external_media` uses
`test-data/unified-ops/0294/video.ppt` (43,520 bytes,
`f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a`) with a
real movie object; and `picture_bullets` uses Apache POI's
`test-data/poi/test-data/slideshow/bug61881.ppt` (74,752 bytes,
`ba2bf551e87f5fd7667dec52ac19005396bdf9af00f0006f9a247a9fede2e8ee`) with a
real `BlipCollection9Container` (record 2040) and `Pictures` stream. The
manifest binds every route to one of these four exact fixture identities, and
each proof repeats its fixture key and digest. Thus every
position/target/kind selector is exercised against a real non-empty owner.
Drop proofs publish nothing; commit proofs require unchanged diagnostics, a
no-op patch, and byte-identical output. The primary proof additionally
requires zero operation count and exact patch before/after endpoints; metadata
requires exact `before_bytes`/`after_bytes` endpoints.

`--mode verify` runs all 28 routes once and emits all four fixture generator
identities plus byte count/SHA-256. `--mode calibrate` doubles a route-local
operation batch to a fixed target and repeats the verified route fixture digest
and route proof.
`--mode timed` runs warmups outside the clock and returns one elapsed batch
sample per request; a postcondition proof is performed outside the timed
interval. The allocator binary uses the existing `allocation_metrics`
system-allocator observer and reports allocation/deallocation/reallocation
calls, bytes, live bytes, and monotonic high-water bytes. The normal binary
reports allocator metrics as unavailable; an uninstrumented run is never
interpreted as zero. The probe rejects counter overflow, live-byte conservation
failures, and peak regressions. A reallocation may legally occur without a
same-interval allocation call, so no invalid `reallocation_calls <=
allocation_calls` invariant is applied.

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
baseline ref/tree and exact lock path, the reviewed relative probe closure, the
four fixture identities and exact route map, the baseline overlay paths and
`src/lib.rs` registration, and the
release build commands. The normal and allocator builds use the package
defaults explicitly recorded as `current-api,current-xlsx-raw-opc` (the
commands omit `--no-default-features` and therefore select those defaults).
It leaves external roots, source archive, build
receipt, toolchain/config identities, binary identities, and the final
candidate unset. `--plan` validates that shape without touching those paths.
At final freeze, fill every placeholder with the coordinator's selected full
candidate ref, source tree, lock, probe, target, binary, archive, and retained
build-receipt identities; the runner rejects a still-parameterized candidate.

The correctness preflight runs both normal endpoints and requires the same
reviewed fixture digest and all 28 route proofs. The timed phase then uses
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

The manifest's retained build receipt must bind the source archive, toolchain,
Cargo configuration, exact source/lock/tree identities, all four fixture
identities and their route map, all four
candidate probe files, the three-file baseline overlay plus the exact
`src/lib.rs` registration diff, commands, environment, release target/profile,
features, target paths, and both resulting executable hashes. Remove only the
named temporary overlay and require the baseline checkout to be clean before
handing it to the runner. Build normal and allocator targets separately, with
one job, incremental compilation off, offline/locked mode, and the
repository's 20-GB memory-capped cargo shim. The binaries are the actual
`<CARGO_TARGET_DIR>/release/ppt_lazy_native` and
`<CARGO_TARGET_DIR>/release/ppt_lazy_native_alloc` outputs (not an
`release/examples` path). The candidate uses the same sources and commands
from its final clean source tip. The build, archive, receipt, and logs remain
outside the repository; this record carries no executable or raw evidence.

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

An authorized compile-only slot type-checked the shared probe through a
temporary external mini-crate that supplied the allocator-metrics interface;
this avoided a release build and did not run the probe. The exact command was
`cargo check --offline --manifest-path /home/zhuhe/code/litchi-worktrees/.coord/ppt-probe-check/Cargo.toml`
with the repository memory-capped Cargo shim, `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, `CARGO_NET_OFFLINE=true`, an
external target directory, and a 20-GB cap. It exited 0; the external target
was 281 MiB. The 6,687-byte log is retained outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs/probe-harness-final.log`
with SHA-256
`012cbf16ad29524ee3a9ca5739f92de10da3bc8b6cdcfeece7bbd5e662ed1cf7`.
