# Change 0353: XLSB source-backed fallback admission boundary

Status: implemented

`performance_claim: none`

## Admission and fallback boundary

Dynamic/source-backed XLSB text no longer retries through an eager full-workbook
fallback after the source owner has admitted the operation. The eager adapter
reader, eager adapter caches, adapter state, and private detector-side duplicate
source/limits state were removed. The original typed source errors, freshness
fences, and eager workbook cache behavior remain. Explicit
`open_xlsb_workbook*` eager APIs and the `DetectedFormat::Xlsb` payload remain
unchanged.

The source owner now walks recognized nonworksheet tabs through filtered
worksheet positions, so they do not cause a false unsupported result. A direct
nonworksheet selection remains a typed refusal, as do sparkline, pivot, slicer,
and timeline selections. Recoverable dynamic/pre-admission source probes may
still use the existing `Workbook::from_bytes` fallback; that path is distinct
from post-admission `UnsupportedFeature`. Change 0304's statement that every
source-owner `UnsupportedFeature` triggers eager fallback is superseded only
after source-owner admission.

## Evidence and resource boundary

The final successful package/scenario-scoped commands, run serially, were:

```sh
ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0353 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi --test xlsb_facade --features xlsb -- --test-threads=1
# => 23/23

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0353 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi-xlsb --test source_backed -- --test-threads=1
# => 40/40
```

The dedicated target reached 647 MiB. Post-run available memory was observed
at approximately 15 GiB while swap remained saturated; the Cargo runs were
serialized. These are constrained-run observations only and are not an RSS,
OOM, total-memory, or constant-memory bound.

## Remaining eager paths and scope

Smart `DetectedFormat` detection, explicit typed eager APIs, unsupported
platforms, dynamic/pre-admission recoverable source probes, and
`Workbook::from_bytes` fallback remain eager where their existing contracts
require it. A selected source-backed worksheet and its required dependencies
still materialize. No latency, throughput, RSS, allocation, physical-I/O,
decompression, concurrency, or benchmark claim follows.
