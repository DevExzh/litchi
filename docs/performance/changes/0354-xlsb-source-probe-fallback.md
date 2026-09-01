# Change 0354: XLSB source-probe error and fallback admission

Status: implemented

`performance_claim: none`

## Probe outcome boundary

The private workbook/XLSB source ingress now distinguishes recoverable
pre-admission outcomes from hard ZIP/OPC/classifier failures. A non-ZIP input,
a ZIP with no XLSB match, or a missing manifest remains eligible for the
compatibility fallback. A hard ZIP, OPC, or classifier failure now returns the
typed `OpcError` outcome; `Workbook::from_bytes` does not retry that outcome by
eagerly parsing the whole workbook. This keeps malformed or hostile packages
from crossing the source-owner admission boundary as an apparently recoverable
format miss.

## FileSource fallback ownership

The pathname `FileSource` path preflights the caller's exact `max_input_bytes`
before allocating fallback storage, and the fallback reader receives that
same exact limit. The source catalog is dropped before the fallback read. The
retained `Bytes` value is moved into the owned source without a clone, with its
pointer/capacity ownership preserved. Known non-XLSB detector variants return
`NotOfficeFile` directly and do not reopen the pathname. Explicit eager APIs,
the public smart detector, and the positive non-ZIP compatibility fallback
remain unchanged.

## Evidence and resource boundary

The final serial commands used one 8 GiB virtual-memory ceiling, one Cargo
job, one test thread, disabled incremental/debug compilation, and one disk
target:

```sh
ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0354 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi-xlsb --lib -- --test-threads=1
# => 51/51; the XLSB-only private source-ingress filter is 7/7

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0354 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi --test xlsb_facade --features xlsb -- --test-threads=1
# => 23/23

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0354 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo check -p litchi --no-default-features --features xlsx
# => success

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0354 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo check -p litchi --no-default-features --features xlsx,xlsb
# => success
```

The target reached 564 MiB. Post-run available memory was observed at
approximately 14 GiB while swap remained saturated; all Cargo commands were
serialized. These are constrained-run observations only and are not an
RSS, OOM, total-memory, latency, or constant-memory bound.

## Remaining scope

An analogous PPTX hard-probe fallback, a public eager-detector/explicit-eager
portable fallback split, and finer constructor ZIP resource mapping remain
open. Full selected worksheets still materialize. No latency, throughput,
allocation, physical-I/O, decompression, or broad XLSB claim follows.
