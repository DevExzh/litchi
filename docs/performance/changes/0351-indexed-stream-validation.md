# Change 0351: indexed-stream validation

Status: correctness and resource hardening only

`performance_claim: none`

## Rejected premise

The initial compressor/zlib premise was rejected. No artifact establishes a
`~65%` residual-zlib result, and the existing bounded `read_to*` path already
uses a fixed-buffer verifier. This change makes no compression, zlib, speed,
or optimization claim.

## Strict sink validation

The strict sink path preflights target encryption, the supported compression
method, and resolved single-disk ZIP64 provenance. It then requires complete
local/central raw name, flags, method, CRC, and size agreement, including
signed and unsigned ZIP32/ZIP64 data-descriptor forms. Every physical span,
including directory spans, participates in the layout proof. Prefixes and gaps
are allowed; overlaps and central-directory intrusion are refused.

The locator validates classic and ZIP64 disk and entry counts, locator and
record offsets, record length and adjacency, short buffers, and prefixed or
suffixed ZIP64 resolved metadata. Physical entry count is bounded by the
declared central size and fixed 46-byte records before fallible retention.
Invalid UTF-8, key, and scratch-buffer allocations are fallible. The strict
layout proof is a single-flight cache of successful immutable-source metadata;
re-entry is rejected and failures may retry. The `ReaderAt` byte-stability
contract remains required.

Store uses its exact source range. Deflate uses the decoder's exact `total_in`.
Strict CRC equality, including zero, applies only to bounded sink
`read_to*`/`read_entry_to*` paths. Ordinary owned reads retain their documented
zero-CRC compatibility, and borrowed access retains its nonempty-zero fallback
behavior.

## Bounded resource boundary

The indexed streaming path uses one 16 KiB scratch buffer for one active
member. This excludes caller-owned source, archive/index, sink/output, cache,
aggregate process memory, and concurrency. It is a local structural invariant,
not a total-memory or physical-I/O bound.

## Validation evidence

The final successful invocations (package/scenario scoped, not workspace-wide)
were:

```sh
cargo fmt --package soapberry-zip -- --check

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0351 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p soapberry-zip --lib -- --test-threads=1
# => 315/315

ulimit -v 8388608
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/change-0351 \
CARGO_BUILD_JOBS=1 \
CARGO_INCREMENTAL=0 \
CARGO_PROFILE_DEV_DEBUG=0 \
CARGO_PROFILE_TEST_DEBUG=0 \
RUST_TEST_THREADS=1 \
cargo test -p litchi-opc --test operation_accounting -- --test-threads=1
# => 13/13
```

The package formatting step succeeded. The final serialized results were:

- `soapberry-zip` library: `315/315`.
- `litchi-opc` `operation_accounting` filter: `13/13`.

All validation used one build job, one test thread, incremental and debug
compilation disabled, one disk target, and an 8 GiB `ulimit`. Independent
source reviews found no remaining P0/P1 after the fixes. Only final successful
results are recorded; transient compile-feedback iterations are excluded.

## Claim boundary

No latency, throughput, RSS, allocation, syscall, physical-I/O, decompression,
or concurrency claim follows. No performance selector, raw result artifact, or
claim-registry entry is added. This record does not claim completion of the
broader GOAL; it records indexed validation and bounded failure/resource
semantics only.
