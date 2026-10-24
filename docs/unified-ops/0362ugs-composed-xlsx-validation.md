# 0362ugs — composed XLSX validation

This is the validation receipt for the composed W8/W4 XLSX source-backed,
catalog, batch, typed-limit, pivot, graph and object families.

The composed integration started at `7ed6fb0ca71f4c15f3c3a3b6b9afdb056e46b610`.
The final source tip for this receipt is
`25fe21158fbe10e9919e0774bbc7eda5c68f31ad`. The XLSX test, doctest and Clippy
runs used `520efdd0ab8ca7e2ddc3a1943fd03f2cd5770420`; the commits after that
tip only repaired merged example/perf consumers and strict lint findings, and
did not alter the XLSX implementation after the tested source tip. The ops
and standalone consumer checks used their final consumer tips as listed below.

## Environment

- Toolchain/build environment: `/home/zhuhe/code/litchi-worktrees/.coord/xlsx-lazy-entry/env.sh`
- Cargo target: `/home/zhuhe/code/litchi-worktrees/.coord/xlsx-lazy-entry/target`
- `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`,
  `CARGO_NET_OFFLINE=true`, `LITCHI_MEM_MAX=20G`
- Final target size: `3,415,574,738` bytes
- No native build, timing, allocator, or performance qualification was run.

The tracked sparse fixture needed by the XLSX corpus oracle was materialized
from `docs/report/spec-gap-validation-evidence/xlsx-form-control-properties`.
No fixture or large evidence artifact was committed.

## Commands and receipts

All commands were run from the repository root after sourcing the environment
above. Logs are retained outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/xlsx-lazy-entry/`.

| Check | Command | Result and external log |
| --- | --- | --- |
| XLSX all targets | `cargo test --offline --locked -p litchi-xlsx --all-features --all-targets --no-fail-fast` | 98 test binaries, 2,558 passed, 0 failed; `xlsx-all-targets-520efdd0ab-final.log`, 258,738 bytes, SHA-256 `5df96e8622758fe25fceaea3739ebd594479dd65b004f67351826c67e48fc9f8` |
| XLSX doctests | `cargo test --offline --locked -p litchi-xlsx --all-features --doc` | 36 passed (35 ordinary doctests and 1 compile-fail doctest), 0 failed; `xlsx-doctests-520efdd0ab.log`, 4,057 bytes, SHA-256 `4e696a7945e711382ea01a926ba694e40c049eeda946a3301b2a4176c53d3da7` |
| Strict XLSX Clippy | `cargo clippy --offline --locked -p litchi-xlsx --all-features --all-targets -- -D warnings` | Exit 0; `xlsx-clippy-520efdd0ab-final2.log`, 168 bytes, SHA-256 `41bb06327ad47e14fa78cec5a2221abe2c3f2894af8137cf022f36f473e93947` |
| Ops all targets/features | `cargo check --offline --locked -p litchi-ops --all-features --all-targets` | Exit 0, including `odf_shared_triad_qualification`; `ops-check-odf-fix3.log`, 166 bytes, SHA-256 `588d9d5a28a8fb632e4420b634546144d98ee1a7a4a9c5e4c757c356d025c070` |
| Perf baseline | `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets --all-features` | Exit 0 at `158b57f3d08f93b792a2cf2cb9e53ea36a350386`; `perf-baseline-check-xlsx-api.log`, 178 bytes, SHA-256 `bd00072f3c71b6a252bc2ca49f7d10c74be4ca5272259b51e8f029a02a01bbc8` |
| Perf baseline F3 | `cargo check --offline --locked --manifest-path tools/perf-baseline-f3/Cargo.toml --all-targets --all-features` | Exit 0; `perf-baseline-f3-check-158b57f3d0.log`, 3,225 bytes, SHA-256 `28480e1b7e1992f55833064ea84f16b5b3468f9e3993b514dc917806ebe2e718` |
| XLSX batch perf consumer | `cargo check --offline --locked --manifest-path tools/perf-baseline/unified-xlsx-batch/Cargo.toml --all-targets --all-features` | Exit 0 at `25fe21158f`; `perf-unified-xlsx-batch-check-final.log`, 203 bytes, SHA-256 `0731123db61173844e110c3010c73120bc155f8ecb8df308c5a3c7525eb89c79` |
| XLSX raw-identity perf consumer | `cargo check --offline --locked --manifest-path tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml --all-targets --all-features` | Exit 0; only existing dead-code warnings in the shared allocation harness; `perf-unified-xlsx-raw-identity-check.log`, 4,490 bytes, SHA-256 `1cc89dfed2e6da947f3b61ca26b9ff5a011901f159c5ef758ff82263a255781f` |
| Unified-ops fuzz consumer | `cargo check --offline --manifest-path crates/litchi-ops/fuzz/Cargo.toml --all-targets --features apply` | Exit 0; `ops-fuzz-check-offline.log`, 1,988 bytes, SHA-256 `696778cc30c5e6070bffaf0467dfdc92cc653eb0bf9dc66483835fd20b83bd3b` |

The standalone `unified-xlsx-batch` lock required an offline refresh because
the composed `litchi-ops` feature graph added the tracked `litchi-sheet-text`
path dependency; the final command above was rerun with `--locked` after that
lock update. The fuzz manifest intentionally has no tracked lockfile; Cargo's
generated lock was moved out of the checkout after its offline check.
