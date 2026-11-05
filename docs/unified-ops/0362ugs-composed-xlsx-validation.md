# 0362ugs — composed XLSX validation

This is the validation receipt for the composed W8/W4 XLSX source-backed,
catalog, batch, typed-limit, pivot, graph and object families.

The composed integration started at `7ed6fb0ca71f4c15f3c3a3b6b9afdb056e46b610`.
The final source tip for this receipt is
`784d0ea83e54936677ebd5fc1d798eacd6d5c88f`; the composed consumer tip before
the planning-guard repair was `25fe21158fbe10e9919e0774bbc7eda5c68f31ad`.
This provenance correction
supersedes the initial receipt commit `6bed5eb062379473994a4482330a35cfc425a2ae`.

This was not one clean-tree run. Each row below records the exact committed
tip and, where applicable, the dirty overlay present while that command ran.
The XLSX all-target run used clean `520efdd0ab8ca7e2ddc3a1943fd03f2cd5770420`.
The later strict Clippy run used that same commit plus the uncommitted overlay
which became `519eb99c6cbd20a6a921752d715c779051a14194`; that overlay changes
XLSX implementation files, so the all-target result is not claimed for the
later lint tip. Subsequent consumer checks likewise identify their overlays
explicitly.

The earlier XLS doctest assignment was transferred to the 0362ugea XLS follow-up.
The XLSX doctest result below remains part of this receipt.

## Environment

- Repository root: `/home/zhuhe/code/litchi-worktrees/uo2-xlsx-lazy-entry`
- Root `Cargo.toml`: 10,672 bytes, SHA-256
  `f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`
- Toolchain/build environment: `(deleted-coord-artifact)/env.sh`
- `env.sh`: 489 bytes, SHA-256
  `85c2df8dedebbf2167037549733f70b9721d5a0eeb99f9b05760e91ad2dc1631`
- Cargo config: `.cargo/config.toml`, 470 bytes, SHA-256
  `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`
- Rust: `rustc 1.95.0 (59807616e 2026-04-14)`, host
  `x86_64-unknown-linux-gnu`, LLVM `22.1.2`
- Cargo: `cargo 1.95.0 (f2d3ce0bd 2026-03-21)`
- Cargo target: `(deleted-coord-artifact)`
- `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`,
  `CARGO_PROFILE_DEV_DEBUG=0`, `CARGO_PROFILE_TEST_DEBUG=0`,
  `CARGO_PROFILE_RELEASE_DEBUG=0`, `CARGO_NET_OFFLINE=true`,
  `LITCHI_MEM_MAX=20G`
- Final target size: `5,092,401,674` bytes
- No native build, timing, allocator, or performance qualification was run.

The tracked sparse fixture needed by the XLSX corpus oracle was materialized
from `docs/report/spec-gap-validation-evidence/xlsx-form-control-properties`.
No fixture or large evidence artifact was committed.

Lock identities used by the commands:

| Lockfile | Bytes | SHA-256 |
| --- | ---: | --- |
| `Cargo.lock` | 139,591 | `76340b069df633ff1c63c98e93e8b10c1e7711164ca2515f490872142ab46c72` |
| `tools/perf-baseline/Cargo.lock` | 57,158 | `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184` |
| `tools/perf-baseline-f3/Cargo.lock` | 34,004 | `e263612d864e4592b80be6ef626948a4eb73b9be31537ac78f5ea00243fc62fb` |
| `tools/perf-baseline/unified-xlsx-batch/Cargo.lock` | 30,990 | `aec8986b4755544318b7926ff9ce3c0b921263309a83d9afdcccbe1801e93c6c` |
| `tools/perf-baseline/unified-xlsx-raw-identity/Cargo.lock` | 30,553 | `8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6` |
| generated fuzz lock (external) `(deleted-coord-artifact)/ops-fuzz-generated-Cargo.lock` | 32,534 | `10708241205e4a90c0120b3f41a77e7548715761801d66da2d55911e59e9ab8a` |

## Commands and receipts

All commands were run from the repository root after sourcing the environment
above. Logs are retained outside the repository under
`(deleted-coord-artifact)`.

| Check | Exact source state | Command | Result and external log |
| --- | --- | --- | --- |
| XLSX all targets | clean `520efdd0ab8ca7e2ddc3a1943fd03f2cd5770420` | `cargo test --offline --locked -p litchi-xlsx --all-features --all-targets --no-fail-fast` | 98 test binaries, 2,558 passed, 0 failed, 0 ignored; `xlsx-all-targets-520efdd0ab-final.log`, 258,738 bytes, SHA-256 `5df96e8622758fe25fceaea3739ebd594479dd65b004f67351826c67e48fc9f8` |
| XLSX doctests | clean `520efdd0ab8ca7e2ddc3a1943fd03f2cd5770420` | `cargo test --offline --locked -p litchi-xlsx --all-features --doc` | 36 passed (35 ordinary doctests and 1 compile-fail doctest), 0 failed; `xlsx-doctests-520efdd0ab.log`, 4,057 bytes, SHA-256 `4e696a7945e711382ea01a926ba694e40c049eeda946a3301b2a4176c53d3da7` |
| Strict XLSX Clippy | `520efdd0ab8ca7e2ddc3a1943fd03f2cd5770420` + dirty overlay equivalent to `519eb99c6cbd20a6a921752d715c779051a14194`; overlay SHA-256 `870ace9802e1f9d3e7da319652a812ab4ce727a686ce00af5d3b3a78dacbe542` | `cargo clippy --offline --locked -p litchi-xlsx --all-features --all-targets -- -D warnings` | Exit 0; `xlsx-clippy-520efdd0ab-final2.log`, 168 bytes, SHA-256 `41bb06327ad47e14fa78cec5a2221abe2c3f2894af8137cf022f36f473e93947` |
| Ops all targets/features | `151cc504a2436958f2cfc7b2b634ca4761efa3f1` + dirty overlay equivalent to `a06d70c2fea76dd2b78115d53be60e4a308e117c`; overlay SHA-256 `f3f9510e8b79daca3795ab8a5d1bf6214cbc481202291307b2822f6805ff23bc` | `cargo check --offline --locked -p litchi-ops --all-features --all-targets` | Exit 0, including `odf_shared_triad_qualification`; `ops-check-odf-fix3.log`, 166 bytes, SHA-256 `588d9d5a28a8fb632e4420b634546144d98ee1a7a4a9c5e4c757c356d025c070` |
| Perf baseline | `a06d70c2fea76dd2b78115d53be60e4a308e117c` + dirty overlay equivalent to `158b57f3d08f93b792a2cf2cb9e53ea36a350386`; overlay SHA-256 `963ea059980706b36e90d96bea2453bd34a6bc25e28c234d3b985df7a0944422` | `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets --all-features` | Exit 0; `perf-baseline-check-xlsx-api.log`, 178 bytes, SHA-256 `bd00072f3c71b6a252bc2ca49f7d10c74be4ca5272259b51e8f029a02a01bbc8` |
| Perf baseline F3 | clean `158b57f3d08f93b792a2cf2cb9e53ea36a350386` | `cargo check --offline --locked --manifest-path tools/perf-baseline-f3/Cargo.toml --all-targets --all-features` | Exit 0; `perf-baseline-f3-check-158b57f3d0.log`, 3,225 bytes, SHA-256 `28480e1b7e1992f55833064ea84f16b5b3468f9e3993b514dc917806ebe2e718` |
| XLSX batch perf consumer | `158b57f3d08f93b792a2cf2cb9e53ea36a350386` + dirty overlay equivalent to `25fe21158fbe10e9919e0774bbc7eda5c68f31ad`; overlay SHA-256 `6b4f4204c246cc8f91e1f97b2836231396a3cd8c42162ac3f65a623d4ff0516e` | `cargo check --offline --locked --manifest-path tools/perf-baseline/unified-xlsx-batch/Cargo.toml --all-targets --all-features` | Exit 0; `perf-unified-xlsx-batch-check-final.log`, 203 bytes, SHA-256 `0731123db61173844e110c3010c73120bc155f8ecb8df308c5a3c7525eb89c79` |
| XLSX raw-identity perf consumer | `158b57f3d08f93b792a2cf2cb9e53ea36a350386` + dirty overlay equivalent to `25fe21158fbe10e9919e0774bbc7eda5c68f31ad`; overlay SHA-256 `6b4f4204c246cc8f91e1f97b2836231396a3cd8c42162ac3f65a623d4ff0516e` | `cargo check --offline --locked --manifest-path tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml --all-targets --all-features` | Exit 0; only existing dead-code warnings in the shared allocation harness; `perf-unified-xlsx-raw-identity-check.log`, 4,490 bytes, SHA-256 `1cc89dfed2e6da947f3b61ca26b9ff5a011901f159c5ef758ff82263a255781f` |
| Unified-ops fuzz consumer | `158b57f3d08f93b792a2cf2cb9e53ea36a350386` + dirty overlay equivalent to `25fe21158fbe10e9919e0774bbc7eda5c68f31ad`; overlay SHA-256 `6b4f4204c246cc8f91e1f97b2836231396a3cd8c42162ac3f65a623d4ff0516e` | `cargo check --offline --manifest-path crates/litchi-ops/fuzz/Cargo.toml --all-targets --features apply` | Exit 0; `ops-fuzz-check-offline.log`, 1,988 bytes, SHA-256 `696778cc30c5e6070bffaf0467dfdc92cc653eb0bf9dc66483835fd20b83bd3b` |
| Planning guard valid route | `b7310e777a5de6d28543ea1bb91ee9fc8604bfc5` + dirty overlay equivalent to `784d0ea83e54936677ebd5fc1d798eacd6d5c88f`; overlay SHA-256 `c9fcd5637c400282a45a57d475935b5df6b8de3456facb641a04b3a5793b8ffa` | `cargo run --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin xlsx_planning_guard -- --shape medium --case valid --warmup 0 --samples 1` | Debug correctness route passed: source unchanged, snapshot values valid, empty commit exact no-op; its `duration_ns` fields are excluded from evidence and no timing claim is made. `xlsx-planning-guard-valid-fixed.log`, 2,854 bytes, SHA-256 `840a5df300cad6d52ac7c4d787bfc599a7b60636be7a7e7157629b75022c7e30` |
| Planning guard late validator | `b7310e777a5de6d28543ea1bb91ee9fc8604bfc5` + dirty overlay equivalent to `784d0ea83e54936677ebd5fc1d798eacd6d5c88f`; overlay SHA-256 `c9fcd5637c400282a45a57d475935b5df6b8de3456facb641a04b3a5793b8ffa` | `cargo run --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin xlsx_planning_guard -- --shape medium --case late-validator --warmup 0 --samples 1` | Debug correctness route passed: exact typed `Invalid`, exact unknown-element message, source unchanged, retry preserved; its `duration_ns` fields are excluded from evidence and no timing claim is made. `xlsx-planning-guard-late-validator-fixed.log`, 3,072 bytes, SHA-256 `184dc0d8382883ca594d08701c02d516ac89e8425efad5716b5f9d4bda773470` |
| Planning guard late raw | `b7310e777a5de6d28543ea1bb91ee9fc8604bfc5` + dirty overlay equivalent to `784d0ea83e54936677ebd5fc1d798eacd6d5c88f`; overlay SHA-256 `c9fcd5637c400282a45a57d475935b5df6b8de3456facb641a04b3a5793b8ffa` | `cargo run --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin xlsx_planning_guard -- --shape medium --case late-raw --warmup 0 --samples 1` | Debug correctness route passed: exact typed `Invalid`, exact boolean message, source unchanged, retry preserved; its `duration_ns` fields are excluded from evidence and no timing claim is made. `xlsx-planning-guard-late-raw-fixed.log`, 2,905 bytes, SHA-256 `c37723c9775232a8b75fe85ddd1930377891c9f99def600c81207dd5bc59fa56` |

The 784 planning-guard fixture correction is intentional, not a weakened
refusal. Its original late-validator cell used `future="1"`, but the current
source contract admits and preserves unknown attributes: the existing
`cell_values/validation.rs` `scan_attributes` contract says every checked
attribute survives, and
`raw/worksheet/tests.rs::eager_cell_attributes_ignore_prefixed_fields_beside_unqualified_fields`
proves an unknown `future` attribute is ignored. The snapshot evidence also
distinguishes preserved unfamiliar attributes from unknown elements inside the
composed `sheetData` span. Therefore first use correctly accepted the old
attribute fixture. The fixture now places `<future/>` inside the modeled cell;
the existing `validate_element` unknown-element branch rejects it with the
recorded typed `Invalid` message. No production acceptance/refusal rule was
changed.

The three guard commands are debug `cargo run` diagnostics and necessarily
populate the schema's `duration_ns` fields. “Debug correctness route” means
those timing fields were excluded from evidence; it does not mean the clock
was not called, and no timing qualification is claimed.

The standalone `unified-xlsx-batch` lock required an offline refresh because
the composed `litchi-ops` feature graph added the tracked `litchi-sheet-text`
path dependency; the final command above was rerun with `--locked` after that
lock update. The fuzz manifest intentionally has no tracked lockfile; Cargo's
generated lock was moved out of the checkout after its offline check.
