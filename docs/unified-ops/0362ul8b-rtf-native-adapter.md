# 0362ul8b — paired RTF native consumer adapter repair

Date: 2026-10-24.  This packet repairs the RTF qualification consumer at the
original source fence `86d708878fa537eab42fd4bd5ffbc0f79c606191` and the
current candidate staging fence `35e37dbede6657dd023b114a1536862c18541946`.
It is focused validation, not final native-performance acceptance; the
coordinator's composed gate remains required.

## Adapter boundary

The active probe is
`tools/perf-baseline/examples/refusal_contract_bench.rs`, SHA-256
`54e1c94472fddef671955126bc35fcd5e0a77afafbdff109b65bc87c4cd66908` at both
legs.  It is a consumer-only adapter selected by standalone package features:

| leg | feature | native refusal API | `Package::apply` result | lock SHA-256 |
| --- | --- | --- | --- | --- |
| before | `rtf-legacy-api` | `edit::Error::{UnsupportedSource,ProtectedDocument}` | `Package` | `57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f` |
| after | `rtf-current-api` | `edit::Error::Refused(Refusal)` | `edit::Commit` | `6dbfc0bcc64cdf8e625a6562f2d5d11c1e98ad7cb6edda299b9adf2a36a70e39` |

The baseline and candidate `tools/perf-baseline/Cargo.toml` overlays are
`af876fbf715edb643d2c7abb8643efb252413c640824890e90b1849207889676` and
`60cde3ec148b08d81806d0b15740808c37ecbc9eed8ecf3f4d32e20ce8631a49`.
No `litchi-rtf` implementation or production feature was overlaid.  The
baseline lock was taken from its pinned 86d tree; the requested external
`rtf-refusals/historicalaudit` directory is absent in this checkout, so no
current lock was copied into the baseline.

The live current workspace example is separately kept at
`crates/litchi-rtf/examples/refusal_contract_bench.rs`, SHA-256
`51d98d9dffe628acf9236929fc6952e05982911280149dcd18399d45d4972485`.  It is
not substituted into the paired baseline adapter because the 86d baseline has
no such workspace example; the standalone probe remains the only paired source
for the native-era comparison.  Its locked current-crate check used the
external pinned 4100 lock, SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.

The adapter matches the actual native variant before normalizing the fixed
contract fields.  The baseline `UnsupportedSource(&'static str)` and
`ProtectedDocument { protection_type }` variants are not relabeled as a
baseline `RefusalKind`; the check-only receipt records their provenance.
Current protection is checked at `commit`, after successful `replace_text`
staging.  All refusal reasons are fixed, content-free text.

## Historical diagnostic retained

The pre-adapter probe (`crates/litchi-rtf/examples/refusal_contract_bench.rs`)
remains historical at SHA-256
`70077fcec439fda2fdf9e9635f5db2f1f7c86ed31f3f972a95c5e5e3a607a585`.  Its
actual pinned check was a current-candidate check at the 35e37 staging fence;
it failed because current `Package::parse` expects `&str` rather than the
provided `&Vec<u8>`, and current `Package::apply` returns `edit::Commit`, not a
`Package` with `.to_bytes()`.  The retained receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/rtf-probe-compile/check-before.json`
(SHA-256 `f3af926353fc3db4f74221f62ed4f4743cc6b31c359ab42415c8f1500e33bda4`),
with log SHA-256
`119d4216675ee79a7e89e17579f019dd4fe3a6d43c08a8bd22d87c37622fd103` and exit
status 101.  This failure is diagnostic evidence, not a passing build.

## Focused validation

The capped environment was `PATH` with `.coord/bin` first,
`LITCHI_CAPPED`/`LITCHI_BLESS` unset, `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, offline and locked, with each
target kept below 20,000,000,000 bytes.  The two live current-crate commands
ran in the source-only scratch checkout
`/home/zhuhe/code/litchi-worktrees/.coord/rtf-live-current-kQPuva`, with the
4100 lock recorded above; that lock was not added to this repair worktree.
The corrected commands were:

```text
cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --no-default-features --features rtf-current-api --example refusal_contract_bench
cargo run --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --no-default-features --features rtf-current-api --example refusal_contract_bench -- --check-only
cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --no-default-features --features rtf-legacy-api --example refusal_contract_bench
cargo run --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --no-default-features --features rtf-legacy-api --example refusal_contract_bench -- --check-only
cargo check --offline --locked -p litchi-rtf --example refusal_contract_bench
cargo run --offline --locked -p litchi-rtf --example refusal_contract_bench -- --check-only
```

All six commands exited 0.  The paired focused check-only outputs were:

```text
before: validation=passed, native_error_api=litchi_rtf::edit::Error::{UnsupportedSource,ProtectedDocument}, apply_result_api=Package, routes=8
after:  validation=passed, native_error_api=litchi_rtf::edit::Error::Refused(Refusal), apply_result_api=edit::Commit, routes=8
```

The live current workspace check-only output was:

```text
validation=passed, native_error_api=litchi_rtf::edit::Error::Refused(Refusal), apply_result_api=litchi_rtf::Package::apply -> litchi_rtf::edit::Commit, routes=8
```

The candidate locked-check log is retained outside the checkout at
`.coord/rtf-probe-compile/adapter-current-check-locked.log` (SHA-256
`bcce8c4403778d68ed4b82507bf8a740d963c3ebc076be246caac13f56fbea5e`); the
candidate check-only log is `adapter-current-check-only.log` (SHA-256
`46f0b0539e5c86cb1a8b47d3cd37bfaec850b51b0b82c979a72d14f6deebd81c`).  The
baseline equivalents are `adapter-baseline-check-locked.log` (SHA-256
`72131d35df8d9815e5b9545782526578dabbef6f8a363d4eb84282e02ed22937`) and
`adapter-baseline-check-only.log` (SHA-256
`37ffb1e7eac667dace125663a9171dc3b07d8398dd3de32d72777ab0065fac34`).  The
live current workspace locked-check log is
`adapter-live-current-check-locked-4100.log` (SHA-256
`75d3102d079b32969ec0e06aa084cc597cef5c3b1101940584f46cd058727d7c`), and
the check-only log is `adapter-live-current-check-only.log` (SHA-256
`37c15e50a03a2646be3ed12f3743fd06e913f5963bc46db43846d27a2058701d`).  Its
direct route receipt is `adapter-live-current-routes.jsonl` (8 lines, SHA-256
`7636523062a48b9754c7e70d84c3ee0cd3a2aaca958ed0f4f02efa977b68203f`), with
empty stderr SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`.
The paired direct route receipts are
`adapter-current-routes.jsonl` (8 lines, SHA-256
`997050b07053de66506a09356e6a6e00e27201debf76d454d67b2317f1dbacfa`) and
`adapter-baseline-routes.jsonl` (8 lines, SHA-256
`54950867a81ba001466f2d9b7515653559b51e410b2cc194767576fec152fc37`); both
stderr streams were empty (the empty-stream SHA-256 is
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`).
Each route receipt contains exactly these eight scenarios:

`open_1k`, `noop_1k`, `text_commit_1k`, `text_commit_save_1k`, `replay_1k`,
`tail_commit_1k`, `unsupported_content`, `protected_commit`.

Correctness checks run outside the timing loops: exact reopen, commit output,
replay output, inverse replay, tail-output reopen, unsupported-source
unchanged bytes, and protected-source unchanged bytes.  The normal timing
loops only execute the operation and black-box its result; they do not assert
refusal kind or reason inside a timed sample.  `--check-only` emits no timing
claim.

## Protocol and gate coverage

`legacy-rtf.json`, `native_qualification_runner.py`, the preparer and validator
now bind the standalone probe path, per-leg feature, per-leg lock hash and
adapter manifest hash; the validator also pins the live current workspace
example separately.  The qualification workload remains five ABBA blocks, 31
rounds, and the same eight route/output/error/fidelity fields.  The validator
still requires positive finite medians and exact fixed refusal fields, while
the source adapter records the native API distinction rather than weakening
that contract.

Build-free protocol tests passed:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_legacy_native_qualification tools.test_native_qualification_runner tools.test_no_3rdparty_references
Ran 42 tests
OK
```

The tests include active probe/lock/adapter pin checks, a stale historical
probe distinction, a mismatched feature negative fixture, real identity
round-trip coverage, manifest/route validation, and the no-`3rdparty` guard.
No glossary or production API row is changed by this consumer-only repair;
the native variant/API-path distinction is recorded here and in the adapter
receipt for the final gate.
