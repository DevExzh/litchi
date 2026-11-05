# 0250 — XLSX W2 owner reshape (ADR 0034 §2–§4)

Status: landed on `uo2/p4-xlsx-w2` (code commit `a3418ede8b`). This record is
step 4 of the coordinator plan (record 0248 §6, Phase B).

## What changed

`litchi_xlsx` had two owners: a mutable physical `Package` (OPC graph plus
encryption provenance, scoped `&mut` editors, `create`, `from_slice`,
`from_bytes(Vec)`, `into_workbook`) and a semantic `Workbook` that held the
state verbs (`new`/`create`, every open route, `edit`, `apply`, `unsigned`,
the outputs and the security verbs). Opening through `Package` validated the
workbook twice: once in `from_opc` and again in `into_workbook`.

Now there is one owner:

- `xlsx::Package` is the immutable, cheap-to-clone snapshot and owns every
  state verb: `new`, `open`, `from_bytes(&[u8])`, `from_reader`,
  `open_with_options` / `from_bytes_with_options` / `from_reader_with_options`
  with a new `OpenOptions` envelope (source limits, leniency, encryption
  limits), the password routes (`*_with_password[_and_options]`,
  `from_decrypted`), `from_opc`, `edit`, `apply`, `unsign`, `same_content`,
  `save`/`save_with_durability`/`save_plain*`, `to_bytes`/`to_plain_bytes`,
  `write_to`/`write_plain_to`, and the encryption verbs. Each open validates
  the workbook once.
- `xlsx::Workbook` is the owned, read-only projection:
  `Package::workbook()` (one `Arc` increment) and the borrowed
  `Package::workbook_ref()`. It keeps the catalog, sheet, cell, style, pivot
  and drawing reads and no state verb.
- `workbook::edit` is a complete family, registered in the glossary gate:
  `Edit`, `Patch`, `Commit`, `Diagnostics`, `Error`, `LimitKind`.
  `Edit::commit() -> Result<Commit>`; `Commit::{package, patch, diagnostics,
  into_package, into_parts}`. `Diagnostics::{changed, changes,
  package_changes, changed_parts}` is content-free. `Patch::is_noop()` now
  returns `bool` and costs time proportional to the changed closure (an
  untouched commit, a structurally empty patch, or a part-only patch whose
  parts equal their sources); whole-snapshot comparison stays available as
  `Package::same_content`. `Patch::len` became `operation_count`.
  `LimitKind::{TransferCells, TransferDependencyScan, SvgInputBytes}` is
  reported as `Error::Limit { kind, observed, maximum }`; those three bounds
  were string `Invalid` errors before. `edit::Error` re-exports the crate
  error; the per-module split (ADR 0034 §6) is not part of this step.
- The scoped `&mut` families (Data Model, Custom Data, surveys, calculation
  metadata, page breaks, custom properties, slicers, timelines, task panes,
  smart tags) now publish through a crate-private `package::Target`: on a
  `Package` the staged OPC graph is revalidated as a complete workbook before
  the owner adopts it, so a failed publication leaves the package unchanged.
  Their public `Transaction::new(&mut OpcPackage)` entries are unchanged.
- Pivot: `Workbook{Transaction,Patch,Commit}` are now
  `PackageEdit`/`PackagePatch`/`PackageCommit` (root aliases
  `*PackageCommit`/`*PackagePatch`). Commits return `Package`, their patches
  apply to `&Package`, and `Package::{edit_pivot_table, apply_pivot_table,
  edit_pivot_table_data[_with_limits], apply_pivot_table_data,
  edit_pivot_cache, apply_pivot_cache}` replace the `Workbook` and
  `*_patch` spellings.
- Durable patches and `History` apply to and retain `Package`.

Removed with no alias: `Package::create`, `Workbook::create` (and the other
`Workbook` constructors), `from_slice*`, `from_bytes(Vec<u8>)`,
`*_with_limits` and `*_with_leniency` ingress, `from_package*`,
`Package::into_workbook`, `Commit::{workbook, into_workbook}`,
`Patch::{is_empty, len}`, `Workbook::unsigned` (use `Package::unsign`). The
plaintext-build `From<Package> for OpcPackage` shim is kept.

Gate: `tools/api_glossary.json` registers 81 new bans in `xlsx-minimal` and
`xlsx-full` and the `litchi_xlsx::workbook::edit` family. `API_GLOSSARY.md`
rows for the owner, projection, edit entry, snapshot accessor, diagnostics,
split commit, application, no-op introspection, `from_bytes(&[u8])`,
`Package::new()` and the configured open flip for XLSX.

## Contracts preserved

- Exact no-op: a commit without changes returns the source snapshot
  (`diagnostics().changed() == false`, `patch().is_noop()`, identical bytes);
  new test `package_owner::an_exact_no_op_commit_is_the_source_byte_for_byte`.
- Signed-write refusal, encryption policy and snapshot immutability: the
  existing suites (`package_signature_policy`, `workbook_unsigned`,
  `encryption`, `calculation_metadata_formula_security`, source-backed
  suites) pass unchanged in meaning; new tests cover source immutability
  after a commit and scoped-family republication.
- Resource budgets: `from_bytes_with_options` checks the input-size limit
  before copying the borrowed bytes (new test); limits otherwise flow as
  before.
- Owner split in ops: `XlsxOwner::{Source(SourceBackedWorkbook),
  Eager(xlsx::Package)}`; source promotion and the eager open now validate
  once instead of twice.
- `from_bytes(&[u8])` copies the input once into the owned source archive,
  so exact-source publication and compressed passthrough are retained. The
  zero-copy owned route remains `Package::from_opc(OpcPackage::from_vec(..))`.

## Performance

Release builds of `tools/perf-baseline` at base `908f9cccdd` and at
`a3418ede8b`. ABBA ×2 (8 processes), 40 samples and 5 warmups per case,
pinned to CPU 27 of a shared 32-core host. Value: median over processes of
each process's p50. Evidence:
`docs/unified-ops/results/0250-xlsx-w2-owner/abba.json` (8 KiB). Binaries are
not committed; SHA-256 base `58ff3427…acc`, candidate `5e8c99f8…570`; build
command `cargo build --release --manifest-path tools/perf-baseline/Cargo.toml
--bin litchi-perf-baseline`.

| Case | Corpus | Base p50 | New p50 | Δ |
|---|---|---:|---:|---:|
| open owned | tiny / medium / dense-wide | 51.4 µs / 102.4 µs / 1.348 ms | 51.7 µs / 102.9 µs / 1.346 ms | +0.6% / +0.5% / −0.2% |
| one-cell commit | tiny / medium / dense-wide | 65.9 µs / 426 µs / 15.09 ms | 66.1 µs / 429 µs / 15.25 ms | +0.3% / +0.6% / +1.1% |
| one-percent commit | tiny / medium / dense-wide | 109.1 µs / 1.670 ms / 33.23 ms | 109.5 µs / 1.675 ms / 33.74 ms | +0.3% / +0.3% / +1.5% |
| no-op commit | tiny / medium / dense-wide | 235 ns / 300 ns / 1.40 µs | 215 ns / 290 ns / 1.29 µs | −8.5% / −3.3% / −8.0% |
| one-cell commit + save | tiny / medium / dense-wide | 102.1 µs / 852 µs / 60.50 ms | 102.6 µs / 856 µs / 60.52 ms | +0.5% / +0.5% / +0.0% |
| no-op commit + save | tiny / medium / dense-wide | 315 ns / 535 ns / 16.9 µs | 305 ns / 540 ns / 17.2 µs | −3.2% / +0.9% / +2.0% |
| ordinary save edit | cell-values medium | 2.376 ms | 2.379 ms | +0.1% |
| bytes leaf open | cell-values medium | 59.2 µs | 61.7 µs | +4.3% |

No case regresses by more than 5%. `xlsx_bytes_leaf_open` times only
`SourceBackedWorkbook` code that this change does not touch; its process
p50s spread 58.9–68.7 µs (base) and 58.7–74.0 µs (candidate), and a second
ABBA ×3 at 100 samples gave base 60.4–68.2 µs and candidate 59.8–72.4 µs.
The shift is within that host noise. `xlsx_open_owned` keeps its owned-move
route through `from_opc(OpcPackage::from_vec(..))`; the canonical
`from_bytes(&[u8])` adds one input copy that this table does not include.
Commits compute no diagnostics unless asked; `Diagnostics::changed` reuses
`Patch::is_noop`.

## Verification (all exit 0)

- `cargo fmt --all --check`
- `cargo clippy -p litchi-xlsx -p litchi-ops -p litchi -p litchi-py
  --all-targets --all-features -- -D warnings`; also `-p litchi-xlsx
  --no-default-features`
- `cargo test -p litchi-xlsx --all-features` (2,421 tests and doctests),
  `cargo test -p litchi-ops --all-features` (1,764),
  `cargo test -p litchi --all-features` (260, 1 ignored). The three
  form-control test targets need the sparse-excluded
  `docs/report/spec-gap-validation-evidence/.../native-corpus.json`; it was
  restored from `HEAD` for the run and the sparse state was reapplied.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xlsx
  -p litchi-ops -p litchi`
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile
  xlsx-minimal` and `--profile xlsx-full`: 89 banned names and 1 family each
- `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py
  verify`
- `crates/pyo3-litchi/tests/run_tests.sh`: 132 tests
- `cargo check` for `tools/perf-baseline` (all targets and features),
  `tools/perf-baseline/unified-xlsx-batch` and `tools/native-resave`
- Capability goldens unchanged (no re-bless).

`cargo check --workspace --all-targets --all-features` fails only on test
targets of untouched crates (soapberry-zip, litchi-opc, litchi-ods) that
`include_bytes!` evidence excluded by this worktree's sparse checkout.

## Remaining work

- The scoped package families still mutate through `&mut self` and keep
  `Transaction`/`*_patch` vocabulary (20 scoped types); their move to
  `edit_x()`/`apply_x()` with `{Edit, Patch, Commit, Diagnostics, Error,
  LimitKind}` is open, as are pivot `Diagnostics` (the pivot commits still
  expose `changed()`).
- `edit::Error` is the crate error; splitting the god-enum per module and
  adding `LimitKind` for the staged-style memory budget (still
  `ResourceLimit(Memory)`) remain.
- `SourceBackedWorkbook` keeps `open_with_limits`/`from_path*` ingress
  names, and the scoped `*_with_limits` read/edit entries remain.
- `from_shared_bytes(Arc<[u8]>)` is not offered: the OPC reader shares only
  `Arc<Vec<u8>>`, so it would not be zero-copy.
- Metadata (`Package::metadata()`/`edit_metadata`) is plan step 6.
- `tools/perf-baseline/src/ordinary_save.rs` keeps the frozen labels
  `litchi_xlsx::Workbook::{save, write_to}` in its route records.
