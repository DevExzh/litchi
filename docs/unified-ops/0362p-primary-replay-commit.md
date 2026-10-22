# 0362p — primary replay owners publish `Commit`

This slice is based on replay `86dc7dee0f26a9249c40f165d0c03dcc0446709e`.
It completes the primary owner replay return shape for the PPTX, DOCX and RTF
families.  Raw record/replay primitives remain raw primitives; the XLS
`revision_records::Snapshot::apply` route is one of those record APIs and was
reviewed without changing its `Result<Snapshot>` return.

## API changes

- `litchi-pptx::package::Package::apply` now returns the opened-family
  `opened::Commit`.  The source-backed primary `Package::apply` now returns its
  shared-core `Commit<Package, Patch>`; durable and topology replay branches
  use the same commit constructor.  Existing same-owner checks, source
  authentication, recipient policy, inverse patches and diagnostics remain in
  their original paths.
- `litchi-docx::source_backed::Package::apply` now returns
  `document::Commit<Self>`.  The former public `apply_document` spelling was
  removed; the package replay implementation is private `apply_package`.
  Every example, test and source-backed consumer unwraps the commit with
  `into_package()` or uses its package projection.
- `litchi-rtf::api::Package::apply_tail_append` now returns the family
  `tail_append::Commit`.  The logical patch keeps its package-producing raw
  `Patch::apply` primitive, while the facade constructs the commit triad with
  exact source/output lengths, inserted bytes and no-op diagnostics.  Source
  freshness, policy and inverse checks remain unchanged.
- Every migrated consumer unwraps the published owner explicitly with
  `Commit::into_package()`.  This includes the native-resave and perf-baseline
  tools, their PPTX inverse/durable controls, DOCX publication route, and the
  RTF logical-tail replay/inverse check.  The detector consumers in those tools
  now pass an explicit `litchi::common::SourceLimits` as required by the raw
  detector API; XLSB's facade probe also imports the public `sheet::Cell`
  projection trait.
- Replay content hashing keeps checked length accounting and maps the private
  hash-sink overflow to the typed capacity error.  The RTF range splitter uses
  checked slices, preserving the bounded source scan and avoiding panic-prone
  indexing.

## Validation

All commands used the capped W8 environment, `CARGO_BUILD_JOBS=2`,
`CARGO_BUILD_PIPELINING=false`, a branch-local target, and a coordinator
temporary directory.  The complete logs are outside the repository.

| command | result | receipt (bytes, SHA-256) |
| --- | --- | --- |
| `cargo test -p litchi-pptx --all-features -- --test-threads=2` | passed; all crate tests, 47 doctests, 1 compile-fail doctest (2 doctests ignored) | `pptx-tests.log` (204224, `f09280ae6c43c9d95cbeac8e96bb75ddf71978f622d1a2278b7f1bf293600df5`) |
| `cargo test -p litchi-docx --all-features -- --test-threads=2` | passed; all crate tests, 168 doctests (31 ignored), 2 compile-fail doctests | `docx-tests-v3.log` (233803, `5f7b6edb17667c1db31bbd8fdb163c010a82571002c68d1df1c1b2f32942d29e`) |
| `cargo test -p litchi-rtf --all-features -- --test-threads=2` | passed; all crate tests, 15 doctests, 13 compile-fail doctests | `rtf-tests.log` (153479, `ec0bf57dede2dff5ced90f647a28ad340e59641c8952e5b5fad52d34d6d0b97e`) |
| `cargo clippy -p litchi-pptx -p litchi-docx -p litchi-rtf --all-targets --all-features -- -D warnings` | passed | `clippy-v3.log` (401, `509473bafd180836c71ab8f2c71ec767bd8a821b78718319bd2464e701922631`) |
| `cargo test -p litchi-pptx --all-features replay_hash -- --test-threads=2` | passed | `pptx-replay-hash.log` (23496, `87e3e03008174e363b016084bb747d5a3507245ff2edfede6343f2f84b7c4486`) |
| `cargo test -p litchi-rtf --all-features --test logical_tail_append -- --test-threads=2` | 23 passed | `rtf-logical-tail.log` (2047, `af28d720cdecefe68c4221f7f61454d199b8406c2beec13ca3a29897d858d75`) |
| `cargo check --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets --all-features` | passed; one existing `save_workbook` dead-code warning | `perf-baseline-check-locked.log` (533 bytes, recorded outside the repository, SHA-256 `f07ec4a28cb824982567207cc678516fe9fe6d4569470bdf56eb440a9bcf1365`) |
| `cargo check --locked --manifest-path tools/native-resave/Cargo.toml --all-targets --all-features` | passed | `native-resave-check-locked.log` (60 bytes, recorded outside the repository, SHA-256 `f694a9e074374d315f58b8a63ce849858b1b8708cf56b3d194218605bd2a9d6e`) |

The earlier package-level check receipts remain available for source
attribution: PPTX `w8-ppt-replay-primary-check3.log` (175 bytes,
`1350656e8213e29228f5a8394651375e7e427aa23f34ddae4c9e77c6cf3a8adb`), DOCX
`w8-docx-replay-primary-check9.log` (174 bytes,
`58d51fa6f9032fd6d657568b8399e84031f9d30a02b54ebb7cf3eeb430ce6881`), and
RTF `w8-rtf-replay-primary-check2.log` (172 bytes,
`84c3534130d5c9bc8d8590a7d95eb580937d02e0b581d40830d39836d66fda24`).

The changed owner calls were migrated in the same source commit.  No
benchmark binaries or large evidence files are part of this record.

The return-shape gate registration for this slice is:

| method | `Result` payload | owner type arguments |
| --- | --- | --- |
| `litchi_pptx::Package::apply` | `litchi_pptx::opened::Commit` | — |
| `litchi_pptx::source::Package::apply` | `litchi_pptx::source::Commit` | — |
| `litchi_docx::source_backed::Package::apply` | `litchi_docx::document::Commit` | `litchi_docx::source_backed::Package` |
| `litchi_rtf::Package::apply_tail_append` | `litchi_rtf::tail_append::Commit` | — |

These exact method/type paths are the rows to add to each format profile's
`result_methods` list when the coordinator applies the return-shape checker;
the DOCX row replaces the removed `source_backed::Package::apply_document`
spelling.
