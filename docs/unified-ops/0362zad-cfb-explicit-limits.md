# 0362zad — canonical CFB limits and consumer migration

- Date: 2026-10-24.
- Base: `e35cede32ae730d351ce9d29d243c35edd321ddf`.
- Scope: the seven public `litchi-cfb` limits rows, the source-layout adoption
  ingress found during review, and every tracked Rust consumer in the
  workspace, including format crates, shared OLE consumers, tests, examples,
  fuzz targets and performance tools.

## Canonical public shapes

The default twins and public `_with_limits` spellings are removed. The
separately audited public `OleWriter::update_stream` duplicate alias is also
removed; its sole test consumer now calls `create_stream`. The canonical raw
entry points now require their typed finite bounds directly:

| Surface | Canonical shape |
|---|---|
| Cursor parser | `OleFile::open(reader, OleFileLimits)` |
| Positional shared parser | `SharedOleFile::open(source, SharedOleFileLimits)` |
| Sealed byte parsers | `SharedOleFile::open_owned(source, version, SharedOleFileLimits)` and `SharedOleFile::open_owned_vec(source, version, SharedOleFileLimits)` |
| Non-mutating validation | `validate_source(source, SharedOleFileLimits, ValidationLimits)` |
| Storage move | `OleWriter::move_storage(source, destination, StorageMoveLimits)` |
| Source-layout adoption | `OleWriter::adopt_source_layout(source, OleFileLimits)` |
| Sequential writer options | consuming `SequentialWriterOptions::limits(SequentialWriterLimits)` and borrowed `limits_value()` |

Every former default consumer passes the corresponding prior default
explicitly. Existing custom bounds are passed unchanged, so validation order,
allocation ceilings, source-fidelity checks, and storage-move behavior remain
the same. Source-layout adoption retains the caller-selected `OleFileLimits`
through `SourceLayout::parse`; object and presentation consumers pass the
profile they retained for the source, while record-only paths pass the prior
explicit default. Private shared-parser helpers retain implementation-oriented
names; none is part of the public API.

The migration includes the CFB crate's unit/integration tests, examples and
fuzz target, all workspace format/shared consumers, signature call sites,
`litchi-ops`, and performance tools. The source scan found no public CFB
`*_with_limits` or `with_limits` declaration, no public `update_stream` alias,
and no remaining limit-less call to a canonical CFB entry point. The focused regression test
`options_limits_round_trip_through_noun_accessors` verifies that the
consuming setter and `<noun>_value()` accessor retain all seven sequential
writer bounds.

## Retained-policy review return

The review found four correctness paths that could silently restore a parser
default after a caller selected a stricter or larger source policy. They are
closed in the same source change:

- `litchi-ops` eager CFB `open_bytes`/`unlock` derive the low-level profile
  from the admitted `OpenOptions` input ceiling; `eager_cfb_profile_follows_explicit_v4_input_ceiling`
  guards the larger explicit input policy.
- `OleWriter::adopt_source_layout` requires `OleFileLimits`; object and PPT
  source-backed consumers retain and pass their source profile, while
  record-only paths pass an explicit prior default.
- Sequential candidate validation stores a profile derived from the planned
  physical output and padded directory bytes; `candidate_validation_profile_tracks_v4_output_budget`
  guards the V4 output budget without imposing the 2 GiB reader default.
- Overlay composition and stream-move inverse replay carry the retained
  `SharedOleFileLimits`; `composed_overlay_reuses_tight_directory_and_allocation_limits`
  proves both tight physical dimensions reject a default-admitted expansion,
  and `inverse_replay_retains_the_custom_physical_profile` guards replay.

These review returns preserve the existing validation order, source-fidelity
fences and allocation policies; they do not add a compatibility alias or a
limit-less escape hatch.

## Glossary and gate coverage

`tools/api_glossary.json` adds the `cfb` profile with suffix and exact
`with_limits` bans. `docs/API_GLOSSARY.md` registers the raw CFB ingress,
source-layout adoption, storage-move and sequential-options shapes and links
this record. The record
0362rd substrate note is narrowed: CFB is closed here, while standalone
soapberry ZIP transport remains separate follow-up scope.

## Validation status

The final focused and workspace checks ran against the uncommitted source
state represented by external patch receipt
`source-state-before-final-record.patch` (547,393 bytes,
SHA-256 `ba0eb0fabfa272fe4eabf16452da9f3e8e203e66bfbec29e1314c3373c91a46f`).
The record-only documentation update below that receipt does not change Rust
source. The source was at base `e35cede32ae730d351ce9d29d243c35edd321ddf`.

The lock origin was the pinned copy
`/home/zhuhe/code/litchi-worktrees/uo2-budget-validation/Cargo.lock`, 139,591
bytes, SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`; the
worktree lock matched byte-for-byte. Commands used the shim PATH, unset
`LITCHI_CAPPED`/`LITCHI_BLESS`, `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`,
`CARGO_NET_OFFLINE=true`, and external target
`(deleted-coord-artifact)` (actual
size 2,355,002,392 bytes, below the 20 GB cap).

The focused CFB command
`cargo test --offline --locked --all-features --all-targets -p litchi-cfb`
passed 495 unit tests and all integration/example targets. Its immutable log
is `cargo-test-cfb-final-rerun6.log` (48,820 bytes, SHA-256
`b4be69509d5f3e2479462e0adcdcf8329db9886e452190d734961b80b0435629`). The
workspace command
`cargo check --offline --locked --all-features --all-targets` also passed; its
log `cargo-check-workspace-review-repairs-rerun5.log` is 15,059 bytes, SHA-256
`b516ba655936afe15d638ac264f3cfa2d9ce16c606ff05d45b272512fba59a37`.
The two focused retained-policy logs are `cargo-test-composed-tight-profile-rerun4.log`
(528 bytes, SHA-256
`df5f0adec88f2412bb64ff33cacd1625825cd7ea0d211c07c2cc79c2c72e38b8`) and
`cargo-test-inverse-profile.log` (412 bytes, SHA-256
`74fb37376d9886952e3459f0b21d02b143c6f02bb8ac00670c9d80b7a1c0c5`). The
glossary unit gate passed 50 tests (`python-test-check-api-glossary.log`, 149
bytes, SHA-256
`fec80aba272b318a6f01a7a930014acc687e158732c97ec0a275482fc3a17d17`), JSON,
diff, public-suffix and arity scans passed, and changed-file rustfmt passed
with an empty log (`rustfmt-changed-files-check.log`, 0 bytes, SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`).
The repository-wide `cargo fmt --all -- --check` was run but remains red only
on two unrelated pre-existing iWork test formatting diffs; no CFB-lane file is
listed in that log (`cargo-fmt-check.log`, 1,749 bytes, SHA-256
`17041870235b8f1f5cfa9a49098f5009d5e0a73f104a38fe06976305ff1bbd59`).
The guarded commit still requires the repository-size and binary checks from
`AGENTS.md`.
