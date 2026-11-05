# 0213 — Canonical PPT Builder and creation/publication

Fresh legacy PowerPoint authoring uses `Builder`. `build()` returns a validated
immutable `Package`; memory, sequential and atomic filesystem output belong to
that package. `Package::new()` creates the minimal valid deck. The old public
`Writer` name and Builder publication methods are removed, while the native
authoring operations remain available. Explicit raw encoding supports codec
fixtures and frozen measurement routes without adding a package reopen to
those measurements. The duplicate filesystem encoder is removed. Its speaker
notes planning, slide links and NotesContainer emission are preserved in the
shared encoder and checked by readback.

Password-to-open authoring requires a move-only, zeroizing
`EncryptionEntropy` token. `generate()` explicitly requests OS randomness;
`from_bytes` supports caller-managed deterministic inputs. Encrypted attempts
consume the token before validation, including failed admission; plain builds
leave pending entropy alone. `Error::Authoring` retains the typed native cause
and redacts its details from outer Debug/Display.

The unified facade gains PPT creation and publication through the same native
owner. Creation adopts the package directly. Unsupported template, OOXML Strict
and initial-metadata options are explicit capability gaps. Publication retains
source checks, signatures unless removal is explicitly requested, atomic
replacement, overwrite/durability options and origin handling. Ordinary
publication of encrypted packages is refused; raw ciphertext export remains an
explicit native operation.

## Verification

The all-feature native suite passes 1,391 tests/doctests (11 existing ignored
cases), and the facade suite passes 1,558 tests/doctests. Six canonical Builder
regressions cover minimal deterministic creation, detached repeat builds,
canonical edit/patch interoperability, retained admission limits, explicit
entropy consumption and typed redacted authoring errors. Seven facade cases
cover creation/options, exact publication, save origins, overwrite refusal,
source changes, signatures, released edit sources and encryption refusals,
including signature removal on an encrypted source.

Reproduce with `cargo test -p litchi-ppt --all-features` and
`cargo test -p litchi-ops --all-features`. Checks use an external offline lock
selecting cached cc 1.6.0; both repository lockfiles remain unchanged.

Warning-denied all-target Clippy and rustdoc pass. Four native minimal-feature
Builder tests and five PPT-only facade creation/publication tests pass. The
PPT glossary profile verifies 35 banned names and one canonical transaction
family. Crate boundaries pass for 66 packages and 268 internal dependencies
with 11 existing debt items. All 15 OfficeCLI parity generator tests pass and
the generated table remains current against the base checkout's source.

The standalone performance harness compiles and both PPT Pictures eager/source
open smoke routes pass after their migration to `raw::encode`. Run
`cargo run --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline -- --case ppt_pictures_eager_open,ppt_pictures_source_backed_open --samples 1 --warmup 0`
to reproduce. This checks harness compatibility only; unoptimized one-sample
runs are not performance evidence. Temporary outputs, external locks, build
artifacts and the three milestone worktrees are removed after integration.

## Remaining scope

Unified PPT navigation/mutation dispatch, complete signed no-op transaction
support, performance qualification, and the all-format/Python/CLI goal remain
open. This creation/publication milestone does not claim timing acceptance.
