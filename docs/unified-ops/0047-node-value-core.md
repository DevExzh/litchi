# 0047 — node projection and shared value coercion cores

## Change and provenance

The operations crate now has a format-independent node/view core, typed paths,
static property specifications and shared scalar coercions. Views stream the
OfficeCLI eight-member JSON shape through serde, including canonical property
rendering, without constructing an intermediate document tree. Length and Color
have one implementation shared by node values and future registry/query users.

Source-only starting candidates were dca30c99ec89255a3f593af5f6a2269ed5582e3d,
6a2057ca0baaa28ececc10206a2aaf3c0ffadbd5 and
3b426848f543ec8ab4891ec457774a705a47460e. No candidate history or historical
measurement evidence was merged. The rewritten base remains an ancestor.

The integration adapts those candidates to the adopted glossary and policies:

- Native identity path construction is explicit under node::raw. Ordinary
  numeric navigation uses Position. Test-support fake nodes exercise raw
  identities; production registries must not expose native identity properties.
- Debug output for nodes, paths, values, previews, properties and predicates
  omits authored content. Value/serializer diagnostics do not echo rejected
  text or external serializer messages, including diagnostic source chains.
  Serializer causes remain available by explicit typed variant matching.
  Known-vocabulary hints are computed
  on demand. Intentional content rendering remains explicit.
- View depth above 64 returns MaxDepthExceeded rather than being clamped. Node count remains
  bounded, and non-finite numbers are refused rather than serialized as null.
  serialize_into preserves typed view errors separately from serializer errors.
- WordSpacing reads bare unsigned twips, converts pt/cm/in and rounds ties to
  even at the twip boundary, checking the u32 range before storing EMU.

Release format variants still carry uninhabited Pending payloads. Test-support
fake nodes provide executable projection fixtures, not usable format adapters.
No File get/find/query/set API, property registry, Python surface or capability
advertisement is added by this record.

## Verification

Golden fixtures cover Word, spreadsheet and presentation view shapes. Tests
cover path escaping, omission/order, depth/node limits, scalar JSON types,
canonical coercions, comparison semantics, hostile input and content redaction.
Allocation counters cover listed success/render/comparison cases and borrowed
error construction. They are not paired format performance measurements.

- cargo test -p litchi-ops --all-features: 502 unit tests, nine allocation
  integration tests and 113 doctests passed.
- cargo test -p litchi-ops --no-default-features --lib: 142 passed.
- cargo clippy -p litchi-ops --all-features --lib -- -D warnings: passed.
- cargo check --workspace --all-features --lib: passed, including Python.
- CARGO_TARGET_DIR=target cargo check --manifest-path
  crates/litchi-ops/fuzz/Cargo.toml --bins: passed.
- CARGO_TARGET_DIR=target cargo run --manifest-path
  crates/litchi-ops/fuzz/Cargo.toml --bin smoke_values: 776 bounded inputs,
  627 successful coercions, no invariant failure. Accepted colors and lengths
  undergo canonical round-trip checks. This is not a sustained fuzz campaign.
- python3 tools/check_crate_boundaries.py: passed, 66 packages, 267 internal
  dependencies and 11 existing explicit debt items.
- python3 tools/check_api_glossary.py: passed, 11 profiles, 144 banned names,
  zero complete transaction families.

Read-only review found no further concrete blocker after the identity, depth
classification and external-error source-chain fixes.

Changed Rust formatting and git diff --check passed. Repository evidence and
binary restrictions are checked before commit. All generated evidence stays
outside the checkout; no benchmark binary is committed.

## Remaining work and explicit limitations

The shared core is a prerequisite for M1/M2, not completion of either milestone.
Concrete adapters, registry dispatch, query compilation, public unified edits,
all-format parity, managed semantic budgets and paired performance acceptance
remain required by the goal.

Word spacing currently rejects comma-grouped input that OfficeCLI's invariant
Double parser accepts. Long grouped numeric-text input (over 128 bytes) uses an
allocating normalization fallback; allocation evidence excludes that path.
Neither limitation is claimed resolved. Regex comparison still follows the
existing unsupported-engine policy. EMU values above 2^53 retain the candidate's
documented floating-point precision limitation. Full OfficeCLI differential
qualification and production property-table closure remain follow-up work.

Borrowed fixtures, generated root/fuzz lockfiles, build output and temporary
session files are removed before commit. No temporary worktree was created;
base, iWork and pre-existing worktrees remain untouched.
