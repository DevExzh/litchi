# 0048 — bounded path resolution core

## Change and provenance

Added the shared M1 path-resolution core and public IntoPath/PathArg inputs.
Text, borrowed parsed paths and owned parsed paths share one resolver. Excel
notation maps to worksheet/cell/row/column steps; relative paths start from a
node, while document-root resolution requires an absolute path. Static family
vocabularies classify aliases, positions, semantic names and raw identities.
The internal Navigate seam uses static dispatch and returns typed lookup states;
get reports a miss and find returns None without constructing a suggestion.

The source-only starting candidate was
85300f74e24cbc74c156727933ecae3b17b1a830. No candidate history, benchmark evidence
or performance claim was imported. The rewritten base remains an ancestor.

Source::Selector preserves the grammar's ParseError, code, kind and original
byte offset. ParseError::offset_by supports suffix parsing without losing the
original input location. Selector Debug and encountered-character diagnostics
now omit authored content. Explicit grammar Display and accessors remain the
content-bearing rendering/data APIs. Resolver diagnostics contain static
vocabulary, offsets and structural counts, not path text or document names.
Known-vocabulary hints run only on errors and cap the inspected input prefix.

Ordinary resolution accepts checked positions and semantic names. After
validating an identity key against its element rule, native keys require
path::raw::path. Unknown keys still report invalid identity rather than a raw
permission error. Native keys include paraId/textId, format object IDs, note and
comment IDs, placeholder idx and numbering ilvl. Semantic name/type/role keys
retain their declared comparison policy. Native NodePath syntax must explicitly
pass through the raw wrapper; production ordinary adapters must emit semantic
or positional paths. The core grammar itself remains a syntax representation.

The original text limit is checked before Excel normalization. Hand-built paths
are checked before root I/O for step, predicate and text limits; unsupported
compound content filters are refused without traversing arbitrary expression
trees. A bounded formatter checks canonical path size without allocating.
Excel sheet notation refuses control characters. Lowercase whole-column input
stays borrowed; its spelling is retained until the adapter supplies a canonical
node path. Caseless names use NFD/full Unicode folding/NFD, matching the adopted
canonical policy rather than lowercase conversion.

## Verification

Tests cover Word, workbook and presentation vocabularies, alias and predicate
rules, named children, spreadsheet coordinates, relative navigation, typed
errors/offsets, misses, suggestions, raw identity admission, bounded typed paths,
redacted source chains and Unicode equivalence. Test-only fake adapters exercise
the navigation contract; production adapters are not implemented by this record.

- cargo test -p litchi-core -p litchi-ops --all-features: 791 unit/integration
  tests and 176 doctests passed, including allocation and compile-fail tests.
- cargo test -p litchi-ops --no-default-features --lib: 178 passed.
- cargo check --workspace --all-features --lib: passed, including Python.
- cargo clippy -p litchi-core -p litchi-ops --all-features --lib -- -D warnings:
  passed.
- CARGO_TARGET_DIR=target cargo run --manifest-path
  crates/litchi-ops/fuzz/Cargo.toml --bin smoke_paths: 909 deterministic inputs,
  435 accepted canonical round trips. This is bounded smoke, not a campaign.

- CARGO_TARGET_DIR=target cargo check --manifest-path
  crates/litchi-ops/fuzz/Cargo.toml --bins: passed.
- RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-core -p litchi-ops
  --all-features --no-deps: passed.
- python3 tools/check_api_glossary.py: passed, 11 profiles, 144 banned names,
  zero complete transaction families.
- python3 tools/check_crate_boundaries.py: passed, 66 packages, 267 internal
  dependencies, 11 existing explicit debt items.

Read-only review identified the bounds, identity, Unicode and diagnostic
adaptations above; regression tests exercise those findings. Changed Rust files
are formatted;
whitespace, staged evidence size and binary restrictions are checked. Borrowed
fixtures, generated root/fuzz lockfiles, target output and session scratch files
are removed. No temporary worktree was created; base, iWork and pre-existing
worktrees remain untouched.

## Remaining scope and limitations

This is an M1 prerequisite, not milestone closure. Concrete format navigation,
File get/find/query, capabilities, registry integration, Python/CLI operations,
all-format read/write and paired performance measurements remain required.

Escaped apostrophes in Excel sheet notation still require an owned normalized
name. Canonical Unicode normalization can buffer long combining sequences.
Allocation checks cover their listed common inputs and do not prove all input
forms allocation-free. Native format name constraints (such as XLSX worksheet
length and forbidden characters) remain the responsibility of future adapters;
this syntax layer checks input bounds and control safety without declaring
arbitrary names valid worksheets. M1 performance acceptance remains unclaimed.

OfficeCLI deltas include explicit raw identity addressing, content-free error
messages, canonical Unicode name equality, refusal of compound content filters
in paths, strict absolute document paths, and retention of lowercase whole-column
spelling. Existing query/regex and record 0047 coercion limitations remain open.
