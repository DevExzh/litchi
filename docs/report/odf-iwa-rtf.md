# Workspace-Split Audit: ODF / iWork / RTF

> Historical audit record. This report preserves the branch names, crate
> paths, source paths, line references, and conclusions observed during the
> original `refactor/workspace-split` review. It does not describe current
> ownership after the ODF family split or later IWA migration. In particular,
> every path below is evidence from that audit unless explicitly labelled
> current.

## 1. Scope

Compared the `refactor/workspace-split` branch against `main` for three sibling
crates carved out of the umbrella `src/`:

| Historical split target     | Source on `main` |
| --------------------------- | ---------------- |
| `crates/litchi-odf/src/`    | `src/odf/`       |
| `crates/litchi-iwa/src/`    | `src/iwa/`       |
| `crates/litchi-rtf/src/`    | `src/rtf/`       |

Plus each crate's `Cargo.toml`, `examples/`, and (for `iwa`) `build.rs` +
`src/protos/`.

## 2. Method

- `git log --diff-filter=D --name-only main..HEAD -- src/{odf,iwa,rtf}/` to
  enumerate the moved files (45 odf, 33 iwa code + 26 protos, 19 rtf).
- For every file pair, computed sha256 of `git show main:<old>` vs the new
  on-disk file. SHA-equal files were checked off; SHA-different files were
  diffed.
- `git diff main -- <old> <new>` filtered with grep to strip the allowed noise
  (path moves, `crate::common`→`litchi_core`, `crate::iwa::`→`crate::`,
  doc-path rewrites, `pub use` re-exports, comments) and inspected the
  residual.
- Verified module-surface parity: `pub mod` / `pub use` listings between old
  `mod.rs` and new `lib.rs`.
- Verified `iwa/build.rs` migration and proto byte-parity (all 26 `.proto`
  files SHA-identical).
- Counted `TODO|FIXME|unimplemented!|panic!` occurrences old vs new per crate.

## 3. Findings

### litchi-iwa  ✅ Pure mechanical

- All 26 `.proto` files under `src/protos/` are byte-identical to
  `main:src/iwa/protos/*` (verified by sha256). prost-build will emit the same
  generated types.
- `build.rs` migrated from root `build.rs` is unchanged in logic — same
  prost-build config, same `enable_type_names()`, same
  `include_file("iwa_protos.rs")`, same proto root walk, same panic-on-error
  path. Only differences: the `#[cfg(feature = "iwa")]` outer gate is dropped
  (the crate itself is gated upstream) and the path is rewritten from
  `src/iwa/protos` to `src/protos`.
- `varint.rs` and `registry.rs` are sha256-identical.
- `snappy.rs` (IWA streaming Snappy decompressor): only diff is
  `crate::iwa::Error` → `crate::Error`. Algorithm unchanged.
- `archive.rs`, `object_index.rs`, `ref_graph.rs`, `structured.rs`,
  `zip_utils.rs`, `protobuf.rs`, `document.rs`, `numbers/{cell,table}.rs`:
  the only diffs are import path rewrites
  (`crate::iwa::protobuf::…` → `crate::protobuf::…`) plus minor cosmetic
  reflow where the shorter path lets a `if let Ok(x) = …::decode(…) {` line
  fit on one line. No struct fields, decode tags, error variants, or branch
  conditions are altered.
- Module surface (`pub mod` / `pub use` in `lib.rs`) is byte-identical to
  the old `mod.rs`.
- `TODO/FIXME/unimplemented!/panic!` count: 1 → 1 (the existing snappy
  comment).
- `examples/{read_iwork,extract_structured}.rs` are new demos (≤80 LOC each)
  with no parser logic; no comparable file existed on `main`.

Representative reviewed files: `crates/litchi-iwa/src/snappy.rs`,
`crates/litchi-iwa/src/archive.rs`, `crates/litchi-iwa/src/object_index.rs`,
`crates/litchi-iwa/src/zip_utils.rs`, `crates/litchi-iwa/build.rs`,
`crates/litchi-iwa/src/protos/TSPMessages.proto`.

### litchi-odf  ✅ Pure mechanical (clippy-style refactors only)

- `core/manifest.rs`, `core/package.rs`, `elements/style.rs`: only
  `crate::common::*` → `litchi_core::*` and `crate::odf::*` → `crate::*`.
- `core/xml.rs`, `ods/parser.rs`, `odp/parser.rs`: collapse
  `match arm => { if cond { … } }` to `match arm if cond => { … }`. Outer
  match in every case has a `_ => {}` catch-all (verified at
  `crates/litchi-odf/src/ods/parser.rs:97` and equivalent in odp parser),
  so guard-fail still falls through to the no-op arm — semantically
  identical. No `else` branches existed in the original; nothing dropped.
- `datatype.rs`: chain reformat (`map_err` closure spans more lines after
  the import shortened) plus two test-only `assert_eq!(x, true)` →
  `assert!(x)` fixups. Production logic byte-identical sans imports.
- `odt/parser.rs`, `odp/parser.rs`: removed redundant `.clone()` in tests on
  `Copy` types (`let t2 = t1;`).
- `elements/table.rs`: 144+/144– diff is a *block relocation* — the
  `pub struct TableElements` impl was hoisted from after the
  `#[cfg(test)] mod tests` block to before it. Token-equal, function bodies
  unchanged. Confirmed by counting added/removed lines == 144 each side and
  by reading both ends of the diff.
- Manifest parsing, content/styles dispatch, table-cell coercion, XML
  reader event handling: all unchanged.
- Module surface in `lib.rs` matches old `mod.rs` exactly except
  `crate::common::RGBColor as Color` → `litchi_core::RGBColor as Color`
  (allowed import-path update).
- `TODO/FIXME/unimplemented!/panic!` count: 44 → 44 (preserved).
- New `examples/{read_odt,read_odp,write_ods}.rs` are demos (≤71 LOC each);
  no comparable files on `main`.

Representative reviewed files: `crates/litchi-odf/src/core/manifest.rs`,
`crates/litchi-odf/src/core/package.rs`, `crates/litchi-odf/src/ods/parser.rs`,
`crates/litchi-odf/src/odp/parser.rs`,
`crates/litchi-odf/src/elements/table.rs`,
`crates/litchi-odf/src/datatype.rs`.

### litchi-rtf  ⚠ 1 minor concern (otherwise mechanical)

- `lexer.rs`, `writer.rs`, `picture.rs` are sha256-identical to `main`.
  Control-word dispatch and binary-blob (`\bin`) extraction therefore
  unchanged.
- `parser.rs`, `document.rs`: only import-path / doc-path rewrites
  (`crate::common::*` → `litchi_core::*`, `litchi::rtf::…` → `litchi_rtf::…`).
- `types.rs`: removed `.clone()` on `Copy` types in two tests.
- `lib.rs` (was `mod.rs`): `pub mod` / `pub use` listing identical.
- `TODO/FIXME/unimplemented!/panic!` count: 8 → 8 (preserved).
- New `examples/{compressed_rtf,inline_demo,parse_rtf,write_rtf}.rs` are
  demos (≤108 LOC each); not present on `main`.

⚠ `crates/litchi-rtf/src/error.rs:9` — `RtfError` enum gained
`#[non_exhaustive]`:

Before (`main:src/rtf/error.rs`):

```rust
/// RTF parsing errors.
#[derive(Debug, Clone)]
pub enum RtfError {
    /// Lexer error during tokenization
    LexerError(String),
```

After (`crates/litchi-rtf/src/error.rs`):

```rust
/// RTF parsing errors.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum RtfError {
    /// Lexer error during tokenization
    LexerError(String),
```

This does not change parser output, error construction, or matching inside
the crate (no internal exhaustive `match` on `RtfError` was found). It is a
public-API tightening that forces external consumers to add `_ =>` arms.
The brief lists "Changed error semantics" as a flag item; technically
runtime semantics are unchanged but the `match` exhaustiveness contract is
narrowed. Listed here as a heads-up rather than a logic regression — no
`panic!`/`assert!` was removed and no error variant was added or removed.

Representative reviewed files: `crates/litchi-rtf/src/lexer.rs`,
`crates/litchi-rtf/src/picture.rs`, `crates/litchi-rtf/src/writer.rs`,
`crates/litchi-rtf/src/parser.rs`, `crates/litchi-rtf/src/error.rs`.

## 4. Summary

**Verdict: MOSTLY MECHANICAL (1 concern)**

- `litchi-odf`: pure mechanical — path rewrites, idiomatic `match`-guard
  refactors that preserve fallthrough into existing `_ => {}` arms, redundant
  `.clone()` removal on `Copy` types, and one in-file impl-block relocation.
- `litchi-iwa`: pure mechanical — proto sources byte-identical, build.rs
  faithfully migrated, archive/snappy/object-index logic untouched.
- `litchi-rtf`: lexer/picture/writer byte-identical; the only non-cosmetic
  change is `#[non_exhaustive]` added to the public `RtfError` enum, which
  affects only external `match` exhaustiveness, not parser behavior.

No algorithm changes, no removed validation/bounds checks, no struct field
reordering, no removed `panic!`/`assert!`/`debug_assert!`, no new
`unimplemented!()`, no silently dropped code paths.
