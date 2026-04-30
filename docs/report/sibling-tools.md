# Sibling Tooling Crates Audit — `refactor/workspace-split`

## 1. Scope

Verify behavioral parity between `main` and `refactor/workspace-split` for the
three sibling tooling crates relocated from repo root to `crates/`:

| New                       | Old (main)        |
| ------------------------- | ----------------- |
| `crates/soapberry-zip/`   | `soapberry-zip/`  |
| `crates/xml-minifier/`    | `xml-minifier/`   |
| `crates/pyo3-litchi/`     | `pyo3-litchi/`    |

Out of scope: rename plumbing, workspace dependency form, `litchi` path-dep
rewiring, formatting, comments, the restored `crates/soapberry-zip/assets/test.zip`
fixture (commit `aae4f76`).

## 2. Method

For each pair, ran:

```
git diff main..HEAD -- crates/<crate> <crate>
```

(`-M`-style rename detection automatic). Eyeballed every non-`//` line and every
file whose stat showed >2 changed lines — particularly
`soapberry-zip/src/{archive,writer,locator}.rs` flagged by the prompt, and
`pyo3-litchi/src/common.rs`.

Cross-checked any non-doc change against the current source tree
(`crates/soapberry-zip/src/office.rs`, `crates/litchi-core/src/detection/types.rs`)
to verify the new code matches an actually-existing API contract.

## 3. Findings

### `soapberry-zip` — verdict: PURE MECHANICAL

Spot-checked:

- ✅ `Cargo.toml` — adds `jiff = "0.2"` to `[dev-dependencies]` only. No effect
  on the published crate; needed by tests/examples that exercise time conversions.
- ✅ `src/archive.rs` (20 lines) — every change is a doc comment crate-name
  rename: `use rawzip::…` → `use soapberry_zip::…` inside `///` blocks. No
  functional code, no signatures, no const, no validation, no error paths
  touched. Bounds, CRC handling, locator interaction, ZIP64 decoding all
  byte-identical.
- ✅ `src/writer.rs` (24 lines) — same pattern: `rawzip::ZipArchiveWriter` →
  `soapberry_zip::ZipArchiveWriter` etc., all inside `///` examples. Encoder,
  data-descriptor sequencing, central-directory finalization, offset accounting
  unchanged.
- ✅ `src/locator.rs` (14 lines) — same pattern: doc-example imports renamed
  from `rawzip` to `soapberry_zip`. EOCD scan algorithm, max-search-space
  semantics, `locate_in_file` / `locate_in_reader` paths untouched.
- ✅ `src/office.rs` (7 lines) — module-level doc example refreshed:
  `ArchiveWriter` → `StreamingArchiveWriter`, `finish()` → `finish_to_bytes()`.
  Verified against current source: lines 284 (`pub struct StreamingArchiveWriter`)
  and 297 (`pub fn finish_to_bytes`). The previous example was simply stale; the
  rename brings docs in sync with the existing public API. The only non-doc
  change is removal of the unused `use std::sync::atomic::{AtomicUsize, Ordering};`
  inside `#[cfg(test)] mod tests`, dead since prior commits.
- ✅ `src/path.rs` (2 lines) — single doc example crate rename.
- ✅ `src/reader_at.rs` (4 lines) — doc example crate rename plus changing the
  fenced block from triple-backtick to triple-backtick `ignore`. No source
  change; the `ignore` annotation just stops the example (which references a
  non-bundled `assets/test-prefix.zip`) from running as a doctest. Behavior of
  `ReaderAt` / `RangeReader` unchanged.
- ✅ `src/time.rs` (6 lines) — doc example crate rename only.

No removed `assert!` / `debug_assert!`, no removed bounds checks, no new
`TODO`/`FIXME`/`unimplemented!()`, no const or default changes anywhere in the
crate. The 20/24/14-line files flagged by the prompt are 100% doc-comment churn.

### `xml-minifier` — verdict: PURE MECHANICAL

- ✅ `Cargo.toml` — only changes are reordering `quick-xml` above `quote`
  (alphabetical, satisfies `cargo sort --check`) and adding a trailing newline.
  No version bumps, no feature changes, no logic-bearing files modified.

### `pyo3-litchi` — verdict: MOSTLY MECHANICAL (2 concerns)

Spot-checked:

- ✅ `Cargo.toml` — reorders deps (PyO3 below `litchi`), points
  `litchi = { path = "../litchi" }` (was `".."`), adds trailing comma in
  features list, drops trailing whitespace. No version / feature changes.
- ✅ `src/document.rs` — adds `#[allow(clippy::arc_with_non_send_sync)]` with
  justification comment, and converts `format!("<Document>")` →
  `"<Document>".to_string()` (clippy `useless_format`). No behavior change.
- ✅ `src/presentation.rs` / `src/sheet.rs` — single line each: same
  `#[allow(clippy::arc_with_non_send_sync)]` annotation. No body changes.
- ⚠ `crates/pyo3-litchi/src/common.rs:29` — `boxed_err_to_py_err` signature tightened.

  Before (`main`):
  ```rust
  pub fn boxed_err_to_py_err(err: Box<dyn std::error::Error>) -> PyErr {
      PyException::new_err(err.to_string())
  }
  ```

  After:
  ```rust
  pub fn boxed_err_to_py_err(err: Box<dyn std::error::Error + Send + Sync>) -> PyErr {
      PyException::new_err(err.to_string())
  }
  ```

  The body is identical (`PyException::new_err(err.to_string())`). The trait
  bound was tightened to match the upstream `litchi::sheet::Workbook::open`
  signature, which now returns `Box<dyn Error + Send + Sync>`. Mechanical
  adaptation forced by the workspace-split rather than a binding logic change,
  but worth noting since it is a signature-level edit. No call site in the
  three pyo3-litchi modules constructs a non-`Send`/`Sync` boxed error.

- ⚠ `crates/pyo3-litchi/src/common.rs:86` — `From<litchi::FileFormat> for FileFormat`
  gains a wildcard fallback arm.

  Before (`main`, exhaustive):
  ```rust
  litchi::FileFormat::Numbers => FileFormat::Numbers,
  litchi::FileFormat::Rtf => FileFormat::Rtf,
  }
  ```

  After:
  ```rust
  litchi::FileFormat::Numbers => FileFormat::Numbers,
  litchi::FileFormat::Rtf => FileFormat::Rtf,
  // `litchi::FileFormat` is `#[non_exhaustive]`; map any future
  // variants we don't yet expose in Python bindings to the closest
  // match. Until a binding is added, fall back to `Doc` rather than
  // panicking — callers can disambiguate via the format-specific
  // detection APIs.
  _ => FileFormat::Doc,
  }
  ```

  Forced by the upstream enum gaining `#[non_exhaustive]`
  (`crates/litchi-core/src/detection/types.rs:5`). All 14 variants present on
  `main` are still mapped one-to-one; the wildcard only fires for *future*
  variants. However, the choice to fall back to `FileFormat::Doc` is a
  conscious behavior decision (vs. e.g. an explicit `Unknown` variant or a
  `PyValueError`), and it is genuinely new logic in the binding layer. For any
  variant added later to `litchi::FileFormat` that is not yet wired into the
  Python `FileFormat` enum, callers will see "Doc" rather than a clear failure.
  Acceptable for a mechanical move (the Python `FileFormat` enum still has the
  same 14 variants as before, so present-day behavior is unchanged), but worth
  surfacing because it is a real new code path rather than a rename.

No removed `panic!`/`assert!`/`debug_assert!`. No new `TODO`/`FIXME`/
`unimplemented!()`. GIL handling, return types, error mapping in `to_py_err`,
and all `#[pymethods]` bodies are otherwise unchanged.

## 4. Summary

**MOSTLY MECHANICAL (2 concerns)** — both in `crates/pyo3-litchi/src/common.rs`,
both forced by upstream changes (added `Send + Sync` bound; added
`#[non_exhaustive]` on `litchi::FileFormat`). Present-day binding behavior is
unchanged; the non-trivial bit is the `_ => FileFormat::Doc` fallback choice
for future enum variants. `soapberry-zip` and `xml-minifier` are pure
mechanical moves: the 20/24/14-line files highlighted by the prompt are
exclusively doc-comment crate-name updates (`rawzip` → `soapberry_zip`),
with one out-of-date doc example brought into sync with existing public
API in `office.rs`. ZIP parser, writer, locator, CRC, header validation,
and time conversion logic are byte-identical to `main`.
