# Workspace-Split Audit: `litchi-eval` and `litchi-formula`

## 1. Scope

Compare the carved-out crates against their pre-split counterparts on `main`:

| New (this branch)              | Old (`main`)            |
| ------------------------------ | ----------------------- |
| `crates/litchi-eval/src/`      | `src/sheet/eval/`       |
| `crates/litchi-formula/src/`   | `src/formula/`          |

Goal: confirm the move is mechanical — same logic, same numerics, same control flow.

## 2. Method

- File-set parity check via `git ls-tree main` vs `find crates/<crate>/src` (paths must align).
- Per-file `diff main:<old>..<new>` with import-path / feature-name normalization (`crate::sheet::eval::engine` → `crate::engine`, `crate::sheet::` → `litchi_core::sheet::`, `crate::formula::` → `crate::`, `eval_engine_web_functions` → `web_functions`).
- Residual hunks then reviewed by hand.
- Confirmed deleted `src/sheet/eval/engine/tests/mod.rs` corresponds to a relocated integration-test crate at `tests/eval_xlsx_integration/` (3 modules promoted to `main.rs`-driven integration test; bodies bit-identical except for the `crate::` → `litchi::` import switch).

## 3. Findings

### 3.1 `litchi-eval` — OK (PURE MECHANICAL)

File-set parity holds (1:1 with `src/sheet/eval/`, plus `mod.rs`→`lib.rs` rename). All non-trivial residual hunks fall into one of these benign buckets:

**a. Sheet-trait promotion to `litchi-core` (per commit `e34a58e`, allowed)**
`use crate::sheet::{CellValue, Result}` → `use litchi_core::sheet::{CellValue, Result}` across every file.

**b. Internal path shortening**
`crate::sheet::eval::engine::*` → `crate::engine::*`, `crate::sheet::eval::parser::*` → `crate::parser::*`, `crate::sheet::eval::BoxFuture` → `crate::BoxFuture`. No re-exports were widened beyond what the old paths already exposed.

**c. MSRV-driven equivalent rewrites** (workspace pins `rust-version = 1.85`; `is_multiple_of` stabilized in 1.87)
- `crates/litchi-eval/src/engine/criteria_aggs.rs:20`, `:235`, `:290`, `:348`
- `crates/litchi-eval/src/engine/logical.rs:30`, `:170`
- `crates/litchi-eval/src/engine/text/excel_formatter.rs:85`
- `crates/litchi-eval/src/engine/text/formatting.rs:505`

  Old (`main`):
  ```rust
  if args.len() < 3 || args.len().is_multiple_of(2) { ... }
  ```
  New:
  ```rust
  if args.len() < 3 || args.len() % 2 == 0 { ... }
  ```
  And `!x.is_multiple_of(2)` → `x % 2 != 0`. Mathematically identical.

**d. Equivalent range-contains rewrites** (same MSRV reason)
- `crates/litchi-eval/src/engine/math/random.rs:77`, `:104`, `:131`, `:145`
- `crates/litchi-eval/src/engine/statistical/distributions.rs:2998`, `:3034`

  ```rust
  // OLD (main)         : assert!((0.0..1.0).contains(&v));
  // NEW (this branch)  : assert!(v >= 0.0 && v < 1.0);
  ```
  Semantically identical assertion. No `assert!` was removed.

**e. Test-module relocation** (allowed; verified not lost)
`crates/litchi-eval/src/engine.rs` drops the inline `mod tests;` declaration; the three xlsx-driven test modules (`aggregate_logical.rs`, `financial.rs`, `lookup_text.rs`) now live in `tests/eval_xlsx_integration/` at the workspace root (`main.rs` is a thin module aggregator). Test bodies are byte-identical apart from the necessary `use crate::…` → `use litchi::…` import flip and the `#![cfg(all(test, …))]` attribute being moved to `main.rs`.

**f. Feature rename**
`#[cfg(feature = "eval_engine_web_functions")]` → `#[cfg(feature = "web_functions")]` inside `litchi-eval`. The umbrella `crates/litchi/Cargo.toml` re-exports it as `eval_engine_web_functions = ["eval_engine", "litchi-eval/web_functions"]`, so the public-facing feature name and gating semantics are preserved.

**g. Rustfmt drift on shorter paths**
`engine/text/modern.rs:107-128`, `:198-210`: shorter `crate::engine::to_number(...)` calls now fit on one line where the old `crate::sheet::eval::engine::to_number(...)` did not. Pure formatting — argument count, order, await points unchanged.

**h. `lib.rs` only**
`#![allow(missing_docs)]` added at top. No re-export surface change vs the old `mod.rs`.

**No flagged items**: function-registry entries, dispatch table, operator precedence, type coercion (`to_number`, `to_bool`, `to_text`), error strings, struct field types, function arity, default constants, panics, assertions, and all `#[cfg]` gating are byte-identical (modulo (a)–(h)). No new `TODO` / `FIXME` / `unimplemented!()`. Spot-checked: `engine.rs` (lib.rs), `engine/dispatch.rs`, `engine/registry.rs`, `engine/financial/cashflows.rs` (largest file), `engine/text/modern.rs`, `parser/expr.rs`, `parser/literal.rs`, `parser/structured_ref.rs`.

### 3.2 `litchi-formula` — OK (PURE MECHANICAL)

File-set parity holds (1:1 with `src/formula/`, plus `mod.rs`→`lib.rs` rename). Residual hunks are exclusively:

**a. Import-path updates**
`crate::formula::*` → `crate::*` everywhere; `crate::common::*` → `litchi_core::*` where present.

**b. Rustfmt drift**
`omml/handlers/accent.rs:3`, `omml/handlers/group_char.rs:3`: shorter import paths now fit on one line. No symbol set changed.

**c. `#[non_exhaustive]` added to public error/info enums**
- `crates/litchi-formula/src/lib.rs:83` (formula info enum)
- `crates/litchi-formula/src/latex/conv/error.rs:5` (LatexConvError)
- `crates/litchi-formula/src/mtef/mod.rs:64` (MtefError)
- `crates/litchi-formula/src/omml/error.rs:3` (OmmlError)

  This is a future-proofing attribute — no variants added, removed, reordered, or renamed; no field types changed. Source-compat-only marker, runtime behavior unchanged.

**d. Match-guard refactor in `omml/utils.rs`** (semantically equivalent)
- `crates/litchi-formula/src/omml/utils.rs:433-438` (Root)
- `crates/litchi-formula/src/omml/utils.rs:474-477` (Fenced)
- `crates/litchi-formula/src/omml/utils.rs:507-526` (validate_element_nesting: Math, Numerator/Denominator, Degree)

  Old:
  ```rust
  super::MathNode::Root { base, .. } => {
      if base.is_empty() {
          return Err(super::OmmlError::MissingRequiredElement(
              "Root base is empty".to_string(),
          ));
      }
  },
  ```
  New:
  ```rust
  super::MathNode::Root { base, .. } if base.is_empty() => {
      return Err(super::OmmlError::MissingRequiredElement(
          "Root base is empty".to_string(),
      ));
  },
  ```
  The terminal wildcard `_ => {},` is preserved in both `validate_math_node` (around `:495`) and `validate_element_nesting` (around `:570`), so when the guard fails the match falls through to the wildcard arm exactly as the old `if`-inside-arm did. Identical behavior; same error variant, same error message.

**e. `lib.rs` only**
`#![allow(missing_docs)]` added at top. Module declarations and re-exports otherwise byte-identical to the old `src/formula/mod.rs`.

**No flagged items**: MathML/MathType element dispatch (`omml/elements.rs`, `omml/handlers/*`), LaTeX template strings (`latex/templates.rs`), operator/symbol tables (`latex/operators.rs`, `latex/symbols.rs`, `mtef/templates.rs`, `mtef/constants.rs`), MTEF binary parser (`mtef/binary/parser.rs`, `mtef/binary/headers.rs`, `mtef/binary/objects.rs`), AST node shapes (`ast/types.rs`, `ast/node.rs`), and matrix/conversion logic are byte-identical. No new `TODO` / `FIXME` / `unimplemented!()`. No removed `panic!` / `assert!` / `debug_assert!`. No changed enum discriminants.

## 4. Summary

**PURE MECHANICAL** — both `litchi-eval` and `litchi-formula` are behaviorally identical to their `main`-branch sources. All residual deltas are:
- import-path adjustments mandated by the crate split;
- `#[non_exhaustive]` source-compat markers on error enums;
- MSRV-equivalence rewrites (`% n == 0` ↔ `is_multiple_of`, `>=`/`<=` ↔ range-contains) compelled by `rust-version = "1.85"`;
- match-guard refactor with preserved wildcard fallthrough;
- rustfmt collapsing now-shorter call sites and use-trees;
- relocation (not deletion) of three xlsx integration-test files;
- one feature rename internal to `litchi-eval` (`eval_engine_web_functions` → `web_functions`) re-aliased by the umbrella to keep the public feature name stable.

Concerns: 0.
