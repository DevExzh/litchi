# 0394 — W2-D: relocate the Numbers drawable-comment example

- Status: bounded W2-D relocation on `uo2/iwa-w2d-numbers-drawable-comment`,
  cut from coordinator source `1bc8f22499d2691e85ebed37686c8381f0c985da`.
- Scope: move one host command example to the already integrated
  selector-first Numbers drawable-comment family. No runtime, parity, native
  queue, ledger-counter, fixture, or host-oracle changes are included.

## Relocation and behavior

| Source | Destination | Leaf route |
|---|---|---|
| `crates/litchi-iwa/examples/edit_numbers_drawable_comment.rs` | `crates/litchi-numbers/examples/edit_sheet_drawable_comment.rs` | `Package::{sheet_drawable_comment, edit_sheet_drawable_comment}`, `SheetSelector`, `DrawableSelector` |

The CLI retains the input/output path, sheet index, drawable index, replacement
text, and `--clear` arguments. It reads the old semantic body, stages a
selector-first set or clear transaction, saves the committed package, reopens
the saved candidate, and prints the old/new semantic bodies. The leaf route
does not expose the host drawable identifier or recreate a host convenience
alias. Existing host tests and native provenance remain untouched.

The tracked source-built parity family already covers the same direct drawable
comment owner for text boxes, shapes, movies and charts; this example is only
the runnable command surface for that landed leaf capability.

## Evidence boundary

This is an example relocation only. It makes no native acceptance or
performance claim and does not alter migration ledger or macOS queue rows.
The host example is removed only after the equivalent leaf example compiles,
runs both set and clear on a tracked Numbers fixture, and reopens its saved
outputs.

## Verification

With `CARGO_BUILD_JOBS=4`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_DEV_DEBUG=0`:

- `cargo fmt --all -- --check`: pass.
- `cargo build --offline --locked -p litchi-numbers --example
  edit_sheet_drawable_comment`: pass; existing vendor warnings only.
- Strict Clippy for the example (`cargo clippy --offline --locked -p
  litchi-numbers --example edit_sheet_drawable_comment --all-features --
  -D warnings`): pass; existing vendor warnings only.
- Warning-denied rustdoc for the example: pass; existing vendor warnings only.
- `python3 -m unittest tools.test_no_3rdparty_references
  tools.test_iwa_exit_ledger`: 17 passed; `iwa_exit_ledger.py --check`: pass.
- `python3 tools/check_crate_boundaries.py`: pass; 70 packages, 313
  declarations, 12 explicit debt items.
- On tracked `test-data/iwork/numbers/image-adjustments-native.numbers`,
  drawable index `0` returned the typed `NotCommentable(table)` refusal. The
  same fixture's drawable index `1` completed the set run and saved/reopened
  with `old=None`, `new=Some("W2-D drawable comment")`, then the clear run
  saved/reopened with `old=Some("W2-D drawable comment")`, `new=None`.
- The example's inverse check restored exact source bytes for both operations.
- Set output SHA-256:
  `fdbcd4393815772f8d3a30da7fb7a4999df8b4aba4778015f464b09863e3a3eb`.
- Clear output SHA-256:
  `fea6baf2ac44cd0e3c2f53f1290dac265ea65af4d84277c4f5b3c53463e76fd9`.
- External run logs and outputs are retained under
  `/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/perf/w2d-0394/run-final`.
