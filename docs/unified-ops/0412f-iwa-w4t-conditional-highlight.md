# 0412f — W4-T: Pages and Keynote conditional highlights on the shared engine

- Status: merge-ready on branch `uo2/iwa-w4t` after
  [0412e](0412e-iwa-w4t-table-topology.md).
- Date: 2026-11-02.
- `performance_claim`: every route uses fewer instructions than its host
  route (§3).

The four Pages and four Keynote `table.conditional_highlight` rows are
migrated.

## 1. The shared engine

`litchi_iwa_table::conditional_highlight` is the Numbers conditional-highlight
engine (`litchi-numbers/src/package/conditional_highlight/`, record 0370d,
volatile rules 0376a), moved with `git mv` and made generic over
`TableLeaf`. Its native graph writer covers:

- the conditional-style list and its entry;
- the style set's pre-pivot and current predicates and their formula
  graphs;
- the per-rule paragraph and cell styles;
- the cell's key and applied rule;
- the calculation engine's volatile owner family.

Its cell helpers are the shared `cell_rich_text` ones (locate, materialize,
replace), and its table geometry is the shared `topology` one. The engine
exposes `rules_of`, `apply_one`, `Staged` and `MAX_RULES`, and a typed
`Error`, `Refusal` and `LimitKind` that each leaf maps into its own family
error.

Numbers keeps its public `cell::conditional_highlight` family: the
selectors, the staged `Edit`, verification and the patch, now in
`package/conditional_highlight.rs`. The Numbers wrappers that only this
module used are deleted: `table_cell_text_style::{locate_cell,
materialize_cell}`, `cell::wire::encode`, and topology's list re-exports.
All `litchi-numbers` tests and the Numbers conditional and volatile parity
cases pass unchanged.

## 2. Facades

`litchi_pages::table::conditional_highlight` and
`litchi_keynote::slide::table::conditional_highlight` provide
`Package::{body,slide}_table_cell_conditional_highlight` (read the ordered
rules), `edit_*_cell_conditional_highlights` and
`apply_*_cell_conditional_highlights`. `Edit::{set, clear}` takes one to
fifteen rules per cell, and several cells per edit. Each change is checked
by reading its rules back. Refusals are typed.

## 3. Parity and performance

The cases mirror the Numbers family's: every non-volatile condition kind
as a single rule on a stored cell and on an empty cell, in-place
replacement with fewer and with more rules, a second cell sharing the list,
clears, and reads. A table's first rules give its model a fresh
conditional-style owner, which is random on the host and deterministic on
the leaf. The writes are therefore tier B, with that owner field dropped
from the model comparison.

| Test (`pages_body_table_` / `keynote_slide_table_` prefix) | Tier |
|---|---|
| `conditional_highlight_set_parity` (28 condition kinds × 2 cells) | B |
| `conditional_highlight_replace_parity`, `_clear_parity` | B |
| `conditional_highlight_read_parity` | A |
| `conditional_highlight_volatile_refused_like_host` (a native owner without the volatile family: both routes refuse) | — |

Callgrind, release parity binary SHA-256
`b0dd07c67e864e1805a1bb226c7e7727798b9a4489e5ade634fc0d09a071bfeb`, CPU 5,
ABBA:

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Pages three-rule set (`pages_body_table_conditional_highlight_cost_abba`) | 3,521,638,712 | 103,230,822 | 0.029 |
| Keynote three-rule set | 966,129,629 | 369,586,881 | 0.383 |
| Numbers set and clear (`numbers_conditional_highlight_cost_abba`, after the lift) | 507,273,133 | 229,221,086 | 0.452 |

Ledger: 4 Pages and 4 Keynote rows migrate. Four native queue rows. The
glossary registers both families.

## 4. Verification

The gate script `(deleted-coord-artifact)` runs on the merge of trunk; its
result is reported in the hand-off.
