# 0376b: W1-F volatile conditional-highlight owner family

- Status: landed on branch `uo2/iwa-w1f` (trunk `b160a6666b` as the base).
- Date: 2026-10-26.
- Input: the coordinator's addition to 0376a. W1-A's Numbers conditional
  highlights refused volatile date rules (`Refusal::VolatileRule`, record
  0370d). The task was to port the host's twelve-owner volatile family into
  `litchi-numbers-wire`, lift the refusal in `cell::conditional_highlight`
  (the call-site hunk was cleared with W1-A through the coordinator), and
  reach parity with the host.
- Base records: [0376](0376-iwa-w1f-formula-cache-engine.md),
  [0376a](0376a-iwa-w1f-formula-authoring-batches-cross-table.md),
  [0370d](0370d-iwa-w1a-conditional-highlights.md).
- `performance_claim`: none. The route was refused before; non-volatile
  rules gain one scan for the conditional-style owner when they replace or
  clear an existing entry.

## 1. What landed

| Crate | What |
|---|---|
| `litchi-numbers-wire` | `conditional_volatile::{ensure, remove}`: a wire planner for the host's `ensure_volatile_owner` and `remove_volatile_host`; typed `Error::{InvalidSource, MissingTableOwner, IncompleteFamily, Rewrite, Allocation}`; unit tests |
| `litchi-numbers` | `conditional_highlight::volatile` publishes the planned objects, UUIDs and engine references; `Refusal::VolatileRule` is removed; `structure::Working::next_identifier` |
| `litchi-iwa` (tests) | `numbers_volatile_highlight_cases.rs`; the `save_token` metadata policy accepts one more leaf stamp (§3) |
| docs | this record; two macOS queue rows |

## 2. What the host does, and what the leaf now does

When a rule set contains a today, yesterday or tomorrow rule (or another
volatile date condition, as `is_volatile_date_condition` defines it):

- **A table without a registered conditional-style owner.** The host creates
  the family: 12 owners in the host's object order, 4 cell-record tiles,
  owner-map entries in internal order, and a formula count of 8. When the
  engine already lists one owner, it reorders the owners into the native
  order. It also adds a `[2, 6]` object-reference field info. The leaf
  builds the same objects, payloads and order: each owner is followed by its
  tile, which is also the UUID-registration order.
- **A table whose family exists.** The host appends one host record to the
  conditional-style owner, its volatile-time set, its table UUID reference
  and its tile, and raises the formula count by one. The leaf does the same
  with the host's coordinate-set merge.
- **A replace or clear.** The host first unregisters the cell
  (`remove_volatile_host`) and registers it again for volatile rules. The
  leaf does the same; removing a cell that is not registered needs no edit.

UUIDs follow the harness's documented policy. The global-name owner's UUID
is random on the host; the leaf derives it from the table's conditional
owner. The created objects take deterministic UUIDs.

## 3. Save token: a documented divergence

The host advances the save tokens of the tile, list and model archives, but
not of `CalculationEngine.iwa`, which its volatile route rewrites. The leaf
stamps that component too, because the metadata codec requires a new token
on every component that gains object UUIDs
(`validate_addition_save_token_selector_coverage`).

The harness's `save_token` policy (`archive_diff.rs`) now accepts this one
shape. It applies only when both sides advanced the document token, and only
to a component where:

- the host left the component at its input token; and
- the leaf stamped it with the new document token.

Any other token difference is still a violation.

## 4. Parity

| Test | Route | Tier | Result |
|---|---|---|---|
| `numbers_conditional_highlight_volatile_parity` | B2 today, C3 yesterday, E7 tomorrow; each is the table's first volatile rule, so it creates the family | B | pass |
| `numbers_conditional_highlight_volatile_lifecycle_parity` | over a host-created family: a second cell (C3 tomorrow), a non-volatile replacement (B2 blank), a volatile replacement (B2 yesterday), a clear (B2) | B | pass |
| `numbers_native_conditional_highlight_volatile_is_refused` | `basic.numbers`, `custom-number-native.numbers`: B2 today | — | pinned: host fails, leaf refuses `UnsupportedDependency` |

The inputs are host `numbers::Builder` documents with 8 × 6 and 12 × 40
tables, as in the host's own volatile tests. No tracked fixture works.
Every Numbers-saved table in `test-data/iwork/numbers/` (37 files) registers
the conditional-style owner without the family's tile:

- the host fails on all of them ("Numbers conditional-style dependency tile
  is missing");
- the leaf returns `IncompleteFamily`, which the litchi-numbers route maps to
  the typed refusal `UnsupportedDependency`.

The known-field decoders drop:

- the table model's conditional-style owner UUID (record 0370d's policy);
- the random UUIDs of owner kinds 3 and 7;
- the engine's owner-map UUIDs.

The ledger rows `set_cell_conditional_highlighting` and
`clear_cell_conditional_highlighting` were already migrated (0370d); their
routes now also cover volatile rules.

The macOS queue gains two W1-F Numbers rows:

- the leaf outputs of the parity cases;
- a Numbers-created volatile document, to show how Numbers itself stores the
  family. This would also give the Numbers-saved shape a native oracle.

## 5. Gaps

| Gap | Note |
|---|---|
| Volatile rules on Numbers-saved tables | refused on both sides (§4); waits for the native-created fixture in the macOS queue |
| Keynote and Pages volatile highlights | their conditional-highlight routes are host-only in the ledger (W1-B, W1-C) |

## 6. Verification

The commands are the gate set of record 0376 §7: fmt, clippy over
`litchi-numbers-wire`, `litchi-numbers` and `litchi-iwa` (all targets, all
features, `-D warnings`), the litchi-ops chunks, `litchi`, `litchi-numbers`,
`litchi-iwa` (parity included), pyo3, perf-baseline, and
`iwa_exit_ledger.py --check`. See the final report for the results.
