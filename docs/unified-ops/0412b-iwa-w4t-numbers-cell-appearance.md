# 0412b — W4-T: Numbers cell appearance; 0412a write measurements

- Status: merge-ready on branch `uo2/iwa-w4t` after
  [0412a](0412a-iwa-w4t-cell-format-closure.md).
- Date: 2026-11-02.
- `performance_claim`: every measured route is at most 4.08% above its host
  route (Keynote custom format), within the 5% bar (§3).

## 1. Numbers cell appearance

`litchi_numbers::table::cell_appearance` (`Package::{table_cell_appearance,
edit_table_cell_appearance, apply_table_cell_appearance}`,
`Edit::{set_fill, reset_fill, set_layout, reset_layout, set_border,
clear_border}`, `Appearance`) is the Pages and Keynote families' shape on
Numbers sheets and tables, over the same engine
(`litchi_iwa_table::cell_appearance`). The engine's plans are published on
the shared structural working set by the new
`cell_appearance::working::publish` (rewrites, field references, the new or
removed style object, UUID/external-reference/removal registrations; root
previews kept; one save-token advance per transaction), so Numbers needs no
private copy of the Keynote workspace publisher. The facade is a
`structural_family!` transaction: exact-source patch and inverse, exact
no-op, typed refusals and limits.

Parity (`crates/litchi-iwa/tests/parity/numbers_cell_appearance_cases.rs`,
the Keynote cases on `basic.numbers`): fill set/reset/rewrite, layout
set/reset, style-variation removal, a three-operation sequence, border
set/restroke/clear — tier B; reads of the first cells on the host-styled
cell and every corpus table both routes read — tier A. The nine
`table.cell_appearance` Numbers rows migrate; two native queue rows.

## 2. 0412a write routes measured

The coordinator asked for the custom and Text write probes of 0412a before
qualifying those routes.

## 3. Performance

Callgrind, release parity binary SHA-256
`dd448e697de0f88e45b5dd91186c7a17c57fcf14c19043f7ce8bdbba039bdaab`, CPU 5,
`pages_cost_probes::{probe_host,probe_leaf}`, ABBA:

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Pages Automatic to custom Number | 3,374,933,785 | 421,251,459 | 0.125 |
| Keynote Automatic to custom Number | 805,517,183 | 838,350,208 | 1.041 |
| Pages Text on an implicit-entry text cell | 3,356,906,367 | 398,391,327 | 0.119 |
| Keynote Text on a text cell | 774,461,003 | 635,728,094 | 0.821 |
| Numbers cell fill set (`numbers_table_cell_appearance_cost_abba`) | 277,896,026 | 171,544,429 | 0.617 |
