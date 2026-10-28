# 0370k — W1-A slice 12: numeral systems on the shared engine, and the Numbers DataFormat facade

- Status: merge-ready on branch `uo2/iwa-w1a`, after slice 11
  ([0370j](0370j-iwa-w1a-sheet-duplicate.md)) and the duplicate-core lift.
- Date: 2026-10-27.
- `performance_claim`: the Numbers numeral-system route is 1.4% above the
  host route, within the 5% instruction bar (§4).

## 1. What landed

### Shared engine: `litchi_iwa_table::cell_format`

The change is additive and keeps the signatures of the existing functions
(coordinator assignment; W1-G owns the module).

- **Numeral systems** are now a display family. The new functions are
  `display::{read_numeral_system_format, read_numeral_system_format_with_budget,
  rewrite_numeral_system_format}`. The payload codec mirrors the host
  editor's:
  - format type 269, with the base, places and minus-sign fields (8, 9,
    10);
  - any other numeric option is invalid;
  - a rewrite in place keeps every other field byte for byte;
  - a new entry is canonical, with fields 1, 8, 9 and 10 in that order.
- **`data_format`:**
  - `read` and `rewrite` dispatch numeral systems.
  - `is_display_family` admits them.
  - `read` now validates the cell, its tile and its format list once for
    all ten display families, through the new
    `display::read_any_display_format`. Before, it ran up to ten family
    readers in turn, each repeating the same validation. The result is
    unchanged: every Pages, Keynote and Numbers format suite passes.

### Numbers: one family for any format

`litchi_numbers::cell::data_format::transaction` provides
`Package::{table_cell_data_format, edit_table_cell_data_format,
apply_table_cell_data_format}`. The read is the shared dispatch. How a
write is published depends on what it touches:

- **Within one family, or between Automatic and a format:** the change
  runs that family's Numbers transaction (Number, Currency, Percentage,
  Scientific, Fraction, Date and Time, Duration, Text, Custom, or the cell
  controls).
- **Between two families:** two transactions run on a private
  intermediate candidate, the first to Automatic.
- **Involving a numeral system:** every step goes through the shared engine
  with appended list keys, as the host editor's compatibility writer
  appends them.

The patch keeps only the source and final artifacts.

Seven ledger rows are migrated:

- Numbers: `table_cell_data_format`, `set_table_cell_data_format`, and the
  three numeral-system routes.
- Pages: `table_cell_numeral_system_format`.
- Keynote: `slide_table_cell_numeral_system_format`.

## 2. Parity

| Test | Routes | Tier |
|---|---|---|
| `numbers_table_cell_data_format_set_parity` | Automatic to three numeral systems and to Number; numeral to numeral; numeral to Automatic; Number to numeral (`custom-retirement-cleared.numbers`, cell B2) | B |
| `numbers_table_cell_numeral_system_format_parity` | the host's numeral set and reset | B |
| `numbers_table_cell_data_format_read_parity` | every decimal family and numeral system the host writes | A |
| `pages_body_table_cell_numeral_system_format_read_parity`, `keynote_slide_table_cell_numeral_system_format_read_parity` | the leaf reads each numeral system the host writes | A |
| `numbers_table_cell_data_format_cross_family_difference` | numeral to Number reads the same on both sides; the list keys differ (documented) | — |

`custom-retirement-cleared.numbers` is the one corpus fixture with a
stored number cell that both routes read as Automatic and that the engine
admits.

## 3. Gaps

| Item | Next step |
|---|---|
| a numeral write on Pages and Keynote native cells | Those cells name an implicit Number format entry, which the engine keeps in place (the W1-G §5 conversion gap). The leaf refuses; the read is covered. |
| a numeral system replaced by another family | Same result as the host, but different list keys: the host interns the new entry first. It needs the engine's conversion path (W1-G §5). |
| leaf cost of the numeral route (§4) | closed by the package-level proof census described in §4; the route is within the performance bar |

## 4. Performance

`cost_probes::numbers_cell_data_format_cost_abba` sets a base-16 numeral
system on B2 of `custom-retirement-cleared.numbers`, in the order host,
leaf, leaf, host:

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| numeral system set | 106,188,551 | 107,635,517 | **+1.36%** (within budget) |

The first version cost 417,826,851 Ir. Two changes brought it down:

- `data_format::read` now validates once instead of up to ten times.
- The facade no longer reads the source format twice.

About 104 M Ir of what remains is the engine's
`control::require_exclusive_selected_model_reference`. It runs once each
for the read, the rewrite and the verification read. The first cold proof
calls `inspect_references_with_policy` on every object in the package; later
proofs replay the admitted census and charge the same reference occurrences
without re-encoding or decoding every header.

The fix keeps strict `RejectUnknownMetadata` admission while avoiding
repeated full-package work. `PackageProofs` memoizes the reference census and
metadata registry for an immutable package. A candidate reopened from an edit
inherits the census for unchanged parsed archives and inspects only changed
ones; each proof still charges the same headers, references and limits. The
unrelated-object unknown-metadata refusal remains covered by
`unknown_metadata_on_an_unrelated_object_still_refuses`.

The post-merge callgrind files remain outside the repository in the
coordinator scratch area: `callgrind-df4-host.out` (140,169 bytes,
SHA-256 `3771f4aa246bfae9943b8cd4ad53a2a6f88c998de970ce3e18546b865a417364`)
and `callgrind-df4-leaf.out` (134,766 bytes,
SHA-256 `d8627c503375041b410e57a9100e43cf599f1d8af820551894da950f64b6c78b`).
They are regenerated by pinning CPU 3 and running the release parity binary
under callgrind with `--toggle-collect` for `host_data_format` and
`leaf_data_format`.

Because the shared proof seam is also used by the non-memoized Pages and
Keynote adapters, their existing Currency replacement probes were rerun after
the trunk merge:

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| Pages `pages_cell_format_cost_abba` | 3,357,413,700 | 398,330,691 | **−88.14%** |
| Keynote `keynote_cell_format_cost_abba` | 785,831,951 | 801,639,685 | **+2.01%** (within budget) |

Those artifacts are also outside the repository: Pages host/leaf are
`callgrind-pages4-host.out` (149,360 bytes,
SHA-256 `177b81acdc94ec51af53d3885d051b2adabe1ca0d5fb0a7dbe3856b2f12f1121`)
and `callgrind-pages4-leaf.out` (147,270 bytes,
SHA-256 `cf5a92c2599d76f1ce104a7185e543aeed859cbf9826308d72c94505e22440cb`);
Keynote host/leaf are `callgrind-keynote4-host.out` (193,471 bytes,
SHA-256 `e4c6d13903b6a62f058614cd131e78179f52218ffcc325ec4b57534d2303036a`)
and `callgrind-keynote4-leaf.out` (159,343 bytes,
SHA-256 `d8b1b7f728cd5646702dd5d28e880e55f9f8251fa046fa3cd85832efe09c4c96`).
All four use the release parity binary, CPU 3, and the corresponding
`pages_cost_probes::{probe_host,probe_leaf}` collection toggle.

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-table -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa-table --all-features`; `cargo test -p litchi-numbers --all-features`; `cargo test -p litchi-pages --all-features --lib`; `cargo test -p litchi-keynote --all-features --lib` | 0 |
| `cargo test -p litchi-iwa --all-features --test parity numeral_system_format_read_parity`; `... numbers_table_cell` | 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` / `pages_unified` / `keynote_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-iwa-table` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 -m unittest tools.test_check_iwork_leaf_public_api tools.test_check_api_glossary` | 0 |
| `cargo check -q --locked --offline --manifest-path crates/{litchi-numbers,litchi-numbers-wire,litchi-pages,litchi-keynote}/fuzz/Cargo.toml --bins` | 0 |
