# 0370k — W1-A slice 12: numeral systems on the shared engine, and the Numbers DataFormat facade

- Status: merge-ready on branch `uo2/iwa-w1a`, after slice 11
  ([0370j](0370j-iwa-w1a-sheet-duplicate.md)), the duplicate-core lift, and
  the W1-G trunk merge.
- Date: 2026-10-28.
- `performance_claim`: the Numbers numeral-system route is 2.4% above the
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
  - `read_with_budget` lets candidate verification reuse the transaction
    ledger that admitted publication; automatic custom-format cells still
    return `DataFormat::Automatic`.

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
| numeral system set | 106,219,490 | 108,707,986 | **+2.34%** (within budget) |

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
ones; each proof precharges the complete object/data reference-occurrence
bound before cold construction, then replays that charge exactly once. The
unrelated-object unknown-metadata refusal remains covered by
`unknown_metadata_on_an_unrelated_object_still_refuses`. Numeral-system
candidate reopen and readback both use the original transaction budget and
`charge_candidate_reopen`; `a_low_reference_budget_refuses_before_cold_census_inspection`
and `candidate_reopen_budget_is_charged_before_readback` cover the two
admission boundaries.

The final integrated callgrind files remain outside the repository in the
coordinator scratch area: `callgrind-df5-host.out` (140,621 bytes,
SHA-256 `d2981178e14cd7cabc95ecc79a04a2f006ab9109cee2a37bb7a54dbbc05b7192`)
and `callgrind-df5-leaf.out` (135,474 bytes,
SHA-256 `d06e4464d9a3201550ef680b62f0053d6e4ad18be8750a9466173882e9ab28e5`).
They are regenerated by pinning CPU 3 and running the release parity binary
under callgrind with `--toggle-collect` for `host_data_format` and
`leaf_data_format`.

Because the shared proof seam is also used by the non-memoized Pages and
Keynote adapters, their existing Currency replacement probes were rerun after
the trunk merge:

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| Pages `pages_cell_format_cost_abba` | 3,359,683,450 | 398,669,377 | **−88.13%** |
| Keynote `keynote_cell_format_cost_abba` | 785,283,262 | 804,420,194 | **+2.44%** (within budget) |

Those artifacts are also outside the repository: Pages host/leaf are
`callgrind-pages5-host.out` (149,109 bytes,
SHA-256 `44fb73aab94cbd0fd069339f9717888364ef5cde991c306f99dde9bb8c220deb`)
and `callgrind-pages5-leaf.out` (147,297 bytes,
SHA-256 `ccbb3739612d5934bb7eaa2b9b7e35c5fbc7eb0ac6ca14ebe01568edf3b2e2c1`);
Keynote host/leaf are `callgrind-keynote5-host.out` (193,510 bytes,
SHA-256 `fcc58763f9482ac4933a08f82c533f345d6be04f692eb8d0858d8a981d79d83a`)
and `callgrind-keynote5-leaf.out` (159,453 bytes,
SHA-256 `a25dcc47333a1c793abe809fafb25b5d3e0ecbabca6553aadf948ad351ef81fe`).
All four use the release parity binary, CPU 3, and the corresponding
`pages_cost_probes::{probe_host,probe_leaf}` collection toggle.

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-table -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa-table --all-features`; `cargo test -p litchi-numbers --all-features --lib`; `cargo test -p litchi-pages --all-features --lib`; `cargo test -p litchi-keynote --all-features --lib` | 0 |
| `cargo test -p litchi-iwa --all-features --test parity numeral_system_format_read_parity`; `cargo test -p litchi-iwa --all-features --test parity numbers_table_cell` | 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` / `pages_unified` / `keynote_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-iwa-table` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 -m unittest tools.test_check_iwork_leaf_public_api tools.test_check_api_glossary` | 0 |
| `cargo check -q --locked --offline --manifest-path crates/{litchi-numbers,litchi-numbers-wire,litchi-pages,litchi-keynote}/fuzz/Cargo.toml --bins` | 0 |
