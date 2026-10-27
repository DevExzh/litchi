# 0370i — W1-A slice 10: Numbers builder table dimensions, language and locale

- Status: landed on branch `uo2/iwa-w1a`, after slice 9
  ([0370h](0370h-iwa-w1a-table-duplicate.md), merged on trunk).
- Date: 2026-10-26.
- `performance_claim`: none. The leaf builder with options runs 3.75 times the host's instructions, over the 5% budget (§4, gap).

## 1. What landed

`litchi_numbers::Builder` gains three options:

| Option | What it sets | Default |
|---|---|---|
| `table_dimensions(rows, columns)` | the sheet table's size, headers included | 10 × 5 |
| `language(tag)` | a BCP 47 tag | `en` |
| `locale(id)` | the locale | `en_US` |

There are two new error variants: `create::Error::InvalidLanguage` and
`InvalidLocale`. A size the table cannot take fails with
`create::Error::Dimensions`.

**Where the values go** (the host builder writes the same fields):

- Dimensions: the build resizes the template's table through
  `table::resize`. It then records the new extent in the table's tile
  (`maxColumn`, `maxRow`).
- Language: `TSA.DocumentArchive.document_language`.
- Locale: the document's `locale_identifier` and
  `creation_locale_identifier`, and the calculation engine's
  `saved_locale_identifier`.

The new `package/document_locale.rs` rewrites these fields on the wire
through the structural working set.

**Host changes:**

- *Defect fixed* (mac-queue row): the host builder's table owner carried no
  table or body extents. Native owners record both, and so have tables
  added by the editor since record 0370f. The Pages builder borrows this
  graph for its first table and keeps the extent-less owner
  (`Builder::without_owner_extents`): the Pages body-table hidden-axes
  route refuses an owner with spanning extents. That is W1-B's to lift.
- *Alignment, not a defect:* the host builder's build-history producer
  string is now the focused crate's `Created by Litchi` (record 0309).

**Documented difference:** a table with no body row or column (1 × 1).
The host builder writes one; the leaf refuses it, as `table::resize`
does, because a native table keeps at least one body row and column.
`numbers_builder_headers_only_refusal` pins this.

Three ledger rows are migrated: `table_dimensions`, `language` and
`locale` → `litchi_numbers::Builder`.

## 2. Parity

`crates/litchi-iwa/tests/parity/numbers_builder_cases.rs` builds with the
host builder and then with the leaf builder under the host output's
identity. That identity is the document, version and private UUIDs read
from its properties.

Both tests are tier B. The host draws these values at random, so they are
compared by form, as a documented policy:

- the table UUID, the owner UUID and the owner-map entry derived from it;
- the row/column UIDs (lengths and permutations);
- the registry's object UUIDs.

| Test | Cases |
|---|---|
| `numbers_builder_table_dimensions_parity` | 10×5, 20×8, 3×2, 2×2, 40×26 |
| `numbers_builder_language_locale_parity` | fr/fr_FR, ja/ja_JP, de-CH/de_CH |

## 3. Gaps

| Item | Next step |
|---|---|
| leaf builder cost (+275% over the host, budget 5%) | build the options in one pass over the template's archives before the single open: synthesize the resized table (model size, UID map, stroke sidecar, owner extents, tile extent) and the locale fields on the wire, with no intermediate publication. Until then, only builds that set options pay the cost; `Package::new` and identity/name-only builds are unchanged |
| no native evidence for builder options | mac queue |

## 4. Performance

`cost_probes::numbers_builder_cost_abba` builds a 20 × 8 `fr`/`fr_FR`
workbook in the order host, leaf, leaf, host:

The absolute cost is small: about 9.3 M instructions, once per new
document. The host writes about 6 KiB of archives straight from its
protobuf values. The leaf opens its template (1.4 M Ir), then publishes
and reopens twice: the resize costs 4.2 M Ir and the language and locale
pass 3.4 M Ir, of which the registry parse in `metadata::apply` is 2.3 M
Ir. The budget is missed; §3 lists the follow-up.

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| builder with options | 2,480,707 | 9,304,911 | **+275.1%** |

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test --all-features -p litchi-pages`, `-p litchi-keynote` | 0 / 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
