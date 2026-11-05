# 0428 — W4-O: deterministic leaf output (custom-format registry UUIDs and the other random UUID sources)

- Status: merge-ready on branch `uo2/iwa-w4o` (cut from trunk
  `feat/unified-ops-api` at `a9ff2679e2`).
- Date: 2026-11-03.
- Trigger: final qualification of `a9ff2679e2` (receipt
  `.coord/final-qual/final-a9ff2679e2-instruction-receipt.md` §2.3). The
  record 0412a cost probes `keynote_slide_table_cell_custom_format_cost_abba`
  and `pages_body_table_cell_custom_format_cost_abba` failed their
  same-length check about 4 runs in 5: the leaf wrote a different package for
  the same input.
- Result: every leaf publication path found drawing randomness now derives
  its UUIDs from package content and position. Both probes pass natively
  20 of 20 runs (10 each) and measure +3.425% (Keynote) and −87.552% (Pages)
  under callgrind (§4).

## 1. Cause and audit

`litchi_iwa_table::cell_format::display::fresh_custom_uuid` called
`litchi_core::id::generate_guid_bytes` (the OS random generator) for the UUID
of a new entry in the document custom-format registry (TSA super field 12 or
the document's field 9). The UUID is varint-encoded, so the random value
changed the bytes and often the length of the compressed output.

Audit of the leaf crates (`litchi-iwa-table`, `litchi-iwa-style`,
`litchi-iwa-text-wire`, `litchi-numbers-wire`, `litchi-pages`,
`litchi-keynote`, `litchi-numbers`, and the other `litchi-iwa-*` support
crates) for `generate_guid*`, `new_v4`, `rand`, `getrandom`, `SystemTime::now`,
`Instant::now`, `Local::now`/`Utc::now`, `RandomState`/`DefaultHasher` and
pointer-derived values:

| # | Site | Kind | Disposition |
|---|---|---|---|
| 1 | `litchi-iwa-table` `cell_format/display.rs` `fresh_custom_uuid` | random registry UUID on every new Custom format (Keynote, Pages, Numbers cell formats) | **fixed** (§2.1) |
| 2 | `litchi-pages` `body_storage_objects::text_attribute_uuid` (callers `body_bookmarks`, `body_date_time_fields`) | random `text_attribute_uuid_string` on a new bookmark or Date & Time field | **fixed** (§2.2) |
| 3 | `litchi-keynote` `slide_audio_creation::allocate_ids` | four random object-map UUIDs, a random build UUID and a build `random_number_seed` taken from it | **fixed** (§2.3) |
| 4 | `litchi-keynote` `slide_drawable_comments::metadata::fresh_storage_uuid_in_working_set` (root create, reply add, cross-component clone) | random comment storage UUID | **fixed** (§2.4) |
| 5 | `litchi-keynote` `slide_drawable_comments::engine::current_apple_date`; `litchi-pages` `text_box_comment_graph::now`; `litchi-numbers` `sheet_drawable_comment_graph` (`SystemTime::now`) | creation date of a new comment storage | **exception**: the wall-clock creation instant, which the host writes the same way |
| 6 | `litchi-keynote` `Identity::generate`, `litchi-pages` `Identity::generate` (`builder.rs`), `litchi-numbers` `create.rs` `Identity::generate` | random document / version / private UUIDs | **exception**: an explicit caller opt-in for a new, independent document identity. `Package::new` and the builders' default use fixed identities, and each crate offers a fixed or derived identity for byte-identical builds |
| 7 | `litchi-iwa-table` `cell_format/control.rs` `archive_key` | archive address as a cache key | not output-visible: used only for equality lookups in a per-transaction proof cache |
| 8 | `SystemTime` in `litchi-iwa-archive`, `litchi-iwa-detect`, `litchi-keynote`/`litchi-pages` `package.rs` | `modified: Option<SystemTime>` fields read from the ZIP container | not a source: read, never minted |

The host crate `litchi-iwa` keeps its random UUIDs and its
`Local::now` relative-date evaluation; it is the parity oracle, not a
publication path, and the parity Tier B custom-format cases already mask
the host's random UUIDs (record 0412a §2). Parity is unchanged.

Not audited exhaustively: iteration order of `HashMap`/`HashSet` values
that reach the output (about 360 declarations across the leaf crates). The
leaves' replay tests compare the bytes of two independent runs, which would
fail on a `RandomState`-ordered emission; none does.

## 2. The fix

All four sites use the derivation of record 0375a §2.1 (record 0298 §2.1):
the SplitMix64 finaliser (`litchi_iwa_style`'s salted `mix` in
`litchi-iwa-table`, `litchi_numbers_wire::table_create::mix` in the Pages and
Keynote leaves) over a seed made of the identifiers the new record names, with RFC 4122 version 4 / variant 1 bits set in the
little-endian byte layout `litchi_core::id::generate_guid_bytes` produced.

### 2.1 Custom-format registry UUID (`0e1ca3b915`)

`derived_custom_uuid(entries, registry_identifier, table_identifier,
format_payload)`: the seed XORs the registry object identifier, the owning
table model, the new entry's index and an FNV-1a digest of the encoded
`TSK.FormatStructArchive`; a counter advances until the UUID is non-zero in
both halves and absent from the registry. The format archive is now encoded
before the UUID (same budget charges). An existing equal entry is still
reused, as before.

### 2.2 Pages smart-field text-attribute UUID (`016a079b41`)

`text_attribute_uuid(identifier, message_type)`: from the new field
object's identifier (fresh above the identifier watermark) and its message
type (2034 Date & Time, 2035 bookmark).

### 2.3 Keynote slide audio creation (`15a211c150`)

Each of the four object-map UUIDs is derived from its object identifier
(`first..first+3`); the build UUID from the build identifier with a separate
salt; `random_number_seed` stays the build UUID's first four bytes.

### 2.4 Keynote comment storage UUID (`15a211c150`)

`fresh_storage_uuid_in_working_set` takes the identifier of the storage it
names (the allocated root or reply identifier, or the clone's new
identifier) and keeps its collision check against every comment storage in
the working set.

## 3. Tests

Each fixed route runs twice on the same input and compares the output:

| Test | Asserts |
|---|---|
| `litchi-keynote` `slide_table_cell_format::custom_format_output_is_deterministic` | byte-identical package; Custom read back |
| `litchi-pages` `body_table_cell_format::custom_format_output_is_deterministic` | byte-identical package; Custom read back |
| `litchi-pages` `body_bookmarks::added_bookmark_output_is_deterministic` | byte-identical package |
| `litchi-pages` `body_date_time_fields::inserted_field_output_is_deterministic` | byte-identical package |
| `litchi-keynote` `slide_audio_creation::audio_creation_output_is_deterministic` | byte-identical package |
| `litchi-keynote` `slide_drawable_comments::created_comment_storage_uuids_are_deterministic` | equal comment signatures (identifier, text, author, storage UUID, replies) with the creation date masked (exception 5) |

## 4. Measurements

Probes, release parity binary `parity-164d681639f564b6` built at
`15a211c150`, native, 10 runs each:
`keynote_cell_format_cases::keynote_slide_table_cell_custom_format_cost_abba`
and `pages_cell_format_cases::pages_body_table_cell_custom_format_cost_abba`
passed 20 of 20.

Callgrind (`qualify_iwork_cost.py`, manifest
`iwork-final-cost-manifest-w4.json`, routes
`w4_0412a_keynote_cell_custom_format` and
`w4_0412a_pages_cell_custom_format`, CPU 3, toggle-isolated inclusive Ir,
exit 0):

| Route | Host Ir | Leaf Ir | Leaf vs host |
|---|---:|---:|---:|
| Keynote custom format | 807,717,583 | 835,382,268 | **+3.425%** |
| Pages custom format | 3,374,411,371 | 420,058,665 | **−87.552%** |

Both within the 5% bar; the Keynote figure matches the final
qualification's passing rerun (+3.456%). Evidence kept outside the
repository: `(deleted-coord-artifact)/results.json` (20,687 bytes,
SHA-256 `a59928040e7580c6e6cb908d1e4651565cebd8b361f898fa3090baf881c6763a`).

## 5. Gates

Run on `15a211c150` plus this record (env of `(deleted-coord-artifact)`).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-table -p litchi-keynote -p litchi-pages -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features keynote / pages / numbers] --all-targets -- -D warnings` (4 runs) | 0 |
| `cargo test -p litchi-iwa-table --all-features` (62) | 0 |
| `cargo test -p litchi-pages --all-features` (781) | 0 |
| `cargo test -p litchi-keynote --all-features` (1224) | 0 |
| release `parity` binary, all non-ignored (1190 passed, 210 ignored) | 0 |
| `cargo test -p litchi-ops --all-features` `--lib` (1032), `--doc` (215), 168 `--test` targets in 6 chunks (1510) | 0 |
| `cargo test -p litchi --all-features` (209) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-table -p litchi-keynote -p litchi-pages` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (73 profiles, 778 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| capability/registry goldens, all features and every feature but `encryption` (unchanged, no bless needed) | 0; 0 |
| tool unit tests (1276); `officecli_capability_delta.py --check`; `officecli_corpus.py --check`; `check_feature_matrix.py` | 0 |
| `python3 tools/unified_bench_matrix.py --check` (after regenerating for the new 0428 row) | 0 |

No public API, capability, glossary or OfficeCLI-corpus change.
