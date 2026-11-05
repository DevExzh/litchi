# 0330 — Registry self-consistency and the XLSX cell link bug (audit WP2)

- Status: implemented on `uo2/registry-consistency`, cut from
  `feat/unified-ops-api` at `7d3dc08d9c`; trunk merged at `54b46a334f`
  (records 0326 and 0331 logs).
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17, §1 item 4, §2 and WP2:
  11 element rows whose operation claims contradict the property flags of
  the same registry, and the XLSX cell `link`/`tooltip` warn-skip in the
  OfficeCLI corpus Bugs table.
- Owned: `crates/litchi-ops/src/format/{pptx,docx,xlsb,xlsx}/spec.rs`,
  `tests/registry.rs`, the XLSX cell write path. Small hunks outside them:
  the PPTX claim lists in `format/pptx.rs`, the three XLSB lines of
  `Element::operations` in `capabilities.rs`, the `pptx` feature of
  `litchi-ops/Cargo.toml` (`dep:chrono`), PPTX comment date admission in
  `format/pptx/comments.rs`, and three checklist recipes in
  `tests/crud_matrix.rs`.
- `performance_claim`: the XLSX coalesced cell run is unchanged within
  noise (§5).

## 1. The rule

`Capabilities::of(format)` answers every element's operations from claim
lists (`format::pptx::{SETTABLE, SHAPE_ADDABLE, REMOVABLE}`, the `matches!`
arms of `Element::operations`), while its property flags come from the
registry tables the write handlers admit with. Nothing tied the two, so they
drifted. The new test `element_operations_agree_with_property_flags`
(`crates/litchi-ops/tests/registry.rs`) asserts, for every enabled format and
element:

1. an element claims `set` exactly when at least one property is settable;
2. an element with an addable property claims `add` (an element may claim
   `add` with no addable property: a slide, row or column is added bare);
3. an element that claims `add`, `set` or `remove` claims `get`, and an
   element with an element gap claims no write and declares no writable
   property.

Contextual refusals (one selection, owner state, one property) stay
operation or property gaps and never turn the element claim off. The test is
feature-agnostic and runs in both golden feature sets (`--all-features`, and
every litchi-ops feature except `encryption`). On the base it lists exactly
the audit's rows (the encrypted DOCX row is the same `cell` row in the
`encryption` build):

```
Docx cell: addable ["text"] without an add claim
Pptx picture: addable ["name", "x", "y", "width", "height", "link", "tooltip", "src"] without an add claim
Pptx tc: settable ["fill", "text", "bold", "italic", "size", "color", "font", "align"] without a set claim
Pptx paragraph: settable ["align", "size", "text", "bold", "italic", "color", "font"] without a set claim
Pptx run: settable ["font.latin", "size", "bold", "italic", "color", "link", "tooltip", "text", "font"] without a set claim
Pptx notes: settable ["text"] without a set claim
Pptx comment: settable ["text", "x", "y"] without a set claim
Pptx comment: addable ["text", "author", "initials", "date", "x", "y"] without an add claim
Xlsb namedrange: settable ["ref"] without a set claim
Xlsb namedrange: addable ["name", "ref", "scope"] without an add claim
```

## 2. The fixes

Each row was checked against the handler and its tests; the side that
matched the implementation was kept.

| format | element | before | after | why |
|---|---|---|---|---|
| pptx | `tc`, `paragraph`, `run` | `gq` | `sgq` | `Pptx::set` routes them through `write/text.rs` (native `set_text_properties`, both owners; record 0252) |
| pptx | `notes` | `gq` | `sgq` | `write/notes.rs` sets speaker-notes text |
| pptx | `comment` | `gq` | `asgq` | `comments.rs` adds and sets legacy comments (`tests/pptx_slide_content.rs`); remove is still unclaimed (WP3) |
| pptx | `picture` | `sgqr` | `asgqr` | `picture.rs` adds from `src` (record 0252) |
| xlsb | `namedrange` | `gq` | `asgqr` | `xlsb/name_write.rs` adds, retargets and removes (record 0261); new operation gap `remove namedrange earlier-name` |
| docx | `cell` | `sgqr`, `text` add+set | `sgqr`, `text` set only | a single cell is never added (operation gap `add cell cell-insertion`); `text` is now `Access::Set` |

No behaviour was widened: every newly claimed verb already worked. The new
XLSB gap states the existing refusal ("only the workbook's last defined name
can be removed"). The PPTX additions gate on the same owner conditions as
the existing PPTX claims (`retains_source`, `supports_shape_add`).

One code fix came out of exercising the new cells: a PPTX comment `date`
that is not an XML `dateTime` was accepted by the registry and then failed as
`corrupt_file` ("invalid PresentationML: invalid XML dateTime") when the
comments part serialized. It is now `invalid_value` on `date`, checked with
the leaf codec's own two forms (RFC 3339, or `%Y-%m-%dT%H:%M:%S%.f`) so the
ops check accepts exactly what the codec accepts. Test: the end of
`comments_are_added_and_set`.

## 3. Golden changes

Re-blessed in both feature sets (`LITCHI_BLESS=1 cargo test -p litchi-ops
--test capabilities --test registry` with `--all-features`, and with every
feature except `encryption`). Every hunk is one of the rows of §2:

| golden | change |
|---|---|
| `capabilities/pptx.json`, `registry/pptx.json` | `picture.add`, `comment.add` false → true; `tc`, `paragraph`, `run`, `notes`, `comment` `set` false → true (7 flags each) |
| `capabilities/xlsb.json`, `registry/xlsb.json` | `namedrange` `add`, `set`, `remove` false → true; one new gap `{operation, remove, namedrange, earlier-name}` |
| `capabilities/docx.json`, `registry/docx.json`, and the `-encryption` copies | `cell.properties.text.add` true → false |

No property was added or removed, no type changed, and no other format's
golden moved.

Regenerated reports:

- `officecli-capability-delta.md`: the four golden hashes; DOCX property
  writes supported 250 → 249 and partial 38 → 39 (`cell.text` is now `-s-`
  against OfficeCLI's `asg`, which is honest: Litchi adds rows with their
  cells, not single cells); the new XLSB operation gap row. Element rows are
  unchanged.
- `results/0289/crud-checklist.{md,tsv}`: see §4.
- `officecli-corpus-deltas.md`: see §4.

## 4. CRUD checklist and corpus

The checklist exercises a claimed cell; an unclaimed one is `n/a` and never
runs. With the claims fixed, the new cells ran and came out ok. Three
`add` cells first failed for want of a value the generator cannot invent
(an image, an author, a sheet-qualified reference), so they gained recipes:
PPTX picture (`src` PNG, `name`) and comment (`author`, `text`) on
`ooxml/pptx/shapes.pptx` slide 1, and XLSB namedrange (`name`, `ref
Sheet1!$A$1:$B$2`) on `create()`. The PPTX recipes use a producer deck
rather than `create() + add slide`: two more created decks pushed the
placeholder write instances past the checklist's ten-instance budget onto
layout placeholders, which refuse writes by design.

| format | claimed | ok | FAIL | n/a |
|---|---:|---:|---:|---:|
| pptx | 262 → **279** | 262 → **279** | 0 | 90 → 73 |
| xlsb | 175 → **180** | 175 → **180** | 0 | 45 → 40 |

New ok cells: pptx `tc`/`paragraph`/`run`/`notes`/`comment` set, edit and
batch; pptx `picture` and `comment` add; xlsb `namedrange` set, add, remove,
edit and batch. The XLSB namedrange `set` replays its current `ref` (marked
"value unchanged: no-op only"); `xlsb_p0_writes.rs` covers a changing
retarget. Other changed TSV rows only name a different "first instance"
fixture label (`shapes.pptx + add picture` instead of `+ add col`).

**XLSX cell link bug.** `excel/cell-formatting.sh:158` sets `value`, `link`,
`tooltip`, `underline` and `font.color` on one cell. The corpus harness
passes `--unsupported warn`, so the command runs as a batch item. In a batch
or edit closure, an XLSX cell `set` first meets the coalesced cell-payload
run (`Xlsx::cell_plan`, `format/xlsx/write.rs`), whose admission table does
not know the cell-link keys. It returned `unsupported_property` for `link`
instead of declining, so the warn loop skipped `link`, then `tooltip`, and
applied the rest. The same failure hit the default refuse policy: any batch
or `edit` carrying a cell link failed, while a direct `set` worked through
`split_cell_link`. `cell_plan` now declines (returns `None`) when a property
resolves in `CELL_LINK_TABLE` (keys and aliases, the same lookup
`split_cell_link` uses), as it already did for `merge`. Test:
`cell_links_in_a_batch_are_applied_not_skipped` (`tests/unsupported_warn.rs`;
it fails on the base with `unsupported_property` in both policies).

Corpus (`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, `tools/officecli_corpus.py`, `--check`): line 158 is now
`ok`; excel passed 1552 → 1553, warned 26 → 25; the Bugs table is empty
("None"), agreeing with the summary's 0 bugs. Two PPTX gap notes now read
"registry lists add/set; refused for this target (operation gap)" instead of
"registry: no add/set":

- `ppt/pictures/pictures-basic.sh:97` adds a picture from a data URI that
  the script's Pillow step should fill. The sandboxed `python3 -I -S` has no
  Pillow, so the payload is empty and the add is refused ("the picture is not
  a PNG, JPEG, GIF, BMP or TIFF image"). That is the harness environment,
  not a claim error.
- `ppt/tables/tables-basic.sh:79` sets a cell `fill=accent1`. The PPTX text
  writer takes RGB fills only, so a scheme colour is refused
  `unsupported_capability`. It has no property gap yet (follow-up F1 below).

## 5. Performance

`cell_plan` is on the coalesced batch path, so it gained one extra lookup per
property: a binary search of a 3-entry table plus 2 aliases. Measurement:
the release build of a scratch binary (`litchi-ops` with `xlsx`; not
committed) times one `File::batch` of 1,000 `set /Sheet1/A{i} value=i`
items on `create(Xlsx)` (the coalesced family). It reports the median of 30
rounds after 5 warm-ups, pinned with `taskset -c 3`, in ABBA order × 4:

| variant | mean of 8 medians |
|---|---:|
| before (`cell_plan` without the link check) | 1,961,445 ns |
| after | 1,963,258 ns |

Ratio 1.0009: no measurable change. The raw medians (`abba.txt`), the
scratch crate (`bench-Cargo.toml`, `bench-main.rs`) and the binaries'
SHA-256 (`binaries.sha256`) are in `results/0330/`. Rebuild with
`cargo build --release` in a copy of the scratch crate, once on this record's
tree (after) and once with `format/xlsx/write.rs` from `8deea88d0e` (before). No leaf
path changed, so a unified-vs-leaf pair was not re-run. Nothing else on a
hot path changed: the claim lists are read by `capabilities()` only, and
the comment date check runs once per comment write.

## 6. Gates

On the merged tree (`18604077f2` plus this record's commit), with the build
env of the rules file; every exit status was checked:

| gate | exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --features <all but encryption> --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features docx/pptx/xlsb/xlsx] --all-targets -- -D warnings` (5 runs) | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, then 6 `--test` chunks of 25) | 0 (after updating `xlsb_remove.rs`'s removable list for `namedrange`) |
| `cargo test -p litchi-ops --features <all but encryption> --test capabilities --test registry` | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus`; `tools/officecli_corpus.py --check` | 0; 0 |
| `tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` (no diff after the merge) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (283 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No public API changed, so no glossary row moved.

## 7. Follow-ups

- F1 (WP3): PPTX table cell `fill` with a scheme colour (`accent1`) is
  refused without a property gap; either accept scheme colours or declare the
  gap.
- F2 (WP3): PPTX comment remove, paragraph/run/tr/tc add and remove remain
  unclaimed and unimplemented, as the audit lists.
- F3: XLSB defined-name removal other than the last needs name-index
  renumbering in formulas (gap `remove namedrange earlier-name`).
- F4: the checklist's XLSB namedrange `set` cell replays the current `ref`;
  a `Text` generator for sheet-qualified references would make it a changing
  set.
