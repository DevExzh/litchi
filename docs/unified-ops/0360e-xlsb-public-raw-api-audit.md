# 0360e — XLSB public raw and record API audit

Status: source audit for the W4b follow-up. This record does not claim the
surfaces below are migrated. It records the public boundaries that still need
an atomic limits migration after the W4b integration candidate.

## Rule applied

The raw XLSB kernel already requires a `raw::Limits` argument on
`Cursor`, `Records`, `Header`, and `Encoder`. The remaining question is the
public API immediately above that kernel. Under 0351 §2.5, a public byte or
record parser, serializer, or detached raw constructor must receive a caller
policy. A fixed-width wire field still needs the required policy argument in
the final API; its fixed width is a proof that no additional variable-resource
budget is needed. It is not a reason to retain a hidden default. The existing
MS-XLSB `MAX_*` constants remain format invariants in addition to caller
limits.

## Confirmed public ingress and output gaps

These entry points either use `raw::Limits::DEFAULT` internally or have no
caller policy at all. Their package-bound callers must pass the retained owner
policy in the same change.

### Standalone BIFF12 and model codecs

| Public surface | Why it is a gap | Main callers to migrate |
| --- | --- | --- |
| `comments::read(bytes)` (`src/comments/codec.rs:79`) and the package-facing `read_recording` path (`:357`) | Parses records, authors, text and runs; `Records`/`Cursor` use `DEFAULT`. | `workbook/source.rs:1889`, `workbook/codec/codec.rs:108`, comments tests. |
| `calc::read(payload)` and `read_with` (`src/calc/codec.rs:22,31`) | Fixed 25/26-byte payload, but the detached record parser still chooses a hidden raw policy. This is policy plumbing rather than a new allocation budget. | `workbook/codec/records.rs:167`, `tests/calc.rs`, reserved-layout tests. |
| `styles::{read,parse_font,parse_num_fmt,parse_fill,parse_direct_color,parse_border,parse_cell_format}` (`src/styles/codec.rs:47,150,190,224,275,320,371`) | `read`, `parse_font`, and `parse_num_fmt` allocate or decode variable strings under `DEFAULT`; the other four are fixed-width but still public byte ingress. | `package/styles_table.rs:63–145`, `cell_values/resources.rs:327,822,825`, style tolerance tests. |
| `auto_filter::read(worksheet)` (`src/auto_filter/mod.rs:20`) and `page_breaks::read(worksheet)` (`src/page_breaks/mod.rs:81`) | Scan an entire worksheet and allocate feature state; the internal record walk uses `DEFAULT`. | `sheet/features.rs:89,96` and feature tests. |
| `cell_watches::worksheet::read(data)` (`src/cell_watches/worksheet.rs:249`) | Detached worksheet stream parser retains typed and opaque records with no caller limits. (`Snapshot::read` is a separate, already-limits-aware path.) | Workbook/package watch loader and worksheet tests. |
| `WorksheetBinaryIndex::from_parts_default` (`src/binary_index.rs:438`) | Default convenience around the explicit `from_parts(..., Limits)` path silently selects a raw policy. | All `from_parts_default` consumers, if any, and API tests. |
| `pivot_view::Part::from_bytes(Vec<u8>)` (`src/pivot_view/codec.rs:25`) | Detached BIFF12 part owns caller bytes without an admission policy. | Pivot-view tests and package/ops adapters. |
| `named_ranges::{parse,Definition::parse}` (`src/named_ranges/codec.rs:34,162`) | Parses names and formula bytes; `entries.rs:44` uses `DEFAULT` while constructing variable definitions. | `cell_values/formula_transfer.rs:191`, `cell_values/structure.rs:4638`, `workbook/source.rs:1348`, `workbook/model.rs:199`. |
| `formula::Parser::{new,with_extra,with_base_cell,with_base_cell_and_extra}` and `ParsedFormula`/`Group` wire parse and output methods (`src/formula/codec/{parser.rs,wire.rs}`) | Token streams, strings and output vectors are variable-size. `MAX_CELL_FORMULA_BYTES` is only the format ceiling; there is no caller raw/output policy. | `package/formula` host consumers, named ranges, cell-value transfer, conditional-format and formula tests. |
| `data_validation::{parse_dval_list,parse_collection_settings}` (`src/data_validation/codec.rs:215,262`) | `parse_dval_list` allocates variable text; collection settings are fixed-width but still public byte ingress. | `package/cells_reader/codec.rs`, writer/data-validation tests, data-validation adapters. |
| `Hyperlink::{parse,try_serialize,serialize}` (`src/hyperlinks/model.rs:138–151`) | Variable wide-string fields are parsed and serialized with no caller input/output policy. | Hyperlink package readers, editor, and hyperlink tests. |
| `slicer::{parse_cache,parse_views,write_cache,write_views}` (`src/slicer/codec.rs:215,393,429,472`) | Variable record collections and strings use default raw readers/writers. | `slicer/package.rs` loaders/stores, transaction adapters and slicer tests. |
| `timeline::{parse_cache,parse_views,write_cache,write_views}` (`src/timeline/codec.rs:308,412,448,521`) | XML node/string collections are bounded only by module constants, not caller policy. | `timeline/package.rs` loaders/stores, transactions and slicer/timeline tests. |
| `shared_workbook::codec::{parse_users,parse_headers,parse_log,write_users,write_headers,write_log}` (`src/shared_workbook/codec.rs:24,90,148,188,234,244`) | Public metadata and opaque-record codecs allocate variable collections and use defaults internally. | `shared_workbook/package.rs`, transaction paths and shared-workbook tests. |
| `comments::threaded::codec::{parse_persons,parse_comments,write_persons,write_comments}` (`src/comments/threaded/codec/mod.rs:32,109,205,223`) | Public XML codecs have hard caps but no caller policy. | `comments/threaded/package` loaders/stores and threaded-comment tests. |

Conditional-formatting has the same split: public parse/serialize methods on
the re-exported model types in
`src/conditional_formatting/codec/binary/records.rs` and
`src/conditional_formatting/edit.rs:151` must take the required policy when
they parse or emit variable extension records. The fixed header readers
(`parse_header`, `parse_classic_header`, and the GUID parser) remain constant
work, but still receive the required policy in the canonical API.

The public `package::styles::{alignment_parser,border_parser}` submodules
also expose `Alignment::parse` and `Border::parse`
(`host/styles/alignment_parser.rs:97`, `host/styles/border_parser.rs:80`).
Their current payload slices are fixed-width wire fields, but these are still
public byte parsers and must follow the required policy signature. The same
applies to the public `MergedCell` codec below.

### Package-public host parsers and writers

The following modules are `pub mod` from `package/mod.rs`, so their public
byte APIs are externally reachable even though they are host adapters:

* `package::chartsheet::parse_chart_sheet_part` (`host/chartsheet/parse.rs:60`);
* `package::connections::parse_connections_part` (`host/connections/parse.rs:86`);
* `package::pivot::{parse_pivot_cache_definition,parse_pivot_cache_records}`
  (`host/pivot/semantic.rs:15,40`);
* `package::table::{parse_table_part,parse_table_part_rel_ids}`
  (`host/table/parse.rs:66,101`);
* `package::scenarios::{parse_worksheet,replace_worksheet}`
  (`host/scenarios/package.rs:128,142`);
* `package::sheet_view::{parse_view,parse_pane,parse_selection,
  write_view_payload,write_pane_payload,write_selection_payload}`
  (`host/sheet_view.rs:124–337`);
* `package::styles_table::StylesTable::from_bytes`
  (`host/styles_table.rs:62`); and
* `package::web_extension_bindings::{parse_xlsb_web_extension_bindings,
  write_xlsb_web_extension_bindings,read,apply}`
  (`host/web_extension_bindings.rs:112,170,597,605`).

The first five scan or allocate variable package data and need raw plus
appropriate family budgets. `sheet_view` has fixed payload readers but
fallible bounded output still needs the output policy. `StylesTable` must
forward the owner policy to all variable style readers. Web-extension
bindings parse application-reference strings, formulas, binding collections,
and retained source bytes; both its input and output paths currently use
`DEFAULT` and need one coherent limits carrier.

Two additional public package boundaries are raw byte owners rather than
record scanners: `package::drawing::parse_drawing_part` (`host/drawing.rs:241`)
and `data_model::{ModelPart::from_bytes,Model::from_bytes}`
(`src/data_model/model.rs:138,177`). They currently rely on module constants
or validation after an infallible `Arc`/`Vec` ownership step. The migration
must admit the input before publication and retain the owner policy; it must
not replace these capabilities with an unbounded copy.

`package::drawing_image::Image::new` (`host/drawing_image.rs:164`) is the
corresponding raw image authoring boundary. It validates against module-wide
constants after converting the caller's bytes into an `Arc`; a required image
limits carrier must pre-admit the bytes before ownership. The image signature
check itself remains a fixed format invariant.

## Deliberate classifications

* `merged_cells::MergedCell::{parse,serialize}` is a public fixed 16-byte
  `BrtMergeCell` codec. It has no variable allocation, but its final signature
  still follows the required raw-policy shape; the policy carries no extra
  merged-cell budget beyond the fixed-width check.
* `package::shared_strings::SharedString::parse`,
  `package::connections::codec::parse_source`, and host `records`,
  `pivot_tables`, `walker`, and `cells_reader` helpers are crate-private
  implementation paths. Their defaults still need internal policy threading
  when reached by a public owner, but they are not separate external APIs.
* `package::Package::from_bytes` is an owner facade. Its default behavior is
  governed by `OpenOptions`, so it is not a detached raw-codec exception.
* Already-compliant examples include `binary_index::{parse,bind,from_parts}`,
  `external_link` parsers, `sparkline::worksheet::read`, `xml_maps` parsers,
  and the `raw` kernel. Their explicit limits are the pattern to preserve.

## Atomic migration and validation scope

Each family needs a required limits argument or a focused limits carrier,
with no default or compatibility twin. Package and owner callers must derive
that carrier from their retained `OpenOptions`/owner policy; tests, fuzzers,
examples, ops adapters, and docs move in the same change. Keep fixed wire
checks and format `MAX_*` constants, and add boundary tests proving a caller
policy tighter than the legacy default is enforced for every allocating or
scanning family. Add output-boundary tests for every public serializer and
preserve exact no-op/unknown-record bytes. The review is source-only; no
qualification or performance result is claimed by this record.
