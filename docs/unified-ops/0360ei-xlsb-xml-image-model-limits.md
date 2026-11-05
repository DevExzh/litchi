# 0360ei — XLSB XML parts, image-limits carrier and data-model payload admission

- Status: source and consumer migration complete on `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open. This slice closes the 0360e audit.
- Base: `f393f11f9c` (the 0360eh host-codec slice).
- Scope: `crates/litchi-xlsb` SpreadsheetDrawing, timeline and threaded-comment XML codecs and their loaders, editors and writers; `package::drawing_image`; `data_model::{ModelPart, Model}::from_bytes`; the crate `SemanticLimits`; tests and this record.

## XML parts: `SemanticLimits::xml_parts`

The three XML part families had only module constants. They now take a required `litchi_core::Limits` (`xml_limits`), the existing hierarchical budget type, whose registered dimensions apply as follows. Every family keeps its format ceiling; the effective bound is the smaller of the two (`semantic_limits::xml_ceiling`).

| Family | Input bytes | Output bytes | Objects | Depth |
|---|---|---|---|---|
| `package::drawing::parse_drawing_part(xml, xml_limits)` | ✔ before UTF-8 decoding | — | — | — |
| `timeline::{parse_cache, parse_views}(xml, xml_limits)` | ✔ | — | ✔ elements | ✔ |
| `timeline::{write_cache, write_views}(value, xml_limits)` | — | ✔ | — | — |
| `comments::threaded::{parse_persons, parse_comments}(xml, xml_limits)` | ✔ | — | — | — |
| `comments::threaded::{write_persons, write_comments}(value, xml_limits)` | — | ✔ | — | — |

`timeline::package::{load_caches, load_views}` take it too. The owner policy is the new `SemanticLimits::xml_parts(litchi_core::Limits)` / `xml_parts_value()` (default `litchi_core::Limits::for_profile(Profile::Server)`), carried by `OpenOptions::semantic_limits`. The opening loader and `Workbook` accessors pass `semantic_limits.xml_parts_value()`; the timeline and threaded-comment editors pass the staged workbook's value (bound before each staging closure); drawing transfer passes the target workbook's; the `Builder` uses the default policy. Refusals keep the families' existing error values.

## Image-limits carrier

`package::drawing_image::Limits { image_bytes, description_bytes }` (default: the format ceilings `MAX_XLSB_WORKSHEET_IMAGE_BYTES`, `MAX_XLSB_IMAGE_DESCRIPTION_BYTES`; fallible `max_image_bytes`, `max_description_bytes` setters refuse a value above the ceiling with `PackageError::InvalidLimit`) and `drawing_image::LimitKind::{ImageBytes, DescriptionBytes}`, wrapped as `package::LimitKind::Image`.

`Image::new(data, format, anchor, &Limits)` now takes `D: AsRef<[u8]> + Into<Arc<[u8]>>`, so the encoded length is admitted (`PackageError::Limit`) **before** the bytes are converted into shared ownership. The image retains its limits; `with_description` checks the retained description bound first. `Image::limits()` exposes them. Worksheet-image transfer between workbooks re-admits under `Limits::DEFAULT` (the bytes were admitted when the source opened).

## Data-model payload

`data_model::ModelPart::from_bytes(data, data_model::SemanticLimits)` and `Model::from_bytes(definition, data, data_model::SemanticLimits)` validate the profile and admit the payload against `max_part_bytes` (`PackageError::Limit` with `LimitKind::DataModel(PayloadBytes)`) before it is retained.

## Deliberate classifications

- `cell_values::worksheet::Patch::{apply, undo, redo}` and `web_extension_bindings::Patch::apply_bytes` compare caller bytes with a retained, already-admitted image and copy it; they decode nothing, so they take no profile.
- `external_link::NameFormula::from_tokens` validates one fixed-grammar external-name Ptg; the external-link family already carries its own `external_link::Limits` on every stream parser.
- `serialize_rule_extension_guid` (0360eg) is a constant-width constructor over a typed GUID.

## Tests

`timeline::tests::xml_limits_bound_timeline_input_elements_depth_and_output` (exact and one-under input, element and depth budgets; exact and one-under output), `package::drawing::tests::xml_limits_bound_drawing_input_bytes`, `package::drawing_image::tests::image_limits_admit_bytes_before_ownership_and_bound_descriptions` (exact and one-under image bytes, invalid ceiling, exact and one-over description bytes) and `tests/data_model.rs::model_payload_is_admitted_under_caller_limits`.

## Validation

Commands were run with `(deleted-coord-artifact)`:

- `cargo fmt --all -- --check`: exit 0.
- `cargo test -p litchi-xlsb --all-features --offline`: exit 0; 1,342 passed, 0 failed.
- `cargo clippy -p litchi-xlsb --all-targets --all-features --offline -- -D warnings`: exit 0.
- `cargo check -p litchi-ops -p litchi -p litchi-cli --all-targets --all-features --offline`: exit 0.

Logs under `(deleted-coord-artifact)`: `xlsb-0360ei-tests.log` (122,046 bytes, SHA-256 `a5c78d047e67eec35385632d231f1b07d681cc1f3923bbc7b4d2b3e384a99544`) and `xlsb-0360ei-clippy.log` (171 bytes, SHA-256 `a4e0f7d0240f4b75a04c3b18c1b9ef2d1ad227a6631fd4f77a153650d6bbee72`).

## 0360e closure

With 0360ea–ei every surface in the 0360e audit takes a required profile: comments and calculation (ea); styles, alignment and border parsers (eb, re-verified here: `Alignment::parse(data, offset, raw::Limits)`, `Border::parse(data, raw::Limits)`); auto-filter, page breaks and cell watches (ec); named ranges (ed); merged cells and data validation (ee); hyperlinks, pivot view and binary index (ef); formula and conditional formatting including `conditional_formatting/edit.rs` (eg); host part codecs, slicer and shared workbook (eh); XML parts, the image-limits carrier and the data-model payload (ei). Glossary registration of the new `xml_parts`, `image_bytes` and `description_bytes` dimensions is part of the W4 glossary rows.

## Performance

Each XML parse gains up to three `min` computations; image construction gains one length comparison before ownership. No allocation is added. Native timing is pending with the rest of W4b.
