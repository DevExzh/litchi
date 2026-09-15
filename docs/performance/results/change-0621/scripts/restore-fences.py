"""Re-adds the observations change 0621 removed, for the negative control."""
import sys, pathlib
direction = sys.argv[1]  # "restore" or "remove"
edits = [
 ("crates/litchi-xls/src/workbook/source.rs",
  "        let expected_version = cfb.source_version().map_err(SourceBackedError::from)?;\n        let file_size",
  "        let expected_version = cfb.source_version().map_err(SourceBackedError::from)?;\n        ensure_current_parts(&source, &cfb, expected_version)?;\n        let file_size"),
 ("crates/litchi-xls/src/workbook/source.rs",
  "    for name in [\"Workbook\", \"Book\"] {",
  "    ensure_current_parts(source, cfb, expected_version)?;\n    for name in [\"Workbook\", \"Book\"] {"),
 ("crates/litchi-xls/src/workbook/source.rs",
  "        check_text_cancellation(execution).map_err(|source| writer.document_error(source))?;",
  "        check_text_state(owner, execution).map_err(|source| writer.document_error(source))?;"),

 ("crates/litchi-xls/src/workbook/source.rs",
  "    }\n\n    let mut first_segment = None;",
  "    }\n    owner.ensure_current()?;\n\n    let mut first_segment = None;"),
 ("crates/litchi-xls/src/workbook/source.rs",
  "        }\n        let source_offset = segment",
  "        }\n        owner.ensure_current()?;\n        let source_offset = segment"),
 ("crates/litchi-xls/src/workbook/source.rs",
  "        Ok(value) => Ok(litchi_core::sheet::CellValue::String(value)),\n        Err(error) => Err(map_shared_string_error(error)),",
  "        Ok(value) => {\n            owner.ensure_current()?;\n            Ok(litchi_core::sheet::CellValue::String(value))\n        },\n        Err(error) => {\n            owner.ensure_current()?;\n            Err(map_shared_string_error(error))\n        },"),
 ("crates/litchi-cfb/src/shared.rs",
  "        let source_length = match source.len() {\n            Ok(length) => length,\n            Err(error) => {\n                Self::refuse_if_changed(source.as_ref(), expected_version)?;\n                return Err(error.into());\n            },\n        };\n        if source_length > limits.max_input_bytes() {\n            Self::refuse_if_changed(source.as_ref(), expected_version)?;",
  "        let source_length = source.len();\n        let observed = source.version()?;\n        if observed != expected_version {\n            return Err(OleError::SourceChanged { expected: expected_version, observed });\n        }\n        let source_length = source_length?;\n        if source_length > limits.max_input_bytes() {"),
]
for path, after, before in edits:
    p = pathlib.Path(path)
    text = p.read_text()
    src, dst = (after, before) if direction == "restore" else (before, after)
    assert text.count(src) == 1, (path, direction, text.count(src), src[:60])
    p.write_text(text.replace(src, dst, 1))
print(f"{direction}: ok")
