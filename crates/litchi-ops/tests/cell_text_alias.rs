//! `text` names a workbook cell's `value` (record 0269): OfficeCLI's XLSX
//! cell handler accepts it (`case "value" or "text"`), so the corpus scripts
//! use it on every data row. XLSB and XLS take the same alias; ODS has no
//! unified cell writer yet. The alias resolves to the same registry entry,
//! so it inherits the `value` writer's no-op, signature, encryption and
//! read-back behaviour; these tests prove the two spellings write the same
//! bytes and read back the same text.

#![cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls"))]

use litchi_core::FileFormat;
use litchi_ops::{File, create, props};

fn cell_text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("cell")
        .view()
        .text()
        .expect("text")
        .map(std::borrow::Cow::into_owned)
}

/// `text` and `value` produce byte-identical files; the text reads back
/// before and after a byte round trip; repeating the write is a no-op.
fn alias_matches_value(format: FileFormat, path: &str) {
    let file = create(format).expect("create");
    let by_value = file
        .set(path, props! { "value" => "Month" })
        .expect("set value");
    let commit = file
        .edit(|edit| edit.set(path, props! { "text" => "Month" }))
        .expect("set text");
    assert!(commit.diagnostics().changed());
    let by_text = commit.file();
    assert_eq!(
        by_text.to_bytes().expect("text bytes"),
        by_value.to_bytes().expect("value bytes"),
        "{format:?}: text and value write the same bytes"
    );
    assert_eq!(cell_text(by_text, path).as_deref(), Some("Month"));
    let reopened = litchi_ops::open_bytes(by_text.to_bytes().expect("bytes")).expect("reopen");
    assert_eq!(cell_text(&reopened, path).as_deref(), Some("Month"));

    let again = by_text
        .edit(|edit| edit.set(path, props! { "text" => "Month" }))
        .expect("repeat");
    assert!(!again.diagnostics().changed(), "{format:?}: exact no-op");

    // The schema export lists the alias on the canonical property.
    let schema = litchi_ops::schema(format);
    let cell = schema.element("cell").expect("cell element");
    let value = cell.property("value").expect("value property");
    assert!(
        value.aliases().any(|alias| alias == "text"),
        "{format:?}: value lists its text alias"
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_cell_text_is_value() {
    alias_matches_value(FileFormat::Xlsx, "/Sheet1/A1");
    // OfficeCLI documents `bold` and `italic` as aliases of the font flags.
    let file = create(FileFormat::Xlsx).expect("create");
    let aliased = file
        .set(
            "/Sheet1/A1",
            props! { "text" => "Month", "bold" => true, "italic" => true },
        )
        .expect("aliases");
    let canonical = file
        .set(
            "/Sheet1/A1",
            props! { "value" => "Month", "font.bold" => true, "font.italic" => true },
        )
        .expect("canonical");
    assert_eq!(
        aliased.to_bytes().expect("aliased bytes"),
        canonical.to_bytes().expect("canonical bytes")
    );
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_cell_text_is_value() {
    alias_matches_value(FileFormat::Xlsb, "/sheet[1]/A1");
}

#[cfg(feature = "xls")]
#[test]
fn xls_cell_text_is_value() {
    alias_matches_value(FileFormat::Xls, "/sheet[1]/A1");
}
