//! Public facade behavior for native protection and inert macro-sheet tabs.
#![cfg(all(feature = "doc", feature = "xls", feature = "xlsx", feature = "xlsb"))]
use litchi::{Code, FileFormat, OpenOptions, SourceMode, props};
fn path(name: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office")
        .join(name)
}
#[test]
fn native_protected_word_reads_but_rejects_facade_body_mutation() {
    let file = path("word/word-document-protection.doc");
    let bytes = std::fs::read(&file).unwrap();
    let p = litchi::open(&file).unwrap();
    assert_eq!(p.format(), Some(FileFormat::Doc));
    assert!(p.text().unwrap().contains("PROTECTION_COMMENT_008"));
    assert_eq!(
        p.set("/body/p[1]", props! {"text"=>"EDIT_PROTECTION_008"})
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_macro_tabs_survive_facade_queries_and_ordinary_cell_edits() {
    for ext in ["xls", "xlsm", "xlsb"] {
        let file = path(&format!("excel/excel-legacy-sheets.{ext}"));
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let p = litchi::open_with(&file, &OpenOptions::new().source_mode(mode)).unwrap();
            assert_eq!(
                p.query("sheet[name=Macro_016]")
                    .unwrap()
                    .collect::<litchi::Result<Vec<_>>>()
                    .unwrap()
                    .len(),
                1
            );
            let changed = p
                .set("Sheet1!D4", props! {"value"=>"Reviewed ordinary worksheet"})
                .unwrap();
            let reopened = litchi::open_bytes(changed.to_bytes().unwrap()).unwrap();
            assert!(
                reopened
                    .text()
                    .unwrap()
                    .contains("Reviewed ordinary worksheet")
            );
            assert!(!reopened.text().unwrap().contains("MACROSHEET_MARKER_016"));
            assert_eq!(
                reopened
                    .get("Sheet1!D4")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("Reviewed ordinary worksheet")
            );
            assert_eq!(
                reopened
                    .query("sheet[name=Macro_016]")
                    .unwrap()
                    .collect::<litchi::Result<Vec<_>>>()
                    .unwrap()
                    .len(),
                1
            );
            assert_eq!(
                reopened.get("Macro_016!A1").unwrap_err().code(),
                if ext == "xls" {
                    Code::UnsupportedCapability
                } else {
                    Code::UnsupportedType
                }
            );
            assert!(p.to_bytes().unwrap() == std::fs::read(&file).unwrap());
        }
    }
}
