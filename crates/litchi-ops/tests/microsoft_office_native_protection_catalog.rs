//! Native protection policy and complete non-worksheet tab catalogs.
#![cfg(all(feature = "doc", feature = "xls", feature = "xlsx", feature = "xlsb"))]
use litchi_ops::{Code, OpenOptions, SourceMode, props};
fn path(name: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office")
        .join(name)
}
#[test]
fn native_doc_protection_refuses_public_edits_and_preserves_readable_comments() {
    let bytes = std::fs::read(path("word/word-document-protection.doc")).unwrap();
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let p = litchi_ops::open_with(
            path("word/word-document-protection.doc"),
            &OpenOptions::new().source_mode(mode),
        )
        .unwrap();
        assert!(p.text().unwrap().contains("PROTECTION_COMMENT_008"));
        let error = p
            .set("/body/p[1]", props! {"text"=>"EDIT_PROTECTION_008"})
            .unwrap_err();
        assert_eq!(error.code(), Code::DocumentProtected);
        assert_eq!(
            p.get("/body/p[1]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("WORD_PROTECTION_008")
        );
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
#[test]
fn native_macro_tabs_remain_queryable_while_cell_access_refuses_explicitly() {
    for ext in ["xls", "xlsm", "xlsb"] {
        let file = path(&format!("excel/excel-legacy-sheets.{ext}"));
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let p = litchi_ops::open_with(&file, &OpenOptions::new().source_mode(mode)).unwrap();
            let text = p.text().unwrap();
            assert!(text.contains("EXCEL_LEGACY_SHEETS_016"));
            // The text projection covers ordinary sheets; the macro body is
            // separately proven present by the owner-level native tests.
            assert!(!text.contains("MACROSHEET_MARKER_016"));
            let tabs = p
                .query("sheet")
                .unwrap()
                .collect::<litchi_ops::Result<Vec<_>>>()
                .unwrap();
            assert_eq!(tabs.len(), 2, "{ext} {mode:?}");
            let macros = p
                .query("sheet[name=Macro_016]")
                .unwrap()
                .collect::<litchi_ops::Result<Vec<_>>>()
                .unwrap();
            assert_eq!(macros.len(), 1);
            assert_eq!(
                p.get(macros[0].path().to_path()).unwrap().path(),
                macros[0].path()
            );
            assert_eq!(
                p.get("Sheet1!A1")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("EXCEL_LEGACY_SHEETS_016")
            );
            let error = p.get("Macro_016!A1").unwrap_err();
            assert_eq!(
                error.code(),
                if ext == "xls" {
                    Code::UnsupportedCapability
                } else {
                    Code::UnsupportedType
                }
            );
        }
    }
}
