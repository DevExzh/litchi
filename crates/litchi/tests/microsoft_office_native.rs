//! Public entry-point regressions using unmodified native Microsoft Office files.
//! The facade-format-tests gate explicitly enables every feature below.
#![cfg(all(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ppt",
    feature = "pptx"
))]

use litchi::{File, FileFormat};

fn native(path: &str, expected: FileFormat) -> File {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office")
        .join(path);
    let file = litchi::open(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()));
    assert_eq!(file.format(), Some(expected));
    file
}

fn word(file: &File) {
    let text = file.text().unwrap();
    for marker in [
        "Native Word text and structure",
        "WORD_NATIVE_END",
        "Unicode: café naïve résumé — Ελληνικά — 中文 — 日本語 — العربية — 😀",
    ] {
        assert!(text.contains(marker), "missing {marker}");
    }
    let paragraphs = file
        .query("p")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap();
    assert!(!paragraphs.is_empty());
    for paragraph in paragraphs {
        assert_eq!(
            file.get(paragraph.path().to_path()).unwrap().kind(),
            paragraph.kind()
        );
    }
    assert!(file.document().unwrap().paragraphs().unwrap().len() >= 12);
}

#[test]
fn native_doc_reads_unicode_paragraphs_through_public_views() {
    word(&native("word/word-text-structure.doc", FileFormat::Doc));
}

#[test]
fn native_docx_reads_unicode_paragraphs_through_public_views() {
    word(&native("word/word-text-structure.docx", FileFormat::Docx));
}

fn workbook(file: &File) {
    assert_eq!(
        file.workbook().unwrap().worksheet_names().unwrap(),
        [
            "Sheet1",
            "Source_A_015",
            "Source_B_015",
            "Consolidated_015",
            "Print_015"
        ]
    );
    let text = file.text().unwrap();
    for marker in [
        "EXCEL_METADATA_015",
        "RICH_TEXT_015 bold italic",
        "日本語_015 東京",
        "CONSOLIDATION_015",
        "END_PRINT_015",
    ] {
        assert!(text.contains(marker), "missing {marker}");
    }
    let cells = file
        .query("cell")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap_or_else(|error| panic!("{error}: {:?}", std::error::Error::source(&error)));
    assert!(!cells.is_empty());
    for cell in cells {
        assert_eq!(file.get(cell.path().to_path()).unwrap().path(), cell.path());
    }
}

#[test]
fn native_xls_keeps_all_five_metadata_worksheets() {
    workbook(&native("excel/excel-metadata.xls", FileFormat::Xls));
}

#[test]
fn native_xlsx_keeps_all_five_metadata_worksheets() {
    workbook(&native("excel/excel-metadata.xlsx", FileFormat::Xlsx));
}

#[test]
fn native_xlsb_keeps_all_five_metadata_worksheets() {
    workbook(&native("excel/excel-metadata.xlsb", FileFormat::Xlsb));
}

fn presentation(file: &File) {
    let view = file.presentation().unwrap();
    assert_eq!(view.slide_count().unwrap(), 3);
    let slides = view.slides().unwrap();
    assert_eq!(slides.len(), 3);
    assert!(
        slides[0]
            .text()
            .unwrap()
            .contains("powerpoint-text-table-001")
    );
    assert!(slides[1].text().unwrap().contains("Third bullet: tea"));
    let table = slides[2].text().unwrap();
    for marker in ["Apples", "3.75", "Pears", "12.50", "Tea", "9.50"] {
        assert!(table.contains(marker), "missing {marker}");
    }
}

#[test]
fn native_ppt_reads_three_slides_and_table_text() {
    presentation(&native(
        "powerpoint/powerpoint-text-table.ppt",
        FileFormat::Ppt,
    ));
}

#[test]
fn native_pptx_reads_three_slides_and_table_text() {
    presentation(&native(
        "powerpoint/powerpoint-text-table.pptx",
        FileFormat::Pptx,
    ));
}

#[test]
fn native_rtf_edit_reopen_and_inverse_keep_the_original_word_source() {
    let source = native("word/word-rtf-minimal.rtf", FileFormat::Rtf);
    assert_eq!(
        source.text().unwrap().trim(),
        "RTF_MINIMAL_006: plain native Microsoft Word paragraph."
    );
    let before = source.to_bytes().unwrap();
    let refusal = source
        .set(
            "/body/p[1]",
            litchi::props! {"text" => "replace mixed formatting"},
        )
        .unwrap_err();
    assert_eq!(refusal.code(), litchi::Code::UnsupportedCapability);
    assert!(
        std::error::Error::source(&refusal)
            .unwrap()
            .to_string()
            .contains("mixed run or paragraph formatting")
    );
    assert_eq!(source.to_bytes().unwrap(), before);
    let commit = source
        .edit(|edit| {
            edit.set(
                "/body/p[1]/r[1]",
                litchi::props! {"text" => "Native Word edited through File"},
            )
        })
        .unwrap_or_else(|error| panic!("{error:?}: {:?}", std::error::Error::source(&error)));
    let reopened = litchi::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.text().unwrap().trim(),
        "Native Word edited through File"
    );
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    let File::Rtf(rtf) = reopened else {
        panic!("expected RTF");
    };
    assert_eq!(
        rtf.text().unwrap().trim(),
        "Native Word edited through File"
    );
}
