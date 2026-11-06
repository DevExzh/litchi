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

#[test]
fn native_rtf_math_opens_through_the_facade_and_unmodeled_grouping_is_typed() {
    let file = native("word/word-math-navigation.rtf", FileFormat::Rtf);
    assert!(file.text().unwrap().contains("MATH_NAVIGATION_END_005"));
    assert_eq!(
        file.to_bytes().unwrap().as_slice(),
        include_bytes!("../../../test-data/microsoft-office/word/word-math-navigation.rtf")
    );
    let error = litchi::open_bytes(
        br"{\rtf1{\mmath{\*\moMathPara{\*\moMath{\mr x}}{\*\moMath{\mr y}}}}}".to_vec(),
    )
    .unwrap_err();
    assert_eq!(error.code(), litchi::Code::UnsupportedCapability);
    assert_eq!(error.format(), Some(FileFormat::Rtf));
    assert!(
        std::error::Error::source(&error)
            .unwrap()
            .to_string()
            .contains("multiple equations in one math paragraph")
    );
}

#[test]
fn native_ppt_table_query_edit_reopen_and_inverse_preserve_neighbor_cells() {
    let file = native("powerpoint/powerpoint-text-table.ppt", FileFormat::Ppt);
    let original = file.to_bytes().unwrap();
    let cells = file
        .query("/slide[3]/table[1] tc")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap();
    assert_eq!(cells.len(), 12);
    for cell in cells {
        assert_eq!(file.get(cell.path().to_path()).unwrap().kind(), cell.kind());
    }
    let text = |file: &File, path: &str| {
        file.get(path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .unwrap()
            .into_owned()
    };
    assert_eq!(text(&file, "/slide[3]/table[1]/tr[2]/tc[1]"), "Apples");
    let commit = file
        .edit(|edit| {
            edit.set(
                "/slide[3]/table[1]/tr[2]/tc[1]",
                litchi::props! {"text"=>"Oranges"},
            )
        })
        .unwrap();
    let reopened = litchi::open_bytes(commit.file().to_bytes().unwrap()).unwrap();
    assert_eq!(text(&reopened, "/slide[3]/table[1]/tr[2]/tc[1]"), "Oranges");
    assert_eq!(text(&reopened, "/slide[3]/table[1]/tr[2]/tc[2]"), "3");
    assert_eq!(text(&reopened, "/slide[3]/table[1]/tr[2]/tc[3]"), "3.75");
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        original
    );
}

#[test]
fn native_pptx_template_show_and_macro_variants_preserve_inert_members_on_public_edits() {
    fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
        let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
            bytes.to_vec(),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        reader
            .member_names()
            .unwrap()
            .into_iter()
            .map(|name| {
                let bytes = reader.read_member(&name).unwrap();
                (name, bytes)
            })
            .collect()
    }
    for name in [
        "powerpoint-visual-objects.potx",
        "powerpoint-visual-objects.ppsx",
        "powerpoint-passive-macro.pptm",
        "powerpoint-passive-macro.potm",
        "powerpoint-passive-macro.ppsm",
    ] {
        let file = native(&format!("powerpoint/{name}"), FileFormat::Pptx);
        let before = file.to_bytes().unwrap();
        let title = file
            .get("/slide[1]/shape[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .unwrap()
            .into_owned();
        assert_eq!(
            title,
            if name.starts_with("powerpoint-visual") {
                "Native PowerPoint visual objects"
            } else {
                "PPT_MACRO_005"
            }
        );
        let commit = file
            .edit(|edit| {
                edit.set(
                    "/slide[1]/shape[1]",
                    litchi::props! {"text"=>"Native variant edited title"},
                )
            })
            .unwrap();
        let after = commit.file().to_bytes().unwrap();
        let reopened = litchi::open_bytes(after.clone()).unwrap();
        assert_eq!(reopened.format(), Some(FileFormat::Pptx));
        assert_eq!(
            reopened
                .get("/slide[1]/shape[1]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Native variant edited title")
        );
        let a = members(&before);
        let b = members(&after);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        for (part, bytes) in a {
            if part != "ppt/slides/slide1.xml" {
                assert!(b[&part] == bytes, "{name}: changed {part}");
            }
        }
        assert!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == before
        );
    }
}

#[test]
fn native_powerpoint_embedded_chart_workbook_exposes_exact_cells_through_facade() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    let parent = litchi::pptx::Package::open(root.join("powerpoint-visual-objects.pptx")).unwrap();
    let part = parent
        .opc()
        .unwrap()
        .get_part(
            &litchi::opc::PackURI::new("/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx").unwrap(),
        )
        .unwrap();
    let child = litchi::open_bytes(part.blob().to_vec()).unwrap();
    assert_eq!(child.format(), Some(FileFormat::Xlsx));
    assert_eq!(
        child.workbook().unwrap().worksheet_names().unwrap(),
        ["Sheet1"]
    );
    let expected = [
        ["Quarter", "North", "South", "West"],
        ["Q1", "10", "15", "8"],
        ["Q2", "20", "18", "12"],
        ["Q3", "30", "21", "16"],
        ["Q4", "40", "24", "20"],
    ];
    for (row, values) in expected.into_iter().enumerate() {
        for (column, value) in ["A", "B", "C", "D"].into_iter().zip(values) {
            let path = format!("Sheet1!{column}{}", row + 1);
            assert_eq!(
                child
                    .get(path.as_str())
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some(value),
                "{path}"
            );
        }
    }
    assert_eq!(child.to_bytes().unwrap(), part.blob());
}
