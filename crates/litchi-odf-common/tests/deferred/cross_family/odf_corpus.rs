//! Real-world corpus compatibility: sample documents copied from the odfpy
//! and odfdo test suites (see `test-data/odf/corpus/`).
//!
//! Every file must open through the high-level API and survive the low-level
//! parse paths without a panic, with sane text and metadata extraction. The
//! one deliberately hostile sample (an XXE probe) must be rejected with
//! typed errors — the external entity is never resolved.

use litchi_odf::{Document, Presentation, Spreadsheet};
use std::path::{Path, PathBuf};

fn corpus() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/corpus")
}

fn assert_typed(error: litchi_core::Error, file: &str, api: &str) {
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(_)),
        "{file}: {api} produced a non-typed error: {error:?}"
    );
}

#[test]
fn corpus_text_documents_parse_and_extract() {
    // (file, minimum paragraph count, minimum table count)
    let cases: &[(&str, usize, usize)] = &[
        ("writer-table.odt", 5, 1),
        ("writer-header-footer.odt", 1, 0),
        ("writer-user-fields.odt", 4, 0),
        ("writer-images-frames.odt", 7, 0),
        ("writer-definition-lists.odt", 14, 0),
        ("writer-paragraph-styles.odt", 10, 0),
        ("writer-table-of-contents.odt", 9, 0),
    ];
    for &(file, min_paragraphs, min_tables) in cases {
        let path = corpus().join(file);
        let document =
            Document::open(&path).unwrap_or_else(|error| panic!("{file}: open failed: {error:?}"));

        let text = document
            .text()
            .unwrap_or_else(|error| panic!("{file}: text() failed: {error:?}"));
        assert!(!text.trim().is_empty(), "{file}: empty text extraction");

        let paragraphs = document
            .paragraphs()
            .unwrap_or_else(|error| panic!("{file}: paragraphs() failed: {error:?}"));
        assert!(
            paragraphs.len() >= min_paragraphs,
            "{file}: {} paragraphs < {min_paragraphs}",
            paragraphs.len()
        );

        let tables = document
            .tables()
            .unwrap_or_else(|error| panic!("{file}: tables() failed: {error:?}"));
        assert!(
            tables.len() >= min_tables,
            "{file}: {} tables < {min_tables}",
            tables.len()
        );

        // Low-level parse paths: no panics, typed results.
        for (api, result) in [
            ("sections", document.sections().map(|_| ())),
            ("forms", document.forms().map(|_| ())),
            ("tracked_changes", document.tracked_changes().map(|_| ())),
            (
                "dynamic_text_fields",
                document.dynamic_text_fields().map(|_| ()),
            ),
            ("text_indexes", document.text_indexes().map(|_| ())),
            ("master_pages", document.master_pages().map(|_| ())),
            ("page_sequence", document.page_sequence().map(|_| ())),
            ("metadata", document.metadata().map(|_| ())),
        ] {
            if let Err(error) = result {
                assert_typed(error, file, api);
            }
        }
    }

    // writer-table-of-contents.odt carries a real text:table-of-content.
    let document = Document::open(corpus().join("writer-table-of-contents.odt")).unwrap();
    let indexes = document.text_indexes().unwrap();
    assert_eq!(indexes.len(), 1, "TOC fixture must expose one index");

    // writer-header-footer.odt defines master-page header/footer content.
    let document = Document::open(corpus().join("writer-header-footer.odt")).unwrap();
    assert!(
        !document.master_pages().unwrap().is_empty(),
        "header/footer fixture must expose master pages"
    );
}

#[test]
fn corpus_spreadsheets_parse_and_extract() {
    // (file, sheet count, minimum csv length)
    let cases: &[(&str, usize, usize)] = &[
        ("calc-formulas.ods", 3, 1),
        ("calc-unicode-chinese.ods", 1, 1),
        ("calc-cell-styles.ods", 1, 1),
        ("calc-two-sheets.ods", 2, 1),
    ];
    for &(file, sheet_count, min_csv) in cases {
        let path = corpus().join(file);
        let mut spreadsheet = Spreadsheet::open(&path)
            .unwrap_or_else(|error| panic!("{file}: open failed: {error:?}"));

        let sheets = spreadsheet
            .sheets()
            .unwrap_or_else(|error| panic!("{file}: sheets() failed: {error:?}"));
        assert_eq!(sheets.len(), sheet_count, "{file}: sheet count");
        assert!(
            sheets.iter().any(|sheet| !sheet.rows.is_empty()),
            "{file}: no populated rows"
        );

        let csv = spreadsheet
            .to_csv()
            .unwrap_or_else(|error| panic!("{file}: to_csv() failed: {error:?}"));
        assert!(csv.len() >= min_csv, "{file}: empty CSV export");

        for (api, result) in [("metadata", spreadsheet.metadata().map(|_| ()))] {
            if let Err(error) = result {
                assert_typed(error, file, api);
            }
        }
        // Cell-style registry (style:map conditional styles) is parsed eagerly.
        let _ = spreadsheet.conditional_cell_styles();
    }

    // calc-formulas.ods keeps formula representations on its cells.
    let mut spreadsheet = Spreadsheet::open(corpus().join("calc-formulas.ods")).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    assert!(
        sheets
            .iter()
            .flat_map(|sheet| &sheet.rows)
            .flat_map(|row| &row.cells)
            .any(|cell| cell.formula.is_some()),
        "formula fixture must expose at least one formula cell"
    );
}

#[test]
fn corpus_presentations_parse_and_extract() {
    // (file, minimum slide count)
    let cases: &[(&str, usize)] = &[
        ("impress-basic.odp", 1),
        ("impress-embedded-spreadsheet.odp", 1),
        ("impress-master-layouts.odp", 5),
    ];
    for &(file, min_slides) in cases {
        let path = corpus().join(file);
        let presentation = Presentation::open(&path)
            .unwrap_or_else(|error| panic!("{file}: open failed: {error:?}"));

        let slides = presentation
            .slides()
            .unwrap_or_else(|error| panic!("{file}: slides() failed: {error:?}"));
        assert!(
            slides.len() >= min_slides,
            "{file}: {} slides < {min_slides}",
            slides.len()
        );

        for (api, result) in [
            ("declarations", presentation.declarations().map(|_| ())),
            ("page_layouts", presentation.page_layouts().map(|_| ())),
            ("forms", presentation.forms().map(|_| ())),
            ("images", presentation.images().map(|_| ())),
            ("metadata", presentation.metadata().map(|_| ())),
        ] {
            if let Err(error) = result {
                assert_typed(error, file, api);
            }
        }
    }

    // impress-embedded-spreadsheet.odp embeds a spreadsheet object whose
    // reference stays inert (never opened or resolved by the corpus test).
    let presentation =
        Presentation::open(corpus().join("impress-embedded-spreadsheet.odp")).unwrap();
    let slides = presentation.slides().unwrap();
    assert!(
        slides.iter().flat_map(|slide| slide.shapes()).count() > 0,
        "embedded-object fixture must expose shapes"
    );
}

/// `writer-minimal-nasty.odt` (odfpy `nasty.odt`) is a deliberate XXE probe:
/// its `content.xml` declares `<!ENTITY nasty SYSTEM "externalcontent.txt">`
/// and references `&nasty;` in a paragraph. Litchi must never resolve that
/// external file — the entity and the DOCTYPE are rejected with typed
/// errors, while the package itself still opens.
#[test]
fn corpus_xxe_probe_is_rejected_without_resolution() {
    let path = corpus().join("writer-minimal-nasty.odt");
    let document = Document::open(&path)
        .unwrap_or_else(|error| panic!("XXE fixture must still open as a package: {error:?}"));

    let error = document
        .text()
        .expect_err("XXE entity must not be resolved by text()");
    assert_typed(error, "writer-minimal-nasty.odt", "text");

    let error = document
        .paragraphs()
        .expect_err("XXE entity must not be resolved by paragraphs()");
    assert_typed(error, "writer-minimal-nasty.odt", "paragraphs");

    let error = document
        .forms()
        .expect_err("DOCTYPE must be rejected by the form parser");
    assert_typed(error, "writer-minimal-nasty.odt", "forms");

    // Structural accessors that do not touch the hostile content still work.
    assert!(document.metadata().is_ok());
}
