//! Structural edits on files from other producers (record 0293).
//!
//! Record 0289 found that structural edits worked mainly on documents Litchi
//! made. Each case below failed before record 0293 and pins one producer
//! construct the leaves now handle; the inverse edit must restore the input
//! bytes. The corpus-wide picture is in `corpus_fidelity.rs` (structure
//! columns) and the ignored `structure_census.rs`. Record 0308 adds the
//! XLS and RTF producer constructs pinned below.

#![cfg(all(
    feature = "xls",
    feature = "pptx",
    feature = "rtf",
    feature = "docx",
    feature = "doc"
))]

use std::path::{Path, PathBuf};

use litchi_ops::{File, NewNode, Placement, Props, open};

fn path(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn input(relative: &str) -> (File, Vec<u8>) {
    let path = path(relative);
    (open(&path).unwrap(), std::fs::read(&path).unwrap())
}

mod xls {
    use super::*;

    /// Inserts a row and a column at the top of the first sheet and removes
    /// them again.
    fn round_trips(relative: &str) {
        round_trips_sheet(relative, 1);
    }

    /// Inserts a row and a column at the top of sheet `sheet` and removes
    /// them again: the inverse restores the input bytes.
    fn round_trips_sheet(relative: &str, sheet: usize) {
        let (file, bytes) = input(relative);
        let parent = format!("/sheet[{sheet}]");
        for kind in ["row", "column"] {
            let added = file
                .add(parent.as_str(), NewNode::new(kind).at(Placement::Start))
                .unwrap_or_else(|error| panic!("{relative} {kind}: {error}"));
            assert_ne!(added.file.to_bytes().unwrap(), bytes, "{relative} {kind}");
            let removed = added.file.remove(&added.path).unwrap();
            assert_eq!(removed.to_bytes().unwrap(), bytes, "{relative} {kind}");
        }
    }

    #[test]
    fn workbook_stream_padding_is_not_a_record() {
        // The Workbook stream ends with zero padding after its final EOF.
        round_trips("ole/xls/Simple.xls");
    }

    #[test]
    fn defined_names_follow_their_sheet() {
        round_trips("libreoffice-core/sc/qa/extras/testdocuments/NamesSheetLocal.xls");
    }

    #[test]
    fn embedded_charts_and_their_series_links_are_kept() {
        round_trips("libreoffice-core/sc/qa/unit/data/xls/embedded-chart.xls");
    }

    #[test]
    fn column_formats_move_with_their_columns() {
        round_trips("libreoffice-core/sc/qa/unit/data/xls/cell-borders.xls");
    }

    #[test]
    fn leaf_refusals_name_their_cause() {
        // A formula with ancillary RgbExtra data is still a typed refusal,
        // and the message is the leaf's, not the code summary.
        let (file, _) = input("ole/xls/FormulaEvalTestData.xls");
        let error = file
            .add("/sheet[1]", NewNode::new("row").at(Placement::Start))
            .map(drop)
            .expect_err("RgbExtra formulas are not shifted");
        assert_eq!(error.code().as_str(), "unsupported_capability");
        assert!(error.message().contains("RgbExtra"), "{}", error.message());
    }

    /// Record 0308: producer constructs the shifts now handle.
    #[test]
    fn worksheets_without_index_or_dbcell_regenerate_their_row_table() {
        round_trips("ole/xls/WithExtendedStyles.xls");
        // DIMENSIONS wider than the cells are kept, shifted.
        round_trips("libreoffice-core/sc/qa/extras/testdocuments/tdf78897.xls");
    }

    #[test]
    fn an_embedded_chart_has_its_own_dimensions() {
        round_trips("ole/xls/SimpleChart.xls");
    }

    #[test]
    fn split_drawings_and_picture_objects_shift() {
        round_trips("ole/xls/SimpleWithImages.xls");
    }

    #[test]
    fn shared_formulas_move_with_their_cells() {
        round_trips("poi/test-data/spreadsheet/44010-TwoCharts.xls");
    }

    #[test]
    fn reference_free_validations_move_with_their_ranges() {
        round_trips("libreoffice-core/sc/qa/unit/data/xls/pivottable_dates_grouping.xls");
    }

    #[test]
    fn a_shifted_shared_formula_reads_back() {
        // Insert a row at the top of a sheet with shared formulas: every
        // formula cell keeps its rendered formula one row lower.
        let (file, _) = input("poi/test-data/spreadsheet/44010-TwoCharts.xls");
        let added = file
            .add("/sheet[1]", NewNode::new("row").at(Placement::Start))
            .unwrap();
        let before = file.query("/sheet[1]/cell").unwrap().count();
        let after = added.file.query("/sheet[1]/cell").unwrap().count();
        assert_eq!(before, after);
    }

    #[test]
    fn locked_drawing_anchors_keep_their_place_over_default_rows() {
        // Comment boxes anchored "don't move or size with cells" over rows
        // of one height keep their cell coordinates (record 0316).
        round_trips("ole/xls/SimpleWithComments.xls");
    }

    #[test]
    fn locked_drawing_anchors_over_uneven_rows_are_refused() {
        let (file, _) = input("libreoffice-core/chart2/qa/extras/data/xls/chart.xls");
        let error = file
            .add("/sheet[1]", NewNode::new("row").at(Placement::Start))
            .map(drop)
            .expect_err("a position-locked anchor over uneven rows needs row sizes");
        assert!(
            error.message().contains("position-locked"),
            "{}",
            error.message()
        );
    }

    /// Record 0316: rule, array and table records, reader repairs and
    /// embedded substreams.
    #[test]
    fn conditional_formats_and_validations_shift_with_their_formulas() {
        // Absolute references in CF rules, and CFEx extensions.
        round_trips("ole/xls/WithConditionalFormatting.xls");
        // PtgRefN offsets in DV formulas.
        round_trips("libreoffice-core/sc/qa/unit/data/xls/validation.xls");
        // CondFmt12/CF12 icon sets, data bars and colour scales.
        round_trips("poi/test-data/spreadsheet/NewStyleConditionalFormattings.xls");
    }

    #[test]
    fn repaired_worksheets_are_shifted() {
        // WINDOW2 page-break zoom 1, read as the default zoom.
        round_trips("ole/xls/SingleLetterRanges.xls");
        // Stale fShrFmla flags on formulas with their own tokens.
        round_trips_sheet("ole/xls/DateTimeToNumberTestCases.xls", 2);
        // An empty sheet whose producer wrote a lone DBCELL.
        round_trips_sheet("ole/xls/SimpleWithColours.xls", 2);
    }

    #[test]
    fn embedded_chart_caches_are_not_cells() {
        // The chart's Number and Label cache records were read as cells A1
        // to A3 by the source-backed reader.
        let (file, _) = input("ole/xls/WithFormattedGraphTitle.xls");
        let text = file
            .get("/sheet[1]/A1")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .map(std::borrow::Cow::into_owned);
        assert_eq!(text.as_deref(), Some("Some Text"));
        round_trips("ole/xls/WithFormattedGraphTitle.xls");
        round_trips_sheet("ole/xls/WithChart.xls", 2);
    }
}

mod pptx {
    use super::*;

    #[test]
    fn producer_slides_are_removed() {
        // Slide 6 carries `p14:creationId`, slide 5 hyperlinks and a jump.
        let (file, _) = input("ooxml/pptx/shapes.pptx");
        for slide in ["/slide[6]", "/slide[5]"] {
            let removed = file.remove(slide).unwrap();
            let reopened = litchi_ops::open_bytes(removed.to_bytes().unwrap()).unwrap();
            assert!(reopened.get("/slide[6]").is_err());
            assert!(reopened.get("/slide[5]").is_ok());
        }
    }

    #[test]
    fn libreoffice_decks_admit_removal() {
        // LibreOffice declares the markup-compatibility namespace on its
        // presentation root.
        let (file, _) = input("office-interop/libreoffice-resaved/shapes-litchi.pptx");
        file.remove("/slide[6]").unwrap();
    }

    #[test]
    fn slides_owning_notes_are_removed_with_them() {
        // Record 0303: the notes slide only this slide owns leaves with it.
        let (file, _) = input("ooxml/pptx/sample.pptx");
        let removed = file.remove("/slide[2]").unwrap();
        let reopened = litchi_ops::open_bytes(removed.to_bytes().unwrap()).unwrap();
        assert!(reopened.get("/slide[1]").is_ok());
        assert!(reopened.get("/slide[2]").is_err());
    }

    #[test]
    fn guide_extensions_admit_slide_moves() {
        // `p15:sldGuideLst` in the presentation extension list.
        let (file, bytes) = input("ooxml/pptx/smartart-autoTxRot.pptx");
        let moved = file
            .move_to("/slide[1]", Placement::after("/slide[2]").unwrap())
            .unwrap();
        let back = moved.move_to("/slide[2]", Placement::Start).unwrap();
        assert_eq!(back.to_bytes().unwrap(), bytes);
    }
}

mod rtf {
    use super::*;

    fn text(file: &File, path: &str) -> Option<String> {
        file.get(path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .map(|text| text.into_owned())
    }

    /// Record 0308: a paragraph added to producer RTF is a scoped group, so
    /// removing it again restores the source bytes.
    #[test]
    fn producer_paragraphs_are_added_and_removed_exactly() {
        let (file, bytes) = input("rtf/testUnicode.rtf");
        let added = file
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Litchi ✓")
                    .at(Placement::Start),
            )
            .unwrap();
        let reopened = litchi_ops::open_bytes(added.file.to_bytes().unwrap()).unwrap();
        assert_eq!(text(&reopened, "/body/p[1]").as_deref(), Some("Litchi ✓"));
        assert_eq!(text(&reopened, "/body/p[2]"), text(&file, "/body/p[1]"));
        let removed = added.file.remove(&added.path).unwrap();
        assert_eq!(removed.to_bytes().unwrap(), bytes);
    }

    /// Record 0308: a paragraph that inherits the root-level state its
    /// neighbour set, and a formatting group holding several paragraphs,
    /// move value-exactly.
    #[test]
    fn inherited_state_and_split_groups_move_value_exactly() {
        for relative in ["rtf/testUnicode.rtf", "rtf/cjklist12.rtf"] {
            let (file, _) = input(relative);
            let first = text(&file, "/body/p[1]");
            let second = text(&file, "/body/p[2]");
            let moved = file
                .move_to("/body/p[1]", Placement::after("/body/p[2]").unwrap())
                .unwrap_or_else(|error| panic!("{relative}: {error}"));
            let reopened = litchi_ops::open_bytes(moved.to_bytes().unwrap()).unwrap();
            assert_eq!(text(&reopened, "/body/p[1]"), second, "{relative}");
            assert_eq!(text(&reopened, "/body/p[2]"), first, "{relative}");
        }
    }

    #[test]
    fn producer_paragraphs_move_and_swap_back_exactly() {
        let (file, bytes) = input("libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf");
        let moved = file
            .move_to("/body/p[1]", Placement::after("/body/p[2]").unwrap())
            .unwrap();
        assert_ne!(moved.to_bytes().unwrap(), bytes);
        let back = moved
            .move_to("/body/p[2]", Placement::before("/body/p[1]").unwrap())
            .unwrap();
        assert_eq!(back.to_bytes().unwrap(), bytes);
        let swapped = file.swap("/body/p[1]", "/body/p[2]").unwrap();
        let back = swapped.swap("/body/p[1]", "/body/p[2]").unwrap();
        assert_eq!(back.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn producer_paragraphs_are_removed() {
        let (file, _) = input("libreoffice-core/sw/qa/extras/rtfexport/data/margmirror.rtf");
        let second = file
            .get("/body/p[2]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .map(|text| text.into_owned());
        let removed = file.remove("/body/p[1]").unwrap();
        let reopened = litchi_ops::open_bytes(removed.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .get("/body/p[1]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .map(|text| text.into_owned()),
            second
        );
    }
}

#[test]
fn docx_alt_chunk_after_the_final_section_is_editable() {
    let relative = "libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx";
    let (file, bytes) = input(relative);
    let commit = file.edit(|_| Ok(())).unwrap();
    assert_eq!(commit.file().to_bytes().unwrap(), bytes);
    // Text edits before the tail are admitted; the tail stays untouched.
    let changed = file
        .set("/body/p[1]", Props::new().with("text", "Litchi"))
        .unwrap();
    let saved = changed.to_bytes().unwrap();
    let reopened = litchi_ops::open_bytes(saved).unwrap();
    assert_eq!(
        reopened
            .get("/body/p[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("Litchi")
    );
}

#[test]
fn doc_word_padded_papx_is_editable() {
    let (file, bytes) = input("poi/test-data/document/test.doc");
    let commit = file.edit(|_| Ok(())).unwrap();
    assert_eq!(commit.file().to_bytes().unwrap(), bytes);
}
