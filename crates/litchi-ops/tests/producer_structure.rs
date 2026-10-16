//! Structural edits on files from other producers (record 0293).
//!
//! Record 0289 found that structural edits worked mainly on documents Litchi
//! made. Each case below failed before record 0293 and pins one producer
//! construct the leaves now handle; the inverse edit must restore the input
//! bytes. The corpus-wide picture is in `corpus_fidelity.rs` (structure
//! columns) and the ignored `structure_census.rs`.

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
        let (file, bytes) = input(relative);
        for kind in ["row", "column"] {
            let added = file
                .add("/sheet[1]", NewNode::new(kind).at(Placement::Start))
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
        // A workbook-wide conditional format is still a typed refusal, and
        // the message is the leaf's, not the code summary.
        let (file, _) = input("ole/xls/WithConditionalFormatting.xls");
        let error = file
            .add("/sheet[1]", NewNode::new("row").at(Placement::Start))
            .map(drop)
            .expect_err("conditional formats are not shifted");
        assert_eq!(error.code().as_str(), "unsupported_capability");
        assert!(error.message().contains("0x01B0"), "{}", error.message());
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
    assert_eq!(commit.file.to_bytes().unwrap(), bytes);
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
    assert_eq!(commit.file.to_bytes().unwrap(), bytes);
}
