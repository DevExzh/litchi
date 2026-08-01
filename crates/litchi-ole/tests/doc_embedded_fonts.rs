//! Tests for the `SttbTtmbd` embedded TrueType font table (MS-DOC 2.9.296).

use litchi_cfb::OleFile;
use litchi_ole::doc::parts::embedded_fonts::DocumentEmbeddedFonts;
use litchi_ole::doc::parts::fib::FileInformationBlock;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse_embedded_fonts(relative: &str) -> Option<DocumentEmbeddedFonts> {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    DocumentEmbeddedFonts::parse(&fib, &table_stream).unwrap()
}

#[test]
fn reads_empty_font_table_with_a_nonstandard_brgbst() {
    // This Word-produced file carries an SttbTtmbd whose brgbst is 26 rather
    // than the recommended 10, with zero embedded fonts.
    let fonts = parse_embedded_fonts("test-data/poi/test-data/hpsf/TestNon4ByteBoundary.doc")
        .expect("document carries an SttbTtmbd");
    assert!(fonts.fonts().is_empty());
}

#[test]
fn documents_without_the_table_report_none() {
    assert!(parse_embedded_fonts("test-data/ole/doc/ThreeColHeadFoot.doc").is_none());
}

#[test]
fn exposes_embedded_fonts_through_the_document_api() {
    let mut package = litchi_ole::doc::Package::from_reader(
        File::open(fixture(
            "test-data/poi/test-data/hpsf/TestNon4ByteBoundary.doc",
        ))
        .unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    let fonts = document.embedded_fonts().expect("SttbTtmbd present");
    assert!(fonts.fonts().is_empty());
}
