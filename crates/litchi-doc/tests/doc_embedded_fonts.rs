#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

//! Tests for the `SttbTtmbd` embedded TrueType font table (MS-DOC 2.9.296).

use litchi_cfb::OleFile;
use litchi_doc::parts::embedded_fonts::DocumentEmbeddedFonts;
use litchi_doc::parts::fib::FileInformationBlock;
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
    let mut package = litchi_doc::Package::from_reader(
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
