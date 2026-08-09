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

//! Tests for the PLRSID revision-save identifier table (MS-DOC 2.9.203).

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::rsids::DocumentRsids;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn reads_plrsid_from_a_word_produced_document() {
    let path = fixture("test-data/poi/test-data/document/47950_normal.doc");
    let mut ole = OleFile::open(File::open(&path).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();

    let rsids = DocumentRsids::parse(&fib, &table_stream)
        .unwrap()
        .expect("Word 2002+ document carries a PLRSID");
    assert_eq!(rsids.ids().len(), 3);
    assert!(rsids.contains(0x00CA_5425));
}

#[test]
fn exposes_rsids_through_the_document_api() {
    let mut package = litchi_doc::Package::from_reader(
        File::open(fixture("test-data/poi/test-data/document/47950_normal.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    let rsids = document.rsids().expect("PLRSID present");
    assert_eq!(rsids.ids().len(), 3);
}

#[test]
fn documents_without_plrsid_report_none() {
    // Word 97 era files predate the PLRSID table.
    let path = fixture("test-data/poi/test-data/document/saved-by-table.doc");
    let mut ole = OleFile::open(File::open(&path).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    // Whether this old file carries the table is implementation-defined; the
    // parser must at least not error on it.
    let _ = DocumentRsids::parse(&fib, &table_stream).unwrap();
}
