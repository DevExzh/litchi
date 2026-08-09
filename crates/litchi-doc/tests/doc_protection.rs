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

//! Tests for Word 2003 range-level protection tables (`SttbfBkmkProt`,
//! `PlcfBkfProt`, `PlcfBklProt`, `SttbProtUser`; MS-DOC 2.9.283 and 2.9.293).
//!
//! No fixture in `test-data/` (or its `3rdparty/` sources) carries these
//! tables — the only candidate hits were encrypted or fuzz-corrupted files —
//! so the typed parsing itself is covered by synthesized table streams in the
//! module's unit tests. These integration tests verify that real documents
//! without the tables parse cleanly and report no protected ranges.

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::protection::Ranges;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse_protection(relative: &str) -> Option<Ranges> {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    Ranges::parse(&fib, &table_stream).unwrap()
}

#[test]
fn word_2002_document_without_protection_tables_reports_none() {
    assert!(parse_protection("test-data/poi/test-data/document/47950_normal.doc").is_none());
}

#[test]
fn word_97_document_predating_the_tables_reports_none() {
    assert!(parse_protection("test-data/poi/test-data/document/saved-by-table.doc").is_none());
}

#[test]
fn exposes_no_protected_ranges_through_the_document_api() {
    let mut package = litchi_doc::Package::from_reader(
        File::open(fixture("test-data/poi/test-data/document/47950_normal.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert!(document.protected_ranges().is_none());
}
