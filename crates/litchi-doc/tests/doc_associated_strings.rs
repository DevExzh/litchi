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

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::{DocumentAssociatedStrings, Package};
use std::fs::File;
use std::path::Path;

#[test]
fn apache_poi_associated_strings_integrate_and_round_trip_exactly() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document/47950_normal.doc");

    let mut package = Package::from_reader(File::open(&path).unwrap()).unwrap();
    let document = package.document().unwrap();
    let integrated = document.associated_strings().unwrap();
    assert_eq!(integrated.author(), "Ross Johnson");
    assert_eq!(integrated.last_revised_by(), "Ross Johnson");

    let mut ole = OleFile::open(File::open(path).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    let parsed = DocumentAssociatedStrings::parse(&fib, &table_stream)
        .unwrap()
        .unwrap();
    let (offset, length) = fib.get_table_pointer(32).unwrap();
    let start = offset as usize;
    let end = start + length as usize;
    assert_eq!(parsed.to_bytes().unwrap(), table_stream[start..end]);
}
