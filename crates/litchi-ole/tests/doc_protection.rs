//! Tests for Word 2003 range-level protection tables (`SttbfBkmkProt`,
//! `PlcfBkfProt`, `PlcfBklProt`, `SttbProtUser`; MS-DOC 2.9.283 and 2.9.293).
//!
//! No fixture in `test-data/` (or its `3rdparty/` sources) carries these
//! tables — the only candidate hits were encrypted or fuzz-corrupted files —
//! so the typed parsing itself is covered by synthesized table streams in the
//! module's unit tests. These integration tests verify that real documents
//! without the tables parse cleanly and report no protected ranges.

use litchi_cfb::OleFile;
use litchi_ole::doc::parts::fib::FileInformationBlock;
use litchi_ole::doc::parts::protection::DocumentProtectedRanges;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse_protection(relative: &str) -> Option<DocumentProtectedRanges> {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() { "1Table" } else { "0Table" };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    DocumentProtectedRanges::parse(&fib, &table_stream).unwrap()
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
    let mut package = litchi_ole::doc::Package::from_reader(
        File::open(fixture("test-data/poi/test-data/document/47950_normal.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert!(document.protected_ranges().is_none());
}
