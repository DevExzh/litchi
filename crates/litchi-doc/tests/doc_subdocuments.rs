//! Tests for the master-document subdocument directory (`PlcfWKB`) and the
//! referenced-file name table (`SttbFnm`), MS-DOC 2.8.34 and 2.9.288.
//!
//! No fixture in `test-data/` or `3rdparty/poi/test-data/document/` carries a
//! non-zero `lcbPlcfWkb` or `lcbSttbFnm` (the one hit elsewhere,
//! `3rdparty/libreoffice-core/sw/qa/extras/ww8export/data/tdf59896.doc`, holds
//! a garbage `lcbSttbFnm` and no table stream at all), so table content is
//! covered by the synthesized unit tests in `parts::subdocuments`; these
//! tests pin down that ordinary documents parse cleanly and report `None`.

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::subdocuments::Collection;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse(relative: &str) -> Option<Collection> {
    let path = fixture(relative);
    let mut ole = OleFile::open(File::open(&path).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    Collection::parse(&fib, &table_stream).unwrap()
}

#[test]
fn word_produced_document_without_master_document_tables_reports_none() {
    assert!(parse("test-data/poi/test-data/document/47950_normal.doc").is_none());
}

#[test]
fn word97_era_document_without_master_document_tables_reports_none() {
    assert!(parse("test-data/poi/test-data/document/saved-by-table.doc").is_none());
}

#[test]
fn exposes_subdocuments_through_the_document_api() {
    let mut package = litchi_doc::Package::from_reader(
        File::open(fixture("test-data/poi/test-data/document/47950_normal.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert!(document.subdocuments().is_none());
}
