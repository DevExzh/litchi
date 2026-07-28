//! Tests for the `PlcfAsumy` AutoSummary priority table (MS-DOC 2.8.4).

use litchi_cfb::OleFile;
use litchi_ole::doc::parts::auto_summary::DocumentAutoSummary;
use litchi_ole::doc::parts::fib::FileInformationBlock;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse_auto_summary(relative: &str) -> Option<DocumentAutoSummary> {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() { "1Table" } else { "0Table" };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    DocumentAutoSummary::parse(&fib, &table_stream).unwrap()
}

#[test]
fn documents_without_asumy_report_none() {
    // None of the checked-in fixtures carry AutoSummary priorities; the
    // parser must report None rather than erroring.
    assert!(parse_auto_summary("test-data/ole/doc/ThreeColHeadFoot.doc").is_none());
    assert!(parse_auto_summary("test-data/poi/test-data/document/47950_normal.doc").is_none());
}

#[test]
fn document_api_reports_none_without_asumy() {
    let mut package = litchi_ole::doc::Package::from_reader(
        File::open(fixture("test-data/ole/doc/ThreeColHeadFoot.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert!(document.auto_summary().is_none());
}
