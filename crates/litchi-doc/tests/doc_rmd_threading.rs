//! Tests for the `RmdThreading` e-mail review threading data (MS-DOC 2.9.230).

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::rmd_threading::DocumentRmdThreading;
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse_threading(relative: &str) -> Option<DocumentRmdThreading> {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    DocumentRmdThreading::parse(&fib, &table_stream).unwrap()
}

#[test]
fn reads_single_author_threading_from_a_word_produced_document() {
    let threading = parse_threading("test-data/ole/doc/ThreeColHeadFoot.doc")
        .expect("Word 2000+ document carries RmdThreading");
    assert_eq!(threading.messages().len(), 1);
    // The sole author authored no e-mail message, so the identifier is empty
    // and the display properties are ignored.
    assert_eq!(threading.messages()[0].message_id(), "");
    assert!(threading.messages()[0].display().is_none());
    assert_eq!(threading.personal_styles().len(), 1);
}

#[test]
fn reads_multi_author_threading_from_a_word_produced_document() {
    let threading = parse_threading(
        "test-data/poi/test-data/document/au.edu.utas.www___data_assets_word_doc_0003_154335_International-Travel-Approval-Request-Form.doc",
    )
    .expect("Word 2000+ document carries RmdThreading");
    assert_eq!(threading.messages().len(), 2);
    assert!(
        threading
            .messages()
            .iter()
            .all(|message| message.message_id().is_empty() && message.display().is_none())
    );
    assert_eq!(threading.personal_styles().len(), 2);
}

#[test]
fn exposes_threading_through_the_document_api() {
    let mut package = litchi_doc::Package::from_reader(
        File::open(fixture("test-data/ole/doc/ThreeColHeadFoot.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    let threading = document.rmd_threading().expect("RmdThreading present");
    assert_eq!(threading.messages().len(), 1);
}
