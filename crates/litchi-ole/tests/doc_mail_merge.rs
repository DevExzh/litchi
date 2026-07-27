//! Tests for the legacy Word mail-merge data-source state: the `Pms`
//! structure (MS-DOC 2.9.205) and the ODSO property set (MS-DOC 2.9.162).

use litchi_cfb::OleFile;
use litchi_ole::doc::parts::fib::FileInformationBlock;
use litchi_ole::doc::parts::mail_merge::{
    DocumentMailMerge, MailMergeDestination, MailMergeDocumentType, MailMergeType,
    MergeDataSourceKind, MergeErrorCheck,
};
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

/// `image-comment-at-char.doc` carries a 136-byte `Pms` (and no ODSO data).
const MERGE_FIXTURE: &str = "test-data/ole/doc/image-comment-at-char.doc";

fn parse_fixture(relative: &str) -> (FileInformationBlock, Vec<u8>) {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    (fib, table_stream)
}

#[test]
fn reads_pms_from_a_mail_merge_document() {
    let (fib, table_stream) = parse_fixture(MERGE_FIXTURE);
    let mail_merge = DocumentMailMerge::parse(&fib, &table_stream)
        .unwrap()
        .expect("document carries a Pms");
    let pms = mail_merge.state().expect("Pms present");

    assert_eq!(pms.state.merge_type, MailMergeType::Letters);
    assert!(pms.state.main_document);
    assert!(!pms.state.data_source);
    assert!(pms.state.suppress_blank_lines);
    assert_eq!(pms.state.destination, MailMergeDestination::None);
    assert_eq!(pms.header_source_index, 0);
    assert_eq!(pms.fetch_source_index, 0);
    assert_eq!(pms.current_record, Some(1));

    for source in &pms.sources {
        assert_eq!(source.source_kind, MergeDataSourceKind::DataFile);
        assert!(source.file_name.is_mail_merge_source());
    }

    assert!(!pms.filter.show_data);
    assert_eq!(pms.filter.error_checking, MergeErrorCheck::PauseAndReport);
    assert!(pms.filter.mail_as_html);
    assert!(pms.strings.is_none());

    // The stored query is inert: exposed verbatim, never executed.
    assert_eq!(
        pms.sql_query.as_deref(),
        Some("SELECT * FROM writer-data-source-ooxml.dbo.Table1$")
    );
    assert_eq!(pms.document_type, Some(MailMergeDocumentType::Letters));

    // The fixture's fcODSO is undefined (lcbODSO is zero).
    assert!(mail_merge.odso_properties().is_empty());
}

#[test]
fn exposes_mail_merge_state_through_the_document_api() {
    let mut package =
        litchi_ole::doc::Package::from_reader(File::open(fixture(MERGE_FIXTURE)).unwrap()).unwrap();
    let document = package.document().unwrap();

    let mail_merge = document.mail_merge().expect("mail-merge state present");
    assert!(mail_merge.state().is_some());

    let pms = document.mail_merge_state().expect("Pms present");
    assert_eq!(pms.state.merge_type, MailMergeType::Letters);
    assert_eq!(pms.current_record, Some(1));

    let odso = document
        .odso_properties()
        .expect("mail-merge state present");
    assert!(odso.is_empty());
}

#[test]
fn documents_without_merge_state_report_none() {
    // Word 97 era file without mail-merge state.
    let mut package = litchi_ole::doc::Package::from_reader(
        File::open(fixture(
            "test-data/poi/test-data/document/saved-by-table.doc",
        ))
        .unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert!(document.mail_merge().is_none());
    assert!(document.mail_merge_state().is_none());
    assert!(document.odso_properties().is_none());
}
