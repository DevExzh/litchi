use super::*;

use crate::document::RevisionKind;
use crate::writer::{RevisionKind as AuthoredRevisionKind, RevisionMetadata};
use litchi_core::Position;
use std::io::Cursor;

const REVISION_DATE_UTC: &str = "2026-07-17T00:00:00.123456+00:00";

fn authored_source_docx() -> Vec<u8> {
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_run_with_text("before ");
        paragraph
            .add_revision(
                AuthoredRevisionKind::Insert,
                authored_metadata("1", "Alice"),
            )
            .add_run_with_text("accepted insertion");
        paragraph
            .add_revision(AuthoredRevisionKind::Insert, authored_metadata("2", "Bob"))
            .add_run_with_text("rejected insertion");
        paragraph
            .add_revision(
                AuthoredRevisionKind::Delete,
                authored_metadata("3", "Carol"),
            )
            .add_run_with_text("accepted deletion");
        paragraph
            .add_revision(AuthoredRevisionKind::Delete, authored_metadata("4", "Dora"))
            .add_run_with_text("restored deletion");
        paragraph.add_run_with_text(" after");
    }
    let mut stream = Cursor::new(Vec::new());
    package.to_plain_stream(&mut stream).unwrap();
    stream.into_inner()
}

fn authored_metadata(id: &str, author: &str) -> RevisionMetadata {
    let mut metadata = RevisionMetadata::new(id, author).unwrap();
    metadata.set_date_utc(Some(REVISION_DATE_UTC)).unwrap();
    metadata
}

fn durable_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(1, 32 * 1024 * 1024, 32 * 1024 * 1024),
        1024 * 1024,
        32,
        8,
        256 * 1024,
        512 * 1024,
    )
}

#[test]
fn authored_revision_actions_round_trip_through_borrowed_package_and_durable_inverse() {
    let source_docx = authored_source_docx();
    let mut package = Package::from_reader(Cursor::new(source_docx.as_slice())).unwrap();
    let source_xml = package.document_snapshot().unwrap().xml_bytes().to_vec();
    let source_text = std::str::from_utf8(&source_xml).unwrap();
    assert_eq!(source_text.matches("<w:ins ").count(), 2);
    assert_eq!(source_text.matches("<w:del ").count(), 2);
    assert_eq!(source_text.matches("w16du:dateUtc=").count(), 4);
    assert!(source_text.contains("mc:Ignorable=\"w16du\""));

    let mut edit = package.edit_document().unwrap();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap()
        .reject_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap()
        .accept_revision(Position::new(0), RevisionKind::Deletion, Position::new(0))
        .unwrap()
        .reject_revision(Position::new(0), RevisionKind::Deletion, Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.patch().operations().len(), 4);
    let target_xml = commit.snapshot().xml_bytes().to_vec();
    let target_text = std::str::from_utf8(&target_xml).unwrap();
    assert!(!target_text.contains("<w:ins"));
    assert!(!target_text.contains("<w:del"));
    assert!(target_text.contains("accepted insertion"));
    assert!(!target_text.contains("rejected insertion"));
    assert!(!target_text.contains("accepted deletion"));
    assert!(target_text.contains("restored deletion"));

    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    package.publish_document_commit(commit).unwrap();

    let changed_file = NamedTempFile::with_suffix(".docx").unwrap();
    package.save(changed_file.path()).unwrap();
    let mut reopened = Package::open(changed_file.path()).unwrap();
    assert_eq!(
        reopened.document_snapshot().unwrap().xml_bytes(),
        target_xml.as_slice()
    );

    let restored = reopened
        .apply_durable_document_patch(&durable.inverse())
        .unwrap();
    assert_eq!(restored.xml_bytes(), source_xml.as_slice());

    let restored_file = NamedTempFile::with_suffix(".docx").unwrap();
    reopened.save(restored_file.path()).unwrap();
    let reopened_restored = Package::open(restored_file.path()).unwrap();
    assert_eq!(
        reopened_restored.document_snapshot().unwrap().xml_bytes(),
        source_xml.as_slice()
    );
}
