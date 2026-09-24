use litchi_core::Position;
use litchi_docx::Package;
use litchi_docx::document::{HistoryLimits, RevisionKind};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use std::io::Cursor;

fn signed_source() -> Vec<u8> {
    signed_document(br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins w:id="1" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>"#)
}

fn signed_document(document: &[u8]) -> Vec<u8> {
    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document.to_vec(),
    )));
    opc.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    )));
    opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    PackageWriter::to_bytes(&opc).unwrap()
}

fn bytes(package: &mut Package) -> Vec<u8> {
    let mut sink = Cursor::new(Vec::new());
    package.to_stream(&mut sink).unwrap();
    sink.into_inner()
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
fn revision_publication_requires_explicit_unsign_and_keeps_noops_exact() {
    let original = signed_source();
    let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    let snapshot = package.document_snapshot().unwrap();
    let noop = snapshot.edit().commit().unwrap();
    package.apply_document_patch(noop.patch()).unwrap();
    assert_eq!(bytes(&mut package), original);

    let mut edit = snapshot.edit();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(package.apply_document_patch(commit.patch()).is_err());
    assert!(package.is_signed());
    assert_eq!(bytes(&mut package), original);
    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    assert!(package.apply_durable_document_patch(&durable).is_err());
    assert!(package.is_signed());
    assert_eq!(bytes(&mut package), original);

    let mut history = snapshot.history(HistoryLimits::new(4, 4 * 1024 * 1024));
    assert!(
        package
            .publish_document_commit_with_history(commit.clone(), &mut history)
            .is_err()
    );
    assert!(!history.can_undo());
    assert_eq!(history.current().xml_bytes(), snapshot.xml_bytes());
    assert_eq!(bytes(&mut package), original);

    package.unsign();
    package.apply_durable_document_patch(&durable).unwrap();
    assert!(!package.is_signed());
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        commit.snapshot().xml_bytes()
    );
    package
        .apply_durable_document_patch(&durable.inverse())
        .unwrap();
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        snapshot.xml_bytes()
    );
}

#[test]
fn signed_undo_and_redo_restore_history_cursor_on_refusal() {
    let original = signed_source();
    let source = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    let snapshot = source.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    let target = signed_document(commit.snapshot().xml_bytes());
    let mut package = Package::from_reader(Cursor::new(target.as_slice())).unwrap();
    let mut history = snapshot.history(HistoryLimits::new(4, 4 * 1024 * 1024));
    history.record(commit.clone()).unwrap();

    assert!(package.undo_document(&mut history).is_err());
    assert!(history.can_undo());
    assert!(!history.can_redo());
    assert_eq!(history.current().xml_bytes(), commit.snapshot().xml_bytes());
    assert_eq!(bytes(&mut package), target);

    assert!(history.undo());
    let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    assert!(package.redo_document(&mut history).is_err());
    assert!(!history.can_undo());
    assert!(history.can_redo());
    assert_eq!(history.current().xml_bytes(), snapshot.xml_bytes());
    assert_eq!(bytes(&mut package), original);

    package.unsign();
    assert!(package.redo_document(&mut history).unwrap());
    assert!(package.undo_document(&mut history).unwrap());
    assert_eq!(history.current().xml_bytes(), snapshot.xml_bytes());
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        snapshot.xml_bytes()
    );
}

#[test]
fn raw_noop_preserves_signatures_and_raw_edits_require_explicit_policy() {
    let original = signed_source();
    let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    package.edit_opc(|_candidate| Ok(())).unwrap();
    assert!(package.is_signed());
    assert_eq!(bytes(&mut package), original);
    let uri = PackURI::new("/word/document.xml").unwrap();
    assert!(package.edit_opc(|candidate| {
        candidate.get_part_mut(&uri)?.set_blob(br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#.to_vec());
        Ok(())
    }).is_err());
    assert_eq!(bytes(&mut package), original);
    assert!(
        package
            .edit_opc(|candidate| {
                candidate.remove_part(&PackURI::new("/_xmlsignatures/origin.sigs").unwrap());
                let ids = candidate
                    .rels()
                    .iter()
                    .filter(|rel| rel.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
                    .map(|rel| rel.r_id().to_owned())
                    .collect::<Vec<_>>();
                for id in ids {
                    candidate.rels_mut().remove(&id);
                }
                Ok(())
            })
            .is_err()
    );
    assert_eq!(bytes(&mut package), original);
    let mut replacement = OpcPackage::from_bytes(&original).unwrap();
    replacement.unsign();
    assert!(
        package
            .edit_opc(|candidate| {
                *candidate = replacement;
                Ok(())
            })
            .is_err()
    );
    assert_eq!(bytes(&mut package), original);
    package
        .edit_opc(|candidate| {
            candidate.unsign();
            Ok(())
        })
        .unwrap();
    assert!(!package.is_signed());
}

#[cfg(feature = "sign")]
#[test]
fn producer_signed_package_keeps_verifiable_signature_after_noop_and_refusal() {
    const ORIGINAL: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx");
    const REPLACEMENT: &[u8] = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>explicit edit</w:t></w:r></w:p></w:body></w:document>"#;
    let mut package = Package::from_reader(Cursor::new(ORIGINAL)).unwrap();
    let policy = litchi_sign::Policy::compatible();
    let reports = package.signatures_with(&policy).unwrap();
    assert_eq!(reports.len(), 1);
    assert_eq!(reports[0].details().integrity(), litchi_sign::Status::Valid);
    assert_eq!(reports[0].details().signature(), litchi_sign::Status::Valid);

    package.edit_opc(|_| Ok(())).unwrap();
    assert_eq!(bytes(&mut package), ORIGINAL);
    assert!(
        package
            .edit_opc(|candidate| {
                let uri = candidate.main_document_part()?.partname().clone();
                candidate.get_part_mut(&uri)?.set_blob(REPLACEMENT.to_vec());
                Ok(())
            })
            .is_err()
    );
    assert_eq!(bytes(&mut package), ORIGINAL);
    let reports = package.signatures_with(&policy).unwrap();
    assert_eq!(reports[0].details().integrity(), litchi_sign::Status::Valid);
    assert_eq!(reports[0].details().signature(), litchi_sign::Status::Valid);

    package.unsign();
    package
        .edit_opc(|candidate| {
            let uri = candidate.main_document_part()?.partname().clone();
            candidate.get_part_mut(&uri)?.set_blob(REPLACEMENT.to_vec());
            Ok(())
        })
        .unwrap();
    let reopened = Package::from_reader(Cursor::new(bytes(&mut package))).unwrap();
    assert!(!reopened.is_signed());
    assert!(reopened.signatures_with(&policy).unwrap().is_empty());
}

#[test]
fn opaque_signature_members_allow_noop_but_refuse_raw_and_semantic_changes() {
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

    let signed = signed_source();
    let archive = ArchiveReader::new(&signed).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        writer
            .write_stored(name, &archive.read(name).unwrap())
            .unwrap();
    }
    writer
        .write_stored("_xmlsignatures/opaque.unknown", b"opaque signature entry")
        .unwrap();
    let original = writer.finish_to_bytes().unwrap();
    let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    assert!(
        package
            .opc_package()
            .non_part_members()
            .iter()
            .any(|member| member.name() == "_xmlsignatures/opaque.unknown")
    );
    package.edit_opc(|_| Ok(())).unwrap();
    assert_eq!(bytes(&mut package), original);
    assert!(
        package
            .edit_opc(|candidate| {
                candidate.unsign();
                Ok(())
            })
            .is_err()
    );
    assert_eq!(bytes(&mut package), original);

    let mut package = Package::from_reader(Cursor::new(original)).unwrap();
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    package.unsign();
    assert!(package.is_signed());
    assert!(package.apply_document_patch(commit.patch()).is_err());
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        snapshot.xml_bytes()
    );
}
