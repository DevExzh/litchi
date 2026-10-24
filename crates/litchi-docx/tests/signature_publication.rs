#![allow(
    clippy::redundant_closure_call,
    reason = "raw OPC test edits stage one closure over the edit's graph"
)]
use litchi_core::Position;
use litchi_docx::Package;
use litchi_docx::document::composition::HistoryLimits;
use litchi_docx::document::{RevisionKind, raw};
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

fn bytes(package: &Package) -> Vec<u8> {
    let mut sink = Cursor::new(Vec::new());
    package.write_to(&mut sink).unwrap();
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
    let noop = package.edit().unwrap().commit().unwrap();
    let unchanged = package.apply(noop.patch()).unwrap().into_package();
    assert!(std::ptr::eq(package.opc_package(), unchanged.opc_package()));
    assert!(package.is_signed());
    assert!(unchanged.is_signed());
    assert_eq!(bytes(&unchanged), original);
    assert_eq!(bytes(&package), original);

    let mut edit = package.clone().unsign().edit().unwrap();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(package.apply(commit.patch()).is_err());
    assert!(package.is_signed());
    assert_eq!(bytes(&package), original);
    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    assert!(package.apply_durable_document(&durable).is_err());
    assert!(package.is_signed());
    assert_eq!(bytes(&package), original);

    let mut history = snapshot.history(HistoryLimits::new(4, 4 * 1024 * 1024));
    assert!(
        package
            .publish_document_commit_with_history(commit.clone(), &mut history)
            .is_err()
    );
    assert!(!history.can_undo());
    assert_eq!(history.current().xml_bytes(), snapshot.xml_bytes());
    assert_eq!(bytes(&package), original);

    let mut unsigned = package.clone();
    unsigned = unsigned.unsign();
    let published = unsigned.apply(commit.patch()).unwrap().into_package();
    assert!(!published.is_signed());
    assert!(package.is_signed());
    assert_eq!(bytes(&package), original);
    package = package.unsign();
    package = package
        .apply_durable_document(&durable)
        .unwrap()
        .into_package();
    assert!(!package.is_signed());
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        commit.patch().target().xml_bytes()
    );
    package = package
        .apply_durable_document(&durable.inverse())
        .unwrap()
        .into_package();
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
    let mut edit = source.clone().unsign().edit().unwrap();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    let target = signed_document(commit.patch().target().xml_bytes());
    let package = Package::from_reader(Cursor::new(target.as_slice())).unwrap();
    let mut history = snapshot.history(HistoryLimits::new(4, 4 * 1024 * 1024));
    history.record(commit.clone()).unwrap();

    assert!(package.undo_document(&mut history).is_err());
    assert!(history.can_undo());
    assert!(!history.can_redo());
    assert_eq!(
        history.current().xml_bytes(),
        commit.patch().target().xml_bytes()
    );
    assert_eq!(bytes(&package), target);

    assert!(history.undo());
    let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    assert!(package.redo_document(&mut history).is_err());
    assert!(!history.can_undo());
    assert!(history.can_redo());
    assert_eq!(history.current().xml_bytes(), snapshot.xml_bytes());
    assert_eq!(bytes(&package), original);

    package = package.unsign();
    let redone = package.redo_document(&mut history).unwrap();
    assert!(redone.diagnostics().changed());
    package = redone.into_package();
    let undone = package.undo_document(&mut history).unwrap();
    assert!(undone.diagnostics().changed());
    package = undone.into_package();
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
    package = {
        let mut edit = package.edit_opc();
        (|_candidate: &mut OpcPackage| -> litchi_docx::Result<_> { Ok(()) })(edit.package_mut())
            .unwrap();
        edit.commit().unwrap().into_package()
    };
    assert!(package.is_signed());
    assert_eq!(bytes(&package), original);
    let uri = PackURI::new("/word/document.xml").unwrap();
    assert!({
        let mut edit = package.edit_opc();
        (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> { {
        candidate.get_part_mut(&uri)?.set_blob(br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#.to_vec());
        Ok(())
    } })(edit.package_mut()).and_then(|_| edit.commit().map(|_| ()))
    }.is_err());
    assert_eq!(bytes(&package), original);
    assert!(
        {
            let mut edit = package.edit_opc();
            (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
                {
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
                }
            })(edit.package_mut())
            .and_then(|_| edit.commit().map(|_| ()))
        }
        .is_err()
    );
    assert_eq!(bytes(&package), original);
    let mut replacement =
        OpcPackage::from_bytes(&original, litchi_opc::SourceLimits::default()).unwrap();
    replacement.unsign();
    assert!(
        {
            let mut edit = package.edit_opc();
            (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
                {
                    *candidate = replacement;
                    Ok(())
                }
            })(edit.package_mut())
            .and_then(|_| edit.commit().map(|_| ()))
        }
        .is_err()
    );
    assert_eq!(bytes(&package), original);
    package = {
        let mut edit = package.edit_opc();
        (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
            {
                candidate.unsign();
                Ok(())
            }
        })(edit.package_mut())
        .unwrap();
        edit.commit().unwrap().into_package()
    };
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

    package = {
        let mut edit = package.edit_opc();
        (|_: &mut OpcPackage| -> litchi_docx::Result<_> { Ok(()) })(edit.package_mut()).unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_eq!(bytes(&package), ORIGINAL);
    assert!(
        {
            let mut edit = package.edit_opc();
            (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
                {
                    let uri = candidate.main_document_part()?.partname().clone();
                    candidate.get_part_mut(&uri)?.set_blob(REPLACEMENT.to_vec());
                    Ok(())
                }
            })(edit.package_mut())
            .and_then(|_| edit.commit().map(|_| ()))
        }
        .is_err()
    );
    assert_eq!(bytes(&package), ORIGINAL);
    let reports = package.signatures_with(&policy).unwrap();
    assert_eq!(reports[0].details().integrity(), litchi_sign::Status::Valid);
    assert_eq!(reports[0].details().signature(), litchi_sign::Status::Valid);

    package = package.unsign();
    package = {
        let mut edit = package.edit_opc();
        (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
            {
                let uri = candidate.main_document_part()?.partname().clone();
                candidate.get_part_mut(&uri)?.set_blob(REPLACEMENT.to_vec());
                Ok(())
            }
        })(edit.package_mut())
        .unwrap();
        edit.commit().unwrap().into_package()
    };
    let reopened = Package::from_reader(Cursor::new(bytes(&package))).unwrap();
    assert!(!reopened.is_signed());
    assert!(reopened.signatures_with(&policy).unwrap().is_empty());
}

#[test]
fn opaque_signature_members_allow_noop_but_refuse_raw_and_semantic_changes() {
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

    let signed = signed_source();
    let archive =
        ArchiveReader::new(&signed, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
    package = {
        let mut edit = package.edit_opc();
        (|_: &mut OpcPackage| -> litchi_docx::Result<_> { Ok(()) })(edit.package_mut()).unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_eq!(bytes(&package), original);
    assert!(
        {
            let mut edit = package.edit_opc();
            (|candidate: &mut OpcPackage| -> litchi_docx::Result<_> {
                {
                    candidate.unsign();
                    Ok(())
                }
            })(edit.package_mut())
            .and_then(|_| edit.commit().map(|_| ()))
        }
        .is_err()
    );
    assert_eq!(bytes(&package), original);

    let mut package = Package::from_reader(Cursor::new(original)).unwrap();
    let snapshot = package.document_snapshot().unwrap();
    let mut edit =
        raw::Tree::from_xml(snapshot.xml_bytes().to_vec(), raw::Limits::default()).unwrap();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let patch = edit.patch().unwrap();
    package = package.unsign();
    assert!(package.is_signed());
    assert!(package.apply(&patch).is_err());
    assert_eq!(
        package.document_snapshot().unwrap().xml_bytes(),
        snapshot.xml_bytes()
    );
}

#[test]
fn owned_commit_enforces_signature_policy_before_returning_a_package() {
    let original = signed_source();
    let mut package = Package::from_reader(original.as_slice()).unwrap();
    let noop = package.edit().unwrap().commit().unwrap().into_package();
    assert!(noop.is_signed());
    assert_eq!(bytes(&noop), original);
    let mut edit = package.edit().unwrap();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(bytes(&package), original);
    package = package.unsign();
    let mut edit = package.edit().unwrap();
    edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let committed = edit.commit().unwrap().into_package();
    assert!(!committed.is_signed());
    assert!(
        committed
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("added")
    );
}
