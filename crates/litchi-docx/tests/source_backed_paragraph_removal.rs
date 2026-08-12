use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{OwnedSource, Position, ReadAt};
use litchi_docx::source_backed::paragraph_copy::Patch as CopyPatch;
use litchi_docx::source_backed::paragraph_remove::{Error, Limits, Patch, Refusal};
use litchi_docx::{Package, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MAIN: &str = "word/document.xml";
const UNUSED: &str = "word/unused.bin";

fn document(paragraphs: &[&str]) -> Vec<u8> {
    let body = paragraphs
        .iter()
        .map(|text| format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#))
        .collect::<String>();
    format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn fixture(document: Vec<u8>) -> OpcPackage {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{MAIN}")).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document,
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED}")).unwrap(),
            "application/octet-stream".to_owned(),
            (0..=255).cycle().take(32 * 1024).collect(),
        )))
        .unwrap();
    package.relate_to(MAIN, rt::OFFICE_DOCUMENT);
    package
}

fn bytes(package: &OpcPackage) -> Vec<u8> {
    PackageWriter::to_bytes(package).unwrap()
}

fn open(source: Vec<u8>) -> source_backed::Package {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source));
    source_backed::Package::from_read_at(source).unwrap()
}

fn reopened_texts(bytes: Vec<u8>) -> Vec<String> {
    Package::from_reader(io::Cursor::new(bytes))
        .unwrap()
        .document_snapshot()
        .unwrap()
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

fn payload_range(zip: &[u8], name: &str) -> std::ops::Range<usize> {
    let name = name.as_bytes();
    for (offset, _) in zip
        .windows(4)
        .enumerate()
        .filter(|(_, signature)| *signature == b"PK\x01\x02")
    {
        if offset + 46 > zip.len() {
            continue;
        }
        let compressed_size =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_length =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        let extra_length =
            u16::from_le_bytes(zip[offset + 30..offset + 32].try_into().unwrap()) as usize;
        if offset + 46 + name_length + extra_length > zip.len()
            || &zip[offset + 46..offset + 46 + name_length] != name
        {
            continue;
        }
        let local_offset =
            u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
        let local_name_length = u16::from_le_bytes(
            zip[local_offset + 26..local_offset + 28]
                .try_into()
                .unwrap(),
        ) as usize;
        let local_extra_length = u16::from_le_bytes(
            zip[local_offset + 28..local_offset + 30]
                .try_into()
                .unwrap(),
        ) as usize;
        let data_start = local_offset + 30 + local_name_length + local_extra_length;
        return data_start..data_start + compressed_size;
    }
    panic!("ZIP member was not found");
}

#[test]
fn removes_first_middle_last_and_the_only_paragraph_exactly() {
    for (position, expected) in [
        (0, vec!["beta", "gamma"]),
        (1, vec!["alpha", "gamma"]),
        (2, vec!["alpha", "beta"]),
    ] {
        let package = open(bytes(&fixture(document(&["alpha", "beta", "gamma"]))));
        let mut edit = package.edit_plain_paragraph_removal().unwrap();
        edit.remove_plain_paragraph(Position::new(position))
            .unwrap();
        let commit = edit.commit();
        assert_eq!(commit.projected().paragraph_count(), 2);
        assert_eq!(commit.effect_report().removed_paragraphs, 1);
        assert!(commit.effect_report().removed_bytes > 0);
        let mut output = Vec::new();
        package
            .publish_plain_paragraph_removal_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(reopened_texts(output), expected);
    }

    let package = open(bytes(&fixture(document(&["only"]))));
    let mut edit = package.edit_plain_paragraph_removal().unwrap();
    edit.remove_plain_paragraph(Position::new(0)).unwrap();
    let commit = edit.commit();
    assert_eq!(commit.projected().paragraph_count(), 0);
    assert!(
        commit
            .projected()
            .xml_bytes()
            .windows(b"<w:body></w:body>".len())
            .any(|value| value == b"<w:body></w:body>")
    );
}

#[test]
fn publication_raw_copies_others_and_both_inverse_paths_restore_exact_artifact() {
    let source = bytes(&fixture(document(&["one", "two", "three"])));
    let retained_range = payload_range(&source, UNUSED);
    let package = open(source.clone());
    let mut edit = package.edit_plain_paragraph_removal().unwrap();
    edit.remove_plain_paragraph(Position::new(1)).unwrap();
    let commit = edit.commit();
    let mut output = Vec::new();
    let publication = package
        .publish_plain_paragraph_removal_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        &output[payload_range(&output, UNUSED)],
        &source[retained_range]
    );

    let inverse_wire = publication.inverse_patch().to_bytes().unwrap();
    let inverse = Patch::from_bytes(&inverse_wire).unwrap();
    let mut durable_restored = Vec::new();
    open(output.clone())
        .publish_plain_paragraph_removal_patch_to_stream(&mut durable_restored, &inverse)
        .unwrap();
    assert_eq!(durable_restored, source);

    let mut exact_restored = Vec::new();
    open(output)
        .publish_plain_paragraph_removal_inverse_to_stream(&mut exact_restored, &publication)
        .unwrap();
    assert_eq!(exact_restored, source);
}

#[test]
fn durable_removal_is_canonical_semantically_closed_and_whole_artifact_bound() {
    let source = bytes(&fixture(document(&["a", "b", "c"])));
    let package = open(source.clone());
    let snapshot = package.plain_paragraph_removal_snapshot().unwrap();
    let mut edit = snapshot.removal_edit();
    edit.remove_plain_paragraph(Position::new(1)).unwrap();
    let commit = edit.commit();
    let wire = commit.patch().to_bytes().unwrap();
    assert!(matches!(
        CopyPatch::from_bytes(&wire),
        Err(Error::InvalidDurable)
    ));
    let durable = Patch::from_bytes(&wire).unwrap();
    assert_eq!(durable.to_bytes().unwrap(), wire);
    assert_eq!(
        durable.apply(&snapshot).unwrap().xml_bytes(),
        commit.projected().xml_bytes()
    );

    let mut foreign = fixture(document(&["a", "b", "c"]));
    foreign
        .get_part_mut(&PackURI::new(format!("/{UNUSED}")).unwrap())
        .unwrap()
        .set_blob(b"different retained member".to_vec());
    let foreign_bytes = bytes(&foreign);
    let stale = open(foreign_bytes.clone())
        .plain_paragraph_removal_snapshot()
        .unwrap();
    assert!(matches!(durable.apply(&stale), Err(Error::StaleSource)));
    let mut stale_output = Vec::new();
    assert!(matches!(
        open(foreign_bytes)
            .publish_plain_paragraph_removal_patch_to_stream(&mut stale_output, &durable),
        Err(Error::StaleSource)
    ));
    assert!(stale_output.is_empty());

    let mut noncanonical_reserved = wire.clone();
    noncanonical_reserved[62..66].copy_from_slice(&1u32.to_le_bytes());
    assert!(matches!(
        Patch::from_bytes(&noncanonical_reserved),
        Err(Error::InvalidDurable)
    ));
    let mut forged_target = wire;
    let last = forged_target.len() - 1;
    forged_target[last] ^= 1;
    assert!(Patch::from_bytes(&forged_target).is_err());
}

#[test]
fn invalid_positions_operation_and_resource_limits_leave_projection_unchanged() {
    let source = bytes(&fixture(document(&["a", "b"])));
    let snapshot = open(source.clone())
        .plain_paragraph_removal_snapshot()
        .unwrap();
    let mut edit = snapshot.removal_edit();
    assert!(matches!(
        edit.remove_plain_paragraph(Position::new(2)),
        Err(Error::OutOfBounds {
            kind: "removal",
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), snapshot.xml_bytes());
    edit.remove_plain_paragraph(Position::new(0)).unwrap();
    let once = edit.projected().xml_bytes().to_vec();
    assert!(matches!(
        edit.remove_plain_paragraph(Position::new(0)),
        Err(Error::Limit {
            resource: "operations",
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), once);

    let paragraph_limited = Limits::new(1024, 1, 100, 8, 2048, 4096).unwrap();
    assert!(matches!(
        open(source.clone()).plain_paragraph_removal_snapshot_with_limits(paragraph_limited),
        Err(Error::Limit {
            resource: "paragraphs",
            ..
        })
    ));

    let output_limited = Limits::new(1024, 4, 100, 8, 1, 4096).unwrap();
    let snapshot = open(source)
        .plain_paragraph_removal_snapshot_with_limits(output_limited)
        .unwrap();
    let mut edit = snapshot.removal_edit();
    assert!(matches!(
        edit.remove_plain_paragraph(Position::new(0)),
        Err(Error::Limit {
            resource: "output bytes",
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), snapshot.xml_bytes());

    let durable_limited = Limits::new(1024, 4, 100, 8, 2048, 128).unwrap();
    let snapshot = open(bytes(&fixture(document(&["a"]))))
        .plain_paragraph_removal_snapshot_with_limits(durable_limited)
        .unwrap();
    let mut edit = snapshot.removal_edit();
    edit.remove_plain_paragraph(Position::new(0)).unwrap();
    assert!(matches!(
        edit.commit().patch().to_bytes(),
        Err(Error::Limit {
            resource: "durable patch bytes",
            ..
        })
    ));
}

#[test]
fn empty_edit_is_exact_noop_and_security_or_complexity_refuses_before_editing() {
    let source = bytes(&fixture(document(&["only"])));
    let package = open(source.clone());
    let commit = package.edit_plain_paragraph_removal().unwrap().commit();
    assert!(commit.patch().is_noop());
    assert!(commit.effect_report().is_noop());
    let wire = commit.patch().to_bytes().unwrap();
    assert_eq!(Patch::from_bytes(&wire).unwrap().to_bytes().unwrap(), wire);
    let mut output = Vec::new();
    package
        .publish_plain_paragraph_removal_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);

    let empty = format!(r#"<w:document xmlns:w="{W}"><w:body></w:body></w:document>"#).into_bytes();
    let mut empty_edit = open(bytes(&fixture(empty)))
        .edit_plain_paragraph_removal()
        .unwrap();
    assert!(matches!(
        empty_edit.remove_plain_paragraph(Position::new(0)),
        Err(Error::OutOfBounds { .. })
    ));

    let mut external = fixture(document(&["plain"]));
    external
        .get_part_mut(&PackURI::new(format!("/{MAIN}")).unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::HYPERLINK.to_owned(),
            "https://example.invalid".to_owned(),
            "rExternal".to_owned(),
            true,
        );
    assert!(matches!(
        open(bytes(&external)).plain_paragraph_removal_snapshot(),
        Err(Error::Refused(Refusal::ExternalRelationship))
    ));

    let complex = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr/><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open(bytes(&fixture(complex))).plain_paragraph_removal_snapshot(),
        Err(Error::Refused(Refusal::ComplexParagraph))
    ));
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset"))?;
        let Some(source) = self.bytes.get(offset..) else {
            return Ok(0);
        };
        let count = source.len().min(output.len());
        output[..count].copy_from_slice(&source[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        Ok(litchi_core::SourceVersion::new(
            91,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn changed_live_source_version_is_rejected_before_publication_output() {
    let source = Arc::new(VersionedSource {
        bytes: bytes(&fixture(document(&["a", "b"]))),
        revision: AtomicU64::new(0),
    });
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_plain_paragraph_removal().unwrap();
    edit.remove_plain_paragraph(Position::new(0)).unwrap();
    let commit = edit.commit();
    source.revision.fetch_add(1, Ordering::SeqCst);
    let mut output = Vec::new();
    assert!(
        package
            .publish_plain_paragraph_removal_to_stream(&mut output, &commit)
            .is_err()
    );
    assert!(output.is_empty());
}
