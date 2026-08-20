use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{OwnedSource, Position, ReadAt};
use litchi_docx::source_backed::paragraph_copy::{Error, Limits, Patch, Refusal};
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

fn fixture(main_path: &str, document: Vec<u8>) -> OpcPackage {
    fixture_with_relationship(main_path, document, rt::OFFICE_DOCUMENT)
}

fn fixture_with_relationship(
    main_path: &str,
    document: Vec<u8>,
    relationship_type: &str,
) -> OpcPackage {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{main_path}")).unwrap(),
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
    package.relate_to(main_path, relationship_type);
    package
}

fn bytes(package: &OpcPackage) -> Vec<u8> {
    PackageWriter::to_bytes(package).unwrap()
}

fn open(source: Vec<u8>) -> source_backed::Package {
    let read_at: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source));
    source_backed::Package::from_read_at(read_at).unwrap()
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
fn copies_exact_fragment_at_first_middle_and_last_source_order_slots() {
    let cases = [
        (0, vec!["beta", "alpha", "beta", "gamma"]),
        (2, vec!["alpha", "beta", "beta", "gamma"]),
        (3, vec!["alpha", "beta", "gamma", "beta"]),
    ];
    for (before, expected) in cases {
        let source = bytes(&fixture(MAIN, document(&["alpha", "beta", "gamma"])));
        let package = open(source);
        let snapshot = package.plain_paragraph_copy_snapshot().unwrap();
        let selected = snapshot
            .xml_bytes()
            .windows(b"<w:p><w:r><w:t>beta</w:t></w:r></w:p>".len())
            .filter(|window| *window == b"<w:p><w:r><w:t>beta</w:t></w:r></w:p>")
            .count();
        assert_eq!(selected, 1);
        let mut edit = snapshot.edit();
        edit.copy_plain_paragraph(Position::new(1), Position::new(before))
            .unwrap();
        let commit = edit.commit();
        assert_eq!(commit.effect_report().copied_paragraphs, 1);
        assert_eq!(
            commit.effect_report().copied_bytes,
            b"<w:p><w:r><w:t>beta</w:t></w:r></w:p>".len()
        );
        let projected_copies = commit
            .projected()
            .xml_bytes()
            .windows(b"<w:p><w:r><w:t>beta</w:t></w:r></w:p>".len())
            .filter(|window| *window == b"<w:p><w:r><w:t>beta</w:t></w:r></w:p>")
            .count();
        assert_eq!(
            projected_copies, 2,
            "exact source fragment must be duplicated"
        );

        let mut output = Vec::new();
        package
            .publish_plain_paragraph_copy_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(reopened_texts(output), expected);
    }
}

#[test]
fn publication_raw_copies_unselected_members_and_inverse_restores_exact_artifact() {
    let source = bytes(&fixture(MAIN, document(&["one", "two"])));
    let unused = payload_range(&source, UNUSED);
    let package = open(source.clone());
    let mut edit = package.edit_plain_paragraph_copy().unwrap();
    edit.copy_plain_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let commit = edit.commit();
    let mut output = Vec::new();
    let publication = package
        .publish_plain_paragraph_copy_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        &output[payload_range(&output, UNUSED)],
        &source[unused],
        "unselected physical payload must be raw-identical"
    );

    let inverse_wire = publication.inverse_patch().to_bytes().unwrap();
    let durable_inverse = Patch::from_bytes(&inverse_wire).unwrap();
    let mut durable_restored = Vec::new();
    open(output.clone())
        .publish_plain_paragraph_copy_patch_to_stream(&mut durable_restored, &durable_inverse)
        .unwrap();
    assert_eq!(durable_restored, source);

    let published = open(output);
    let mut restored = Vec::new();
    published
        .publish_plain_paragraph_copy_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
}

#[test]
fn durable_forward_inverse_are_canonical_and_reject_whole_artifact_staleness() {
    let source = bytes(&fixture(MAIN, document(&["a", "b", "c"])));
    let package = open(source.clone());
    let snapshot = package.plain_paragraph_copy_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.copy_plain_paragraph(Position::new(2), Position::new(1))
        .unwrap();
    let commit = edit.commit();
    let wire = commit.patch().to_bytes().unwrap();
    let live_applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(live_applied.xml_bytes(), commit.projected().xml_bytes());
    assert_eq!(wire, commit.patch().to_bytes().unwrap());
    let durable = Patch::from_bytes(&wire).unwrap();
    assert_eq!(durable.to_bytes().unwrap(), wire);
    let applied = durable.apply(&snapshot).unwrap();
    assert_eq!(applied.xml_bytes(), commit.projected().xml_bytes());
    let mut published = Vec::new();
    open(source.clone())
        .publish_plain_paragraph_copy_patch_to_stream(&mut published, &durable)
        .unwrap();
    assert_eq!(reopened_texts(published), ["a", "c", "b", "c"]);
    let inverse_wire = durable.inverse().to_bytes().unwrap();
    let inverse = Patch::from_bytes(&inverse_wire).unwrap();
    let restored = inverse.apply(&applied).unwrap();
    assert_eq!(restored.xml_bytes(), snapshot.xml_bytes());

    let mut foreign = fixture(MAIN, document(&["a", "b", "c"]));
    foreign
        .get_part_mut(&PackURI::new(format!("/{UNUSED}")).unwrap())
        .unwrap()
        .set_blob(b"different retained member".to_vec());
    let stale = open(bytes(&foreign))
        .plain_paragraph_copy_snapshot()
        .unwrap();
    assert!(matches!(durable.apply(&stale), Err(Error::StaleSource)));

    let mut tampered = wire;
    tampered[58..62].copy_from_slice(&0u32.to_le_bytes());
    assert!(matches!(
        Patch::from_bytes(&tampered),
        Err(Error::InvalidDurable)
    ));
}

#[test]
fn empty_edit_is_an_exact_noop_and_stale_inverse_writes_nothing() {
    let source = bytes(&fixture(MAIN, document(&["only"])));
    let package = open(source.clone());
    let commit = package.edit_plain_paragraph_copy().unwrap().commit();
    assert!(commit.patch().is_noop());
    assert!(commit.effect_report().is_noop());
    let mut output = Vec::new();
    let publication = package
        .publish_plain_paragraph_copy_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);

    let foreign = open(bytes(&fixture(MAIN, document(&["foreign"]))));
    let mut inverse_output = Vec::new();
    assert!(matches!(
        foreign.publish_plain_paragraph_copy_inverse_to_stream(&mut inverse_output, &publication),
        Err(Error::StaleSource)
    ));
    assert!(inverse_output.is_empty());
}

#[test]
fn positions_operation_count_and_resource_limits_are_failure_atomic() {
    let source = bytes(&fixture(MAIN, document(&["a", "b"])));
    let package = open(source.clone());
    let snapshot = package.plain_paragraph_copy_snapshot().unwrap();
    let mut edit = snapshot.edit();
    assert!(matches!(
        edit.copy_plain_paragraph(Position::new(2), Position::new(0)),
        Err(Error::OutOfBounds { kind: "source", .. })
    ));
    assert_eq!(edit.projected().xml_bytes(), snapshot.xml_bytes());
    assert!(matches!(
        edit.copy_plain_paragraph(Position::new(0), Position::new(3)),
        Err(Error::OutOfBounds {
            kind: "insertion",
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), snapshot.xml_bytes());
    edit.copy_plain_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    let once = edit.projected().xml_bytes().to_vec();
    assert!(matches!(
        edit.copy_plain_paragraph(Position::new(0), Position::new(1)),
        Err(Error::Limit {
            resource: "operations",
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), once);

    let paragraph_limited = Limits::new(1024, 1, 100, 8, 2048, 4096).unwrap();
    assert!(matches!(
        open(source.clone()).plain_paragraph_copy_snapshot_with_limits(paragraph_limited),
        Err(Error::Limit {
            resource: "paragraphs",
            ..
        })
    ));
    let output_limited = Limits::new(1024, 4, 100, 8, document(&["a", "b"]).len(), 4096).unwrap();
    let snapshot = open(source)
        .plain_paragraph_copy_snapshot_with_limits(output_limited)
        .unwrap();
    let mut edit = snapshot.edit();
    assert!(matches!(
        edit.copy_plain_paragraph(Position::new(0), Position::new(0)),
        Err(Error::Limit {
            resource: "output bytes",
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), snapshot.xml_bytes());

    let durable_limited = Limits::new(1024, 4, 100, 8, 2048, 128).unwrap();
    let snapshot = open(bytes(&fixture(MAIN, document(&["a"]))))
        .plain_paragraph_copy_snapshot_with_limits(durable_limited)
        .unwrap();
    let mut edit = snapshot.edit();
    edit.copy_plain_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    assert!(matches!(
        edit.commit().patch().to_bytes(),
        Err(Error::Limit {
            resource: "durable patch bytes",
            ..
        })
    ));
}

#[test]
fn relationships_dependencies_signatures_macros_and_paths_are_refused() {
    let external = {
        let mut package = fixture(MAIN, document(&["plain"]));
        package
            .get_part_mut(&PackURI::new(format!("/{MAIN}")).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::HYPERLINK.to_owned(),
                "https://example.invalid".to_owned(),
                "rExternal".to_owned(),
                true,
            );
        bytes(&package)
    };
    assert!(matches!(
        open(external).plain_paragraph_copy_snapshot(),
        Err(Error::Refused(Refusal::ExternalRelationship))
    ));

    let custom_xml = {
        let mut package = fixture(MAIN, document(&["plain"]));
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/customXml/item1.xml").unwrap(),
                "application/xml".to_owned(),
                b"<x/>".to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new(format!("/{MAIN}")).unwrap())
            .unwrap()
            .relate_to("../customXml/item1.xml", rt::CUSTOM_XML);
        bytes(&package)
    };
    assert!(matches!(
        open(custom_xml).plain_paragraph_copy_snapshot(),
        Err(Error::Refused(Refusal::UnsupportedDependency))
    ));

    for strict_relationship in [
        "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk",
    ] {
        let mut package = fixture(MAIN, document(&["plain"]));
        package
            .get_part_mut(&PackURI::new(format!("/{MAIN}")).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                strict_relationship.to_owned(),
                "missing.xml".to_owned(),
                "rStrictDependency".to_owned(),
                false,
            );
        assert!(matches!(
            open(bytes(&package)).plain_paragraph_copy_snapshot(),
            Err(Error::Refused(Refusal::UnsupportedDependency))
        ));
    }

    let signed = {
        let mut package = fixture(MAIN, document(&["plain"]));
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"origin".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        bytes(&package)
    };
    assert!(matches!(
        open(signed).plain_paragraph_copy_snapshot(),
        Err(Error::Refused(Refusal::SignatureInfrastructure))
    ));

    let mut macro_package = fixture(MAIN, document(&["plain"]));
    macro_package
        .get_part_mut(&PackURI::new(format!("/{MAIN}")).unwrap())
        .unwrap()
        .set_content_type(ct::WML_DOCUMENT_MACRO_MAIN.to_owned())
        .unwrap();
    assert!(matches!(
        open(bytes(&macro_package)).plain_paragraph_copy_snapshot(),
        Err(Error::Refused(Refusal::MacroEnabled))
    ));

    let hostile_path = bytes(&fixture("word/sub/document.xml", document(&["plain"])));
    assert!(matches!(
        open(hostile_path).plain_paragraph_copy_snapshot(),
        Err(Error::Refused(Refusal::MainDocumentShape))
    ));
}

#[test]
fn structured_paragraphs_sections_and_hostile_namespaces_are_refused() {
    let cases = [
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:hyperlink><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:bookmarkStart w:id="1" w:name="x"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:fldSimple w:instr="DATE"><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:drawing/></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:pict><w:txbxContent><w:p/></w:txbxContent></w:pict></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:x="urn:hostile"><w:body><w:p><w:r><w:t>x</w:t><x:unknown/></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        ),
    ];
    for xml in cases {
        assert!(matches!(
            open(bytes(&fixture(MAIN, xml.into_bytes()))).plain_paragraph_copy_snapshot(),
            Err(Error::Refused(
                Refusal::ComplexDocument | Refusal::ComplexParagraph
            ))
        ));
    }
}

#[test]
fn strict_and_ordinary_producer_namespace_declarations_are_supported() {
    let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    let strict_document = format!(
        r#"<w:document xmlns:w="{strict}" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><w:body><w:p><w:r><w:t>strict</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let package = open(bytes(&fixture_with_relationship(
        MAIN,
        strict_document,
        rt::STRICT_OFFICE_DOCUMENT,
    )));
    let mut edit = package.edit_plain_paragraph_copy().unwrap();
    edit.copy_plain_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    let commit = edit.commit();
    let mut output = Vec::new();
    package
        .publish_plain_paragraph_copy_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(reopened_texts(output), ["strict", "strict"]);

    let ordinary = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:r><w:t>ordinary</w:t></w:r></w:p></w:body></w:document>"#
    );
    assert_eq!(
        open(bytes(&fixture(MAIN, ordinary.into_bytes())))
            .plain_paragraph_copy_snapshot()
            .unwrap()
            .paragraph_count(),
        1
    );

    let escaped = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>a &amp; b</w:t></w:r></w:p></w:body></w:document>"#
    );
    assert_eq!(
        open(bytes(&fixture(MAIN, escaped.into_bytes())))
            .plain_paragraph_copy_snapshot()
            .unwrap()
            .paragraph_count(),
        1
    );
}

#[test]
fn empty_body_and_empty_or_sole_paragraph_slots_are_unambiguous() {
    let empty_body =
        format!(r#"<w:document xmlns:w="{W}"><w:body></w:body></w:document>"#).into_bytes();
    let snapshot = open(bytes(&fixture(MAIN, empty_body)))
        .plain_paragraph_copy_snapshot()
        .unwrap();
    assert_eq!(snapshot.paragraph_count(), 0);
    let mut edit = snapshot.edit();
    assert!(matches!(
        edit.copy_plain_paragraph(Position::new(0), Position::new(0)),
        Err(Error::OutOfBounds { kind: "source", .. })
    ));

    for before in [0, 1] {
        let xml = format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#)
            .into_bytes();
        let package = open(bytes(&fixture(MAIN, xml)));
        let mut edit = package.edit_plain_paragraph_copy().unwrap();
        edit.copy_plain_paragraph(Position::new(0), Position::new(before))
            .unwrap();
        let commit = edit.commit();
        assert_eq!(commit.projected().paragraph_count(), 2);
        assert_eq!(commit.effect_report().copied_bytes, b"<w:p/>".len());
    }
}

#[test]
fn enforced_protection_is_refused() {
    for owner in [
        r#"<w:documentProtection w:edit="readOnly" w:enforcement="1"/>"#,
        r#"<w:trackRevisions/>"#,
    ] {
        let mut package = fixture(MAIN, document(&["plain"]));
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/settings.xml").unwrap(),
                ct::WML_SETTINGS.to_owned(),
                format!(r#"<w:settings xmlns:w="{W}">{owner}</w:settings>"#).into_bytes(),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new(format!("/{MAIN}")).unwrap())
            .unwrap()
            .relate_to("settings.xml", rt::SETTINGS);
        assert!(matches!(
            open(bytes(&package)).plain_paragraph_copy_snapshot(),
            Err(Error::Refused(Refusal::Protection))
        ));
    }
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
            77,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn changed_source_version_is_rejected_before_publication_output() {
    let source = Arc::new(VersionedSource {
        bytes: bytes(&fixture(MAIN, document(&["a"]))),
        revision: AtomicU64::new(0),
    });
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_plain_paragraph_copy().unwrap();
    edit.copy_plain_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    let commit = edit.commit();
    source.revision.fetch_add(1, Ordering::SeqCst);
    let mut output = Vec::new();
    assert!(
        package
            .publish_plain_paragraph_copy_to_stream(&mut output, &commit)
            .is_err()
    );
    assert!(output.is_empty());
}

struct PartialSink {
    bytes: Vec<u8>,
    chunk: usize,
}

impl io::Write for PartialSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.chunk);
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ZeroSink;

impl io::Write for ZeroSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

impl io::Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted == self.limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let count = bytes.len().min(self.limit - self.accepted);
        self.accepted += count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn partial_sinks_complete_and_write_zero_fails_without_false_progress() {
    let source = bytes(&fixture(MAIN, document(&["a", "b"])));
    let package = open(source.clone());
    let mut edit = package.edit_plain_paragraph_copy().unwrap();
    edit.copy_plain_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let commit = edit.commit();
    let mut partial = PartialSink {
        bytes: Vec::new(),
        chunk: 7,
    };
    package
        .publish_plain_paragraph_copy_to_stream(&mut partial, &commit)
        .unwrap();
    assert_eq!(reopened_texts(partial.bytes), ["a", "b", "a"]);

    let package = open(source);
    let mut edit = package.edit_plain_paragraph_copy().unwrap();
    edit.copy_plain_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let commit = edit.commit();
    let Err(error) = package.publish_plain_paragraph_copy_to_stream(ZeroSink, &commit) else {
        panic!("write-zero sink must fail");
    };
    assert!(
        matches!(
            error,
            Error::Document(litchi_docx::Error::Opc(litchi_opc::OpcError::IoError(_)))
        ),
        "unexpected write-zero error: {error:?}"
    );

    let source = bytes(&fixture(MAIN, document(&["a", "b"])));
    let package = open(source);
    let mut edit = package.edit_plain_paragraph_copy().unwrap();
    edit.copy_plain_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let commit = edit.commit();
    let mut failing = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        package.publish_plain_paragraph_copy_to_stream(&mut failing, &commit),
        Err(Error::Document(litchi_docx::Error::Opc(
            litchi_opc::OpcError::IncompleteOutput { written: 128, .. }
        )))
    ));
}
