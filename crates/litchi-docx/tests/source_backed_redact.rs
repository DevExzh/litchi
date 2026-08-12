use std::io::{self, Write};
use std::sync::Arc;

use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_docx::{Error, Package, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, TargetMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/word/document.xml";
const KEEP: &str = "/word/keep.bin";

#[derive(Clone)]
struct Link {
    id: String,
    target: String,
}

fn fixture(document: Vec<u8>, links: Vec<Link>, non_main_link: bool) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new(MAIN).unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document,
    );
    for link in links {
        main.rels_mut()
            .try_add_relationship(
                rt::HYPERLINK.to_owned(),
                link.target,
                link.id,
                TargetMode::External,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(KEEP).unwrap(),
            "application/octet-stream".to_owned(),
            b"untouched\0opaque\xffbytes".to_vec(),
        )))
        .unwrap();
    if non_main_link {
        let mut header = BlobPart::new(
            PackURI::new("/word/header1.xml").unwrap(),
            ct::WML_HEADER.to_owned(),
            format!(r#"<w:hdr xmlns:w="{W}"/>"#).into_bytes(),
        );
        header
            .rels_mut()
            .try_add_relationship(
                rt::HYPERLINK.to_owned(),
                "https://shared.invalid/".to_owned(),
                "rShared".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        package.try_add_part(Box::new(header)).unwrap();
    }
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn sign(bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            b"<origin/>".to_vec(),
        )))
        .unwrap();
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    PackageWriter::to_bytes(&package).unwrap()
}

fn add_external_owner(bytes: &[u8], package_level: bool, relationship_type: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    if package_level {
        package
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_owned(),
                "https://one.invalid/".to_owned(),
                "rExternalOwner".to_owned(),
                TargetMode::External,
            )
            .unwrap();
    } else {
        let mut header = BlobPart::new(
            PackURI::new("/word/header2.xml").unwrap(),
            ct::WML_HEADER.to_owned(),
            format!(r#"<w:hdr xmlns:w="{W}"/>"#).into_bytes(),
        );
        header
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_owned(),
                "https://one.invalid/".to_owned(),
                "rExternalOwner".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        package.try_add_part(Box::new(header)).unwrap();
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn open(bytes: Vec<u8>) -> source_backed::Package {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    source_backed::Package::from_read_at(source).unwrap()
}

struct CollidingSource(Vec<u8>);

impl ReadAt for CollidingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.0.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if offset >= self.0.len() {
            return Ok(0);
        }
        let count = output.len().min(self.0.len() - offset);
        output[..count].copy_from_slice(&self.0[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(8_087, 0))
    }
}

fn open_colliding(bytes: Vec<u8>) -> source_backed::Package {
    let source: Arc<dyn ReadAt> = Arc::new(CollidingSource(bytes));
    source_backed::Package::from_read_at(source).unwrap()
}

fn link(id: &str, target: &str) -> Link {
    Link {
        id: id.to_owned(),
        target: target.to_owned(),
    }
}

fn document() -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rFirst"><w:r><w:t>first</w:t></w:r></w:hyperlink><w:r><w:t>|</w:t></w:r><w:hyperlink r:id="rMiddle"><w:r><w:t>middle</w:t></w:r></w:hyperlink><w:r><w:t>|</w:t></w:r><w:hyperlink r:id="rLast"><w:r><w:t>last</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn one_document() -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rFirst"><w:r><w:t>visible</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn contains_bytes(haystack: &[u8], needle: &[u8]) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}

#[test]
fn inventories_and_irreversibly_redacts_duplicate_target_ids_first_middle_last() {
    let bytes = fixture(
        document(),
        vec![
            link("rFirst", "https://remove.invalid/"),
            link("rMiddle", "https://keep.invalid/"),
            link("rLast", "https://remove.invalid/"),
        ],
        false,
    );
    let package = open(bytes);
    let snapshot = package.external_hyperlink_redaction_snapshot().unwrap();
    assert_eq!(snapshot.relationships().len(), 3);
    assert_eq!(snapshot.relationships()[0].relationship_id(), "rFirst");
    assert_eq!(snapshot.relationships()[2].relationship_id(), "rMiddle");
    assert!(
        snapshot
            .relationships()
            .iter()
            .all(|item| item.wrapper_count() == 1)
    );
    let before = snapshot.document_xml().to_vec();
    let plan = snapshot
        .plan_target_urls(&["https://remove.invalid/", "https://remove.invalid/"])
        .unwrap();
    assert_eq!(snapshot.document_xml(), before);
    assert_eq!(plan.effect_report().selected_targets(), 1);
    assert_eq!(plan.effect_report().removed_relationships(), 2);
    assert_eq!(plan.effect_report().unwrapped_hyperlinks(), 2);
    let commit = plan.apply().unwrap();
    assert!(!commit.patch().is_noop());

    let mut output = Vec::new();
    package
        .publish_external_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    assert!(!contains_bytes(&output, b"https://remove.invalid/"));
    let reopened = OpcPackage::from_bytes(&output).unwrap();
    let main = reopened.main_document_part().unwrap();
    assert!(main.rels().get("rFirst").is_none());
    assert!(main.rels().get("rLast").is_none());
    assert_eq!(
        main.rels().get("rMiddle").unwrap().target_ref(),
        "https://keep.invalid/"
    );
    assert_eq!(
        reopened
            .get_part(&PackURI::new(KEEP).unwrap())
            .unwrap()
            .blob(),
        b"untouched\0opaque\xffbytes"
    );
    let xml = std::str::from_utf8(main.blob()).unwrap();
    assert!(!xml.contains("rFirst"));
    assert!(xml.contains("<w:hyperlink r:id=\"rMiddle\">"));
    assert!(!xml.contains("rLast"));
    assert!(xml.contains("<w:t>first</w:t>"));
    assert!(xml.contains("<w:t>middle</w:t>"));
    assert!(xml.contains("<w:t>last</w:t>"));
}

#[test]
fn zero_is_an_exact_noop_and_one_removes_the_only_link() {
    let bytes = fixture(
        format!(r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rOne"><w:r><w:t>visible</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#).into_bytes(),
        vec![link("rOne", "https://one.invalid/")],
        false,
    );
    let package = open(bytes.clone());
    let noop = package
        .external_hyperlink_redaction_snapshot()
        .unwrap()
        .plan_target_urls(&[])
        .unwrap()
        .apply()
        .unwrap();
    assert!(noop.patch().is_noop());
    let mut identical = Vec::new();
    package
        .publish_external_hyperlink_redaction_to_stream(&mut identical, &noop)
        .unwrap();
    assert_eq!(identical, bytes);

    let package = open(bytes);
    let commit = package
        .plan_external_hyperlink_redaction(&["https://one.invalid/"])
        .unwrap()
        .apply()
        .unwrap();
    let mut output = Vec::new();
    package
        .publish_external_hyperlink_redaction_to_stream(&mut output, &commit)
        .unwrap();
    let reopened = OpcPackage::from_bytes(&output).unwrap();
    let main = reopened.main_document_part().unwrap();
    assert!(main.rels().is_empty());
    let reopened = Package::from_opc_package(reopened).unwrap();
    assert_eq!(reopened.document().unwrap().text().unwrap(), "visible");
}

#[test]
fn selection_bounds_unknown_targets_and_foreign_sources_refuse_before_output() {
    let links = (0..65)
        .map(|index| link(&format!("r{index}"), &format!("https://{index}.invalid/")))
        .collect::<Vec<_>>();
    let wrappers = (0..65)
        .map(|index| format!(r#"<w:hyperlink r:id="r{index}"/>"#))
        .collect::<String>();
    let bytes = fixture(
        format!(r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p>{wrappers}</w:p></w:body></w:document>"#).into_bytes(),
        links,
        false,
    );
    let package = open(bytes.clone());
    let snapshot = package.external_hyperlink_redaction_snapshot().unwrap();
    let targets = (0..65)
        .map(|index| format!("https://{index}.invalid/"))
        .collect::<Vec<_>>();
    let refs = targets.iter().map(String::as_str).collect::<Vec<_>>();
    assert!(matches!(
        snapshot.plan_target_urls(&refs),
        Err(Error::ExternalHyperlinkRedactionLimit {
            resource: "target URL selectors",
            maximum: 64,
            actual: 65,
        })
    ));
    let exact = open(bytes.clone())
        .external_hyperlink_redaction_snapshot()
        .unwrap();
    assert!(exact.plan_target_urls(&refs[..64]).is_ok());
    let oversized = "x".repeat(64 * 4 * 1024 + 1);
    assert!(matches!(
        snapshot.plan_target_urls(&[oversized.as_str()]),
        Err(Error::ExternalHyperlinkRedactionLimit {
            resource: "target URL selector bytes",
            maximum: 262_144,
            actual: 262_145,
        })
    ));
    assert!(
        snapshot
            .plan_target_urls(&["https://missing.invalid/"])
            .is_err()
    );

    let commit = snapshot
        .plan_target_urls(&["https://0.invalid/"])
        .unwrap()
        .apply()
        .unwrap();
    let foreign = open(bytes);
    let mut output = Vec::new();
    assert!(matches!(
        foreign.publish_external_hyperlink_redaction_to_stream(&mut output, &commit),
        Err(Error::ExternalHyperlinkRedactionConflict)
    ));
    assert!(output.is_empty());
}

#[test]
fn full_artifact_fingerprint_defeats_colliding_custom_source_versions() {
    let links = vec![link("rFirst", "https://one.invalid/")];
    let original = fixture(one_document(), links, false);
    let package = open_colliding(original.clone());
    let commit = package
        .plan_external_hyperlink_redaction(&["https://one.invalid/"])
        .unwrap()
        .apply()
        .unwrap();

    let mut foreign = OpcPackage::from_bytes(&original).unwrap();
    foreign
        .get_part_mut(&PackURI::new(KEEP).unwrap())
        .unwrap()
        .set_blob(b"different unrelated artifact bytes".to_vec());
    let foreign = PackageWriter::to_bytes(&foreign).unwrap();
    let colliding = open_colliding(foreign);
    let mut output = Vec::new();
    assert!(matches!(
        colliding.publish_external_hyperlink_redaction_to_stream(&mut output, &commit),
        Err(Error::ExternalHyperlinkRedactionConflict)
    ));
    assert!(output.is_empty());
}

#[test]
fn refuses_non_main_shared_owners_fields_and_partial_sink_is_typed() {
    let one = vec![link("rFirst", "https://shared.invalid/")];
    assert!(matches!(
        open(fixture(one_document(), one.clone(), true)).external_hyperlink_redaction_snapshot(),
        Err(Error::UnsafeEdit {
            operation: "external_hyperlink_redaction",
            ..
        })
    ));
    let field = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:fldSimple w:instr="HYPERLINK https://field.invalid/"/><w:hyperlink r:id="rFirst"/></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open(fixture(field, one.clone(), false)).external_hyperlink_redaction_snapshot(),
        Err(Error::UnsafeEdit {
            operation: "external_hyperlink_redaction",
            ..
        })
    ));
    let unknown_relationship_attribute = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:drawing r:embed="rFirst"/><w:hyperlink r:id="rFirst"/></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert!(matches!(
        open(fixture(unknown_relationship_attribute, one.clone(), false,))
            .external_hyperlink_redaction_snapshot(),
        Err(Error::UnsafeEdit {
            operation: "external_hyperlink_redaction",
            ..
        })
    ));
    let base = fixture(one_document(), one.clone(), false);
    for foreign_owner in [
        add_external_owner(&base, false, rt::IMAGE),
        add_external_owner(&base, true, rt::HYPERLINK),
    ] {
        assert!(matches!(
            open(foreign_owner).external_hyperlink_redaction_snapshot(),
            Err(Error::UnsafeEdit {
                operation: "external_hyperlink_redaction",
                ..
            })
        ));
    }

    let package = open(fixture(one_document(), one, false));
    let commit = package
        .plan_external_hyperlink_redaction(&["https://shared.invalid/"])
        .unwrap()
        .apply()
        .unwrap();
    let mut sink = FailingSink {
        accepted: 0,
        maximum: 128,
    };
    assert!(matches!(
        package.publish_external_hyperlink_redaction_to_stream(&mut sink, &commit),
        Err(Error::Opc(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);

    let signed = sign(&fixture(
        one_document(),
        vec![link("rFirst", "https://shared.invalid/")],
        false,
    ));
    assert!(matches!(
        open(signed).external_hyperlink_redaction_snapshot(),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
}

struct FailingSink {
    accepted: usize,
    maximum: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.maximum {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.maximum - self.accepted);
        self.accepted += written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
