#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused publication assertions intentionally panic on fixture errors"
)]

//! Change 0754: a part replaced through `Part::set_blob_verified` carries the
//! proof that its exact allocation passed the source publication audit, and
//! the eager writer skips its own audit of that part only while the bytes it
//! publishes are exactly the proof's. A later replacement, a copy, or a custom
//! part that pairs a borrowed proof with other bytes is audited and refused
//! exactly as before.

use std::sync::Arc;

use litchi_opc::part::PayloadHandle;
use litchi_opc::{OpcError, OpcPackage, PackURI, PackageWriter, Part, Relationships, XmlPart};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use xml_minifier::audit::{Limits, VerifiedSource, verify_source};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT: &str = "/word/document.xml";
const CONTENT_TYPE: &str = "application/xml";

const ORIGINAL: &[u8] = b"<?xml version=\"1.0\"?>\n<document>\n  <before/>\n</document>\n";
const EDITED: &[u8] = b"<?xml version=\"1.0\"?>\n<document>\n  <after/>\n</document>\n";
/// Well-formed to the tokenizer, refused by the source audit: an undeclared
/// prefix.
const REFUSED: &[u8] = b"<document><x:undeclared/></document>";

fn archive(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="{CONTENT_TYPE}"/></Types>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .unwrap();
    writer.write_stored("word/document.xml", document).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn document_uri() -> PackURI {
    PackURI::new(DOCUMENT).unwrap()
}

fn package() -> OpcPackage {
    OpcPackage::from_bytes(&archive(ORIGINAL)).unwrap()
}

fn proof(bytes: &[u8]) -> VerifiedSource {
    VerifiedSource::verify(Arc::new(bytes.to_vec()), Limits::default()).unwrap()
}

fn published_document(package: &OpcPackage) -> Vec<u8> {
    let output = PackageWriter::to_bytes(package).unwrap();
    ArchiveReader::new(&output)
        .unwrap()
        .read("word/document.xml")
        .unwrap()
}

fn assert_refused(package: &OpcPackage, bytes: &[u8]) {
    match PackageWriter::to_bytes(package) {
        Err(OpcError::XmlPublication { part, source }) => {
            assert_eq!(part, DOCUMENT);
            assert_eq!(source, verify_source(bytes, Limits::default()).unwrap_err());
        },
        other => panic!("expected the writer's audit to refuse, got {other:?}"),
    }
}

#[test]
fn a_verified_payload_publishes_exactly_what_the_unproven_route_publishes() {
    let proven_bytes = proof(EDITED);
    let mut proven = package();
    proven
        .get_part_mut(&document_uri())
        .unwrap()
        .set_blob_verified(proven_bytes.clone());
    // The built-in part adopted the proof's own allocation.
    assert!(std::ptr::eq(
        proven.get_part(&document_uri()).unwrap().blob(),
        proven_bytes.bytes().as_slice()
    ));

    let mut unproven = package();
    unproven
        .get_part_mut(&document_uri())
        .unwrap()
        .set_blob(EDITED.to_vec());

    assert_eq!(published_document(&proven), EDITED);
    assert_eq!(
        PackageWriter::to_bytes(&proven).unwrap(),
        PackageWriter::to_bytes(&unproven).unwrap()
    );
}

#[test]
fn a_payload_replaced_after_its_proof_is_audited() {
    let mut package = package();
    let part = package.get_part_mut(&document_uri()).unwrap();
    part.set_blob_verified(proof(EDITED));
    part.set_blob(REFUSED.to_vec());
    assert_refused(&package, REFUSED);

    // `set_blob_shared` drops the proof too.
    let part = package.get_part_mut(&document_uri()).unwrap();
    part.set_blob_verified(proof(EDITED));
    part.set_blob_shared(Arc::new(REFUSED.to_vec()));
    assert_refused(&package, REFUSED);
}

/// A custom part that hands the writer a payload handle borrowed from a
/// verified built-in part, but publishes bytes of its own.
#[derive(Clone)]
struct Substituting {
    partname: PackURI,
    handle: PayloadHandle,
    bytes: Arc<Vec<u8>>,
    rels: Relationships,
}

impl Substituting {
    fn new(handle: PayloadHandle, bytes: Arc<Vec<u8>>) -> Self {
        Self {
            partname: document_uri(),
            handle,
            bytes,
            rels: Relationships::new(DOCUMENT.to_owned()),
        }
    }
}

impl Part for Substituting {
    fn blob(&self) -> &[u8] {
        &self.bytes
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.bytes)
    }

    fn content_type(&self) -> &str {
        CONTENT_TYPE
    }

    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn payload_handle(&self) -> PayloadHandle {
        self.handle.clone()
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.bytes = Arc::new(blob);
    }
}

#[test]
fn a_custom_part_cannot_borrow_a_proof_for_other_bytes() {
    let proven = proof(EDITED);
    let mut donor = XmlPart::new(document_uri(), CONTENT_TYPE.to_owned(), Vec::new());
    donor.set_blob_verified(proven.clone());
    let handle = donor.payload_handle();

    // Refused bytes behind a borrowed proof are audited and refused.
    let mut package = package();
    package.add_part(Box::new(Substituting::new(
        handle.clone(),
        Arc::new(REFUSED.to_vec()),
    )));
    assert_refused(&package, REFUSED);

    // So are bytes equal to the proven ones but in another allocation: they
    // are audited (and pass), which the writer cannot tell apart from the
    // proven route except by their content, so this checks the output.
    let mut package = self::package();
    package.add_part(Box::new(Substituting::new(
        handle.clone(),
        Arc::new(EDITED.to_vec()),
    )));
    assert_eq!(published_document(&package), EDITED);

    // The proof's own allocation publishes.
    let mut package = self::package();
    package.add_part(Box::new(Substituting::new(
        handle,
        Arc::clone(proven.bytes()),
    )));
    assert_eq!(published_document(&package), EDITED);
}

#[test]
fn a_custom_part_without_its_own_verified_storage_drops_the_proof() {
    // The trait's default `set_blob_verified` adopts the bytes and drops the
    // proof, so a custom part is always audited.
    let mut package = package();
    let mut custom = Substituting::new(
        XmlPart::new(document_uri(), CONTENT_TYPE.to_owned(), Vec::new()).payload_handle(),
        Arc::new(Vec::new()),
    );
    custom.set_blob_verified(proof(EDITED));
    assert_eq!(custom.blob(), EDITED);
    package.add_part(Box::new(custom));
    assert_eq!(published_document(&package), EDITED);
}
