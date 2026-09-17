#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "payload-reuse fixtures and assertions panic on malformed test setup"
)]

use super::*;
use crate::part::{BlobPart, Part};
use crate::{PackageWriter, ReadResource, TargetMode};
use soapberry_zip::office::StreamingArchiveWriter;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

const XML_URI: &str = "/word/document.xml";
const BINARY_URI: &str = "/word/media/image.bin";
const XML_PAYLOAD: &[u8] =
    br#"<?xml version="1.0"?><document><body><p>payload reuse</p></body></document>"#;
const BINARY_PAYLOAD: &[u8] = b"payload-reuse-binary-sentinel";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).expect("test URI must be valid")
}

fn payload_archive(comment: &[u8], binary_payload: &[u8]) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated(
            "[Content_Types].xml",
            br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/xml"/></Types>"#,
        )
        .expect("content types fixture");
    writer
        .write_deflated(
            "_rels/.rels",
            br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRoot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        )
        .expect("package relationships fixture");
    writer
        .write_deflated("word/document.xml", XML_PAYLOAD)
        .expect("XML payload fixture");
    writer
        .write_deflated(
            "word/_rels/document.xml.rels",
            br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.bin"/></Relationships>"#,
        )
        .expect("part relationships fixture");
    writer
        .write_stored("word/media/image.bin", binary_payload)
        .expect("binary payload fixture");
    with_eocd_comment(
        writer.finish_to_bytes().expect("payload archive fixture"),
        comment,
    )
}

fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
    let comment_length = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
    let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
    assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
    archive[eocd + 20..eocd + 22].copy_from_slice(&comment_length.to_le_bytes());
    archive.extend_from_slice(comment);
    archive
}

/// Assert what the reused package holds for one part.
///
/// `expect_donation` says whether the donor's allocation is adoptable. Since
/// change 0661 a donor opened through an ordinary owned door holds its
/// payloads in the retained source archive and materializes one through the
/// positional reader on first access. That reader sizes a **stored** member's
/// buffer one byte above its length, so the donation guard — which refuses a
/// donor allocation larger than the freshly decoded one, so that reuse can
/// never raise retention — correctly declines it. A deflated member is sized
/// identically on both paths and is still donated.
fn assert_payload_internals(
    donor: &OpcPackage,
    reused: &OpcPackage,
    ordinary: &OpcPackage,
    partname: &PackURI,
    expect_donation: bool,
) {
    let donor_blob = donor.get_part(partname).expect("donor part").blob_arc();
    let reused_blob = reused.get_part(partname).expect("reused part").blob_arc();
    let decoded_blob = ordinary
        .get_part(partname)
        .expect("ordinary decoded part")
        .blob_arc();

    assert_eq!(
        Arc::ptr_eq(&donor_blob, &reused_blob),
        expect_donation,
        "donation for {partname}"
    );
    assert_eq!(donor_blob.as_slice(), reused_blob.as_slice());
    assert!(reused_blob.capacity() <= decoded_blob.capacity());

    let preserved = reused
        .preservation
        .as_ref()
        .expect("owned reuse must retain preservation provenance")
        .parts
        .get(partname)
        .expect("preserved part payload");
    assert!(Arc::ptr_eq(
        &reused_blob,
        preserved
            .blob
            .decoded()
            .expect("eager reuse keeps a payload")
    ));

    if partname == &pack(XML_URI) {
        let source_xml = reused
            .source_xml_parts
            .get(partname)
            .expect("XML source payload");
        assert!(Arc::ptr_eq(
            &reused_blob,
            source_xml.decoded().expect("eager reuse keeps a payload")
        ));
    }
}

#[test]
fn matching_binary_and_xml_payloads_share_all_owned_views_and_target_output() {
    let donor_bytes = payload_archive(b"donor archive", BINARY_PAYLOAD);
    let target_bytes = payload_archive(b"target archive", BINARY_PAYLOAD);
    let mut donor = OpcPackage::from_vec(donor_bytes).expect("open donor");

    // Donor graph state is deliberately different. Reuse may borrow payloads,
    // but it must parse all target metadata independently.
    donor
        .rels_mut()
        .try_add_relationship(
            "urn:test:donor-root".to_owned(),
            "https://donor.invalid/root".to_owned(),
            "rIdDonorRoot".to_owned(),
            TargetMode::External,
        )
        .expect("donor root metadata");
    donor
        .get_part_mut(&pack(XML_URI))
        .expect("donor XML part")
        .rels_mut()
        .try_add_relationship(
            "urn:test:donor-part".to_owned(),
            "https://donor.invalid/part".to_owned(),
            "rIdDonorPart".to_owned(),
            TargetMode::External,
        )
        .expect("donor part metadata");
    donor.set_save_options(SaveOptions {
        fonts: FontEmbedding::Full,
    });

    let ordinary = OpcPackage::from_vec(target_bytes.clone()).expect("ordinary target parse");
    let reused =
        OpcPackage::from_vec_reusing_payloads(target_bytes.clone(), ReadLimits::default(), &donor)
            .expect("matching payloads should be reusable");

    assert_payload_internals(&donor, &reused, &ordinary, &pack(XML_URI), true);
    assert_payload_internals(&donor, &reused, &ordinary, &pack(BINARY_URI), false);
    assert_eq!(reused.get_part(&pack(XML_URI)).unwrap().blob(), XML_PAYLOAD);
    assert_eq!(
        reused.get_part(&pack(BINARY_URI)).unwrap().blob(),
        BINARY_PAYLOAD
    );
    assert!(reused.rels().get("rIdDonorRoot").is_none());
    assert!(
        reused
            .get_part(&pack(XML_URI))
            .unwrap()
            .rels()
            .get("rIdDonorPart")
            .is_none()
    );
    assert_eq!(reused.save_options().fonts, FontEmbedding::None);
    assert!(reused.is_unmodified_owned_source());
    assert_eq!(PackageWriter::to_bytes(&reused).unwrap(), target_bytes);
}

#[test]
fn mismatched_bytes_and_larger_donor_capacity_keep_decoded_payload() {
    let donor_bytes = payload_archive(b"donor", BINARY_PAYLOAD);
    let target_bytes = payload_archive(b"target", BINARY_PAYLOAD);
    let binary = pack(BINARY_URI);
    let target = OpcPackage::from_vec(target_bytes.clone()).expect("ordinary target parse");
    let decoded = target.get_part(&binary).expect("target binary part");
    let decoded_payload = decoded.blob().to_vec();
    let decoded_capacity = decoded.blob_arc().capacity();

    let mut mismatched = decoded_payload.clone();
    mismatched[0] ^= 0xff;
    let mismatched_blob = Arc::new(mismatched);
    let mut byte_mismatch_donor = OpcPackage::from_vec(donor_bytes.clone()).expect("open donor");
    byte_mismatch_donor.add_part(Box::new(BlobPart::new_shared(
        binary.clone(),
        "application/octet-stream".to_owned(),
        Arc::clone(&mismatched_blob),
    )));
    let byte_mismatch = OpcPackage::from_vec_reusing_payloads(
        target_bytes.clone(),
        ReadLimits::default(),
        &byte_mismatch_donor,
    )
    .expect("mismatched donor bytes must be ignored");
    let byte_mismatch_part = byte_mismatch.get_part(&binary).unwrap();
    assert_eq!(byte_mismatch_part.blob(), decoded_payload.as_slice());
    assert!(!Arc::ptr_eq(
        &byte_mismatch_part.blob_arc(),
        &mismatched_blob
    ));

    let oversized_capacity = decoded_capacity
        .checked_add(1)
        .expect("test payload capacity has room for a larger donor");
    let mut oversized = Vec::with_capacity(oversized_capacity);
    oversized.extend_from_slice(&decoded_payload);
    assert!(oversized.capacity() > decoded_capacity);
    let oversized_blob = Arc::new(oversized);
    let mut capacity_donor = OpcPackage::from_vec(donor_bytes).expect("open donor");
    capacity_donor.add_part(Box::new(BlobPart::new_shared(
        binary.clone(),
        "application/octet-stream".to_owned(),
        Arc::clone(&oversized_blob),
    )));
    let capacity_mismatch =
        OpcPackage::from_vec_reusing_payloads(target_bytes, ReadLimits::default(), &capacity_donor)
            .expect("an oversized donor allocation must be ignored");
    let capacity_part = capacity_mismatch.get_part(&binary).unwrap();
    assert_eq!(capacity_part.blob(), decoded_payload.as_slice());
    assert!(!Arc::ptr_eq(&capacity_part.blob_arc(), &oversized_blob));
}

#[test]
fn custom_donor_blob_arc_cannot_inject_payload_bytes() {
    let donor_bytes = payload_archive(b"donor", BINARY_PAYLOAD);
    let target_bytes = payload_archive(b"target", BINARY_PAYLOAD);
    let binary = pack(BINARY_URI);
    let mut donor = OpcPackage::from_vec(donor_bytes).expect("open donor");
    let visible = BINARY_PAYLOAD.to_vec();
    let injected = Arc::new(b"custom donor payload must not escape".to_vec());
    donor.add_part(Box::new(MismatchedBlobArcPart::new(
        binary.clone(),
        "application/octet-stream".to_owned(),
        visible,
        Arc::clone(&injected),
    )));

    let reused = OpcPackage::from_vec_reusing_payloads(target_bytes, ReadLimits::default(), &donor)
        .expect("custom donor metadata must not block target parsing");
    let part = reused.get_part(&binary).expect("target binary part");
    assert_eq!(part.blob(), BINARY_PAYLOAD);
    assert!(!Arc::ptr_eq(&part.blob_arc(), &injected));
    assert!(
        !reused
            .get_part(&binary)
            .unwrap()
            .blob()
            .eq(injected.as_slice())
    );
}

#[test]
fn reused_payloads_isolate_donor_output_and_clone_mutations() {
    let donor_bytes = payload_archive(b"donor", BINARY_PAYLOAD);
    let target_bytes = payload_archive(b"target", BINARY_PAYLOAD);
    let binary = pack(BINARY_URI);
    let donor = OpcPackage::from_vec(donor_bytes).expect("open donor");
    let reused =
        OpcPackage::from_vec_reusing_payloads(target_bytes.clone(), ReadLimits::default(), &donor)
            .expect("reuse target payloads");
    assert!(donor.is_unmodified_owned_source());
    assert!(reused.is_unmodified_owned_source());

    let mut output = PackageWriter::to_bytes(&reused).expect("serialize reused output");
    output[0] ^= 0xff;
    assert_eq!(PackageWriter::to_bytes(&reused).unwrap(), target_bytes);

    let mut target_clone = reused.clone();
    target_clone
        .get_part_mut(&binary)
        .expect("clone binary part")
        .set_blob(b"clone mutation".to_vec());
    assert!(!target_clone.is_unmodified_owned_source());
    assert_eq!(reused.get_part(&binary).unwrap().blob(), BINARY_PAYLOAD);
    assert!(reused.is_unmodified_owned_source());

    let mut donor_clone = donor.clone();
    donor_clone
        .get_part_mut(&binary)
        .expect("donor clone binary part")
        .set_blob(b"donor mutation".to_vec());
    assert!(!donor_clone.is_unmodified_owned_source());
    assert_eq!(donor.get_part(&binary).unwrap().blob(), BINARY_PAYLOAD);
    assert_eq!(reused.get_part(&binary).unwrap().blob(), BINARY_PAYLOAD);
    assert!(donor.is_unmodified_owned_source());
}

#[test]
fn custom_donor_cannot_claim_an_arc_inconsistent_with_its_visible_payload() {
    let binary = pack(BINARY_URI);
    let offered = Arc::new(BINARY_PAYLOAD.to_vec());
    let mut donor = OpcPackage::new();
    donor.add_part(Box::new(MismatchedBlobArcPart::new(
        binary.clone(),
        "application/octet-stream".to_owned(),
        b"different visible payload".to_vec(),
        Arc::clone(&offered),
    )));
    let bytes = payload_archive(b"target", BINARY_PAYLOAD);
    let reused =
        OpcPackage::from_vec_reusing_payloads(bytes.clone(), ReadLimits::default(), &donor)
            .expect("inconsistent donation must fall back to independently decoded storage");
    let payload = reused.get_part(&binary).expect("target binary").blob_arc();
    assert_eq!(payload.as_slice(), BINARY_PAYLOAD);
    assert!(!Arc::ptr_eq(&payload, &offered));
    let mut output = Vec::new();
    reused.to_stream(&mut output).expect("exact publication");
    assert_eq!(output, bytes);
}

#[test]
fn input_limit_is_refused_before_donor_payload_callbacks() {
    let donor_bytes = payload_archive(b"donor", BINARY_PAYLOAD);
    let target_bytes = payload_archive(b"target", BINARY_PAYLOAD);
    let calls = Arc::new(AtomicUsize::new(0));
    let mut donor = OpcPackage::from_vec(donor_bytes).expect("open donor");
    donor.add_part(Box::new(CallbackPart::new(
        pack(BINARY_URI),
        Arc::clone(&calls),
    )));
    OpcPackage::from_vec_reusing_payloads(target_bytes.clone(), ReadLimits::default(), &donor)
        .expect("valid input must reach the matching donor callback");
    assert!(calls.load(Ordering::SeqCst) > 0);
    calls.store(0, Ordering::SeqCst);
    let maximum = u64::try_from(target_bytes.len() - 1).expect("fixture length fits u64");
    let limits = ReadLimits::builder()
        .max_input_bytes(maximum)
        .expect("input limit")
        .build()
        .expect("input limit profile");

    let error = OpcPackage::from_vec_reusing_payloads(target_bytes, limits, &donor)
        .expect_err("input limit must reject the target archive");
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::InputBytes,
            ..
        }
    ));
    assert_eq!(calls.load(Ordering::SeqCst), 0);
}

#[test]
fn malformed_reuse_input_keeps_the_typed_zip_error() {
    let donor =
        OpcPackage::from_vec(payload_archive(b"donor", BINARY_PAYLOAD)).expect("open donor");
    let error = OpcPackage::from_vec_reusing_payloads(
        b"not an OPC ZIP".to_vec(),
        ReadLimits::default(),
        &donor,
    )
    .expect_err("malformed input must be rejected");
    assert!(matches!(error, OpcError::ZipError(_)));
}

#[derive(Clone, Debug)]
struct MismatchedBlobArcPart {
    partname: PackURI,
    content_type: String,
    visible: Vec<u8>,
    arc: Arc<Vec<u8>>,
    rels: Relationships,
}

impl MismatchedBlobArcPart {
    fn new(partname: PackURI, content_type: String, visible: Vec<u8>, arc: Arc<Vec<u8>>) -> Self {
        Self {
            rels: Relationships::for_source(&partname),
            partname,
            content_type,
            visible,
            arc,
        }
    }
}

impl Part for MismatchedBlobArcPart {
    fn blob(&self) -> &[u8] {
        &self.visible
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.arc)
    }

    fn content_type(&self) -> &str {
        &self.content_type
    }

    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.visible = blob;
    }

    fn set_content_type(&mut self, content_type: String) -> Result<()> {
        self.content_type = content_type;
        Ok(())
    }
}

#[derive(Clone, Debug)]
struct CallbackPart {
    inner: BlobPart,
    calls: Arc<AtomicUsize>,
}

impl CallbackPart {
    fn new(partname: PackURI, calls: Arc<AtomicUsize>) -> Self {
        Self {
            inner: BlobPart::new(
                partname,
                "application/octet-stream".to_owned(),
                b"callback payload".to_vec(),
            ),
            calls,
        }
    }
}

impl Part for CallbackPart {
    fn blob(&self) -> &[u8] {
        self.inner.blob()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        self.calls.fetch_add(1, Ordering::SeqCst);
        self.inner.blob_arc()
    }

    fn content_type(&self) -> &str {
        self.inner.content_type()
    }

    fn partname(&self) -> &PackURI {
        self.inner.partname()
    }

    fn rels(&self) -> &Relationships {
        self.inner.rels()
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        self.inner.rels_mut()
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.inner.set_blob(blob);
    }

    fn set_content_type(&mut self, content_type: String) -> Result<()> {
        self.inner.set_content_type(content_type)
    }
}
