use std::io;
use std::io::Write;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{Position, ReadAt, SourceVersion};
use litchi_docx::alt::{Data, MAX_DATA_BYTES};
use litchi_docx::document::{Snapshot, TransactionError};
use litchi_docx::{Error, Package, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::rel::TargetMode;
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part};

const MAIN_DOCUMENT: &str = "word/document.xml";
const UNUSED_PART: &str = "word/000-unused.bin";
const TRAILING_PART: &str = "word/zzz-trailing.bin";
const ALT_TARGET: &str = "word/altChunk.html";
const ALT_VENDOR_PART: &str = "word/vendor-alt.bin";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

struct ObservedSource {
    bytes: Vec<u8>,
    main_payload: std::ops::Range<usize>,
    unused_payload: std::ops::Range<usize>,
    main_reads: AtomicUsize,
    unused_reads: AtomicUsize,
    revision: AtomicU64,
}

impl ObservedSource {
    fn new(
        bytes: Vec<u8>,
        main_payload: std::ops::Range<usize>,
        unused_payload: std::ops::Range<usize>,
    ) -> Self {
        Self {
            bytes,
            main_payload,
            unused_payload,
            main_reads: AtomicUsize::new(0),
            unused_reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for ObservedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let requested = offset..end;
        if overlaps(&requested, &self.main_payload) {
            self.main_reads.fetch_add(1, Ordering::SeqCst);
        }
        if overlaps(&requested, &self.unused_payload) {
            self.unused_reads.fetch_add(1, Ordering::SeqCst);
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(17, self.revision.load(Ordering::SeqCst)))
    }
}

fn overlaps(left: &std::ops::Range<usize>, right: &std::ops::Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
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
        if offset + 46 + name_length + extra_length > zip.len() {
            continue;
        }
        if &zip[offset + 46..offset + 46 + name_length] == name {
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
            assert!(data_start + compressed_size <= zip.len());
            return data_start..data_start + compressed_size;
        }
    }
    panic!("ZIP member {name:?} was not found")
}

fn incompressible_bytes() -> Vec<u8> {
    let mut state = 0x5EED_C0DE_u32;
    (0..1_048_576)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            (state >> 24) as u8
        })
        .collect()
}

fn fixture_for_document(
    document: Vec<u8>,
    main_relationship: &str,
    signed: bool,
) -> (Vec<u8>, std::ops::Range<usize>, std::ops::Range<usize>) {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{MAIN_DOCUMENT}")).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_string(),
            document,
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED_PART}")).unwrap(),
            "application/octet-stream".to_string(),
            incompressible_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{TRAILING_PART}")).unwrap(),
            "application/octet-stream".to_string(),
            incompressible_bytes(),
        )))
        .unwrap();
    package.relate_to(MAIN_DOCUMENT, main_relationship);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_string(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    let zip = PackageWriter::to_bytes(&package).unwrap();
    let main = payload_range(&zip, MAIN_DOCUMENT);
    let unused = payload_range(&zip, UNUSED_PART);
    (zip, main, unused)
}

fn fixture() -> (Vec<u8>, std::ops::Range<usize>, std::ops::Range<usize>) {
    fixture_for_document(
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:x="urn:litchi:test"><w:body><w:p><w:r><w:t>alpha</w:t><x:opaque value="preserve"/></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p></w:body></w:document>"#
        )
        .into_bytes(),
        rt::OFFICE_DOCUMENT,
        false,
    )
}

fn alt_chunk_fixture(signed: bool, shared_target: bool) -> Vec<u8> {
    alt_chunk_fixture_with_layout(signed, shared_target, false)
}

fn alt_chunk_fixture_with_layout(
    signed: bool,
    shared_target: bool,
    nested_anchor: bool,
) -> Vec<u8> {
    let second = if shared_target {
        r#"<w:altChunk r:id="rIdAltChunk2"/>"#
    } else {
        ""
    };
    let anchors = if nested_anchor {
        format!(r#"<w:p><w:altChunk r:id="rIdAltChunk1"/>{second}</w:p>"#)
    } else {
        format!(r#"<w:altChunk r:id="rIdAltChunk1"/>{second}"#)
    };
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p>{anchors}<w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();

    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new(format!("/{MAIN_DOCUMENT}")).unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document,
    );
    main.rels_mut()
        .try_add_relationship(
            rt::ALTERNATIVE_FORMAT_IMPORT.to_owned(),
            "altChunk.html".to_owned(),
            "rIdAltChunk1".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    if shared_target {
        main.rels_mut()
            .try_add_relationship(
                rt::ALTERNATIVE_FORMAT_IMPORT.to_owned(),
                "altChunk.html".to_owned(),
                "rIdAltChunk2".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{ALT_TARGET}")).unwrap(),
            "text/html".to_owned(),
            b"<html><body>original</body></html>".to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{ALT_VENDOR_PART}")).unwrap(),
            "application/octet-stream".to_owned(),
            (0_u8..=u8::MAX).cycle().take(8192).collect(),
        )))
        .unwrap();
    package.relate_to(MAIN_DOCUMENT, rt::OFFICE_DOCUMENT);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

/// Return one raw local ZIP member span without using the package reader.
///
/// This intentionally compares local framing and compressed bytes, while the
/// central-directory relative offset (which may move after one replacement)
/// remains outside the member identity check.
fn raw_local_member(zip: &[u8], name: &str) -> Vec<u8> {
    let name = name.as_bytes();
    for (offset, signature) in zip.windows(4).enumerate() {
        if signature != b"PK\x01\x02" || offset + 46 > zip.len() {
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
        let end = local_offset + 30 + local_name_length + local_extra_length + compressed_size;
        return zip[local_offset..end].to_vec();
    }
    panic!("ZIP member {name:?} was not found")
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
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
fn opening_stays_cold_and_document_query_loads_only_main_document() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));

    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).expect("open source-backed DOCX");
    assert_eq!(source.main_reads.load(Ordering::SeqCst), 0);
    assert_eq!(source.unused_reads.load(Ordering::SeqCst), 0);

    let document = package.document().expect("load main document");
    assert!(source.main_reads.load(Ordering::SeqCst) > 0);
    assert_eq!(source.unused_reads.load(Ordering::SeqCst), 0);
    assert_eq!(document.extract_text().unwrap(), "alphabeta");
    assert_eq!(document.paragraph_count().unwrap(), 2);
    assert_eq!(
        document.paragraph(1).unwrap().unwrap().text().unwrap(),
        "beta"
    );
    assert!(document.paragraph(2).unwrap().is_none());
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(paragraphs[0].text().unwrap(), "alpha");
    assert_eq!(paragraphs[1].text().unwrap(), "beta");
}

#[test]
fn source_changes_remain_visible_before_the_first_payload_read() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).expect("open source-backed DOCX");

    source.changed();
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
}

#[test]
fn read_limits_are_returned_as_the_original_opc_error() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let read_at: Arc<dyn ReadAt> = source;

    assert!(matches!(
        source_backed::Package::from_read_at_with_limits(read_at, limits),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));
}

#[test]
fn document_commit_edits_only_the_main_part_and_reopens() {
    let (source_bytes, main, unused) = fixture();
    let trailing = payload_range(&source_bytes, TRAILING_PART);
    let source = Arc::new(ObservedSource::new(
        source_bytes.clone(),
        main,
        unused.clone(),
    ));
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(1), "edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    let published = package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();

    assert_eq!(published.xml_bytes(), commit.snapshot().xml_bytes());
    assert_eq!(
        &output[payload_range(&output, UNUSED_PART)],
        &source_bytes[unused]
    );
    assert_eq!(
        &output[payload_range(&output, TRAILING_PART)],
        &source_bytes[trailing]
    );
    let reopened = Package::from_reader(io::Cursor::new(output)).unwrap();
    let snapshot = reopened.document_snapshot().unwrap();
    assert_eq!(snapshot.paragraph_count(), 2);
    assert_eq!(
        snapshot
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "alpha"
    );
    assert_eq!(
        snapshot
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "edited"
    );
    assert!(
        std::str::from_utf8(snapshot.xml_bytes())
            .unwrap()
            .contains("<x:opaque value=\"preserve\"/>")
    );
}

#[test]
fn exact_noop_reproduces_every_source_byte() {
    let (source_bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(source_bytes.clone(), main, unused));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let commit = package.edit_document().unwrap().commit().unwrap();
    let mut output = Vec::new();

    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source_bytes);
}

#[test]
fn rolled_back_document_operations_use_the_exact_signed_source_path() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:x="urn:litchi:test"><w:body><w:p><w:r><w:t>alpha</w:t><x:opaque value="preserve"/></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let (source_bytes, main, unused) = fixture_for_document(document, rt::OFFICE_DOCUMENT, true);
    let source = Arc::new(ObservedSource::new(source_bytes.clone(), main, unused));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "temporary")
        .unwrap();
    edit.replace_paragraph_text(Position::new(0), "alpha")
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.patch().operations().len(), 2);
    assert!(!commit.patch().changed());
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source_bytes);
}

#[test]
fn stale_commit_and_changed_source_fail_before_output() {
    let (source_bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(source_bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(1), "edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_document_commit_to_stream(&mut output, &commit),
        Err(TransactionError::Document(Error::Opc(
            OpcError::SourceChanged { .. }
        )))
    ));
    assert!(output.is_empty());

    let stale_document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>different</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let (stale_bytes, stale_main, stale_unused) =
        fixture_for_document(stale_document, rt::OFFICE_DOCUMENT, false);
    let stale_source = Arc::new(ObservedSource::new(stale_bytes, stale_main, stale_unused));
    let stale_read_at: Arc<dyn ReadAt> = stale_source;
    let stale_package = source_backed::Package::from_read_at(stale_read_at).unwrap();
    assert!(matches!(
        stale_package.publish_document_commit_to_stream(&mut output, &commit),
        Err(TransactionError::StaleSource)
    ));
    assert!(output.is_empty());
}

#[test]
fn markup_compatibility_branch_selection_is_refused_before_output() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
    )
    .into_bytes();
    let commit = Snapshot::from_xml(document.clone())
        .unwrap()
        .edit()
        .commit()
        .unwrap();
    let (bytes, main, unused) = fixture_for_document(document, rt::OFFICE_DOCUMENT, false);
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut output = Vec::new();

    assert!(matches!(
        package.publish_document_commit_to_stream(&mut output, &commit),
        Err(TransactionError::Document(Error::UnsafeEdit { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn signed_changes_are_refused_but_signed_noops_are_exact() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>signed</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let (signed_bytes, main, unused) = fixture_for_document(document, rt::OFFICE_DOCUMENT, true);
    let source = Arc::new(ObservedSource::new(
        signed_bytes.clone(),
        main.clone(),
        unused.clone(),
    ));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let noop = package.edit_document().unwrap().commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, signed_bytes);

    let source = Arc::new(ObservedSource::new(signed_bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    output.clear();
    assert!(matches!(
        package.publish_document_commit_to_stream(&mut output, &commit),
        Err(TransactionError::Document(Error::Opc(
            OpcError::SignedSourceRequiresExplicitPolicy
        )))
    ));
    assert!(output.is_empty());
}

#[test]
fn sequential_sink_failure_is_typed_as_incomplete_output() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(1), "edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };

    assert!(matches!(
        package.publish_document_commit_to_stream(&mut sink, &commit),
        Err(TransactionError::Document(Error::Opc(
            OpcError::IncompleteOutput { .. }
        )))
    ));
    assert_eq!(sink.accepted, 128);
}

#[test]
fn strict_main_document_relationship_uses_the_same_overlay_contract() {
    const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    let document = format!(
        r#"<w:document xmlns:w="{STRICT_W}"><w:body><w:p><w:r><w:t>strict</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let (bytes, main, unused) = fixture_for_document(document, rt::STRICT_OFFICE_DOCUMENT, false);
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source;
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "strict edited")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();

    let reopened = Package::from_reader(io::Cursor::new(output)).unwrap();
    assert_eq!(
        reopened
            .document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "strict edited"
    );
}

#[test]
fn alt_chunk_payload_overlay_is_one_edit_and_preserves_raw_unselected_members() {
    let source_bytes = alt_chunk_fixture(false, false);
    let source = Arc::new(ObservedSource::new(source_bytes.clone(), 0..0, 0..0));
    let package = source_backed::Package::from_read_at(source).unwrap();
    let replacement = b"<html><body>replacement &amp; inert</body></html>".to_vec();
    let mut output = Vec::new();

    package
        .publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(replacement.clone()),
        )
        .unwrap();

    assert_ne!(output, source_bytes);
    assert_eq!(
        raw_local_member(&output, MAIN_DOCUMENT),
        raw_local_member(&source_bytes, MAIN_DOCUMENT)
    );
    assert_eq!(
        raw_local_member(&output, ALT_VENDOR_PART),
        raw_local_member(&source_bytes, ALT_VENDOR_PART)
    );
    let reopened = OpcPackage::from_reader(io::Cursor::new(output)).unwrap();
    let target = reopened
        .get_part(&PackURI::new(format!("/{ALT_TARGET}")).unwrap())
        .unwrap();
    assert_eq!(target.blob(), replacement.as_slice());
}

#[test]
fn alt_chunk_real_libreoffice_fixture_preserves_untouched_zip_members() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-html.docx"
    );
    let source_bytes = std::fs::read(path).unwrap();
    let source = Arc::new(ObservedSource::new(source_bytes.clone(), 0..0, 0..0));
    let package = source_backed::Package::from_read_at(source).unwrap();
    let replacement = b"<html><body>changed from a real producer</body></html>".to_vec();
    let mut output = Vec::new();

    package
        .publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(replacement.clone()),
        )
        .unwrap();

    for member in [
        "[Content_Types].xml",
        "word/document.xml",
        "word/_rels/document.xml.rels",
    ] {
        assert_eq!(
            raw_local_member(&output, member),
            raw_local_member(&source_bytes, member),
            "untouched member {member} changed"
        );
    }
    let reopened = OpcPackage::from_reader(io::Cursor::new(output)).unwrap();
    assert_eq!(
        reopened
            .get_part(&PackURI::new("/word/altChunk.html").unwrap())
            .unwrap()
            .blob(),
        replacement.as_slice()
    );
}

#[test]
fn alt_chunk_overlay_refuses_shared_targets_and_signed_changes_but_keeps_signed_noops_exact() {
    let shared_bytes = alt_chunk_fixture(false, true);
    let shared_source = Arc::new(ObservedSource::new(shared_bytes, 0..0, 0..0));
    let shared_package = source_backed::Package::from_read_at(shared_source).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        shared_package.publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(b"<html><body>ambiguous</body></html>".to_vec()),
        ),
        Err(Error::UnsafeEdit { .. })
    ));
    assert!(output.is_empty());

    let signed_bytes = alt_chunk_fixture(true, false);
    let signed_source = Arc::new(ObservedSource::new(signed_bytes.clone(), 0..0, 0..0));
    let signed_package = source_backed::Package::from_read_at(signed_source).unwrap();
    signed_package
        .publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(b"<html><body>original</body></html>".to_vec()),
        )
        .unwrap();
    assert_eq!(output, signed_bytes);

    let signed_source = Arc::new(ObservedSource::new(signed_bytes, 0..0, 0..0));
    let signed_package = source_backed::Package::from_read_at(signed_source).unwrap();
    output.clear();
    assert!(matches!(
        signed_package.publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(b"<html><body>signed change</body></html>".to_vec()),
        ),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn alt_chunk_overlay_refuses_non_body_anchor_layouts() {
    let source_bytes = alt_chunk_fixture_with_layout(false, false, true);
    let source = Arc::new(ObservedSource::new(source_bytes, 0..0, 0..0));
    let package = source_backed::Package::from_read_at(source).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(b"<html><body>nested</body></html>".to_vec()),
        ),
        Err(Error::OutOfBounds { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn alt_chunk_overlay_rejects_oversized_payload_before_reading_source() {
    let source_bytes = alt_chunk_fixture(false, false);
    let source = Arc::new(ObservedSource::new(source_bytes, 0..0, 0..0));
    let package = source_backed::Package::from_read_at(source).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_alt_chunk_to_stream(
            &mut output,
            source_backed::AltChunkSelector::index(0),
            Data::Html(vec![b'x'; MAX_DATA_BYTES + 1]),
        ),
        Err(Error::Invalid(_))
    ));
    assert!(output.is_empty());
}
