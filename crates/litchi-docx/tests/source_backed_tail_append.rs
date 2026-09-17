#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused source-backed append fixtures intentionally panic on invalid test data"
)]

//! Integration coverage for the bounded source-backed DOCX tail append.
//!
//! The fixture builders in this file deliberately stay on the public OPC and
//! DOCX boundaries.  In particular, the tests do not inspect the private
//! scanner or ZIP preservation index.  The main XML is supplied in both Store
//! and Deflate members, while an opaque member gives the publication oracle a
//! physical byte-preservation check.

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as CoreLimits, OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_docx::{Error as DocumentError, Package, ReadLimits, source_backed};
use litchi_ooxml_common::mce::Error as MceError;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcError, PackURI, SourceBackedPackage};
use soapberry_zip::ZipArchive;
use soapberry_zip::office::StreamingArchiveWriter;

use litchi_docx::source_backed::tail_append::{Error as TailAppendError, Refusal};
use litchi_docx::source_backed::{TailAppendLimits, TailAppendOptions, TailAppendPlan};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const OFFICE_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_OFFICE_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SETTINGS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const STRICT_SETTINGS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
const ATTACHED_TEMPLATE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";
const STRICT_ATTACHED_TEMPLATE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate";
const MAIL_MERGE_SOURCE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource";
const STRICT_MAIL_MERGE_SOURCE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeSource";
const RECIPIENT_DATA_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/recipientData";
const STRICT_RECIPIENT_DATA_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/recipientData";
const EXTERNAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const COMMENTS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const SIGNATURE_ORIGIN_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const MAIN: &str = "word/document.xml";
const MAIN_URI: &str = "/word/document.xml";
const MAIN_RELS: &str = "word/_rels/document.xml.rels";
const OPAQUE: &str = "word/opaque.bin";
const SETTINGS: &str = "word/settings.xml";
const SETTINGS_RELS: &str = "word/_rels/settings.xml.rels";
const COMMENTS: &str = "word/comments.xml";
const SIGNATURE: &str = "_xmlsignatures/origin.sigs";
const MAIL_MERGE_SOURCE: &str = "word/mailMergeSource.bin";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Compression {
    Store,
    Deflate,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Topology {
    Plain,
    Protected,
    External,
    UnsupportedDependency,
    Signed,
}

fn word_document(body: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{WORD}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn simple_document(strict: bool) -> Vec<u8> {
    let (prefix, namespace) = if strict {
        ("s", STRICT_WORD)
    } else {
        ("w", WORD)
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><{prefix}:document xmlns:{prefix}="{namespace}"><{prefix}:body><{prefix}:p><{prefix}:r><{prefix}:t>seed</{prefix}:t></{prefix}:r></{prefix}:p></{prefix}:body></{prefix}:document>"#
    )
    .into_bytes()
}

fn relationships(strict: bool, topology: Topology) -> Vec<u8> {
    let office = if strict {
        rt::STRICT_OFFICE_DOCUMENT
    } else {
        rt::OFFICE_DOCUMENT
    };
    let mut output = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rIdDocument" Type="{office}" Target="{MAIN}"/>"#
    );
    if topology == Topology::Signed {
        output.push_str(&format!(
            r#"<Relationship Id="rIdSignature" Type="{SIGNATURE_ORIGIN_REL}" Target="{SIGNATURE}"/>"#
        ));
    }
    output.push_str("</Relationships>");
    output.into_bytes()
}

fn main_relationships(strict: bool, topology: Topology) -> Option<Vec<u8>> {
    let settings_rel = if strict {
        STRICT_SETTINGS_REL
    } else {
        SETTINGS_REL
    };
    let value = match topology {
        Topology::Protected => format!(
            r#"<Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rIdSettings" Type="{settings_rel}" Target="settings.xml"/></Relationships>"#
        ),
        Topology::External => format!(
            r#"<Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rIdExternal" Type="{EXTERNAL_REL}" Target="https://example.invalid/linked" TargetMode="External"/></Relationships>"#
        ),
        Topology::UnsupportedDependency => format!(
            r#"<Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rIdComments" Type="{COMMENTS_REL}" Target="comments.xml"/></Relationships>"#
        ),
        Topology::Plain | Topology::Signed => return None,
    };
    Some(value.into_bytes())
}

fn content_types(topology: Topology) -> Vec<u8> {
    let mut output = format!(
        r#"<Types xmlns="{CONTENT_TYPES}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/{MAIN}" ContentType="{}"/>"#,
        ct::WML_DOCUMENT_MAIN,
    );
    if matches!(topology, Topology::Protected) {
        output.push_str(&format!(
            r#"<Override PartName="/{SETTINGS}" ContentType="{}"/>"#,
            ct::WML_SETTINGS,
        ));
    }
    if matches!(topology, Topology::UnsupportedDependency) {
        output.push_str(&format!(
            r#"<Override PartName="/{COMMENTS}" ContentType="{}"/>"#,
            ct::WML_COMMENTS,
        ));
    }
    if matches!(topology, Topology::Signed) {
        output.push_str(&format!(
            r#"<Override PartName="/{SIGNATURE}" ContentType="{}"/>"#,
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        ));
    }
    output.push_str("</Types>");
    output.into_bytes()
}

fn settings_xml(strict: bool) -> Vec<u8> {
    let namespace = if strict { STRICT_WORD } else { WORD };
    format!(
        r#"<w:settings xmlns:w="{namespace}"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>"#
    )
    .into_bytes()
}

fn settings_xml_with_body(strict: bool, body: &str) -> Vec<u8> {
    let namespace = if strict { STRICT_WORD } else { WORD };
    format!(r#"<w:settings xmlns:w="{namespace}">{body}</w:settings>"#).into_bytes()
}

fn settings_xml_with_relationships(strict: bool, body: &str) -> Vec<u8> {
    let namespace = if strict { STRICT_WORD } else { WORD };
    let relationships = if strict {
        STRICT_OFFICE_RELATIONSHIPS
    } else {
        OFFICE_RELATIONSHIPS
    };
    format!(r#"<w:settings xmlns:w="{namespace}" xmlns:r="{relationships}">{body}</w:settings>"#)
        .into_bytes()
}

fn adversarial_settings_limits(source_xml_bytes: usize, max_token_bytes: u64) -> TailAppendLimits {
    let mut limits = limits(source_xml_bytes);
    // Keep these fixtures below the source and token ceilings so the asserted
    // failure comes from the settings owner envelope itself.
    limits.max_settings_xml_bytes = 64 * 1024;
    limits.max_token_bytes = max_token_bytes;
    limits.max_workspace_bytes = 2 * 1024 * 1024;
    limits
}

fn settings_with_many_unused_namespaces(strict: bool, extension_count: usize) -> Vec<u8> {
    let namespace = if strict { STRICT_WORD } else { WORD };
    let mut xml = format!(r#"<w:settings xmlns:w="{namespace}" xmlns:x="urn:settings-extension""#);
    for index in 0..20 {
        xml.push_str(&format!(
            r#" xmlns:u{index}="urn:unused-{index}-{}""#,
            "n".repeat(120)
        ));
    }
    xml.push('>');
    for _ in 0..extension_count {
        xml.push_str("<x:opaque/>");
    }
    xml.push_str("</w:settings>");
    xml.into_bytes()
}

fn settings_with_long_mce_directives(strict: bool, target_count: usize) -> Vec<u8> {
    let namespace = if strict { STRICT_WORD } else { WORD };
    let long_uri = format!("urn:settings-mce-{}", "m".repeat(3000));
    let mut targets = String::from("u&#x3a;tag u:a&#32;u:b u&#x3a;inner");
    for index in 0..target_count {
        targets.push_str(&format!(" u:t{index}"));
    }
    format!(
        r#"<w:settings xmlns:w="{namespace}" xmlns:mc="{MCE}" xmlns:u="{long_uri}" mc:Ignorable="u"><u:outer mc:Ignorable="u" mc:PreserveElements="u&#x3a;outer" mc:ProcessContent="{targets}"><u:inner/></u:outer></w:settings>"#
    )
    .into_bytes()
}

fn settings_relationships_xml(entries: &str) -> Vec<u8> {
    format!(r#"<Relationships xmlns="{RELATIONSHIPS}">{entries}</Relationships>"#).into_bytes()
}

fn archive(xml: &[u8], compression: Compression, strict: bool, topology: Topology) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", &content_types(topology))
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", &relationships(strict, topology))
        .expect("package relationships fixture must be writable");
    if let Some(rels) = main_relationships(strict, topology) {
        writer
            .write_stored(MAIN_RELS, &rels)
            .expect("main relationships fixture must be writable");
    }
    writer
        .write_stored(OPAQUE, b"opaque physical bytes\0\xff\x01")
        .expect("opaque member fixture must be writable");
    match compression {
        Compression::Store => writer
            .write_stored(MAIN, xml)
            .expect("stored main fixture must be writable"),
        Compression::Deflate => writer
            .write_deflated_sized(MAIN, xml)
            .expect("Deflate main fixture must be writable"),
    }
    match topology {
        Topology::Protected => writer
            .write_stored(SETTINGS, &settings_xml(strict))
            .expect("settings fixture must be writable"),
        Topology::UnsupportedDependency => writer
            .write_stored(
                COMMENTS,
                format!(r#"<w:comments xmlns:w="{WORD}"/>"#).as_bytes(),
            )
            .expect("comments fixture must be writable"),
        Topology::Signed => writer
            .write_stored(SIGNATURE, b"signature-origin")
            .expect("signature fixture must be writable"),
        Topology::Plain | Topology::External => {},
    }
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn settings_archive(
    xml: &[u8],
    compression: Compression,
    strict: bool,
    settings: &[u8],
    settings_relationships: Option<&[u8]>,
    extra_member: Option<(&str, &[u8])>,
) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", &content_types(Topology::Protected))
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", &relationships(strict, Topology::Plain))
        .expect("package relationships fixture must be writable");
    writer
        .write_stored(
            MAIN_RELS,
            &main_relationships(strict, Topology::Protected)
                .expect("settings relationship fixture must exist"),
        )
        .expect("main relationships fixture must be writable");
    writer
        .write_stored(OPAQUE, b"opaque physical bytes\0\xff\x01")
        .expect("opaque member fixture must be writable");
    match compression {
        Compression::Store => writer
            .write_stored(MAIN, xml)
            .expect("stored main fixture must be writable"),
        Compression::Deflate => writer
            .write_deflated_sized(MAIN, xml)
            .expect("Deflate main fixture must be writable"),
    }
    writer
        .write_stored(SETTINGS, settings)
        .expect("settings fixture must be writable");
    if let Some(relationships) = settings_relationships {
        writer
            .write_stored(SETTINGS_RELS, relationships)
            .expect("settings relationships fixture must be writable");
    }
    if let Some((path, bytes)) = extra_member {
        writer
            .write_stored(path, bytes)
            .expect("settings target fixture must be writable");
    }
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn part_bytes(archive: &[u8], uri: &str) -> Vec<u8> {
    SourceBackedPackage::from_vec(archive.to_vec())
        .expect("published archive must open")
        .part(&PackURI::new(uri).expect("part URI must be valid"))
        .expect("requested part must exist")
        .data()
        .expect("part must decode")
        .as_bytes()
        .to_vec()
}

fn main_bytes(archive: &[u8]) -> Vec<u8> {
    part_bytes(archive, MAIN_URI)
}

fn settings_bytes(archive: &[u8]) -> Vec<u8> {
    part_bytes(archive, "/word/settings.xml")
}

#[derive(Debug, PartialEq, Eq)]
struct PhysicalMemberRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn physical_member_record(archive_bytes: &[u8], member: &str) -> PhysicalMemberRecord {
    let archive = ZipArchive::from_slice(archive_bytes).expect("archive index must parse");
    for result in archive.entries() {
        let entry = result.expect("archive entry must parse");
        if entry.is_dir() {
            continue;
        }
        let path = entry
            .file_path()
            .try_normalize()
            .expect("member path must normalize");
        if path.as_str() != member {
            continue;
        }
        let zip_entry = archive
            .get_entry(entry.wayfinder())
            .expect("indexed entry must resolve");
        let (_, compressed_end) = zip_entry.compressed_data_range();
        let local_payload_end =
            usize::try_from(compressed_end).expect("compressed payload end must fit usize");
        let descriptor_bytes = if entry.has_data_descriptor() {
            let descriptor = archive_bytes
                .get(local_payload_end..)
                .expect("data descriptor must start inside archive");
            let with_signature = descriptor.starts_with(b"PK\x07\x08");
            if entry.is_zip64() {
                if with_signature { 24 } else { 20 }
            } else if with_signature {
                16
            } else {
                12
            }
        } else {
            0
        };
        let local_start = usize::try_from(entry.local_header_offset())
            .expect("local header offset must fit usize");
        let local_end = local_payload_end
            .checked_add(descriptor_bytes)
            .expect("local record end must fit usize");
        let central_start = usize::try_from(entry.central_directory_offset())
            .expect("central record offset must fit usize");
        let central_size = usize::try_from(entry.metadata_size_hint())
            .expect("central metadata length must fit usize")
            .checked_add(46)
            .expect("central record size must fit usize");
        let central_end = central_start
            .checked_add(central_size)
            .expect("central record end must fit usize");
        return PhysicalMemberRecord {
            local: archive_bytes
                .get(local_start..local_end)
                .expect("local record must fit archive")
                .to_vec(),
            central: archive_bytes
                .get(central_start..central_end)
                .expect("central record must fit archive")
                .to_vec(),
        };
    }
    panic!("ZIP member {member:?} is missing");
}

fn compression_method(archive: &[u8], member: &str) -> u16 {
    let record = physical_member_record(archive, member);
    u16::from_le_bytes(record.local[8..10].try_into().expect("compression method"))
}

fn semantic_texts(archive: &[u8]) -> Vec<String> {
    let package = Package::from_reader(io::Cursor::new(archive.to_vec())).expect("DOCX reopens");
    package
        .document_snapshot()
        .expect("document reopens")
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text().expect("paragraph text decodes"))
        .collect()
}

#[derive(Debug)]
struct PrefixFailSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl PrefixFailSink {
    fn after(bytes: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining: bytes,
        }
    }
}

impl Write for PrefixFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test sink stopped",
            ));
        }
        let accepted = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            48_303,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn managed_context(memory: u64, work: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-tail-append-managed-test",
        CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory.max(1)).expect("managed memory limit must be nonzero"),
        0,
    )
    .expect("managed execution limits must be valid");
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn managed_package(
    archive: Vec<u8>,
    memory: u64,
    work: u64,
) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) = managed_context(memory, work);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        ReadLimits::default(),
        context,
    )
    .expect("managed source-backed DOCX fixture must open");
    (budget, cancellation_source, package)
}

#[derive(Debug)]
struct MemoryObservingSource {
    inner: OwnedSource,
    budget: Budget,
    armed: AtomicBool,
    peak: AtomicU64,
}

impl MemoryObservingSource {
    fn new(bytes: Vec<u8>, budget: Budget) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            budget,
            armed: AtomicBool::new(false),
            peak: AtomicU64::new(0),
        }
    }

    fn arm(&self) {
        self.armed.store(true, Ordering::Release);
    }

    fn peak(&self) -> u64 {
        self.peak.load(Ordering::Acquire)
    }

    fn observe(&self) {
        if !self.armed.load(Ordering::Acquire) {
            return;
        }
        let observed = self.budget.used(Resource::Memory);
        let _ = self
            .peak
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |peak| {
                Some(peak.max(observed))
            });
    }
}

impl ReadAt for MemoryObservingSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.observe();
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug)]
struct OneByteSource {
    inner: OwnedSource,
}

impl OneByteSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
        }
    }
}

impl ReadAt for OneByteSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let count = output.len().min(1);
        self.inner.read_at(offset, &mut output[..count])
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug)]
struct InterruptedOnceSource {
    inner: OwnedSource,
    armed: AtomicBool,
    interrupted: AtomicBool,
}

impl InterruptedOnceSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            armed: AtomicBool::new(false),
            interrupted: AtomicBool::new(true),
        }
    }

    fn arm(&self) {
        self.interrupted.store(false, Ordering::Release);
        self.armed.store(true, Ordering::Release);
    }
}

impl ReadAt for InterruptedOnceSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.armed.load(Ordering::Acquire) && !self.interrupted.swap(true, Ordering::AcqRel) {
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug)]
struct ReadCountingSource {
    inner: OwnedSource,
    armed: AtomicBool,
    reads: AtomicU64,
}

impl ReadCountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            armed: AtomicBool::new(false),
            reads: AtomicU64::new(0),
        }
    }

    fn arm(&self) {
        self.reads.store(0, Ordering::Release);
        self.armed.store(true, Ordering::Release);
    }

    fn reads(&self) -> u64 {
        self.reads.load(Ordering::Acquire)
    }
}

impl ReadAt for ReadCountingSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.armed.load(Ordering::Acquire) {
            self.reads.fetch_add(1, Ordering::AcqRel);
        }
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn little_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(
        bytes[offset..offset + 2]
            .try_into()
            .expect("fixture u16 must fit archive"),
    )
}

fn little_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(
        bytes[offset..offset + 4]
            .try_into()
            .expect("fixture u32 must fit archive"),
    )
}

fn mark_member_encrypted(mut archive: Vec<u8>, member: &str) -> Vec<u8> {
    let eocd = archive
        .windows(4)
        .rposition(|window| window == b"PK\x05\x06")
        .expect("fixture EOCD must be present");
    let count = usize::from(little_u16(&archive, eocd + 10));
    let mut central = usize::try_from(little_u32(&archive, eocd + 16))
        .expect("fixture central offset must fit usize");
    for _ in 0..count {
        assert_eq!(&archive[central..central + 4], b"PK\x01\x02");
        let name_len = usize::from(little_u16(&archive, central + 28));
        let extra_len = usize::from(little_u16(&archive, central + 30));
        let comment_len = usize::from(little_u16(&archive, central + 32));
        let name_start = central + 46;
        let name_end = name_start
            .checked_add(name_len)
            .expect("fixture member name end must fit usize");
        if archive
            .get(name_start..name_end)
            .expect("fixture member name must fit archive")
            == member.as_bytes()
        {
            let local = usize::try_from(little_u32(&archive, central + 42))
                .expect("fixture local offset must fit usize");
            let local_flags = little_u16(&archive, local + 6) | 1;
            let central_flags = little_u16(&archive, central + 8) | 1;
            archive[local + 6..local + 8].copy_from_slice(&local_flags.to_le_bytes());
            archive[central + 8..central + 10].copy_from_slice(&central_flags.to_le_bytes());
            return archive;
        }
        central = central
            .checked_add(46 + name_len + extra_len + comment_len)
            .expect("fixture central record end must fit usize");
    }
    panic!("fixture ZIP member {member:?} is missing");
}

#[derive(Debug)]
struct CancelOnShortWriteSink {
    cancellation: CancellationSource,
    bytes: Vec<u8>,
    chunk: usize,
    cancelled: bool,
}

impl Write for CancelOnShortWriteSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.cancelled {
            self.cancelled = true;
            self.cancellation.cancel();
        }
        let accepted = bytes.len().min(self.chunk.max(1));
        self.bytes.extend_from_slice(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn assert_cancelled_publication_source(error: &OpcError) {
    match error {
        OpcError::Cancelled | OpcError::Execution(ExecutionError::Cancelled) => {},
        OpcError::IncompleteOutput { source, .. } => {
            assert_cancelled_publication_source(source);
        },
        _ => panic!("publication failure did not retain typed cancellation"),
    }
}

fn section_xml() -> &'static str {
    r#"<q:sectPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:vendor-word-like" v:flag="keep" q:rsidR="007"><v:opaque a="&quot;"/><q:pgSz q:w="11906" q:h="16838"/></q:sectPr>"#
}

fn limits(source_xml_bytes: usize) -> TailAppendLimits {
    let source_xml_bytes = u64::try_from(source_xml_bytes).expect("fixture XML length fits u64");
    TailAppendLimits {
        max_source_xml_bytes: source_xml_bytes.saturating_add(4096),
        max_text_bytes: 64 * 1024,
        max_fragment_bytes: 65_536,
        max_candidate_xml_bytes: source_xml_bytes.saturating_add(64 * 1024),
        // The OPC XML audit derives a checked workspace envelope from these
        // values, including namespace and expanded-attribute scratch.  The
        // finite 4 KiB/depth-32 profile needs about 1.6 MiB.
        max_events: 4096,
        max_depth: 32,
        max_paragraphs: 65_536,
        max_settings_xml_bytes: 4096,
        max_workspace_bytes: 2 * 1024 * 1024,
        max_output_bytes: 2 * 1024 * 1024,
        max_token_bytes: 4096,
    }
}

fn bounded_package(bytes: Vec<u8>) -> source_backed::Package {
    source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source-backed DOCX fixture must open")
}

fn append_plan<'a>(
    package: &'a source_backed::Package,
    text: &str,
    source_xml_bytes: usize,
) -> TailAppendPlan<'a> {
    package
        .tail_append_plain_paragraph(text)
        .with_limits(limits(source_xml_bytes))
        .prepare()
        .expect("tail append plan must prepare")
}

#[test]
fn appends_before_exact_opaque_section_and_preserves_untouched_store_and_deflate_members() {
    let section = section_xml();
    let source_xml = word_document(&format!(
        r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>{section}"#
    ));
    for (compression, strict) in [
        (Compression::Store, false),
        (Compression::Deflate, false),
        (Compression::Store, true),
        (Compression::Deflate, true),
    ] {
        let (source_xml, expected_section) = if strict {
            (
                format!(
                    r#"<?xml version="1.0"?><s:document xmlns:s="{STRICT_WORD}"><s:body><s:p><s:r><s:t>seed</s:t></s:r></s:p><q:sectPr xmlns:q="{STRICT_WORD}" xmlns:v="urn:vendor-word-like" v:flag="keep" q:rsidR="007"><v:opaque a="&quot;"/><q:pgSz q:w="11906" q:h="16838"/></q:sectPr></s:body></s:document>"#
                )
                .into_bytes(),
                r#"<q:sectPr xmlns:q="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:v="urn:vendor-word-like" v:flag="keep" q:rsidR="007"><v:opaque a="&quot;"/><q:pgSz q:w="11906" q:h="16838"/></q:sectPr>"#,
            )
        } else {
            (source_xml.clone(), section)
        };
        let source_archive = archive(&source_xml, compression, strict, Topology::Plain);
        let source_main = main_bytes(&source_archive);
        let package = bounded_package(source_archive.clone());
        let plan = append_plan(&package, "tail & < > \" '", source_xml.len());
        let source_proof = plan.source_proof();
        let candidate_proof = plan.candidate_proof();
        let splice_proof = plan.splice_proof();
        assert_eq!(
            splice_proof.insertion_offset, source_proof.insertion_offset,
            "the splice anchor must be the source semantic anchor"
        );
        assert_eq!(
            candidate_proof.generated_offset, source_proof.insertion_offset,
            "the generated paragraph must start at the source insertion anchor"
        );
        assert_eq!(
            candidate_proof.candidate_len,
            source_proof
                .source_len
                .checked_add(splice_proof.fragment_len)
                .expect("candidate length must fit u64")
        );
        let mut output = Vec::new();
        let publication = plan
            .write_to_stream(&mut output)
            .expect("bounded tail append must publish");
        assert_ne!(output, source_archive);
        assert_eq!(semantic_texts(&output), ["seed", "tail & < > \" '"]);
        assert_eq!(
            physical_member_record(&source_archive, OPAQUE),
            physical_member_record(&output, OPAQUE)
        );
        assert_eq!(
            compression_method(&source_archive, MAIN),
            compression_method(&output, MAIN)
        );

        let output_main = main_bytes(&output);
        let source_section_at = source_main
            .windows(expected_section.len())
            .position(|window| window == expected_section.as_bytes());
        let output_section_at = output_main
            .windows(expected_section.len())
            .position(|window| window == expected_section.as_bytes());
        if let (Some(source_section_at), Some(output_section_at)) =
            (source_section_at, output_section_at)
        {
            assert_eq!(
                u64::try_from(source_section_at).expect("source anchor fits u64"),
                source_proof.insertion_offset,
                "source proof must identify the final section opening"
            );
            assert_eq!(
                u64::try_from(output_section_at).expect("candidate anchor fits u64"),
                source_proof
                    .insertion_offset
                    .checked_add(splice_proof.fragment_len)
                    .expect("candidate anchor must fit u64"),
                "candidate section anchor must shift by the fragment"
            );
            assert_eq!(
                &source_main[..source_section_at],
                &output_main[..source_section_at],
                "the source prefix through the insertion anchor must remain lexical"
            );
            assert_eq!(
                &output_main[output_section_at..output_section_at + expected_section.len()],
                expected_section.as_bytes(),
                "the final sectPr span must be copied byte-for-byte"
            );
            assert!(output_section_at > source_section_at);
        } else {
            panic!("fixture section-properties span must survive publication");
        }

        let mut inverse = Vec::new();
        let candidate = bounded_package(output.clone());
        publication
            .write_inverse_to_stream(&candidate, &mut inverse)
            .expect("immediate inverse must restore exact source");
        assert_eq!(inverse, source_archive);
        drop(package);
    }
}

#[test]
fn default_word_namespace_and_shadowed_local_prefix_are_bound_at_insertion() {
    let default_xml =
        format!(r#"<document xmlns="{WORD}"><body><p><r><t>seed</t></r></p></body></document>"#)
            .into_bytes();
    let shadowed_xml = format!(
        r#"<a:document xmlns:a="{WORD}"><b:body xmlns:b="{WORD}"><b:p><b:r><b:t>seed</b:t></b:r></b:p></b:body></a:document>"#
    )
    .into_bytes();
    for xml in [default_xml, shadowed_xml] {
        let source_archive = archive(&xml, Compression::Store, false, Topology::Plain);
        let package = bounded_package(source_archive);
        let plan = append_plan(&package, "tail", xml.len());
        let mut output = Vec::new();
        plan.write_to_stream(&mut output)
            .expect("namespace-bound tail append must publish");
        assert_eq!(semantic_texts(&output), ["seed", "tail"]);
    }
}

#[test]
fn append_without_a_section_properties_element_targets_the_body_close() {
    let source_xml = word_document("");
    let source_archive = archive(&source_xml, Compression::Deflate, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let plan = append_plan(&package, "tail", source_xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("tail append without sectPr must publish");
    assert_eq!(semantic_texts(&output), ["tail"]);
}

#[test]
fn insertion_survives_distinct_body_prefix_and_word_prefix_shadowing_in_section() {
    let xml = format!(
        r#"<a:document xmlns:a="{WORD}" xmlns:w="urn:root-word-shadow"><b:body xmlns:b="{WORD}"><a:p><a:r><a:t>seed</a:t></a:r></a:p><s:sectPr xmlns:s="{WORD}" xmlns:w="urn:vendor-word-like"><w:opaque/></s:sectPr></b:body></a:document>"#
    )
    .into_bytes();
    let source_archive = archive(&xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let plan = append_plan(&package, "tail", xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("shadowed namespace append must publish");
    assert_eq!(semantic_texts(&output), ["seed", "tail"]);

    let output_main = main_bytes(&output);
    let section = b"<s:sectPr xmlns:s=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w=\"urn:vendor-word-like\"><w:opaque/></s:sectPr>";
    let section_at = output_main
        .windows(section.len())
        .position(|window| window == section)
        .expect("shadowed section must remain byte-identical");
    let source_section_at = xml
        .windows(section.len())
        .position(|window| window == section)
        .expect("source shadowed section must be present");
    assert_eq!(
        &output_main[section_at..section_at + section.len()],
        &xml[source_section_at..source_section_at + section.len()]
    );
}

#[test]
fn empty_text_is_a_real_empty_paragraph_and_is_not_the_exact_noop_branch() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Deflate, false, Topology::Plain);
    let package = bounded_package(source_archive.clone());
    let plan = append_plan(&package, "", source_xml.len());
    let mut output = Vec::new();
    let publication = plan
        .write_to_stream(&mut output)
        .expect("empty authored text still publishes one paragraph");
    assert_ne!(output, source_archive);
    assert_eq!(semantic_texts(&output), ["seed", ""]);
    let _ = publication;
}

#[test]
fn explicit_noop_is_exact_and_reversible_without_a_generated_paragraph() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(source_archive.clone());
    let plan = package
        .tail_append_noop()
        .with_limits(limits(source_xml.len()))
        .prepare()
        .expect("explicit no-op must prepare");
    assert!(plan.is_noop());
    let mut output = Vec::new();
    let publication = plan
        .write_to_stream(&mut output)
        .expect("explicit no-op must publish");
    assert!(publication.is_noop());
    assert_eq!(output, source_archive);

    let mut inverse = Vec::new();
    publication
        .write_inverse_to_stream(&package, &mut inverse)
        .expect("no-op inverse must publish");
    assert_eq!(inverse, source_archive);
}

#[test]
fn cancellation_is_checked_before_preparation_and_before_publication_output() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);

    let package = bounded_package(source_archive.clone());
    let (cancel_source, cancellation) = CancellationSource::pair();
    cancel_source.cancel();
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .with_options(TailAppendOptions::default().with_cancellation_token(&cancellation))
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    assert!(result.is_err());
    assert!(
        output.is_empty(),
        "cancelled preparation must emit no archive"
    );

    let package = bounded_package(source_archive);
    let (cancel_source, cancellation) = CancellationSource::pair();
    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .with_options(TailAppendOptions::default().with_cancellation_token(&cancellation))
        .prepare()
        .expect("uncancelled plan must prepare");
    cancel_source.cancel();
    let mut output = Vec::new();
    assert!(plan.write_to_stream(&mut output).is_err());
    assert!(
        output.is_empty(),
        "cancelled publication must emit no archive"
    );
}

#[test]
fn authored_text_uses_xml_space_and_rejects_streaming_writer_control_rules() {
    let source_xml =
        word_document(r#"<w:p><w:r><w:t xml:space="preserve"> seed </w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let plan = append_plan(&package, " leading & <tail> ", source_xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("escaped text publication must succeed");
    assert_eq!(semantic_texts(&output), [" seed ", " leading & <tail> "]);
    let output_main = main_bytes(&output);
    let xml_space_count = output_main
        .windows(b"xml:space=\"preserve\"".len())
        .filter(|&window| window == b"xml:space=\"preserve\"")
        .count();
    assert!(
        xml_space_count >= 2,
        "leading/trailing authored spaces require xml:space preservation"
    );

    for invalid in [
        "tab\tvalue",
        "line\nvalue",
        "line\rvalue",
        "control\u{0001}value",
    ] {
        let package = bounded_package(archive(
            &source_xml,
            Compression::Store,
            false,
            Topology::Plain,
        ));
        let mut output = Vec::new();
        let result = package
            .tail_append_plain_paragraph(invalid)
            .with_limits(limits(source_xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output));
        assert!(
            result.is_err(),
            "invalid authored text must refuse: {invalid:?}"
        );
        assert!(output.is_empty());
    }
}

#[test]
fn many_xml_space_paragraphs_use_the_aggregate_attribute_budget() {
    const PARAGRAPHS: usize = 4_200;
    let mut body = String::new();
    for index in 0..PARAGRAPHS {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t xml:space="preserve">p{index}</w:t></w:r></w:p>"#
        ));
    }
    let source_xml = word_document(&body);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let mut policy = limits(source_xml.len());
    policy.max_events = u64::try_from(PARAGRAPHS)
        .expect("paragraph count fits u64")
        .saturating_mul(8)
        .saturating_add(64);
    policy.max_depth = 8;
    policy.max_paragraphs = 65_536;
    policy.max_token_bytes = 4 * 1024;
    policy.max_candidate_xml_bytes = u64::try_from(source_xml.len())
        .expect("source XML length fits u64")
        .saturating_add(64 * 1024);

    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(policy)
        .prepare()
        .expect("many small attributed paragraphs must fit aggregate audit limits");
    let source_proof = plan.source_proof();
    let candidate_proof = plan.candidate_proof();
    let splice_proof = plan.splice_proof();
    assert_eq!(source_proof.paragraph_count, PARAGRAPHS as u64);
    assert_eq!(candidate_proof.paragraph_count, (PARAGRAPHS + 1) as u64);
    assert_eq!(
        candidate_proof.generated_offset,
        source_proof.insertion_offset
    );
    assert!(candidate_proof.generated_once);
    assert_eq!(
        candidate_proof.candidate_len,
        source_proof
            .source_len
            .checked_add(splice_proof.fragment_len)
            .expect("candidate length fits u64")
    );

    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("aggregate XML attributes must not be confused with token scratch");
    let texts = semantic_texts(&output);
    assert_eq!(texts.len(), PARAGRAPHS + 1);
    assert_eq!(texts[0], "p0");
    assert_eq!(texts[PARAGRAPHS - 1], "p4199");
    assert_eq!(texts[PARAGRAPHS], "tail");
}

#[test]
fn strict_tail_append_preserves_opaque_section_with_vendor_word_like_descendants() {
    let xml = format!(
        r#"<s:document xmlns:s="{STRICT_WORD}"><s:body><s:p><s:r><s:t>seed</s:t></s:r></s:p><q:sectPr xmlns:q="{STRICT_WORD}" xmlns:v="urn:vendor-word-like" v:flag="keep"><v:sectPr><v:opaque/></v:sectPr><q:pgMar q:top="720"/></q:sectPr></s:body></s:document>"#
    )
    .into_bytes();
    let source_archive = archive(&xml, Compression::Store, true, Topology::Plain);
    let package = bounded_package(source_archive.clone());
    let plan = append_plan(&package, "strict-tail", xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("strict source with opaque section must publish");
    assert_eq!(semantic_texts(&output), ["seed", "strict-tail"]);
    let output_main = main_bytes(&output);
    let section_start = xml
        .windows(b"<q:sectPr".len())
        .position(|window| window == b"<q:sectPr")
        .expect("section start");
    let section_end = xml
        .windows(b"</q:sectPr>".len())
        .position(|window| window == b"</q:sectPr>")
        .expect("section end")
        + b"</q:sectPr>".len();
    let output_start = output_main
        .windows(b"<q:sectPr".len())
        .position(|window| window == b"<q:sectPr")
        .expect("published section start");
    assert_eq!(
        &output_main[output_start..output_start + section_end - section_start],
        &xml[section_start..section_end]
    );
}

#[test]
fn opaque_section_allows_an_explicit_empty_default_namespace_and_preserves_it() {
    let section = format!(r#"<q:sectPr xmlns:q="{WORD}" xmlns=""><opaque/></q:sectPr>"#);
    let xml = word_document(&format!(
        r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>{section}"#
    ));
    let source_archive = archive(&xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let plan = append_plan(&package, "tail", xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("an explicitly empty default namespace is valid opaque XML");
    assert_eq!(semantic_texts(&output), ["seed", "tail"]);
    let output_main = main_bytes(&output);
    assert!(
        output_main
            .windows(section.len())
            .any(|window| window == section.as_bytes()),
        "the opaque section with xmlns=\"\" must remain byte-for-byte intact"
    );
}

#[test]
fn opaque_section_namespace_attribute_and_event_policy_refuses_before_output() {
    let cases = [
        (
            "unbound prefixed element",
            format!(r#"<q:sectPr xmlns:q="{WORD}"><u:opaque/></q:sectPr>"#),
        ),
        (
            "reserved xml namespace rebinding",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns:xml="urn:invalid-xml"><q:pgSz/></q:sectPr>"#
            ),
        ),
        (
            "reserved xmlns namespace rebinding",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns:xmlns="urn:invalid-xmlns"><q:pgSz/></q:sectPr>"#
            ),
        ),
        (
            "duplicate expanded attributes",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns:x="urn:same" xmlns:y="urn:same" x:flag="one" y:flag="two"/>"#
            ),
        ),
        (
            "duplicate expanded attributes after namespace-value unescape",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns:a="urn:x" xmlns:b="urn:&#120;" a:n="one" b:n="two"/>"#
            ),
        ),
        (
            "default namespace bound to reserved XML URI",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns="http://www.w3.org/XML/1998/namespace"><opaque/></q:sectPr>"#
            ),
        ),
        (
            "default namespace bound to reserved XMLNS URI",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns="http://www.w3.org/2000/xmlns/"><opaque/></q:sectPr>"#
            ),
        ),
        (
            "invalid opaque element QName",
            format!(r#"<q:sectPr xmlns:q="{WORD}" xmlns:v="urn:vendor"><v:1opaque/></q:sectPr>"#),
        ),
        (
            "invalid opaque attribute QName",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns:v="urn:vendor"><v:opaque v:1flag="keep"/></q:sectPr>"#
            ),
        ),
        (
            "invalid namespace prefix QName",
            format!(
                r#"<q:sectPr xmlns:q="{WORD}" xmlns:1vendor="urn:vendor"><q:pgSz/></q:sectPr>"#
            ),
        ),
        (
            "comment event",
            format!(r#"<q:sectPr xmlns:q="{WORD}"><!--unsupported--></q:sectPr>"#),
        ),
        (
            "CDATA event",
            format!(r#"<q:sectPr xmlns:q="{WORD}"><![CDATA[unsupported]]></q:sectPr>"#),
        ),
        (
            "processing instruction event",
            format!(r#"<q:sectPr xmlns:q="{WORD}"><?vendor unsupported?></q:sectPr>"#),
        ),
    ];
    for (label, section) in cases {
        let xml = word_document(&format!(
            r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>{section}"#
        ));
        let package = bounded_package(archive(&xml, Compression::Store, false, Topology::Plain));
        let mut output = Vec::new();
        let result = package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output));
        assert!(result.is_err(), "opaque {label} must be refused");
        assert!(
            output.is_empty(),
            "opaque {label} must emit no archive prefix"
        );
    }
}

#[test]
fn opaque_section_preserves_unicode_qnames() {
    let section = format!(
        r#"<q:sectPr xmlns:q="{WORD}" xmlns:é="urn:vendor"><é:opaqueé é:flag="keep"/></q:sectPr>"#
    );
    let xml = word_document(&format!(
        r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>{section}"#
    ));
    let package = bounded_package(archive(&xml, Compression::Store, false, Topology::Plain));
    let plan = append_plan(&package, "tail", xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("Unicode XML NCNames are valid opaque section names");
    assert_eq!(semantic_texts(&output), ["seed", "tail"]);
    let output_main = main_bytes(&output);
    assert!(
        output_main
            .windows(section.len())
            .any(|window| window == section.as_bytes()),
        "opaque Unicode QNames must remain byte-for-byte intact"
    );
}

fn assert_refused(xml: Vec<u8>, topology: Topology) {
    assert_refused_with_text(xml, topology, "tail");
}

fn assert_refused_with_text(xml: Vec<u8>, topology: Topology, text: &str) {
    let archive = archive(&xml, Compression::Store, false, topology);
    let package = bounded_package(archive);
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph(text)
        .with_limits(limits(xml.len()))
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    assert!(result.is_err(), "fixture must be refused: {topology:?}");
    assert!(output.is_empty(), "refusal must happen before ZIP output");
}

#[test]
fn grammar_refusals_happen_before_publication() {
    let section = section_xml();
    let cases = [
        word_document(&format!(r#"{section}<w:p/>"#)),
        word_document(&format!(r#"{section}{section}"#)),
        word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p><w:sectPr><w:sectPr/></w:sectPr>"#),
        word_document(&format!(r#"<w:p>{section}</w:p>"#)),
        word_document(r#"<w:tbl/>"#),
        word_document(r#"<w:unknown/>"#),
        word_document(r#"<w:p><w:r><w:tab/></w:r></w:p>"#),
        word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p><!-- comment -->"#),
        word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p><?pi value?>"#),
        b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>".to_vec(),
    ];
    for xml in cases {
        assert_refused(xml, Topology::Plain);
    }
}

#[test]
fn xml_declaration_and_encoding_grammar_refuse_before_publication() {
    let duplicate_declaration = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert_refused(duplicate_declaration, Topology::Plain);

    let xml_11 = format!(
        r#"<?xml version="1.1" encoding="UTF-8"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert_refused(xml_11, Topology::Plain);

    let non_utf8_declaration = format!(
        r#"<?xml version="1.0" encoding="ISO-8859-1"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>café</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    assert_refused_with_text(non_utf8_declaration, Topology::Plain, "café");

    for declaration in [
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" version="1.0"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<?xml encoding="UTF-8" version="1.0"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<?xml version="1.0" standalone="maybe"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes" extra="no"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
        ),
    ] {
        assert_refused(declaration.into_bytes(), Topology::Plain);
    }

    for standalone in ["yes", "no"] {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="{standalone}"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
        )
        .into_bytes();
        let package = bounded_package(archive(&xml, Compression::Store, false, Topology::Plain));
        let mut output = Vec::new();
        package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output))
            .expect("XML 1.0 standalone yes/no declarations are admitted");
        assert_eq!(semantic_texts(&output), ["seed", "tail"]);
    }
}

#[test]
fn protected_external_unsupported_and_signed_topologies_refuse_before_output() {
    let xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    for topology in [
        Topology::Protected,
        Topology::External,
        Topology::UnsupportedDependency,
        Topology::Signed,
    ] {
        assert_refused(xml.clone(), topology);
    }
}

#[test]
fn topology_refusals_keep_their_public_reason_and_emit_no_prefix() {
    let xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let cases = [
        (Topology::Protected, Refusal::Protection),
        (Topology::External, Refusal::ExternalRelationship),
        (
            Topology::UnsupportedDependency,
            Refusal::UnsupportedDependency,
        ),
        (Topology::Signed, Refusal::SignatureInfrastructure),
    ];
    for (topology, expected) in cases {
        let package = bounded_package(archive(&xml, Compression::Store, false, topology));
        let mut output = Vec::new();
        let result = package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output));
        match result {
            Err(TailAppendError::Refused(actual)) => assert_eq!(actual, expected),
            Err(error) => panic!("unexpected topology error: {error}"),
            Ok(_) => panic!("topology was admitted unexpectedly"),
        }
        assert!(output.is_empty());
    }
}

#[test]
fn settings_protection_flags_honor_explicit_on_and_off_values() {
    let cases = [
        (
            "documentProtection enabled",
            r#"<w:documentProtection w:edit="readOnly" w:enforcement="1"/>"#,
            true,
        ),
        (
            "documentProtection disabled",
            r#"<w:documentProtection w:edit="readOnly" w:enforcement="0"/>"#,
            false,
        ),
        (
            "writeProtection enabled",
            r#"<w:writeProtection w:val="1"/>"#,
            true,
        ),
        (
            "writeProtection disabled",
            r#"<w:writeProtection w:val="0"/>"#,
            false,
        ),
        (
            "trackRevisions enabled",
            r#"<w:trackRevisions w:val="1"/>"#,
            true,
        ),
        (
            "trackRevisions disabled",
            r#"<w:trackRevisions w:val="0"/>"#,
            false,
        ),
    ];
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        for (label, body, refused) in cases {
            let settings = settings_xml_with_body(strict, body);
            let source_archive = settings_archive(
                &source_xml,
                Compression::Store,
                strict,
                &settings,
                None,
                None,
            );
            let package = bounded_package(source_archive);
            let mut output = Vec::new();
            let result = package
                .tail_append_plain_paragraph("tail")
                .with_limits(limits(source_xml.len()))
                .prepare()
                .and_then(|plan| plan.write_to_stream(&mut output));
            if refused {
                match result {
                    Err(TailAppendError::Refused(Refusal::Protection)) => {},
                    Err(error) => panic!("{label} returned the wrong refusal: {error}"),
                    Ok(_) => panic!("{label} was admitted unexpectedly"),
                }
                assert!(output.is_empty(), "{label} must emit no archive prefix");
            } else {
                result.expect("an explicitly disabled protection flag must be admitted");
                assert_eq!(semantic_texts(&output), ["seed", "tail"]);
            }
        }
    }
}

#[test]
fn supported_settings_mce_fallback_is_admitted_at_exact_source_limit() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let namespace = if strict { STRICT_WORD } else { WORD };
        let settings = format!(
            r#"<w:settings xmlns:w="{namespace}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><w:docVars/></mc:Choice><mc:Fallback><w:docVars/></mc:Fallback></mc:AlternateContent></w:settings>"#
        )
        .into_bytes();
        let source_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &settings,
            None,
            None,
        );
        let package = bounded_package(source_archive.clone());
        let mut output = Vec::new();
        package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(source_xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output))
            .expect("unsupported Choice must select the valid MCE fallback");
        assert_eq!(semantic_texts(&output), ["seed", "tail"]);
        assert_eq!(
            settings_bytes(&source_archive),
            settings_bytes(&output),
            "MCE validation must leave the source settings member byte-identical"
        );

        // The scoped namespace codec no longer redeclares every inherited
        // binding on every element, so this fallback fits the source ceiling.
        let mut exact_output = Vec::new();
        let mut exact_limits = limits(source_xml.len());
        exact_limits.max_settings_xml_bytes =
            u64::try_from(settings.len()).expect("settings fixture length fits u64");
        bounded_package(source_archive.clone())
            .tail_append_plain_paragraph("tail")
            .with_limits(exact_limits)
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut exact_output))
            .expect("scoped namespace output fits the exact source-size ceiling");
        assert_eq!(semantic_texts(&exact_output), ["seed", "tail"]);
        assert_eq!(
            settings_bytes(&source_archive),
            settings_bytes(&exact_output)
        );

        let mut output_limited = Vec::new();
        exact_limits.max_settings_xml_bytes -= 1;
        let result = bounded_package(source_archive)
            .tail_append_plain_paragraph("tail")
            .with_limits(exact_limits)
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output_limited));
        assert!(
            matches!(
                &result,
                Err(TailAppendError::Limit {
                    resource: "settings XML bytes",
                    ..
                })
            ),
            "source-size ceiling must still reject before publication: {:?}",
            result.as_ref().err()
        );
        assert!(
            output_limited.is_empty(),
            "settings refusal must precede output"
        );
    }
}

#[test]
fn settings_mce_must_understand_unknown_namespace_refuses_before_publication() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let namespace = if strict { STRICT_WORD } else { WORD };
        let settings = format!(
            r#"<w:settings xmlns:w="{namespace}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported" mc:MustUnderstand="u"><w:docVars/></w:settings>"#
        )
        .into_bytes();
        let source_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &settings,
            None,
            None,
        );
        let package = bounded_package(source_archive);
        let mut output = Vec::new();
        let result = package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(source_xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output));
        assert!(
            matches!(
                &result,
                Err(TailAppendError::Document(DocumentError::Mce(MceError::MustUnderstand(required_namespace))))
                    if required_namespace == "urn:unsupported"
            ),
            "unknown MustUnderstand namespace must fail through the MCE document error: {:?}",
            result.as_ref().err()
        );
        assert!(
            output.is_empty(),
            "MCE MustUnderstand refusal must precede ZIP output"
        );
    }
}

#[test]
fn settings_relationship_closure_is_checked_without_staging_relationships() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let mail_merge_source = if strict {
            STRICT_MAIL_MERGE_SOURCE_REL
        } else {
            MAIL_MERGE_SOURCE_REL
        };
        let attached_template = if strict {
            STRICT_ATTACHED_TEMPLATE_REL
        } else {
            ATTACHED_TEMPLATE_REL
        };
        let recipient_data = if strict {
            STRICT_RECIPIENT_DATA_REL
        } else {
            RECIPIENT_DATA_REL
        };

        let valid_settings = settings_xml_with_relationships(
            strict,
            r#"<w:mailMerge><w:dataSource r:id="rIdSource"/></w:mailMerge>"#,
        );
        let valid_relationships = settings_relationships_xml(&format!(
            r#"<Relationship Id="rIdSource" Type="{mail_merge_source}" Target="mailMergeSource.bin"/>"#
        ));
        let source_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &valid_settings,
            Some(&valid_relationships),
            Some((MAIL_MERGE_SOURCE, b"inert mail merge source")),
        );
        let package = bounded_package(source_archive);
        let mut output = Vec::new();
        package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(source_xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output))
            .expect("valid mail-merge relationship closure must be admitted");
        assert_eq!(semantic_texts(&output), ["seed", "tail"]);

        let invalid_cases = [
            (
                "attachedTemplate missing relationship",
                settings_xml_with_relationships(
                    strict,
                    r#"<w:attachedTemplate r:id="rIdTemplate"/>"#,
                ),
                None,
            ),
            (
                "attachedTemplate wrong relationship type",
                settings_xml_with_relationships(
                    strict,
                    r#"<w:attachedTemplate r:id="rIdTemplate"/>"#,
                ),
                Some(settings_relationships_xml(&format!(
                    r#"<Relationship Id="rIdTemplate" Type="{mail_merge_source}" Target="mailMergeSource.bin"/>"#
                ))),
            ),
            (
                "attachedTemplate internal relationship",
                settings_xml_with_relationships(
                    strict,
                    r#"<w:attachedTemplate r:id="rIdTemplate"/>"#,
                ),
                Some(settings_relationships_xml(&format!(
                    r#"<Relationship Id="rIdTemplate" Type="{attached_template}" Target="mailMergeSource.bin"/>"#
                ))),
            ),
            (
                "mailMerge missing relationship",
                settings_xml_with_relationships(
                    strict,
                    r#"<w:mailMerge><w:dataSource r:id="rIdSource"/></w:mailMerge>"#,
                ),
                None,
            ),
            (
                "mailMerge wrong relationship type",
                settings_xml_with_relationships(
                    strict,
                    r#"<w:mailMerge><w:dataSource r:id="rIdSource"/></w:mailMerge>"#,
                ),
                Some(settings_relationships_xml(&format!(
                    r#"<Relationship Id="rIdSource" Type="{attached_template}" Target="mailMergeSource.bin"/>"#
                ))),
            ),
            (
                "orphan recipientData relationship",
                settings_xml_with_relationships(strict, ""),
                Some(settings_relationships_xml(&format!(
                    r#"<Relationship Id="rIdRecipients" Type="{recipient_data}" Target="mailMergeSource.bin"/>"#
                ))),
            ),
        ];
        for (label, settings, relationships) in invalid_cases {
            let source_archive = settings_archive(
                &source_xml,
                Compression::Store,
                strict,
                &settings,
                relationships.as_deref(),
                Some((MAIL_MERGE_SOURCE, b"inert mail merge source")),
            );
            let package = bounded_package(source_archive);
            let mut output = Vec::new();
            let result = package
                .tail_append_plain_paragraph("tail")
                .with_limits(limits(source_xml.len()))
                .prepare()
                .and_then(|plan| plan.write_to_stream(&mut output));
            assert!(
                matches!(&result, Err(TailAppendError::Document(_))),
                "{label} must fail through settings relationship validation: {:?}",
                result.as_ref().err()
            );
            assert!(output.is_empty(), "{label} must emit no archive prefix");
        }
    }
}

#[test]
fn empty_settings_root_is_admitted_in_both_word_dialects() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let namespace = if strict { STRICT_WORD } else { WORD };
        let settings = format!(r#"<w:settings xmlns:w="{namespace}"/>"#).into_bytes();
        let source_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &settings,
            None,
            None,
        );
        let package = bounded_package(source_archive);
        let mut output = Vec::new();
        package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(source_xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output))
            .expect("a self-closing settings root must be a valid empty settings part");
        assert_eq!(semantic_texts(&output), ["seed", "tail"]);
    }
}

#[test]
fn settings_xml_declaration_inside_root_refuses_before_publication() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let namespace = if strict { STRICT_WORD } else { WORD };
        let settings = format!(
            r#"<w:settings xmlns:w="{namespace}"><?xml version="1.0" encoding="UTF-8"?></w:settings>"#
        )
        .into_bytes();
        let source_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &settings,
            None,
            None,
        );
        let package = bounded_package(source_archive);
        let mut output = Vec::new();
        let result = package
            .tail_append_plain_paragraph("tail")
            .with_limits(limits(source_xml.len()))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut output));
        assert!(
            result.is_err(),
            "settings declaration inside root must refuse"
        );
        assert!(
            output.is_empty(),
            "settings declaration refusal must emit no output"
        );
    }
}

#[test]
fn inherited_settings_namespace_copies_hit_the_workspace_ceiling() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let small_settings = settings_with_many_unused_namespaces(strict, 1);
        let small_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &small_settings,
            None,
            None,
        );
        let small_package = bounded_package(small_archive);
        let mut small_output = Vec::new();
        small_package
            .tail_append_plain_paragraph("tail")
            .with_limits(adversarial_settings_limits(source_xml.len(), 4096))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut small_output))
            .expect("one extension with long inherited namespaces fits the workspace envelope");
        assert_eq!(semantic_texts(&small_output), ["seed", "tail"]);

        let large_settings = settings_with_many_unused_namespaces(strict, 128);
        let large_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &large_settings,
            None,
            None,
        );
        let large_package = bounded_package(large_archive);
        let mut large_output = Vec::new();
        let result = large_package
            .tail_append_plain_paragraph("tail")
            .with_limits(adversarial_settings_limits(source_xml.len(), 4096))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut large_output));
        match result {
            Err(TailAppendError::Limit {
                resource: "settings XML workspace",
                ..
            }) => {},
            Err(error) => panic!(
                "inherited namespace copies must hit the settings workspace ceiling: {error}"
            ),
            Ok(_) => panic!("oversized inherited namespace fixture was admitted"),
        }
        assert!(large_output.is_empty());
    }
}

#[test]
fn nested_long_mce_directives_account_for_decoded_expanded_names() {
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        let small_settings = settings_with_long_mce_directives(strict, 4);
        let small_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &small_settings,
            None,
            None,
        );
        let small_package = bounded_package(small_archive);
        let mut small_output = Vec::new();
        small_package
            .tail_append_plain_paragraph("tail")
            .with_limits(adversarial_settings_limits(source_xml.len(), 4096))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut small_output))
            .expect("small nested MCE directive fixture must remain admitted");
        assert_eq!(semantic_texts(&small_output), ["seed", "tail"]);

        let large_settings = settings_with_long_mce_directives(strict, 400);
        let large_archive = settings_archive(
            &source_xml,
            Compression::Store,
            strict,
            &large_settings,
            None,
            None,
        );
        let large_package = bounded_package(large_archive);
        let mut large_output = Vec::new();
        let result = large_package
            .tail_append_plain_paragraph("tail")
            .with_limits(adversarial_settings_limits(source_xml.len(), 4096))
            .prepare()
            .and_then(|plan| plan.write_to_stream(&mut large_output));
        match result {
            Err(TailAppendError::Limit {
                resource: "settings XML workspace",
                ..
            }) => {},
            Err(error) => panic!(
                "expanded MCE directive names must hit the settings workspace ceiling: {error}"
            ),
            Ok(_) => panic!("oversized nested MCE directive fixture was admitted"),
        }
        assert!(large_output.is_empty());
    }
}

#[test]
fn settings_scalar_duplicates_and_malformed_on_off_values_refuse_before_output() {
    let cases = [
        (
            "duplicate documentProtection enforcement attributes",
            r#"<w:documentProtection w:enforcement="0" w:enforcement="0"/>"#,
            true,
        ),
        (
            "duplicate trackRevisions val attributes",
            r#"<w:trackRevisions w:val="0" w:val="0"/>"#,
            true,
        ),
        (
            "duplicate writeProtection val attributes",
            r#"<w:writeProtection w:val="0" w:val="0"/>"#,
            true,
        ),
        (
            "documentProtection enabled lexical value",
            r#"<w:documentProtection w:enforcement="enabled"/>"#,
            false,
        ),
        (
            "documentProtection disabled lexical value",
            r#"<w:documentProtection w:enforcement="disabled"/>"#,
            false,
        ),
        (
            "trackRevisions enabled lexical value",
            r#"<w:trackRevisions w:val="enabled"/>"#,
            false,
        ),
        (
            "trackRevisions disabled lexical value",
            r#"<w:trackRevisions w:val="disabled"/>"#,
            false,
        ),
        (
            "writeProtection enabled lexical value",
            r#"<w:writeProtection w:val="enabled"/>"#,
            false,
        ),
        (
            "writeProtection disabled lexical value",
            r#"<w:writeProtection w:val="disabled"/>"#,
            false,
        ),
    ];
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        for (label, body, duplicate_attribute) in cases {
            let settings = settings_xml_with_body(strict, body);
            let source_archive = settings_archive(
                &source_xml,
                Compression::Store,
                strict,
                &settings,
                None,
                None,
            );
            let package = bounded_package(source_archive);
            let mut output = Vec::new();
            let result = package
                .tail_append_plain_paragraph("tail")
                .with_limits(limits(source_xml.len()))
                .prepare()
                .and_then(|plan| plan.write_to_stream(&mut output));
            let complete_parser_error = matches!(&result, Err(TailAppendError::Document(_)));
            let guarded_duplicate_error =
                duplicate_attribute && matches!(&result, Err(TailAppendError::Scan(_)));
            assert!(
                complete_parser_error || guarded_duplicate_error,
                "{label} must be rejected without weakening lexical-value checks: {:?}",
                result.as_ref().err()
            );
            assert!(output.is_empty(), "{label} must emit no archive prefix");
        }
    }
}

#[test]
fn repeated_protection_and_tracking_elements_refuse_before_output() {
    let cases = [
        (
            "documentProtection enabled then disabled",
            r#"<w:documentProtection w:enforcement="1"/><w:documentProtection w:enforcement="0"/>"#,
        ),
        (
            "documentProtection disabled then enabled",
            r#"<w:documentProtection w:enforcement="0"/><w:documentProtection w:enforcement="1"/>"#,
        ),
        (
            "documentProtection same value",
            r#"<w:documentProtection w:enforcement="1"/><w:documentProtection w:enforcement="1"/>"#,
        ),
        (
            "trackRevisions enabled then disabled",
            r#"<w:trackRevisions w:val="1"/><w:trackRevisions w:val="0"/>"#,
        ),
        (
            "trackRevisions disabled then enabled",
            r#"<w:trackRevisions w:val="0"/><w:trackRevisions w:val="1"/>"#,
        ),
        (
            "trackRevisions same value",
            r#"<w:trackRevisions w:val="1"/><w:trackRevisions w:val="1"/>"#,
        ),
    ];
    for strict in [false, true] {
        let source_xml = simple_document(strict);
        for (label, body) in cases {
            let settings = settings_xml_with_body(strict, body);
            let source_archive = settings_archive(
                &source_xml,
                Compression::Store,
                strict,
                &settings,
                None,
                None,
            );
            let package = bounded_package(source_archive);
            let mut output = Vec::new();
            let result = package
                .tail_append_plain_paragraph("tail")
                .with_limits(limits(source_xml.len()))
                .prepare()
                .and_then(|plan| plan.write_to_stream(&mut output));
            assert!(
                matches!(&result, Err(TailAppendError::Document(_))),
                "{label} must be rejected by settings cardinality validation: {:?}",
                result.as_ref().err()
            );
            assert!(output.is_empty(), "{label} must emit no archive prefix");
        }
    }
}

#[test]
fn malformed_mail_merge_schema_refuses_before_output() {
    for strict in [false, true] {
        let namespace = if strict { STRICT_WORD } else { WORD };
        let cases = [
            (
                "mailMerge namespace spoof",
                format!(
                    r#"<w:settings xmlns:w="{namespace}" xmlns:x="urn:spoof"><x:mailMerge/></w:settings>"#
                )
                .into_bytes(),
            ),
            (
                "duplicate direct mailMerge structural owner",
                settings_xml_with_body(strict, r#"<w:mailMerge/><w:mailMerge/>"#),
            ),
            (
                "duplicate odso structural owner",
                settings_xml_with_body(
                    strict,
                    r#"<w:mailMerge><w:odso/><w:odso/></w:mailMerge>"#,
                ),
            ),
            (
                "duplicate mailMerge child",
                settings_xml_with_body(
                    strict,
                    r#"<w:mailMerge><w:linkToQuery/><w:linkToQuery/></w:mailMerge>"#,
                ),
            ),
            (
                "mailMerge child namespace spoof",
                format!(
                    r#"<w:settings xmlns:w="{namespace}" xmlns:x="urn:spoof"><w:mailMerge><x:linkToQuery/></w:mailMerge></w:settings>"#
                )
                .into_bytes(),
            ),
            (
                "mailMerge child out of schema order",
                settings_xml_with_relationships(
                    strict,
                    r#"<w:mailMerge><w:dataSource r:id="rIdSource"/><w:linkToQuery/></w:mailMerge>"#,
                ),
            ),
            (
                "mailMerge malformed on/off value",
                settings_xml_with_body(
                    strict,
                    r#"<w:mailMerge><w:linkToQuery w:val="enabled"/></w:mailMerge>"#,
                ),
            ),
        ];
        let source_xml = simple_document(strict);
        for (label, settings) in cases {
            let source_archive = settings_archive(
                &source_xml,
                Compression::Store,
                strict,
                &settings,
                None,
                None,
            );
            let package = bounded_package(source_archive);
            let mut output = Vec::new();
            let result = package
                .tail_append_plain_paragraph("tail")
                .with_limits(limits(source_xml.len()))
                .prepare()
                .and_then(|plan| plan.write_to_stream(&mut output));
            assert!(
                matches!(&result, Err(TailAppendError::Document(_))),
                "{label} must be rejected by the complete mailMerge parser: {:?}",
                result.as_ref().err()
            );
            assert!(output.is_empty(), "{label} must emit no archive prefix");
        }
    }
}

#[test]
fn append_and_source_limits_are_checked_before_output() {
    let xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let archive = archive(&xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(archive.clone());

    let mut too_small_text = limits(xml.len());
    too_small_text.max_text_bytes = 1;
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("text exceeds limit")
        .with_limits(too_small_text)
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    assert!(result.is_err());
    assert!(output.is_empty());

    let mut too_small_source = limits(xml.len());
    too_small_source.max_source_xml_bytes = 1;
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("tail")
        .with_limits(too_small_source)
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    assert!(result.is_err());
    assert!(output.is_empty());
}

#[test]
fn one_large_token_is_rejected_by_the_token_ceiling_before_output() {
    let long_text = "x".repeat(8193);
    let source_xml = word_document(&format!(r#"<w:p><w:r><w:t>{long_text}</w:t></w:r></w:p>"#));
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let mut token_limited = limits(source_xml.len());
    token_limited.max_token_bytes = 8192;
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("tail")
        .with_limits(token_limited)
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    assert!(result.is_err());
    assert!(
        output.is_empty(),
        "overlong token must fail before ZIP output"
    );
}

#[test]
fn bounded_prepare_and_publication_keep_main_xml_out_of_the_payload_cache() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Deflate, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let before = package.cache_diagnostics();
    let plan = append_plan(&package, "tail", source_xml.len());
    let prepared = package.cache_diagnostics();
    assert_eq!(prepared.cold_loads, before.cold_loads);
    assert_eq!(prepared.successful_loads, before.successful_loads);
    assert_eq!(prepared.retained_entries, before.retained_entries);

    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("bounded publication must succeed");
    let after = package.cache_diagnostics();
    assert_eq!(after.cold_loads, before.cold_loads);
    assert_eq!(after.successful_loads, before.successful_loads);
    assert_eq!(after.retained_entries, before.retained_entries);
}

#[test]
fn source_change_after_preparation_is_rejected_without_output() {
    let xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let archive = archive(&xml, Compression::Deflate, false, Topology::Plain);
    let source = Arc::new(VersionedSource::new(archive.clone()));
    let source_reader: Arc<dyn ReadAt> = source.clone();
    let package =
        source_backed::Package::from_read_at(source_reader).expect("versioned source must open");
    let plan = append_plan(&package, "tail", xml.len());
    source.bump();
    let mut output = Vec::new();
    assert!(plan.write_to_stream(&mut output).is_err());
    assert!(output.is_empty(), "stale source must fail before output");
}

#[test]
fn encrypted_metadata_refusal_precedes_main_payload_reads_and_output() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let encrypted_archive = mark_member_encrypted(
        archive(&source_xml, Compression::Store, false, Topology::Plain),
        MAIN,
    );
    let source = Arc::new(ReadCountingSource::new(encrypted_archive));
    let source_reader: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(source_reader)
        .expect("encrypted metadata fixture must open without payload reads");
    source.arm();
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    assert!(result.is_err(), "encrypted metadata must refuse an append");
    assert!(
        output.is_empty(),
        "encrypted refusal must precede archive output"
    );
    assert_eq!(
        source.reads(),
        0,
        "encrypted metadata refusal must not read the main payload"
    );
}

#[test]
fn sink_failure_reports_partial_output_after_preflight() {
    let xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let archive = archive(&xml, Compression::Store, false, Topology::Plain);
    let package = bounded_package(archive);
    let plan = append_plan(&package, "tail", xml.len());
    let mut sink = PrefixFailSink::after(19);
    let result = plan.write_to_stream(&mut sink);
    assert!(result.is_err(), "short sink must fail after output starts");
    assert!(!sink.bytes.is_empty());
}

#[test]
fn managed_prepare_charges_semantic_memory_and_releases_plan_workspace() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Deflate, false, Topology::Plain);
    let (budget, _cancellation_source, context) = managed_context(8 * 1024 * 1024, u64::MAX);
    let source = Arc::new(MemoryObservingSource::new(source_archive, budget.clone()));
    let source_reader: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at_with_execution_context(
        source_reader,
        ReadLimits::default(),
        context,
    )
    .expect("managed source-backed DOCX fixture must open");
    let baseline = budget.used(Resource::Memory);
    source.arm();

    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .prepare()
        .expect("managed semantic passes must prepare");
    assert!(
        source.peak() > baseline,
        "the source callback must observe a live semantic or reader reservation"
    );
    assert!(
        budget.used(Resource::Memory) > baseline,
        "the retained splice plan must keep its bounded workspace charged"
    );

    drop(plan);
    assert_eq!(budget.used(Resource::Memory), baseline);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_memory_limit_refuses_before_archive_output() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);
    let (budget, _cancellation_source, package) =
        managed_package(source_archive, 8 * 1024, u64::MAX);
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    let resource = match result {
        Err(TailAppendError::Execution(ExecutionError::ResourceLimit(limit))) => {
            Some(limit.resource)
        },
        Err(TailAppendError::Document(DocumentError::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit),
        )))) => Some(limit.resource),
        _ => None,
    };
    assert_eq!(resource, Some(Resource::Memory));
    assert!(output.is_empty(), "memory refusal must precede ZIP output");
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_work_limit_exhaustion_refuses_before_archive_output() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);
    let source_work = u64::try_from(source_xml.len()).expect("fixture XML length fits u64");
    let (budget, _cancellation_source, package) = managed_package(
        source_archive,
        8 * 1024 * 1024,
        source_work.saturating_add(1),
    );
    let mut output = Vec::new();
    let result = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .prepare()
        .and_then(|plan| plan.write_to_stream(&mut output));
    let resource = match result {
        Err(TailAppendError::Execution(ExecutionError::ResourceLimit(limit))) => {
            Some(limit.resource)
        },
        Err(TailAppendError::Document(DocumentError::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit),
        )))) => Some(limit.resource),
        _ => None,
    };
    assert_eq!(resource, Some(Resource::Work));
    assert!(output.is_empty(), "work refusal must precede ZIP output");
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_during_short_publication_is_typed_after_a_prefix() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Deflate, false, Topology::Plain);
    let (_budget, cancellation_source, package) =
        managed_package(source_archive, 8 * 1024 * 1024, u64::MAX);
    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .prepare()
        .expect("managed publication plan must prepare");
    let mut sink = CancelOnShortWriteSink {
        cancellation: cancellation_source,
        bytes: Vec::new(),
        chunk: 1,
        cancelled: false,
    };
    let result = plan.write_to_stream(&mut sink);
    match result {
        Err(TailAppendError::Document(DocumentError::Opc(OpcError::IncompleteOutput {
            written,
            source,
        }))) => {
            assert_eq!(
                written,
                u64::try_from(sink.bytes.len()).expect("sink prefix length fits u64")
            );
            assert_cancelled_publication_source(&source);
        },
        _ => panic!("short-write cancellation must be typed"),
    }
    assert!(sink.cancelled);
    assert!(
        !sink.bytes.is_empty(),
        "publication cancellation must report the accepted prefix"
    );
}

#[test]
fn options_cancellation_during_short_publication_is_typed_after_a_prefix() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Deflate, false, Topology::Plain);
    let package = bounded_package(source_archive);
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(limits(source_xml.len()))
        .with_options(TailAppendOptions::default().with_cancellation_token(&cancellation))
        .prepare()
        .expect("options-only cancellation plan must prepare");
    let mut sink = CancelOnShortWriteSink {
        cancellation: cancellation_source,
        bytes: Vec::new(),
        chunk: 1,
        cancelled: false,
    };
    let result = plan.write_to_stream(&mut sink);
    match result {
        Err(TailAppendError::Document(DocumentError::Opc(OpcError::IncompleteOutput {
            written,
            source,
        }))) => {
            assert_eq!(
                written,
                u64::try_from(sink.bytes.len()).expect("sink prefix length fits u64")
            );
            assert_cancelled_publication_source(&source);
        },
        _ => panic!("options-only short-write cancellation must be typed"),
    }
    assert!(sink.cancelled);
    assert!(
        !sink.bytes.is_empty(),
        "options-only cancellation must report the accepted prefix"
    );
}

#[test]
fn source_backed_append_accepts_one_byte_reads_and_retries_interrupted_once() {
    let source_xml = word_document(r#"<w:p><w:r><w:t>seed</w:t></w:r></w:p>"#);
    let source_archive = archive(&source_xml, Compression::Store, false, Topology::Plain);

    let one_byte: Arc<dyn ReadAt> = Arc::new(OneByteSource::new(source_archive.clone()));
    let package =
        source_backed::Package::from_read_at(one_byte).expect("one-byte source must open");
    let plan = append_plan(&package, "tail", source_xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("one-byte source must publish");
    assert_eq!(semantic_texts(&output), ["seed", "tail"]);

    let interrupted = Arc::new(InterruptedOnceSource::new(source_archive));
    let interrupted_reader: Arc<dyn ReadAt> = interrupted.clone();
    let package = source_backed::Package::from_read_at(interrupted_reader)
        .expect("interrupted source must open before it is armed");
    interrupted.arm();
    let plan = append_plan(&package, "tail", source_xml.len());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("one interrupted source read must be retried");
    assert_eq!(semantic_texts(&output), ["seed", "tail"]);
}
