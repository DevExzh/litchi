#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use std::io;
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, AtomicUsize, Ordering},
};

use litchi_core::{Error, ReadAt, SourceVersion};
use litchi_odf_common::core::{PackageWriter, Profile, SourcePackageLimits};
use litchi_odp::model::Reference;
use litchi_odp::{Builder, Presentation, SourceBackedPresentation};
use soapberry_zip::office::StreamingArchiveWriter;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: AtomicUsize,
    bytes_read: AtomicUsize,
    revision: AtomicU64,
    versions: AtomicUsize,
    ranges: Mutex<Vec<(u64, u64)>>,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            bytes_read: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            versions: AtomicUsize::new(0),
            ranges: Mutex::new(Vec::new()),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Relaxed)
    }

    fn bytes_read(&self) -> usize {
        self.bytes_read.load(Ordering::Relaxed)
    }

    fn versions(&self) -> usize {
        self.versions.load(Ordering::Relaxed)
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }

    fn ranges(&self) -> Vec<(u64, u64)> {
        self.ranges.lock().unwrap().clone()
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("test source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let amount = input.len().min(output.len());
        output[..amount].copy_from_slice(&input[..amount]);
        self.bytes_read.fetch_add(amount, Ordering::Relaxed);
        if amount != 0 {
            let end = offset + u64::try_from(amount).unwrap();
            self.ranges.lock().unwrap().push((offset, end));
        }
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::Relaxed);
        Ok(SourceVersion::new(
            0x4f44_5001,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

fn media_package() -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation><draw:page draw:name="one"><draw:frame><draw:text-box><text:p>source backed</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer
        .add_file(
            "styles.xml",
            format!(r#"<office:document-styles xmlns:office="{OFFICE}"/>"#).as_bytes(),
        )
        .unwrap();
    writer
        .add_file("Media/clip.bin", &incompressible_media())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn text_package() -> Vec<u8> {
    let mut builder = Builder::new();
    builder.add_slide_with_title("First", "alpha").unwrap();
    builder.add_slide_with_title("Second", "beta").unwrap();
    builder.build().unwrap()
}

fn malformed_text_package() -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation><draw:page><draw:frame><draw:text-box><text:p><text:p>nested</text:p></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn oversized_text_package() -> Vec<u8> {
    let paragraph = "x".repeat(1024 * 1024);
    let mut content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation><draw:page>"#
    );
    for _ in 0..17 {
        content.push_str("<draw:frame><draw:text-box><text:p>");
        content.push_str(&paragraph);
        content.push_str("</text:p></draw:text-box></draw:frame>");
    }
    content.push_str("</draw:page></office:presentation></office:body></office:document-content>");
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
    );
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", manifest.as_bytes())
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

fn password_package(password: &str) -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:presentation/></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption(password, Profile::compatible())
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn incompressible_media() -> Vec<u8> {
    let mut state = 0x9e37_79b9_u32;
    let mut bytes = Vec::with_capacity(256 * 1024);
    for _ in 0..(256 * 1024) {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        bytes.push(state as u8);
    }
    bytes
}

fn media_range(bytes: &[u8]) -> (u64, u64) {
    member_range(bytes, b"Media/clip.bin")
}

fn member_range(bytes: &[u8], path: &[u8]) -> (u64, u64) {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes).unwrap();
    archive
        .entries()
        .filter_map(|entry| {
            let entry = entry.ok()?;
            if entry.file_path().as_ref() != path {
                return None;
            }
            Some(
                archive
                    .get_entry(entry.wayfinder())
                    .unwrap()
                    .compressed_data_range(),
            )
        })
        .next()
        .unwrap()
}

fn content_with_body(body: &str) -> String {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"#
    )
}

fn oversized_content_package() -> Vec<u8> {
    let content = content_with_body(
        r#"<draw:page><draw:frame><draw:text-box><text:p>small</text:p></draw:text-box></draw:frame></draw:page>"#,
    );
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
    );
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", manifest.as_bytes())
        .unwrap();
    let mut bytes = archive.finish_to_bytes().unwrap();
    let name = b"content.xml";
    let declared_size = u32::try_from(256 * 1024 * 1024 + 1).unwrap();
    let central_offset = bytes
        .windows(46 + name.len())
        .position(|window| {
            window.starts_with(b"PK\x01\x02")
                && u16::from_le_bytes([window[28], window[29]]) as usize == name.len()
                && &window[46..] == name
        })
        .unwrap();
    bytes[central_offset + 24..central_offset + 28].copy_from_slice(&declared_size.to_le_bytes());
    bytes
}

fn encrypted_oversized_content_package() -> Vec<u8> {
    let declared_size = 256_u64 * 1024 * 1024 + 1;
    let manifest = format!(
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="{MIME}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="{declared_size}"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="http://www.w3.org/2001/04/xmlenc#sha256" m:key-size="32"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#
    );
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive.write_stored("content.xml", &[0_u8; 32]).unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", manifest.as_bytes())
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

fn overlaps(left: (u64, u64), right: (u64, u64)) -> bool {
    left.0 < right.1 && right.0 < left.1
}

#[test]
fn source_facade_matches_owned_semantics_and_defers_media() {
    let bytes = media_package();
    let media_range = media_range(&bytes);
    let eager = Presentation::from_bytes(bytes.clone()).unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let source_presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();

    assert_eq!(
        source_presentation.content_xml().unwrap(),
        eager.content_xml()
    );
    assert_eq!(
        source_presentation.styles_xml().unwrap(),
        eager.styles_xml()
    );
    assert_eq!(
        source_presentation.slide_count().unwrap(),
        eager.slide_count().unwrap()
    );
    assert_eq!(
        source_presentation.slides().unwrap(),
        eager.slides().unwrap()
    );
    assert_eq!(source_presentation.text().unwrap(), eager.text().unwrap());

    let before_media = source.bytes_read();
    assert!(
        source
            .ranges()
            .into_iter()
            .all(|range| !overlaps(range, media_range)),
        "open and semantic reads must not touch selected media range"
    );
    let reference = Reference::new("Media/clip.bin").unwrap();
    assert_eq!(
        source_presentation.media_data(&reference).unwrap(),
        Some(incompressible_media())
    );
    assert!(source.bytes_read() > before_media);
    assert!(
        source
            .ranges()
            .into_iter()
            .any(|range| overlaps(range, media_range)),
        "selected media read must touch selected media range"
    );
}

#[test]
fn source_facade_text_cache_preserves_projection_and_checks_freshness() {
    let bytes = text_package();
    let eager = Presentation::from_bytes(bytes.clone()).unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let expected = eager.text().unwrap();
    let reads_before_text = source.reads();

    assert_eq!(presentation.text().unwrap(), expected);
    let versions_before_threshold_call = source.versions();
    assert_eq!(presentation.text().unwrap(), expected);
    assert!(source.versions() - versions_before_threshold_call >= 3);
    let versions_before_cache_hit = source.versions();
    assert_eq!(presentation.text().unwrap(), expected);
    assert_eq!(
        source.versions() - versions_before_cache_hit,
        2,
        "a retained text cache hit only observes source freshness before and after cloning"
    );
    assert_eq!(source.reads(), reads_before_text);
    assert_eq!(expected, "First\nalpha\n\nSecond\nbeta");

    source.bump_revision();
    assert!(matches!(
        presentation.text(),
        Err(Error::SourceChanged { .. })
    ));
}

#[test]
fn source_facade_text_cache_is_safe_for_concurrent_first_construction() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SourceBackedPresentation>();

    let bytes = text_package();
    let expected = Presentation::from_bytes(bytes.clone())
        .unwrap()
        .text()
        .unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let presentation = Arc::new(SourceBackedPresentation::from_read_at(source.clone()).unwrap());

    let handles: Vec<_> = (0..8)
        .map(|_| {
            let presentation = Arc::clone(&presentation);
            std::thread::spawn(move || presentation.text())
        })
        .collect();
    for handle in handles {
        assert_eq!(handle.join().unwrap().unwrap(), expected);
    }

    let versions_before_cache_hit = source.versions();
    assert_eq!(presentation.text().unwrap(), expected);
    assert_eq!(source.versions() - versions_before_cache_hit, 2);
}

#[test]
fn source_facade_text_parse_errors_are_not_cached() {
    let source = Arc::new(CountingSource::new(malformed_text_package()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();

    assert!(presentation.text().is_err());
    let versions_before_retry = source.versions();
    assert!(presentation.text().is_err());
    assert_eq!(
        source.versions() - versions_before_retry,
        2,
        "a parse error must leave the thresholded cache uninitialized so the parser is retried"
    );
}

#[test]
fn source_facade_oversized_text_falls_back_without_retention() {
    let source = Arc::new(CountingSource::new(oversized_text_package()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();

    let first = presentation.text().unwrap();
    assert!(first.len() > 16 * 1024 * 1024);
    assert_eq!(presentation.text().unwrap(), first);

    let versions_before_fallback = source.versions();
    assert_eq!(presentation.text().unwrap(), first);
    assert!(
        source.versions() - versions_before_fallback >= 3,
        "an oversized projection uses the parser fallback instead of retaining text"
    );
}

#[test]
fn source_facade_reports_typed_stale_source_errors() {
    let bytes = Builder::new().build().unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    source.bump_revision();

    assert!(matches!(
        presentation.slide_count(),
        Err(Error::SourceChanged { .. })
    ));
}

#[test]
fn source_facade_applies_physical_source_limits() {
    let bytes = media_package();
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let limits = SourcePackageLimits::new(
        u64::try_from(bytes.len() - 1).unwrap(),
        SourcePackageLimits::default().archive_limits(),
    );

    assert!(SourceBackedPresentation::from_read_at_with_limits(source, limits).is_err());
}

#[test]
fn source_facade_rejects_oversized_content_before_materialization() {
    let bytes = oversized_content_package();
    let content_range = member_range(&bytes, b"content.xml");
    let source = Arc::new(CountingSource::new(bytes));

    assert!(matches!(
        SourceBackedPresentation::from_read_at(source.clone()),
        Err(Error::InvalidFormat(message)) if message.contains("content.xml exceeds the family limit")
    ));
    assert!(
        source
            .ranges()
            .into_iter()
            .all(|range| !overlaps(range, content_range)),
        "declared oversized content must be rejected before its payload is read"
    );
}

#[test]
fn source_facade_rejects_encrypted_oversized_plaintext_before_materialization() {
    let bytes = encrypted_oversized_content_package();
    let content_range = member_range(&bytes, b"content.xml");
    let source = Arc::new(CountingSource::new(bytes));

    assert!(matches!(
        SourceBackedPresentation::from_read_at(source.clone()),
        Err(Error::InvalidFormat(message)) if message.contains("content.xml exceeds the family limit")
    ));
    assert!(
        source
            .ranges()
            .into_iter()
            .all(|range| !overlaps(range, content_range)),
        "oversized encrypted plaintext must be rejected before ciphertext is read"
    );
}

#[test]
fn source_facade_validates_metadata_like_the_owned_family() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:presentation/></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer
        .add_file(
            "meta.xml",
            br#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><o:meta><m:document-statistic m:page-count="-1"/></o:meta></o:document-meta>"#,
        )
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    assert!(Presentation::from_bytes(bytes.clone()).is_err());
    assert!(SourceBackedPresentation::from_read_at(Arc::new(CountingSource::new(bytes))).is_err());
}

#[test]
fn source_facade_can_materialize_the_existing_mutable_owner() {
    let bytes = Builder::new().build().unwrap();
    let source =
        SourceBackedPresentation::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();
    let materialized = source.materialize().unwrap();
    assert_eq!(materialized.slide_count().unwrap(), 0);
    assert_eq!(materialized.text().unwrap(), "");
}

#[test]
fn source_facade_supports_password_protected_content() {
    let password = "source-password";
    let bytes = password_package(password);
    let eager = Presentation::from_bytes_with_password(bytes.clone(), password).unwrap();
    let source = SourceBackedPresentation::from_read_at_with_password(
        Arc::new(CountingSource::new(bytes.clone())),
        password,
    )
    .unwrap();
    assert_eq!(source.content_xml().unwrap(), eager.content_xml());
    assert_eq!(source.text().unwrap(), eager.text().unwrap());
    assert!(
        SourceBackedPresentation::from_read_at_with_password(
            Arc::new(CountingSource::new(bytes)),
            "wrong-password",
        )
        .is_err()
    );
}
