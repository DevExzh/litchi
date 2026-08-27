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
use litchi_odf_common::core::{PackageWriter, SourcePackageLimits};
use litchi_odp::model::Reference;
use litchi_odp::{Presentation, SourceBackedPresentationCatalog};

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: AtomicUsize,
    bytes_read: AtomicUsize,
    revision: AtomicU64,
    ranges: Mutex<Vec<(u64, u64)>>,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            bytes_read: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            ranges: Mutex::new(Vec::new()),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Relaxed)
    }

    fn bytes_read(&self) -> usize {
        self.bytes_read.load(Ordering::Relaxed)
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
            self.ranges
                .lock()
                .unwrap()
                .push((offset, offset + u64::try_from(amount).unwrap()));
        }
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5002,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

fn package_with_content(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer
        .add_file("Media/clip.bin", b"catalog media payload")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn raw_package_with_content(content: &str) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer.write_stored("mimetype", MIME.as_bytes()).unwrap();
    writer
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    writer
        .write_deflated(
            "META-INF/manifest.xml",
            br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn package() -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation><draw:page draw:name="First"><draw:frame><draw:text-box><text:p>alpha</text:p></draw:text-box></draw:frame></draw:page><draw:page draw:name="Second"><draw:frame><draw:text-box><text:p>beta</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#
    );
    package_with_content(&content)
}

fn content_with_body(body: &str) -> String {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"#
    )
}

fn media_range(bytes: &[u8]) -> (u64, u64) {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes).unwrap();
    archive
        .entries()
        .filter_map(|entry| {
            let entry = entry.ok()?;
            if entry.file_path().as_ref() != b"Media/clip.bin" {
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

fn overlaps(left: (u64, u64), right: (u64, u64)) -> bool {
    left.0 < right.1 && right.0 < left.1
}

#[test]
fn catalog_retains_order_and_count_without_slide_payloads() {
    let source = Arc::new(CountingSource::new(package()));
    let catalog = SourceBackedPresentationCatalog::from_read_at(source).unwrap();

    let entries = catalog.catalog().unwrap();
    assert_eq!(entries.len(), 2);
    assert_eq!(entries[0].index(), 0);
    assert_eq!(entries[0].name(), Some("First"));
    assert_eq!(entries[1].index(), 1);
    assert_eq!(entries[1].name(), Some("Second"));
    assert_eq!(catalog.slide_count().unwrap(), 2);
}

#[test]
fn catalog_accepts_custom_namespace_prefixes_and_decodes_page_names() {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:d="{DRAW}" xmlns:t="{TEXT}"><o:body><o:presentation><d:page d:name="First &amp; &quot;slide&quot;"><d:frame><d:text-box><t:p>alpha</t:p></d:text-box></d:frame></d:page></o:presentation></o:body></o:document-content>"#
    );
    let source = Arc::new(CountingSource::new(package_with_content(&content)));
    let catalog = SourceBackedPresentationCatalog::from_read_at(source).unwrap();

    assert_eq!(catalog.slide_count().unwrap(), 1);
    assert_eq!(
        catalog.catalog().unwrap()[0].name(),
        Some("First & \"slide\"")
    );
}

#[test]
fn catalog_rejects_malformed_document_prologue_and_entities() {
    let valid = content_with_body(
        r#"<draw:page><draw:frame><draw:text-box><text:p>alpha</text:p></draw:text-box></draw:frame></draw:page>"#,
    );
    let malformed = [
        format!("leading text{valid}"),
        format!("{valid}trailing text"),
        format!("<![CDATA[leading text]]>{valid}"),
        format!("{valid}<?xml version=\"1.0\"?>"),
        valid.replace("alpha", "alpha &bogus;"),
    ];

    for content in malformed {
        let source = Arc::new(CountingSource::new(raw_package_with_content(&content)));
        assert!(matches!(
            SourceBackedPresentationCatalog::from_read_at(source),
            Err(Error::InvalidFormat(_))
        ));
    }
}

#[test]
fn selected_slide_rereads_content_and_matches_owned_semantics() {
    let bytes = package();
    let eager = Presentation::from_bytes(bytes.clone()).unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let catalog = SourceBackedPresentationCatalog::from_read_at(source.clone()).unwrap();
    let before = source.reads();

    let expected = eager.slides().unwrap().into_iter().nth(1).unwrap();
    let selected = catalog.slide_at(1).unwrap().unwrap();

    assert_eq!(selected, expected);
    assert!(source.reads() > before, "selection must reread content.xml");
}

#[test]
fn media_stays_cold_until_explicit_selection() {
    let bytes = package();
    let media_range = media_range(&bytes);
    let source = Arc::new(CountingSource::new(bytes));
    let catalog = SourceBackedPresentationCatalog::from_read_at(source.clone()).unwrap();

    catalog.slide_at(0).unwrap();
    assert!(
        source
            .ranges()
            .into_iter()
            .all(|range| !overlaps(range, media_range)),
        "catalog and selected slide must not read media"
    );
    let before_media = source.bytes_read();

    let reference = Reference::new("Media/clip.bin").unwrap();
    assert_eq!(
        catalog.media_data(&reference).unwrap(),
        Some(b"catalog media payload".to_vec())
    );
    assert!(source.bytes_read() > before_media);
    assert!(
        source
            .ranges()
            .into_iter()
            .any(|range| overlaps(range, media_range)),
        "explicit media selection must read the media range"
    );
}

#[test]
fn out_of_range_selection_does_not_read_payload() {
    let source = Arc::new(CountingSource::new(package()));
    let catalog = SourceBackedPresentationCatalog::from_read_at(source.clone()).unwrap();
    let before = source.reads();

    assert!(catalog.slide_at(2).unwrap().is_none());
    assert_eq!(source.reads(), before);
}

#[test]
fn catalog_reports_source_changes() {
    let source = Arc::new(CountingSource::new(package()));
    let catalog = SourceBackedPresentationCatalog::from_read_at(source.clone()).unwrap();
    source.bump_revision();

    assert!(matches!(
        catalog.catalog(),
        Err(Error::SourceChanged { .. })
    ));
    assert!(matches!(
        catalog.slide_at(0),
        Err(Error::SourceChanged { .. })
    ));
}

#[test]
fn catalog_applies_source_limits_before_content_scan() {
    let source = Arc::new(CountingSource::new(package()));
    let limits = SourcePackageLimits::default().with_max_source_bytes(1);

    assert!(SourceBackedPresentationCatalog::from_read_at_with_limits(source, limits).is_err());
}
