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
