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
use litchi_odf_common::core::{OwnedPackage, PackageWriter, Profile, SourcePackageLimits};
use litchi_odt::{Document, ReadLimits, SourceBackedDocument};
use soapberry_zip::office::StreamingArchiveWriter;

const MIME: &str = "application/vnd.oasis.opendocument.text";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const DC: &str = "http://purl.org/dc/elements/1.1/";
const META: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

fn content_xml() -> String {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}" xmlns:table="{TABLE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}" xmlns:style="{STYLE}"><office:automatic-styles><style:style style:name="Body" style:family="paragraph"/></office:automatic-styles><office:body><office:text><text:h text:outline-level="1">Heading</text:h><text:p text:style-name="Body">hello <text:span>world</text:span></text:p><table:table table:name="Table1"><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="2"><text:p>cell</text:p></table:table-cell></table:table-row></table:table><draw:frame><draw:image xlink:href="Pictures/photo.bin"/></draw:frame><text:p>tail</text:p></office:text></office:body></office:document-content>"#
    )
}

fn styles_xml() -> String {
    format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:fo="{FO}"><office:styles><style:style style:name="Named" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style></office:styles></office:document-styles>"#
    )
}

fn valid_meta_xml() -> String {
    format!(
        r#"<office:document-meta xmlns:office="{OFFICE}" xmlns:dc="{DC}" xmlns:meta="{META}"><office:meta><dc:title>Source title</dc:title><dc:creator>Source author</dc:creator><meta:generator>test</meta:generator></office:meta></office:document-meta>"#
    )
}

fn package_with(
    content: &[u8],
    styles: Option<&[u8]>,
    meta: Option<&[u8]>,
    media: Option<&[u8]>,
) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content).unwrap();
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles).unwrap();
    }
    if let Some(meta) = meta {
        writer.add_file("meta.xml", meta).unwrap();
    }
    if let Some(media) = media {
        writer
            .add_file_with_media_type("Pictures/photo.bin", media, "application/octet-stream")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn package() -> Vec<u8> {
    package_with(
        content_xml().as_bytes(),
        Some(styles_xml().as_bytes()),
        Some(valid_meta_xml().as_bytes()),
        Some(&incompressible_media()),
    )
}

fn malformed_text_package() -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}"><office:body><office:text><text:p>A<text:s text:c="1000001"/></text:p></office:text></office:body></office:document-content>"#
    );
    package_with(content.as_bytes(), None, None, None)
}

fn oversized_text_package() -> Vec<u8> {
    let paragraph = "x".repeat(1024 * 1024);
    let mut content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}"><office:body><office:text><text:p>"#
    );
    for _ in 0..17 {
        content.push_str(&paragraph);
    }
    content.push_str("</text:p></office:text></office:body></office:document-content>");
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
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption(password, Profile::compatible())
        .unwrap();
    writer
        .add_file("content.xml", content_xml().as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn incompressible_media() -> Vec<u8> {
    let mut state = 0x9e37_79b9_u32;
    let mut bytes = Vec::with_capacity(64 * 1024);
    for _ in 0..(64 * 1024) {
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
            if entry.file_path().as_ref() != b"Pictures/photo.bin" {
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

    fn bytes_read(&self) -> usize {
        self.bytes_read.load(Ordering::Relaxed)
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Relaxed)
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
            self.ranges
                .lock()
                .unwrap()
                .push((offset, offset + u64::try_from(amount).unwrap()));
        }
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::Relaxed);
        Ok(SourceVersion::new(
            0x4f44_5401,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn source_facade_matches_owned_semantics_and_can_materialize() {
    let bytes = package();
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    let source = SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();

    assert_eq!(source.text().unwrap(), eager.text().unwrap());
    assert_eq!(
        source.paragraph_count().unwrap(),
        eager.paragraph_count().unwrap()
    );
    assert_eq!(
        source
            .paragraphs()
            .unwrap()
            .into_iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>(),
        eager
            .paragraphs()
            .unwrap()
            .into_iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>()
    );
    assert_eq!(
        source.tables().unwrap().len(),
        eager.tables().unwrap().len()
    );
    assert_eq!(
        source.elements().unwrap().len(),
        eager.elements().unwrap().len()
    );
    assert_eq!(
        source.tables_expanded().unwrap().len(),
        eager.tables_expanded().unwrap().len()
    );
    assert_eq!(
        source.metadata().unwrap().title,
        eager.metadata().unwrap().title
    );
    assert_eq!(
        source.metadata().unwrap().author,
        eager.metadata().unwrap().author
    );
    assert_eq!(
        source
            .odf_metadata()
            .unwrap()
            .and_then(|metadata| metadata.title),
        eager
            .odf_metadata()
            .unwrap()
            .and_then(|metadata| metadata.title)
    );
    assert_eq!(
        source.styles().unwrap().styles.len(),
        eager.styles().styles.len()
    );
    assert_eq!(
        source
            .get_style_properties("Named")
            .unwrap()
            .paragraph
            .text_align,
        eager.get_style_properties("Named").paragraph.text_align
    );
    assert_eq!(source.protection().unwrap(), eager.protection().unwrap());
    assert!(
        source
            .files()
            .unwrap()
            .iter()
            .any(|path| path == "content.xml")
    );

    let materialized = source.materialize().unwrap();
    assert_eq!(materialized.text().unwrap(), eager.text().unwrap());
}

#[test]
fn source_facade_text_cache_has_thresholded_freshness_vector_and_fresh_hits() {
    let bytes = package();
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();
    let expected = eager.text().unwrap();
    let reads_before_text = source.reads();
    let versions_before_text = source.versions();

    assert_eq!(document.text().unwrap(), expected);
    let after_first = source.versions();
    assert_eq!(after_first - versions_before_text, 2);

    assert_eq!(document.text().unwrap(), expected);
    let after_publication = source.versions();
    assert_eq!(after_publication - after_first, 4);

    let cached_first = document.text().unwrap();
    let after_first_hit = source.versions();
    assert_eq!(after_first_hit - after_publication, 2);
    let cached_second = document.text().unwrap();
    assert_eq!(source.versions() - after_first_hit, 2);
    assert_ne!(cached_first.as_ptr(), cached_second.as_ptr());
    assert_eq!(cached_first, expected);
    assert_eq!(cached_second, expected);
    assert_eq!(source.reads(), reads_before_text);
}

#[test]
fn source_facade_text_cache_is_safe_for_concurrent_first_construction() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SourceBackedDocument>();

    let bytes = package();
    let expected = Document::from_bytes(bytes.clone()).unwrap().text().unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let document = Arc::new(SourceBackedDocument::from_read_at(source.clone()).unwrap());

    let handles: Vec<_> = (0..8)
        .map(|_| {
            let document = Arc::clone(&document);
            std::thread::spawn(move || document.text())
        })
        .collect();
    for handle in handles {
        assert_eq!(handle.join().unwrap().unwrap(), expected);
    }

    let versions_before_cache_hit = source.versions();
    assert_eq!(document.text().unwrap(), expected);
    assert_eq!(source.versions() - versions_before_cache_hit, 2);
}

#[test]
fn source_facade_text_parse_errors_are_not_cached() {
    let source = Arc::new(CountingSource::new(malformed_text_package()));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();

    assert!(document.text().is_err());
    let versions_before_retry = source.versions();
    assert!(document.text().is_err());
    assert_eq!(
        source.versions() - versions_before_retry,
        2,
        "a parse error must leave the thresholded cache uninitialized so the parser is retried"
    );
}

#[test]
fn source_facade_oversized_text_falls_back_without_retention() {
    let source = Arc::new(CountingSource::new(oversized_text_package()));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();

    let first = document.text().unwrap();
    assert!(first.len() > 16 * 1024 * 1024);
    assert_eq!(document.text().unwrap(), first);

    let versions_before_fallback = source.versions();
    assert_eq!(document.text().unwrap(), first);
    assert_eq!(
        source.versions() - versions_before_fallback,
        3,
        "an oversized projection uses the parser fallback instead of retaining text"
    );
}

#[test]
fn source_facade_text_cache_refuses_stale_revision_after_publication() {
    let source = Arc::new(CountingSource::new(package()));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();
    assert!(document.text().is_ok());
    assert!(document.text().is_ok());

    source.bump_revision();
    assert!(matches!(document.text(), Err(Error::SourceChanged { .. })));
}

#[test]
fn source_facade_defers_unrelated_media_until_selected_image() {
    let bytes = package();
    let media_range = media_range(&bytes);
    let source = Arc::new(CountingSource::new(bytes));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();
    let before_media = source.bytes_read();

    assert!(
        source
            .ranges()
            .into_iter()
            .all(|range| !overlaps(range, media_range)),
        "open must not read selected media"
    );
    let images = document.images().unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(source.bytes_read(), before_media);
    assert_eq!(
        document.image_bytes(&images[0]).unwrap(),
        Some(incompressible_media())
    );
    assert!(source.bytes_read() > before_media);
    assert!(
        source
            .ranges()
            .into_iter()
            .any(|range| overlaps(range, media_range)),
        "selected image must read its compressed package range"
    );
}

#[test]
fn source_facade_refuses_forged_image_paths_without_reading_administrative_members() {
    let bytes = package();
    let source = Arc::new(CountingSource::new(bytes));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();
    let image = document.images().unwrap().pop().unwrap();
    let mut forged = image.clone();

    {
        let path = "content.xml";
        forged.source = litchi_odf_common::media::Source::PackagePart {
            href: path.to_string(),
            path: path.to_string(),
            manifest_media_type: None,
        };
        let before = source.bytes_read();
        let result = document.image_bytes(&forged);
        assert_eq!(result.unwrap(), None);
        assert_eq!(source.bytes_read(), before, "forged path {path} was read");
    }

    for path in [
        "mimetype",
        "Pictures/../mimetype",
        "META-INF/manifest.xml",
        "../outside.bin",
    ] {
        forged.source = litchi_odf_common::media::Source::PackagePart {
            href: path.to_string(),
            path: path.to_string(),
            manifest_media_type: None,
        };
        let before = source.bytes_read();
        assert!(document.image_bytes(&forged).is_err());
        assert_eq!(source.bytes_read(), before, "forged path {path} was read");
    }
}

#[test]
fn source_facade_reports_stale_source_before_queries_and_materialization() {
    let source = Arc::new(CountingSource::new(package()));
    let document = SourceBackedDocument::from_read_at(source.clone()).unwrap();
    source.bump_revision();

    assert!(matches!(document.text(), Err(Error::SourceChanged { .. })));
    assert!(matches!(
        document.images(),
        Err(Error::SourceChanged { .. })
    ));
    assert!(matches!(
        document.member_data("Pictures/photo.bin"),
        Err(Error::SourceChanged { .. })
    ));
    assert!(matches!(
        document.materialize(),
        Err(Error::SourceChanged { .. })
    ));
}

#[test]
fn source_facade_enforces_source_limits_and_family_validation() {
    let bytes = package();
    let limits = ReadLimits::default().with_max_source_bytes((bytes.len() - 1) as u64);
    assert!(matches!(
        SourceBackedDocument::from_read_at_with_limits(
            Arc::new(CountingSource::new(bytes.clone())),
            limits,
        ),
        Err(Error::ResourceLimit(_))
    ));

    let wrong_mime = {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        writer
            .add_file("content.xml", content_xml().as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    };
    assert!(SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(wrong_mime))).is_err());

    let wrong_root = package_with(
        br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
        None,
        None,
        None,
    );
    assert!(SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(wrong_root))).is_err());

    let invalid_content_utf8 = replace_member(
        &package_with(
            content_xml().as_bytes(),
            None,
            Some(valid_meta_xml().as_bytes()),
            None,
        ),
        "content.xml",
        &[0xff],
    );
    assert!(
        SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(invalid_content_utf8)))
            .is_err()
    );

    let invalid_styles_utf8 = replace_member(
        &package_with(
            content_xml().as_bytes(),
            Some(styles_xml().as_bytes()),
            Some(valid_meta_xml().as_bytes()),
            None,
        ),
        "styles.xml",
        &[0xff],
    );
    assert!(
        SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(invalid_styles_utf8)))
            .is_err()
    );

    let invalid_meta_utf8 = replace_member(
        &package_with(
            content_xml().as_bytes(),
            None,
            Some(valid_meta_xml().as_bytes()),
            None,
        ),
        "meta.xml",
        &[0xff],
    );
    assert!(
        SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(invalid_meta_utf8)))
            .is_err()
    );
}

#[test]
fn source_facade_retains_lazy_semantic_meta_errors() {
    let malformed_but_utf8 = format!(
        r#"<office:document-meta xmlns:office="{OFFICE}" xmlns:meta="{META}"><office:meta><meta:document-statistic meta:page-count="-1"/></office:meta></office:document-meta>"#
    );
    let bytes = package_with(
        content_xml().as_bytes(),
        None,
        Some(malformed_but_utf8.as_bytes()),
        None,
    );
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    let source = SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();
    assert!(eager.metadata().is_err());
    assert!(source.metadata().is_err());
    assert!(source.odf_metadata().is_err());
}

#[test]
fn source_facade_supports_passwords_and_exported_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SourceBackedDocument>();

    let password = "source-password";
    let bytes = password_package(password);
    let eager = Document::from_bytes_with_password(bytes.clone(), password).unwrap();
    let source = SourceBackedDocument::from_read_at_with_password(
        Arc::new(CountingSource::new(bytes.clone())),
        password,
    )
    .unwrap();
    assert_eq!(source.text().unwrap(), eager.text().unwrap());
    assert!(
        SourceBackedDocument::from_read_at_with_password(
            Arc::new(CountingSource::new(bytes)),
            "wrong-password",
        )
        .is_err()
    );

    let _: Option<ReadLimits> = None;
}

#[cfg(any(unix, windows))]
#[test]
fn source_facade_opens_filesystem_paths_without_changing_semantics() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.odt");
    let bytes = package();
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    std::fs::write(&path, bytes).unwrap();
    let source = SourceBackedDocument::from_path(&path).unwrap();
    assert_eq!(source.text().unwrap(), eager.text().unwrap());
}

#[test]
fn source_facade_rejects_unsafe_member_paths() {
    let source =
        SourceBackedDocument::from_read_at(Arc::new(CountingSource::new(package()))).unwrap();
    assert!(source.member_data("../outside.bin").is_err());
    assert!(source.media_data("../outside.bin").is_err());
}

#[allow(dead_code)]
fn _limits_type_is_shared() {
    let _: SourcePackageLimits = ReadLimits::default();
}

fn replace_member(bytes: &[u8], target: &str, replacement: &[u8]) -> Vec<u8> {
    let package = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    let names = package.files().unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in names {
        let data = if name == target {
            replacement.to_vec()
        } else {
            package.get_file(&name).unwrap()
        };
        writer.write_stored(&name, &data).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}
