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
use litchi_odf_common::core::{PackageWriter, Profile};
use litchi_odt::{
    ReadLimits, SourceBackedDocumentCatalog,
    elements::text::{Kind, TextElements},
};
use soapberry_zip::office::StreamingArchiveWriter;

const MIME: &str = "application/vnd.oasis.opendocument.text";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

fn content_xml() -> String {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}" xmlns:table="{TABLE}" xmlns:draw="{DRAW}"><office:automatic-styles/><office:body><office:text><text:p>outer <draw:frame><draw:text-box><text:h>nested heading</text:h><text:p>nested paragraph</text:p></draw:text-box></draw:frame></text:p><text:h>heading</text:h><text:list><text:list-item><text:p>list paragraph</text:p></text:list-item></text:list><table:table table:name="Table1"><table:table-row><table:table-cell><text:p>cell paragraph</text:p></table:table-cell></table:table-row></table:table><text:tracked-changes><text:p>hidden tracked change</text:p></text:tracked-changes><text:note-body><text:p>hidden note</text:p></text:note-body><text:ruby-text><text:p>hidden ruby</text:p></text:ruby-text><text:p>tail</text:p></office:text></office:body></office:document-content>"#
    )
}

fn aliased_content_xml() -> String {
    format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><o:text><t:p>first</t:p><t:h>second</t:h><t:p>third</t:p></o:text></o:body></o:document-content>"#
    )
}

fn package_with(content: &[u8], extras: bool) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content).unwrap();
    if extras {
        let styles = format!(
            r#"<office:document-styles xmlns:office="{OFFICE}"><office:styles/></office:document-styles>"#
        );
        let meta = format!(
            r#"<office:document-meta xmlns:office="{OFFICE}"><office:meta/></office:document-meta>"#
        );
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
        writer.add_file("meta.xml", meta.as_bytes()).unwrap();
        writer
            .add_file_with_media_type(
                "Pictures/opaque.bin",
                b"opaque media payload",
                "application/octet-stream",
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn raw_package_with(content: &[u8]) -> Vec<u8> {
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
    );
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive.write_deflated("content.xml", content).unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", manifest.as_bytes())
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

fn oversized_declared_content_package() -> (Vec<u8>, (u64, u64)) {
    let mut bytes = raw_package_with(b"<office:document-content/>");
    let content_range = member_range(&bytes, "content.xml");
    let declared_size = (256 * 1024 * 1024 + 1) as u32;
    let mut offset = 0;
    let mut patched = false;
    while let Some(relative) = bytes[offset..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = offset + relative;
        let name_length =
            u16::from_le_bytes(bytes[central + 28..central + 30].try_into().unwrap()) as usize;
        let extra_length =
            u16::from_le_bytes(bytes[central + 30..central + 32].try_into().unwrap()) as usize;
        let comment_length =
            u16::from_le_bytes(bytes[central + 32..central + 34].try_into().unwrap()) as usize;
        let name_start = central + 46;
        let name_end = name_start + name_length;
        if &bytes[name_start..name_end] == b"content.xml" {
            bytes[central + 24..central + 28].copy_from_slice(&declared_size.to_le_bytes());
            patched = true;
            break;
        }
        offset = name_end + extra_length + comment_length;
    }
    assert!(patched, "raw test package has no content.xml central entry");
    (bytes, content_range)
}

fn password_package(content: &[u8], password: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption(password, Profile::compatible())
        .unwrap();
    writer.add_file("content.xml", content).unwrap();
    writer.finish_to_bytes().unwrap()
}

struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: AtomicUsize,
    revision: AtomicU64,
    ranges: Mutex<Vec<(u64, u64)>>,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            ranges: Mutex::new(Vec::new()),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Relaxed)
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
            0x4f44_5403,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

fn member_range(bytes: &[u8], name: &str) -> (u64, u64) {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes).unwrap();
    archive
        .entries()
        .filter_map(|entry| {
            let entry = entry.ok()?;
            if entry.file_path().as_ref() != name.as_bytes() {
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
fn catalog_matches_existing_mixed_nested_text_block_order() {
    let content = content_xml();
    let expected = TextElements::parse(&content).unwrap();
    let bytes = package_with(content.as_bytes(), false);
    let catalog =
        SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();

    let entries = catalog.catalog().unwrap();
    assert_eq!(entries.len(), expected.len());
    assert_eq!(catalog.text_block_count().unwrap(), expected.len());
    for (index, (entry, block)) in entries.iter().zip(expected.iter()).enumerate() {
        assert_eq!(entry.index(), index);
        assert_eq!(entry.kind(), block.kind());
        let selected = catalog.block_at(index).unwrap().unwrap();
        assert_eq!(selected.kind(), block.kind());
        assert_eq!(selected.text().unwrap(), block.text().unwrap());
    }
    assert!(entries.iter().any(|entry| entry.kind() == Kind::Heading));
}

#[test]
fn catalog_resolves_namespace_aliases() {
    let content = aliased_content_xml();
    let expected = TextElements::parse(&content).unwrap();
    let catalog = SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(
        package_with(content.as_bytes(), false),
    )))
    .unwrap();

    assert_eq!(catalog.text_block_count().unwrap(), 3);
    for (index, block) in expected.iter().enumerate() {
        assert_eq!(
            catalog.block_at(index).unwrap().unwrap().text().unwrap(),
            block.text().unwrap()
        );
    }
}

#[test]
fn catalog_matches_top_level_note_and_ruby_block_order() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}"><office:body><office:text><text:note-body><text:p>note body</text:p></text:note-body><text:ruby-text><text:h>ruby text</text:h></text:ruby-text><text:p>tail</text:p></office:text></office:body></office:document-content>"#
    );
    let expected = TextElements::parse(&content).unwrap();
    let catalog = SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(
        package_with(content.as_bytes(), false),
    )))
    .unwrap();

    assert_eq!(catalog.text_block_count().unwrap(), 3);
    assert_eq!(catalog.catalog().unwrap()[0].kind(), Kind::Paragraph);
    assert_eq!(catalog.catalog().unwrap()[1].kind(), Kind::Heading);
    assert_eq!(catalog.catalog().unwrap()[2].kind(), Kind::Paragraph);
    for (index, block) in expected.iter().enumerate() {
        assert_eq!(catalog.catalog().unwrap()[index].kind(), block.kind());
        assert_eq!(
            catalog.block_at(index).unwrap().unwrap().text().unwrap(),
            block.text().unwrap()
        );
    }
}

#[test]
fn catalog_matches_nested_note_and_ruby_suppression() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}"><office:body><office:text><text:p>before<text:note-body><text:p>hidden note</text:p></text:note-body><text:ruby-text><text:h>hidden ruby</text:h></text:ruby-text>after</text:p><text:h>tail heading</text:h></office:text></office:body></office:document-content>"#
    );
    let expected = TextElements::parse(&content).unwrap();
    let catalog = SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(
        package_with(content.as_bytes(), false),
    )))
    .unwrap();

    assert_eq!(catalog.text_block_count().unwrap(), expected.len());
    for (index, block) in expected.iter().enumerate() {
        assert_eq!(catalog.catalog().unwrap()[index].kind(), block.kind());
        let selected = catalog.block_at(index).unwrap().unwrap();
        assert_eq!(selected.text().unwrap(), block.text().unwrap());
    }
}

#[test]
fn selection_rereads_content_and_missing_selection_does_not_read_payload() {
    let source = Arc::new(CountingSource::new(package_with(
        content_xml().as_bytes(),
        false,
    )));
    let catalog = SourceBackedDocumentCatalog::from_read_at(source.clone()).unwrap();
    let before = source.reads();
    assert_eq!(
        catalog.block_at(0).unwrap().unwrap().text().unwrap(),
        "outer "
    );
    assert!(source.reads() > before);

    let before_missing = source.reads();
    assert!(catalog.block_at(usize::MAX).unwrap().is_none());
    assert_eq!(source.reads(), before_missing);
}

#[test]
fn styles_metadata_and_media_remain_cold_until_explicit_member_read() {
    let bytes = package_with(content_xml().as_bytes(), true);
    let styles = member_range(&bytes, "styles.xml");
    let meta = member_range(&bytes, "meta.xml");
    let media = member_range(&bytes, "Pictures/opaque.bin");
    let source = Arc::new(CountingSource::new(bytes));
    let catalog = SourceBackedDocumentCatalog::from_read_at(source.clone()).unwrap();
    let _ = catalog.block_at(0).unwrap();

    let ranges = source.ranges();
    assert!(ranges.iter().all(|range| !overlaps(*range, styles)));
    assert!(ranges.iter().all(|range| !overlaps(*range, meta)));
    assert!(ranges.iter().all(|range| !overlaps(*range, media)));

    assert!(catalog.member_data("styles.xml").unwrap().is_some());
    assert!(source.ranges().iter().any(|range| overlaps(*range, styles)));
    assert!(catalog.member_data("meta.xml").unwrap().is_some());
    assert!(source.ranges().iter().any(|range| overlaps(*range, meta)));
    assert_eq!(
        catalog.media_data("Pictures/opaque.bin").unwrap(),
        Some(b"opaque media payload".to_vec())
    );
    assert!(source.ranges().iter().any(|range| overlaps(*range, media)));
}

#[test]
fn catalog_reports_source_changes() {
    let source = Arc::new(CountingSource::new(package_with(
        content_xml().as_bytes(),
        false,
    )));
    let catalog = SourceBackedDocumentCatalog::from_read_at(source.clone()).unwrap();
    source.bump_revision();
    let reads_after_revision = source.reads();
    assert!(matches!(
        catalog.catalog(),
        Err(Error::SourceChanged { .. })
    ));
    assert_eq!(source.reads(), reads_after_revision);
    assert!(matches!(
        catalog.block_at(0),
        Err(Error::SourceChanged { .. })
    ));
    assert_eq!(source.reads(), reads_after_revision);
}

#[test]
fn catalog_rejects_malformed_document_and_defers_bad_selected_text() {
    let valid = content_xml();
    let malformed = format!("{valid}trailing text");
    let error = SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(
        raw_package_with(malformed.as_bytes()),
    )))
    .unwrap_err();
    assert!(matches!(
        error,
        Error::InvalidFormat(message)
            if message.contains("ODT content.xml") && message.contains("outside its root")
    ));

    let bad_text = valid.replace(
        "<text:h>heading</text:h>",
        &format!(
            r#"<text:h xmlns:alias="{TEXT}" text:style-name="first" alias:style-name="second">heading</text:h>"#
        ),
    );
    let bad_index = TextElements::parse(&valid)
        .unwrap()
        .iter()
        .position(|block| block.text().unwrap() == "heading")
        .unwrap();
    let catalog = SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(
        raw_package_with(bad_text.as_bytes()),
    )))
    .unwrap();
    let error = catalog.block_at(bad_index).unwrap_err();
    assert!(matches!(
        error,
        Error::InvalidFormat(message)
            if message.contains("duplicate ODF text attribute 'text:style-name'")
    ));
}

#[test]
fn catalog_enforces_package_limits_and_materializes_exactly() {
    let content = content_xml();
    let bytes = package_with(content.as_bytes(), false);
    let manifest_limited = ReadLimits::default().with_max_manifest_bytes(1);
    assert!(
        SourceBackedDocumentCatalog::from_read_at_with_limits(
            Arc::new(CountingSource::new(bytes.clone())),
            manifest_limited,
        )
        .is_err()
    );
    let source_limited = ReadLimits::default().with_max_source_bytes(bytes.len() as u64 - 1);
    assert!(
        SourceBackedDocumentCatalog::from_read_at_with_limits(
            Arc::new(CountingSource::new(bytes.clone())),
            source_limited,
        )
        .is_err()
    );

    let catalog =
        SourceBackedDocumentCatalog::from_read_at(Arc::new(CountingSource::new(bytes.clone())))
            .unwrap();
    let document = catalog.materialize().unwrap();
    assert_eq!(document.original_bytes(), bytes.as_slice());
}

#[test]
fn catalog_rejects_oversized_declared_content_before_decompression() {
    let (bytes, content_range) = oversized_declared_content_package();
    let source = Arc::new(CountingSource::new(bytes));
    let error = SourceBackedDocumentCatalog::from_read_at(source.clone()).unwrap_err();
    assert!(matches!(
        error,
        Error::InvalidFormat(message)
            if message == "ODT content.xml exceeds the family limit"
    ));
    assert!(
        source
            .ranges()
            .iter()
            .all(|range| !overlaps(*range, content_range))
    );
}

#[test]
fn catalog_preserves_password_for_selection_and_materialization() {
    let password = "catalog-password";
    let bytes = password_package(content_xml().as_bytes(), password);
    let catalog = SourceBackedDocumentCatalog::from_read_at_with_password(
        Arc::new(CountingSource::new(bytes.clone())),
        password,
    )
    .unwrap();
    assert_eq!(
        catalog.block_at(0).unwrap().unwrap().kind(),
        Kind::Paragraph
    );
    let document = catalog.materialize().unwrap();
    assert_eq!(document.original_bytes(), bytes.as_slice());
}
