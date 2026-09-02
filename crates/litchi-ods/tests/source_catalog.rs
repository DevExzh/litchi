use std::io::{Cursor, Read, Write};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};

use litchi_core::{Error, OwnedSource, ReadAt, SourceVersion};
use litchi_odf_common::core::{PackageWriter, Profile};
use litchi_odf_common::signature::{DocumentSigner, SignatureAlgorithm};
use litchi_ods::{
    ReadLimits, SourceBackedSpreadsheet, SourceBackedSpreadsheetCatalog, Spreadsheet,
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

mod support;

const MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DDE_CONTENT: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Base"><office:dde-source office:dde-application="soffice" office:dde-topic="file:///never-contacted.ods" office:dde-item="Sheet1.A1:B2" office:name="Reference" office:conversion-mode="keep-text" office:automatic-update="true"/><table:table-row><table:table-cell office:value-type="string"><text:p>Pre <text:span text:style-name="Bold">styled <text:a xlink:href="https://never-fetched.invalid/" xlink:type="simple">link</text:a></text:span> tail</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Scenario"><table:scenario table:scenario-ranges="$Scenario.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4" table:is-active="true" table:display-border="false" table:border-color="#12AbEF" table:copy-back="true" table:copy-styles="false" table:copy-formulas="true" table:comment="Best &amp; worst" table:protected="false"/><table:table-row/></table:table><table:dde-links><table:dde-link><office:dde-source office:dde-application="calc" office:dde-topic="file:///never-opened.ods" office:dde-item="Prices.A1"/><table:table><table:table-row><table:table-cell office:value-type="float" office:value="42"/></table:table-row></table:table></table:dde-link></table:dde-links></office:spreadsheet></office:body></office:document-content>"##;
const RSA_KEY: &[u8] =
    include_bytes!("../../litchi-odf-common/tests/fixtures/signatures/rsa-key.pk8");
const RSA_CERT: &[u8] =
    include_bytes!("../../litchi-odf-common/tests/fixtures/signatures/rsa-cert.der");

fn package() -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="First"><table:table-row><table:table-cell office:value-type="string"><text:p>first</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Selected"><table:table-row><table:table-cell office:value-type="string"><text:p>middle</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="42"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).expect("ODS MIME");
    writer
        .set_document_signer(
            DocumentSigner::from_pkcs8_der(
                SignatureAlgorithm::RsaSha256,
                RSA_KEY,
                vec![RSA_CERT.to_vec()],
                "2026-08-21T12:00:00Z",
            )
            .expect("document signer"),
        )
        .expect("configure document signer");
    writer
        .add_file("content.xml", content.as_bytes())
        .expect("content.xml");
    // These members deliberately contain bytes that the catalog lifecycle
    // never decodes. Their presence proves that catalog open is not the
    // existing all-members semantic owner.
    writer
        .add_file("styles.xml", b"<styles/>")
        .expect("styles member");
    let mut media = Vec::with_capacity(512 * 1024);
    media.extend((0..(512 * 1024)).map(|index| (index as u8).wrapping_mul(31)));
    writer
        .add_file_with_media_type("Pictures/opaque.bin", &media, "application/octet-stream")
        .expect("opaque media");
    writer.finish_to_bytes().expect("package bytes")
}

fn aliased_package() -> Vec<u8> {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}"><o:body><o:spreadsheet><t:table t:name="Alias"><t:table-row><t:table-cell o:value-type="string"><x:p>alias</x:p></t:table-cell></t:table-row></t:table><t:table t:name="Empty"/></o:spreadsheet></o:body></o:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).expect("ODS MIME");
    writer
        .add_file("content.xml", content.as_bytes())
        .expect("content.xml");
    writer.finish_to_bytes().expect("aliased package bytes")
}

fn package_with_content(content: &str) -> Vec<u8> {
    support::raw_package(&[("content.xml", content.as_bytes(), "text/xml")])
}

fn oversized_declared_content_package() -> (Vec<u8>, (u64, u64)) {
    let mut bytes = raw_deflated_package(b"<office:document-content/>");
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

fn raw_deflated_package(content: &[u8]) -> Vec<u8> {
    let manifest = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
    );
    let mut output = Cursor::new(Vec::new());
    {
        let mut archive = ZipWriter::new(&mut output);
        let stored =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        let deflated =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        archive
            .start_file("mimetype", stored)
            .expect("raw mimetype");
        archive
            .write_all(MIME.as_bytes())
            .expect("raw mimetype bytes");
        archive
            .start_file("content.xml", deflated)
            .expect("raw content");
        archive.write_all(content).expect("raw content bytes");
        archive
            .start_file("META-INF/manifest.xml", deflated)
            .expect("raw manifest");
        archive
            .write_all(manifest.as_bytes())
            .expect("raw manifest bytes");
        archive.finish().expect("raw package finish");
    }
    output.into_inner()
}

fn oversized_encrypted_content_package() -> (Vec<u8>, (u64, u64)) {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet/></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).expect("ODS MIME");
    writer
        .set_encryption("catalog-password", Profile::compatible())
        .expect("encryption profile");
    writer
        .add_file("content.xml", content.as_bytes())
        .expect("encrypted content");
    let bytes = rewrite_manifest_size(
        &writer.finish_to_bytes().expect("encrypted package"),
        u64::from((256 * 1024 * 1024 + 1) as u32),
    );
    let content_range = member_range(&bytes, "content.xml");
    (bytes, content_range)
}

fn rewrite_manifest_size(bytes: &[u8], size: u64) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("encrypted ZIP");
    let mut output = Cursor::new(Vec::new());
    {
        let mut writer = ZipWriter::new(&mut output);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        for index in 0..archive.len() {
            let (name, mut data) = {
                let mut entry = archive.by_index(index).expect("ZIP entry");
                let name = entry.name().to_owned();
                let mut data = Vec::new();
                entry.read_to_end(&mut data).expect("ZIP entry bytes");
                (name, data)
            };
            if name == "META-INF/manifest.xml" {
                let manifest = String::from_utf8(data).expect("manifest UTF-8");
                let marker = r#"manifest:size=""#;
                let value_start =
                    manifest.find(marker).expect("encrypted size metadata") + marker.len();
                let value_end = value_start
                    + manifest[value_start..]
                        .find('"')
                        .expect("encrypted size terminator");
                let mut rewritten = String::new();
                rewritten.push_str(&manifest[..value_start]);
                rewritten.push_str(&size.to_string());
                rewritten.push_str(&manifest[value_end..]);
                data = rewritten.into_bytes();
            }
            writer.start_file(name, options).expect("ZIP entry");
            writer.write_all(&data).expect("ZIP entry bytes");
        }
        writer.finish().expect("rewritten ZIP");
    }
    output.into_inner()
}

fn member_range(bytes: &[u8], name: &str) -> (u64, u64) {
    let mut offset = 0;
    while let Some(relative) = bytes[offset..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = offset + relative;
        let compressed_size = u64::from(u32::from_le_bytes(
            bytes[central + 20..central + 24].try_into().unwrap(),
        ));
        let name_length =
            u16::from_le_bytes(bytes[central + 28..central + 30].try_into().unwrap()) as usize;
        let extra_length =
            u16::from_le_bytes(bytes[central + 30..central + 32].try_into().unwrap()) as usize;
        let comment_length =
            u16::from_le_bytes(bytes[central + 32..central + 34].try_into().unwrap()) as usize;
        let name_start = central + 46;
        let name_end = name_start + name_length;
        if &bytes[name_start..name_end] == name.as_bytes() {
            let local = usize::try_from(u32::from_le_bytes(
                bytes[central + 42..central + 46].try_into().unwrap(),
            ))
            .unwrap();
            let local_name_length =
                u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
            let local_extra_length =
                u16::from_le_bytes(bytes[local + 28..local + 30].try_into().unwrap()) as usize;
            let start = u64::try_from(local + 30 + local_name_length + local_extra_length).unwrap();
            return (start, start + compressed_size);
        }
        offset = name_end + extra_length + comment_length;
    }
    panic!("ZIP member {name} not found")
}

fn overlaps(left: (u64, u64), right: (u64, u64)) -> bool {
    left.0 < right.1 && right.0 < left.1
}

struct ProbeSource {
    source: OwnedSource,
    revision: AtomicU64,
    ranges: Mutex<Vec<(u64, u64)>>,
}

impl ProbeSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            source: OwnedSource::new(bytes),
            revision: AtomicU64::new(0),
            ranges: Mutex::new(Vec::new()),
        })
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }

    fn ranges(&self) -> Vec<(u64, u64)> {
        self.ranges.lock().expect("range lock").clone()
    }
}

impl ReadAt for ProbeSource {
    fn len(&self) -> std::io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let read = self.source.read_at(offset, output)?;
        if read != 0 {
            self.ranges
                .lock()
                .map_err(|_| std::io::Error::other("range lock poisoned"))?
                .push((
                    offset,
                    offset + u64::try_from(read).expect("read length fits u64"),
                ));
        }
        Ok(read)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5301,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn catalog_open_and_selected_sheet_are_lazy_and_semantically_equal() {
    let bytes = package();
    let source = ProbeSource::new(bytes.clone());
    let catalog =
        SourceBackedSpreadsheetCatalog::from_read_at(source.clone()).expect("catalog open");

    assert_eq!(catalog.sheet_names().expect("names"), ["First", "Selected"]);
    assert_eq!(catalog.sheet_count().expect("count"), 2);
    assert_eq!(catalog.catalog().expect("catalog")[1].name(), "Selected");
    assert_eq!(catalog.catalog().expect("catalog")[1].index(), 1);

    let opened = catalog.source_read_metrics().expect("open metrics");
    assert!(opened.read_calls > 0);
    assert!(opened.read_bytes < bytes.len() as u64 / 2);

    catalog.reset_source_read_metrics().expect("reset metrics");
    assert!(catalog.sheet_at(99).expect("missing sheet").is_none());
    assert_eq!(
        catalog.source_read_metrics().expect("missing metrics"),
        Default::default()
    );

    let selected = catalog
        .sheet("Selected")
        .expect("selected sheet read")
        .expect("selected sheet exists");
    assert_eq!(selected.name, "Selected");
    assert_eq!(selected.rows[0].cells[0].text, "middle");
    assert!(
        catalog
            .source_read_metrics()
            .expect("query metrics")
            .read_bytes
            > 0
    );

    let eager = Spreadsheet::from_bytes(bytes).expect("eager owner");
    let expected = eager.sheets().get(1).expect("eager selected sheet").clone();
    assert_eq!(selected, expected);
}

#[test]
fn catalog_keeps_opaque_members_and_signatures_deferred_until_selected() {
    let bytes = package();
    let catalog =
        SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
            .expect("catalog open with opaque members");

    let before = catalog.source_read_metrics().expect("open metrics");
    assert!(
        catalog
            .media_files()
            .expect("media list")
            .iter()
            .any(|path| path == "Pictures/opaque.bin")
    );
    assert_eq!(catalog.source_read_metrics().expect("list metrics"), before);

    let signatures = catalog.digital_signatures().expect("signature read");
    assert_eq!(signatures.document_signatures.len(), 1);
    let media = catalog
        .media_data("Pictures/opaque.bin")
        .expect("media read")
        .expect("media member");
    assert_eq!(media.len(), 512 * 1024);

    let materialized = catalog.materialize().expect("explicit materialization");
    let snapshot = materialized.document_snapshot().expect("snapshot");
    assert_eq!(snapshot.as_bytes(), bytes.as_slice());
}

#[test]
fn catalog_preserves_limits_and_source_version_lifecycle() {
    let bytes = package();
    let limited = ReadLimits::default().with_max_manifest_bytes(1);
    let error = SourceBackedSpreadsheetCatalog::from_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes.clone())),
        limited,
    )
    .expect_err("manifest limit must be enforced during catalog open");
    assert!(matches!(error, Error::ResourceLimit(_)));

    let source = ProbeSource::new(bytes);
    let catalog =
        SourceBackedSpreadsheetCatalog::from_read_at(source.clone()).expect("catalog open");
    source.bump();
    assert!(matches!(
        catalog.sheet_count(),
        Err(Error::SourceChanged { .. })
    ));
    assert!(matches!(
        catalog.sheet_at(0),
        Err(Error::SourceChanged { .. })
    ));
}

#[test]
fn source_owners_reject_oversized_content_before_payload_read() {
    let (bytes, content_range) = oversized_declared_content_package();
    let source = ProbeSource::new(bytes);
    let error = SourceBackedSpreadsheet::from_read_at(source.clone())
        .expect_err("full source owner must reject oversized content");
    assert!(matches!(
        error,
        Error::InvalidFormat(message) if message == "ODS content.xml exceeds the family limit"
    ));
    assert!(
        source
            .ranges()
            .iter()
            .all(|range| !overlaps(*range, content_range)),
        "full source open must not read oversized content.xml"
    );

    let (bytes, content_range) = oversized_declared_content_package();
    let source = ProbeSource::new(bytes);
    let error = SourceBackedSpreadsheetCatalog::from_read_at(source.clone())
        .expect_err("catalog source owner must reject oversized content");
    assert!(matches!(
        error,
        Error::InvalidFormat(message) if message == "ODS content.xml exceeds the family limit"
    ));
    assert!(
        source
            .ranges()
            .iter()
            .all(|range| !overlaps(*range, content_range)),
        "catalog open must not read oversized content.xml"
    );

    let (bytes, content_range) = oversized_encrypted_content_package();
    let source = ProbeSource::new(bytes);
    let error =
        SourceBackedSpreadsheet::from_read_at_with_password(source.clone(), "catalog-password")
            .expect_err("full source owner must reject oversized encrypted content");
    assert!(matches!(
        error,
        Error::InvalidFormat(message) if message == "ODS content.xml exceeds the family limit"
    ));
    assert!(
        source
            .ranges()
            .iter()
            .all(|range| !overlaps(*range, content_range)),
        "full source open must not read oversized encrypted content.xml"
    );

    let (bytes, content_range) = oversized_encrypted_content_package();
    let source = ProbeSource::new(bytes);
    let error = SourceBackedSpreadsheetCatalog::from_read_at_with_password(
        source.clone(),
        "catalog-password",
    )
    .expect_err("catalog source owner must reject oversized encrypted content");
    assert!(matches!(
        error,
        Error::InvalidFormat(message) if message == "ODS content.xml exceeds the family limit"
    ));
    assert!(
        source
            .ranges()
            .iter()
            .all(|range| !overlaps(*range, content_range)),
        "catalog open must not read oversized encrypted content.xml"
    );
}

#[test]
fn catalog_handles_namespace_aliases_and_empty_selected_worksheets() {
    let catalog =
        SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(aliased_package())))
            .expect("aliased catalog open");
    assert_eq!(
        catalog.sheet_names().expect("aliased names"),
        ["Alias", "Empty"]
    );
    let selected = catalog
        .sheet_at(1)
        .expect("empty selected worksheet")
        .expect("empty worksheet exists");
    assert_eq!(selected.name, "Empty");
    assert!(selected.rows.is_empty());
}

#[test]
fn catalog_excludes_inert_dde_cached_tables_and_matches_eager_sheets() {
    let bytes = package_with_content(DDE_CONTENT);
    let eager = Spreadsheet::from_bytes(bytes.clone()).expect("eager DDE owner");
    let eager_sheets = eager.sheets();
    let catalog = SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("DDE catalog open");

    assert_eq!(
        catalog.sheet_names().expect("DDE sheet names"),
        ["Base", "Scenario"]
    );
    assert_eq!(
        catalog.sheet_count().expect("DDE sheet count"),
        eager_sheets.len()
    );
    for (index, expected) in eager_sheets.iter().enumerate() {
        let selected = catalog
            .sheet_at(index)
            .expect("DDE selected sheet")
            .expect("DDE selected sheet exists");
        assert_eq!(selected, expected.clone());
    }
    assert!(
        catalog
            .sheet_at(eager_sheets.len())
            .expect("DDE cache is not a sheet")
            .is_none()
    );
}

#[test]
fn catalog_excludes_self_closing_dde_cache_before_real_sheet() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:dde-links><table:dde-link><table:table/></table:dde-link></table:dde-links><table:table table:name="Real"><table:table-row/></table:table></office:spreadsheet></office:body></office:document-content>"#
    );
    let bytes = package_with_content(&content);
    let eager = Spreadsheet::from_bytes(bytes.clone()).expect("eager self-closing DDE owner");
    assert_eq!(
        eager
            .sheets()
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Real"]
    );

    let catalog = SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("self-closing DDE catalog open");
    assert_eq!(
        catalog.sheet_names().expect("self-closing DDE names"),
        ["Real"]
    );
    assert_eq!(catalog.sheet_count().expect("self-closing DDE count"), 1);
    assert_eq!(
        catalog
            .sheet_at(0)
            .expect("self-closing DDE selected sheet")
            .expect("real worksheet exists"),
        eager.sheets()[0].clone()
    );
}

#[test]
fn catalog_preserves_validation_precedence_for_malformed_xml() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:table table:name="Broken"></office:spreadsheet>"#
    );
    let error = SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(
        package_with_content(&content),
    )))
    .expect_err("malformed content must be refused");
    assert!(
        matches!(error, Error::InvalidFormat(message) if message.contains("invalid ODS content.xml"))
    );
}

#[test]
fn catalog_rejects_duplicate_and_nested_worksheet_entries() {
    let duplicate = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:table table:name="Same"/><table:table table:name="Same"/></office:spreadsheet></office:body></office:document-content>"#
    );
    let duplicate_error = SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(
        package_with_content(&duplicate),
    )))
    .expect_err("duplicate sheet names must be refused");
    assert!(
        matches!(duplicate_error, Error::InvalidFormat(message) if message.contains("duplicated"))
    );

    let nested = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:table table:name="Outer"><table:table table:name="Inner"/></table:table></office:spreadsheet></office:body></office:document-content>"#
    );
    let nested_error = SourceBackedSpreadsheetCatalog::from_read_at(Arc::new(OwnedSource::new(
        package_with_content(&nested),
    )))
    .expect_err("nested worksheet entries must be refused");
    assert!(
        matches!(nested_error, Error::InvalidFormat(message) if message.contains("direct child"))
    );
}
