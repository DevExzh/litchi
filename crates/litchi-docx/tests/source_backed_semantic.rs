use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_docx::{Block, Element, Error, Package, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, SourceBackedPackage,
};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_CORE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
const STRICT_CORE_NAMESPACE: &str = "http://purl.oclc.org/ooxml/package/metadata/core-properties";

fn package_bytes(
    document: impl Into<Vec<u8>>,
    core: Option<(&str, &str)>,
    external_core: bool,
) -> Vec<u8> {
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document.into(),
    );
    main.rels_mut().add_relationship(
        rt::MS_ALTERNATIVE_FORMAT_IMPORT.to_owned(),
        "chunk.html".to_owned(),
        "chunk".to_owned(),
        false,
    );

    let mut package = OpcPackage::new();
    package.add_part(Box::new(main));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/chunk.html").unwrap(),
        "text/html".to_owned(),
        b"<p>chunk</p>".to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/unrelated.bin").unwrap(),
        "application/octet-stream".to_owned(),
        vec![0xA5; 32 * 1024],
    )));

    if let Some((xml, relationship)) = core {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_owned(),
            xml.as_bytes().to_vec(),
        )));
        if external_core {
            package.relate_to_external("https://example.invalid/core.xml", relationship);
        } else {
            package.relate_to("docProps/core.xml", relationship);
        }
    }
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn semantic_fixture() -> Vec<u8> {
    package_bytes(
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="{MC}" xmlns:x="urn:litchi:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>inactive</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><x:opaque value="retained"/><w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl><w:altChunk r:id="chunk"/><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:body></w:document>"#
        ),
        None,
        false,
    )
}

fn raw_document_fixture(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{}"/></Types>"#,
        ct::WML_DOCUMENT_MAIN
    );
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{}" Target="word/document.xml"/></Relationships>"#,
        rt::OFFICE_DOCUMENT
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

fn managed_raw_document(document: &[u8]) -> source_backed::Package {
    let bytes = raw_document_fixture(document);
    let memory = 16 * 1024 * 1024;
    let budget = litchi_core::Budget::root(
        "docx-managed-semantic-preflight-test",
        litchi_core::Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = litchi_core::CancellationSource::pair();
    let execution_limits = litchi_core::ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        litchi_docx::ReadLimits::default(),
        litchi_core::ExecutionContext::new(budget, cancellation, execution_limits),
    )
    .unwrap()
}

fn table_signature(table: &litchi_docx::Table) -> litchi_docx::Result<String> {
    let mut cells = Vec::new();
    for row in table.rows()? {
        for cell in row.cells()? {
            cells.push(cell.text()?);
        }
    }
    Ok(cells.join("|"))
}

fn eager_block_signature(
    document: &litchi_docx::document::Document<'_>,
) -> litchi_docx::Result<Vec<String>> {
    document
        .blocks()?
        .into_iter()
        .map(|block| match block {
            Block::Paragraph(paragraph) => Ok(format!("p:{}", paragraph.text()?)),
            Block::Table(table) => Ok(format!("t:{}", table_signature(&table)?)),
            Block::Alt(chunk) => Ok(format!("a:{}", chunk.relationship().as_str())),
            Block::Unknown(block) => {
                Ok(format!("u:{}", String::from_utf8_lossy(block.xml_bytes())))
            },
        })
        .collect()
}

fn source_block_signature(document: &source_backed::Document) -> litchi_docx::Result<Vec<String>> {
    document
        .blocks()?
        .into_iter()
        .map(|block| match block {
            Block::Paragraph(paragraph) => Ok(format!("p:{}", paragraph.text()?)),
            Block::Table(table) => Ok(format!("t:{}", table_signature(&table)?)),
            Block::Alt(chunk) => Ok(format!("a:{}", chunk.relationship().as_str())),
            Block::Unknown(block) => {
                Ok(format!("u:{}", String::from_utf8_lossy(block.xml_bytes())))
            },
        })
        .collect()
}

fn eager_element_signature(
    document: &litchi_docx::document::Document<'_>,
) -> litchi_docx::Result<Vec<String>> {
    document
        .elements()?
        .into_iter()
        .map(|element| match element {
            Element::Paragraph(paragraph) => Ok(format!("p:{}", paragraph.text()?)),
            Element::Table(table) => Ok(format!("t:{}", table_signature(&table)?)),
            Element::Unknown(block) => {
                Ok(format!("u:{}", String::from_utf8_lossy(block.xml_bytes())))
            },
        })
        .collect()
}

fn source_element_signature(
    document: &source_backed::Document,
) -> litchi_docx::Result<Vec<String>> {
    document
        .elements()?
        .into_iter()
        .map(|element| match element {
            Element::Paragraph(paragraph) => Ok(format!("p:{}", paragraph.text()?)),
            Element::Table(table) => Ok(format!("t:{}", table_signature(&table)?)),
            Element::Unknown(block) => {
                Ok(format!("u:{}", String::from_utf8_lossy(block.xml_bytes())))
            },
        })
        .collect()
}

#[test]
fn source_semantic_queries_match_eager_mce_order_and_unknown_retention() {
    let bytes = semantic_fixture();
    let eager = Package::from_reader(io::Cursor::new(bytes.clone())).unwrap();
    let eager_document = eager.document().unwrap();
    let source = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let source_document = source.document().unwrap();

    assert_eq!(
        source_document.extract_text().unwrap(),
        eager_document.text().unwrap()
    );
    assert_eq!(
        source_document.paragraph_count().unwrap(),
        eager_document.paragraph_count().unwrap()
    );
    assert_eq!(
        source_document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>(),
        eager_document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>()
    );
    let eager_tables = eager_document
        .tables()
        .unwrap()
        .iter()
        .map(table_signature)
        .collect::<litchi_docx::Result<Vec<_>>>()
        .unwrap();
    let source_tables = source_document
        .tables()
        .unwrap()
        .iter()
        .map(table_signature)
        .collect::<litchi_docx::Result<Vec<_>>>()
        .unwrap();
    assert_eq!(source_tables, eager_tables);
    assert_eq!(source_tables, ["outernested"]);

    assert_eq!(
        source_block_signature(&source_document).unwrap(),
        eager_block_signature(&eager_document).unwrap()
    );
    let blocks = source_block_signature(&source_document).unwrap();
    assert_eq!(blocks[0], "p:fallback");
    assert!(blocks[1].starts_with("u:") && blocks[1].contains("retained"));
    assert_eq!(blocks[2], "t:outernested");
    assert_eq!(blocks[3], "a:chunk");
    assert_eq!(blocks[4], "p:tail");
    assert_eq!(
        source_element_signature(&source_document).unwrap(),
        eager_element_signature(&eager_document).unwrap()
    );
    let elements = source_element_signature(&source_document).unwrap();
    assert_eq!(elements[0], "p:fallback");
    assert!(elements[1].starts_with("u:") && elements[1].contains("retained"));
    assert_eq!(elements[2], "t:outernested");
    assert_eq!(elements[3], "p:tail");
}

fn transitional_core(title: &str) -> String {
    format!(
        r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title><dc:creator>Ada</dc:creator><cp:revision>7</cp:revision></cp:coreProperties>"#
    )
}

fn strict_core(title: &str) -> String {
    format!(
        r#"<cp:coreProperties xmlns:cp="{STRICT_CORE_NAMESPACE}" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title><dc:creator>Ada</dc:creator></cp:coreProperties>"#
    )
}

#[test]
fn source_metadata_matches_dialects_absence_and_graph_failures() {
    let document = format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#);
    let transitional = package_bytes(
        document.clone(),
        Some((
            transitional_core("Transitional").as_str(),
            rt::CORE_PROPERTIES,
        )),
        false,
    );
    let eager = Package::from_reader(io::Cursor::new(transitional.clone())).unwrap();
    let eager_props = eager.props().unwrap();
    let source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(transitional))).unwrap();
    let metadata = source.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), eager_props.title.as_deref());
    assert_eq!(metadata.author.as_deref(), eager_props.creator.as_deref());
    assert_eq!(
        metadata.revision.as_deref(),
        eager_props.revision.as_deref()
    );

    let strict = package_bytes(
        document.clone(),
        Some((strict_core("Strict").as_str(), STRICT_CORE_RELATIONSHIP)),
        false,
    );
    let strict_source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(strict))).unwrap();
    assert_eq!(
        strict_source.metadata().unwrap().title.as_deref(),
        Some("Strict")
    );

    let absent = package_bytes(document.clone(), None, false);
    let absent_source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(absent))).unwrap();
    assert!(!absent_source.metadata().unwrap().has_data());

    let wrong_dialect = package_bytes(
        document.clone(),
        Some((strict_core("wrong").as_str(), rt::CORE_PROPERTIES)),
        false,
    );
    let wrong_source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(wrong_dialect))).unwrap();
    assert!(matches!(wrong_source.metadata(), Err(Error::Common(_))));

    let external = package_bytes(
        document,
        Some((transitional_core("external").as_str(), rt::CORE_PROPERTIES)),
        true,
    );
    let external_source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(external))).unwrap();
    assert!(matches!(external_source.metadata(), Err(Error::Common(_))));
}

struct ProbeSource {
    bytes: Vec<u8>,
    main: std::ops::Range<usize>,
    core: Option<std::ops::Range<usize>>,
    unrelated: std::ops::Range<usize>,
    main_reads: AtomicUsize,
    core_reads: AtomicUsize,
    unrelated_reads: AtomicUsize,
    revision: AtomicU64,
    change_on_main: AtomicBool,
    change_on_core: AtomicBool,
}

impl ProbeSource {
    fn new(bytes: Vec<u8>, core: Option<std::ops::Range<usize>>) -> Self {
        let main = payload_range(&bytes, "word/document.xml");
        let unrelated = payload_range(&bytes, "word/unrelated.bin");
        Self {
            bytes,
            main,
            core,
            unrelated,
            main_reads: AtomicUsize::new(0),
            core_reads: AtomicUsize::new(0),
            unrelated_reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            change_on_main: AtomicBool::new(false),
            change_on_core: AtomicBool::new(false),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn arm_main_change(&self) {
        self.change_on_main.store(true, Ordering::SeqCst);
    }

    fn arm_core_change(&self) {
        self.change_on_core.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for ProbeSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let range = offset..end;
        if overlaps(&range, &self.main) {
            self.main_reads.fetch_add(1, Ordering::SeqCst);
            if self.change_on_main.swap(false, Ordering::SeqCst) {
                self.changed();
            }
        }
        if let Some(core) = &self.core {
            if overlaps(&range, core) {
                self.core_reads.fetch_add(1, Ordering::SeqCst);
                if self.change_on_core.swap(false, Ordering::SeqCst) {
                    self.changed();
                }
            }
        }
        if overlaps(&range, &self.unrelated) {
            self.unrelated_reads.fetch_add(1, Ordering::SeqCst);
        }
        output[..end - offset].copy_from_slice(&self.bytes[range]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(91, self.revision.load(Ordering::SeqCst)))
    }
}

fn overlaps(left: &std::ops::Range<usize>, right: &std::ops::Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

fn payload_range(zip: &[u8], name: &str) -> std::ops::Range<usize> {
    for (offset, _) in zip
        .windows(4)
        .enumerate()
        .filter(|(_, signature)| *signature == b"PK\x01\x02")
    {
        if offset + 46 > zip.len() {
            continue;
        }
        let compressed =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_length =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        let extra_length =
            u16::from_le_bytes(zip[offset + 30..offset + 32].try_into().unwrap()) as usize;
        if offset + 46 + name_length + extra_length > zip.len()
            || &zip[offset + 46..offset + 46 + name_length] != name.as_bytes()
        {
            continue;
        }
        let local = u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
        let local_name =
            u16::from_le_bytes(zip[local + 26..local + 28].try_into().unwrap()) as usize;
        let local_extra =
            u16::from_le_bytes(zip[local + 28..local + 30].try_into().unwrap()) as usize;
        let start = local + 30 + local_name + local_extra;
        return start..start + compressed;
    }
    panic!("ZIP payload not found: {name}")
}

#[test]
fn source_queries_read_only_selected_payloads_and_preserve_stale_precedence() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>main</w:t></w:r></w:p></w:body></w:document>"#
    );
    let core_xml = transitional_core("selected");
    let bytes = package_bytes(document, Some((&core_xml, rt::CORE_PROPERTIES)), false);
    let core_range = payload_range(&bytes, "docProps/core.xml");
    let source = Arc::new(ProbeSource::new(bytes, Some(core_range)));
    let package = source_backed::Package::from_read_at(source.clone()).unwrap();
    assert_eq!(source.main_reads.load(Ordering::SeqCst), 0);
    assert_eq!(source.core_reads.load(Ordering::SeqCst), 0);
    assert_eq!(source.unrelated_reads.load(Ordering::SeqCst), 0);

    assert_eq!(
        package.metadata().unwrap().title.as_deref(),
        Some("selected")
    );
    assert!(source.core_reads.load(Ordering::SeqCst) > 0);
    assert_eq!(source.main_reads.load(Ordering::SeqCst), 0);
    assert_eq!(source.unrelated_reads.load(Ordering::SeqCst), 0);
    let document = package.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), "main");
    assert!(source.main_reads.load(Ordering::SeqCst) > 0);
    assert_eq!(source.unrelated_reads.load(Ordering::SeqCst), 0);

    // A source revision is checked before graph inspection and payload decode.
    let stale_bytes = package_bytes(
        format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#),
        Some((&core_xml, rt::CORE_PROPERTIES)),
        false,
    );
    let stale_core_range = payload_range(&stale_bytes, "docProps/core.xml");
    let stale_source = Arc::new(ProbeSource::new(stale_bytes, Some(stale_core_range)));
    let stale_package = source_backed::Package::from_read_at(stale_source.clone()).unwrap();
    stale_source.changed();
    let stale_result = stale_package.metadata();
    assert!(
        matches!(
            stale_result,
            Err(Error::Opc(OpcError::SourceChanged { .. }))
        ),
        "{stale_result:?}"
    );

    let during_core_bytes = package_bytes(
        format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#),
        Some((&core_xml, rt::CORE_PROPERTIES)),
        false,
    );
    let during_core_range = payload_range(&during_core_bytes, "docProps/core.xml");
    let during_core_source = Arc::new(ProbeSource::new(during_core_bytes, Some(during_core_range)));
    let during_core_package =
        source_backed::Package::from_read_at(during_core_source.clone()).unwrap();
    during_core_source.arm_core_change();
    assert!(matches!(
        during_core_package.metadata(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));

    let during_main_bytes = package_bytes(
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#),
        None,
        false,
    );
    let during_main_source = Arc::new(ProbeSource::new(during_main_bytes, None));
    let during_main_package =
        source_backed::Package::from_read_at(during_main_source.clone()).unwrap();
    during_main_source.arm_main_change();
    assert!(matches!(
        during_main_package.document(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));

    let during_snapshot_bytes = package_bytes(
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#),
        None,
        false,
    );
    let during_snapshot_source = Arc::new(ProbeSource::new(during_snapshot_bytes, None));
    let during_snapshot_package =
        source_backed::Package::from_read_at(during_snapshot_source.clone()).unwrap();
    during_snapshot_source.arm_main_change();
    assert!(matches!(
        during_snapshot_package.document_snapshot(),
        Err(litchi_docx::document::TransactionError::Document(
            Error::Opc(OpcError::SourceChanged { .. })
        ))
    ));
}

#[test]
fn source_package_materializes_through_the_owned_docx_facade_on_explicit_request() {
    let bytes = semantic_fixture();
    let source = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let owned = source.to_owned_package().unwrap();
    assert_eq!(
        owned.document().unwrap().text().unwrap(),
        "fallbackouternestedtail"
    );

    let adopted_opc =
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(semantic_fixture()))).unwrap();
    let adopted = source_backed::Package::from_source_backed_package(adopted_opc).unwrap();
    assert_eq!(
        adopted.document().unwrap().extract_text().unwrap(),
        "fallbackouternestedtail"
    );
}

fn assert_send_sync<T: Send + Sync>() {}

#[test]
fn source_document_is_send_sync_and_managed_arc_views_refuse_consistently() {
    assert_send_sync::<source_backed::Package>();
    assert_send_sync::<source_backed::Document>();

    let xml = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>managed</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#
    );
    let bytes = package_bytes(xml, None, false);
    let budget = litchi_core::Budget::root(
        "docx-semantic-managed-test",
        litchi_core::Limits::new(
            bytes.len() as u64 * 4,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = litchi_core::CancellationSource::pair();
    let limits = litchi_core::ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(bytes.len() as u64 * 4).unwrap(),
        0,
    )
    .unwrap();
    let context = litchi_core::ExecutionContext::new(budget, cancellation, limits);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        litchi_docx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), "managedcell");
    assert_eq!(
        document.paragraph_text(0).unwrap().as_deref(),
        Some("managed")
    );
    assert!(matches!(
        package.to_owned_package(),
        Err(Error::Opc(OpcError::ManagedPackageMaterialization))
    ));
    for error in [
        document.paragraphs().unwrap_err(),
        document.tables().unwrap_err(),
        document.blocks().unwrap_err(),
        document.elements().unwrap_err(),
    ] {
        assert!(matches!(error, Error::UnsafeEdit { .. }));
    }
    cancellation_source.cancel();
    assert!(matches!(
        document.extract_text(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
}

#[test]
fn managed_source_preflight_rejects_special_events_and_extra_roots() {
    let cases = [
        (
            format!(r#"<w:document xmlns:w="{W}"><w:body><w:p></w:document>"#),
            "mismatched end tag",
        ),
        (
            format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>trailing"#),
            "trailing text",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body/></w:document>{}"#,
                '\u{000b}'
            ),
            "vertical-tab trailing text",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body/></w:document>{}"#,
                '\u{000c}'
            ),
            "form-feed trailing text",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body/></w:document><w:document xmlns:w="{W}"/>"#
            ),
            "second root",
        ),
        (
            format!(
                r#"<!DOCTYPE w:document [<!ELEMENT document ANY>]><w:document xmlns:w="{W}"><w:body/></w:document>"#
            ),
            "DTD",
        ),
        (
            format!(r#"<?custom instruction?><w:document xmlns:w="{W}"><w:body/></w:document>"#),
            "processing instruction",
        ),
        (
            format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document><?xml version="1.0"?>"#),
            "late XML declaration",
        ),
        (
            format!(
                r#"<?xml version="1.0"?><?xml version="1.0"?><w:document xmlns:w="{W}"><w:body/></w:document>"#
            ),
            "duplicate XML declaration",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&custom;</w:t></w:r></w:p></w:body></w:document>"#
            ),
            "custom entity",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p x:value="&custom;" xmlns:x="urn:test"/></w:body></w:document>"#
            ),
            "custom attribute entity",
        ),
    ];
    for (xml, label) in cases {
        let package = managed_raw_document(xml.as_bytes());
        let error = match package.document() {
            Ok(_) => panic!("{label}: managed preflight unexpectedly succeeded"),
            Err(error) => error,
        };
        assert!(
            matches!(error, Error::UnsafeEdit { .. } | Error::InvalidFormat(_)),
            "{label}: {error:?}"
        );
    }

    let predefined = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>ok &amp; &#x41;</w:t></w:r></w:p></w:body></w:document>"#
    );
    let package = managed_raw_document(predefined.as_bytes());
    let document = package.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), "ok & A");

    let declaration = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>declared</w:t></w:r></w:p></w:body></w:document>"#
    );
    let package = managed_raw_document(declaration.as_bytes());
    assert_eq!(
        package.document().unwrap().extract_text().unwrap(),
        "declared"
    );
}

#[test]
fn source_section_scan_rejects_non_xml_outer_whitespace_and_late_declarations() {
    let cases = [
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>{}"#,
                '\u{000b}'
            ),
            "vertical-tab trailing text",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>{}"#,
                '\u{000c}'
            ),
            "form-feed trailing text",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>&amp;"#
            ),
            "trailing predefined entity",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>&#65;"#
            ),
            "trailing numeric entity",
        ),
        (
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document><?xml version="1.0"?>"#
            ),
            "late XML declaration",
        ),
        (
            format!(
                r#"<?xml version="1.0"?><?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>"#
            ),
            "duplicate XML declaration",
        ),
    ];
    for (xml, label) in cases {
        let package = managed_raw_document(xml.as_bytes());
        assert!(
            package.section_inventory_snapshot().is_err(),
            "{label}: section preflight unexpectedly succeeded"
        );
    }

    let valid = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>"#
    );
    let package = managed_raw_document(valid.as_bytes());
    assert!(package.section_inventory_snapshot().is_ok());
}

#[test]
fn malformed_source_document_errors_at_first_semantic_read() {
    let bytes = package_bytes(
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="invalid"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
        ),
        None,
        false,
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let document = package.document().unwrap();
    let table = document.tables().unwrap().pop().unwrap();
    let cell = table
        .rows()
        .unwrap()
        .pop()
        .unwrap()
        .cells()
        .unwrap()
        .pop()
        .unwrap();
    assert!(matches!(cell.v_merge(), Err(Error::InvalidFormat(_))));
}
