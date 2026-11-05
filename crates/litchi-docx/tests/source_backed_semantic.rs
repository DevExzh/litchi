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
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
    source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_docx::OpenOptions::default()
            .source_limits(litchi_docx::SourceLimits::default())
            .execution_context(litchi_core::ExecutionContext::new(
                budget,
                cancellation,
                execution_limits,
            )),
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
    document: &litchi_docx::document::Document,
) -> litchi_docx::Result<Vec<String>> {
    document
        .blocks()?
        .into_iter()
        .map(|block| match block {
            Block::Paragraph(paragraph) => Ok(format!("p:{}", paragraph.text()?)),
            Block::Table(table) => Ok(format!("t:{}", table_signature(&table)?)),
            Block::Alt(chunk) => Ok(format!("a:{}", chunk.relationship().as_str())),
            Block::Sdt(control) => Ok(format!("s:{}", control.text()?)),
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
            Block::Sdt(control) => Ok(format!("s:{}", control.text()?)),
            Block::Unknown(block) => {
                Ok(format!("u:{}", String::from_utf8_lossy(block.xml_bytes())))
            },
        })
        .collect()
}

fn eager_element_signature(
    document: &litchi_docx::document::Document,
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
    let eager_props = eager.metadata().unwrap();
    let source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(transitional))).unwrap();
    let metadata = source.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), eager_props.title.as_deref());
    assert_eq!(metadata.author.as_deref(), eager_props.author.as_deref());
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
    // A core-properties namespace that disagrees with its relationship type
    // is repaired and recorded by the OPC/ooxml-common reader (leniency F8.1),
    // so the read succeeds instead of failing.
    assert_eq!(
        wrong_source.metadata().unwrap().title.as_deref(),
        Some("wrong")
    );

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
        Err(litchi_docx::document::Error::Document(Error::Opc(
            OpcError::SourceChanged { .. }
        )))
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

    let adopted_opc = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(semantic_fixture())),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let adopted = source_backed::Package::from_source_backed_package(adopted_opc).unwrap();
    assert_eq!(
        adopted.document().unwrap().extract_text().unwrap(),
        "fallbackouternestedtail"
    );
}

#[test]
fn source_edit_budget_refuses_before_main_payload_capture() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>payload</w:t></w:r></w:p></w:body></w:document>"#
    );
    let bytes = package_bytes(document.clone(), None, false);
    let source = Arc::new(ProbeSource::new(bytes, None));
    let package = source_backed::Package::from_read_at(source.clone()).unwrap();
    let maximum = document.len().saturating_sub(1);

    let edit = package
        .edit_with_options(
            litchi_docx::document::EditOptions::default().max_snapshot_bytes(maximum),
        )
        .expect("the lazy constructor retains the edit policy");
    let result = edit.source().map(|_| ());
    assert!(matches!(
        result,
        Err(litchi_docx::document::Error::Limit {
            kind: litchi_docx::document::LimitKind::ProjectedXmlBytes,
            observed,
            maximum: actual,
        }) if observed == document.len() && actual == maximum
    ));
    assert_eq!(source.main_reads.load(Ordering::SeqCst), 0);
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
    let memory = 1 << 20;
    let budget = litchi_core::Budget::root(
        "docx-semantic-managed-test",
        litchi_core::Limits::new(
            memory,
            64 * 1024 * 1024,
            64 * 1024 * 1024,
            1_000_000,
            1024,
            1 << 30,
        ),
    );
    let (cancellation_source, cancellation) = litchi_core::CancellationSource::pair();
    let limits = litchi_core::ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let context = litchi_core::ExecutionContext::new(budget, cancellation, limits);
    let package = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_docx::OpenOptions::default()
            .source_limits(litchi_docx::SourceLimits::default())
            .execution_context(context),
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
        assert!(matches!(error, Error::Refused(_)));
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
            matches!(error, Error::Refused(_) | Error::InvalidFormat(_)),
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

fn source_paragraph_texts(document: &source_backed::Document) -> Vec<String> {
    document
        .paragraphs()
        .unwrap()
        .iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

#[test]
fn source_paragraph_queries_are_order_and_thread_independent() {
    let bytes = semantic_fixture();
    let source = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();

    // A caller that extracts text first and a caller that selects paragraphs
    // first must observe exactly the same paragraph projection: the unmanaged
    // index is built on the first paragraph query, not at `document()`.
    let text_first = source.document().unwrap();
    let text = text_first.extract_text().unwrap();
    let after_text = (
        text_first.paragraph_count().unwrap(),
        source_paragraph_texts(&text_first),
        text_first.paragraph_text(0).unwrap(),
        text_first
            .paragraph(0)
            .unwrap()
            .map(|paragraph| paragraph.text().unwrap()),
    );

    let paragraph_first = source.document().unwrap();
    let before_text = (
        paragraph_first.paragraph_count().unwrap(),
        source_paragraph_texts(&paragraph_first),
        paragraph_first.paragraph_text(0).unwrap(),
        paragraph_first
            .paragraph(0)
            .unwrap()
            .map(|paragraph| paragraph.text().unwrap()),
    );
    assert_eq!(after_text, before_text);
    assert_eq!(paragraph_first.extract_text().unwrap(), text);
    assert_eq!(after_text.0, 4);
    assert_eq!(after_text.3.as_deref(), Some("fallback"));

    // Repeating every query on one document answers from the cached ranges.
    assert_eq!(
        (
            text_first.paragraph_count().unwrap(),
            source_paragraph_texts(&text_first),
            text_first.paragraph_text(0).unwrap(),
        ),
        (after_text.0, after_text.1.clone(), after_text.2.clone())
    );

    // Concurrent first queries on one shared view agree on one index.
    let shared = source.document().unwrap();
    let shared = &shared;
    std::thread::scope(|scope| {
        let handles: Vec<_> = (0..8)
            .map(|_| {
                scope.spawn(move || {
                    (
                        shared.paragraph_count().unwrap(),
                        source_paragraph_texts(shared),
                        shared.paragraph_text(0).unwrap(),
                    )
                })
            })
            .collect();
        for handle in handles {
            let observed = handle.join().unwrap();
            assert_eq!(
                observed,
                (after_text.0, after_text.1.clone(), after_text.2.clone())
            );
        }
    });
}

#[test]
fn managed_source_paragraph_queries_keep_their_eager_index() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body></w:document>"#
    );
    let managed = managed_raw_document(document.as_bytes());
    let view = managed.document().unwrap();

    assert_eq!(view.paragraph_count().unwrap(), 2);
    assert_eq!(view.paragraph_text(0).unwrap().as_deref(), Some("one"));
    assert_eq!(view.paragraph_text(1).unwrap().as_deref(), Some("two"));
    assert_eq!(view.paragraph_text(2).unwrap(), None);
    assert_eq!(view.extract_text().unwrap(), "onetwo");
    // Arc-backed selective views stay refused on the managed route.
    assert!(matches!(view.paragraphs().unwrap_err(), Error::Refused(_)));
}

fn effective_xml(text: &str) -> Vec<u8> {
    format!(r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>"#).into_bytes()
}

#[test]
fn effective_owner_reads_publishes_and_materializes_current_graph() {
    let physical = SourceBackedPackage::from_vec(
        package_bytes(effective_xml("before"), None, false),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
    let old_view = owner.document().unwrap();
    assert_eq!(old_view.paragraph_count().unwrap(), 1);
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_replace_part(
        PackURI::new("/word/document.xml").unwrap(),
        effective_xml("after"),
    )
    .unwrap();
    plan.try_remove_part(PackURI::new("/word/unrelated.bin").unwrap())
        .unwrap();
    let candidate = physical.clone().prepare_topology(plan).unwrap();
    let changed = owner.with_prepared_topology(candidate).unwrap();
    let view = changed.document().unwrap();
    assert_eq!(view.paragraph_text(0).unwrap().as_deref(), Some("after"));
    assert_eq!(
        changed
            .to_owned_package()
            .unwrap()
            .document()
            .unwrap()
            .text()
            .unwrap(),
        "after"
    );
    let mut bytes = Vec::new();
    changed.write_to(&mut bytes).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        bytes,
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert!(
        reopened
            .part(&PackURI::new("/word/unrelated.bin").unwrap())
            .is_err()
    );
    assert_eq!(
        source_backed::Package::from_source_backed_package(reopened)
            .unwrap()
            .document()
            .unwrap()
            .extract_text()
            .unwrap(),
        "after"
    );
    assert_eq!(old_view.extract_text().unwrap(), "before");
    assert_eq!(owner.document().unwrap().extract_text().unwrap(), "before");
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    let two = String::from_utf8(effective_xml("second")).unwrap().replace(
        "</w:body>",
        "<w:p><w:r><w:t>third</w:t></w:r></w:p></w:body>",
    );
    plan.try_replace_part(
        PackURI::new("/word/document.xml").unwrap(),
        two.into_bytes(),
    )
    .unwrap();
    let later = changed
        .with_prepared_topology(physical.prepare_topology(plan).unwrap())
        .unwrap();
    assert_eq!(later.document().unwrap().paragraph_count().unwrap(), 2);
    assert_eq!(view.paragraph_count().unwrap(), 1);
    drop(later);
    drop(changed);
    drop(owner);
    assert_eq!(view.extract_text().unwrap(), "after");
}

#[test]
fn effective_owner_refuses_unmigrated_operations_and_foreign_candidates() {
    let physical = SourceBackedPackage::from_vec(
        package_bytes(effective_xml("before"), None, false),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
    let commit = owner.edit().unwrap().commit().unwrap();
    for changed in [false, true] {
        let mut plan = litchi_opc::SourceTopologyPlan::new();
        if changed {
            plan.try_replace_part(
                PackURI::new("/word/document.xml").unwrap(),
                effective_xml("after"),
            )
            .unwrap();
        }
        let effective = owner
            .with_prepared_topology(physical.clone().prepare_topology(plan).unwrap())
            .unwrap();
        assert!(effective.source_backed_package().is_err());
        assert!(effective.document_snapshot().is_ok());
        assert!(effective.edit().is_ok());
        assert!(effective.metadata().is_ok());
        assert!(effective.embedded().is_err());
        assert!(effective.ink().is_err());
        assert!(effective.section_layout_snapshot().is_err());
        assert!(effective.document_variables_snapshot().is_err());
        assert!(
            effective
                .external_hyperlink_sanitization_snapshot()
                .is_err()
        );
        assert!(effective.story_hyperlinks_only_snapshot().is_err());
        // Record 0355a: a detached patch replays by content. The no-op
        // patch of "before" applies to the unchanged effective owner (and
        // writes its artifact) and is refused as stale by the changed one.
        let applied = effective.apply(commit.patch());
        if changed {
            assert!(applied.is_err());
        } else {
            let applied = applied.unwrap();
            let (mut written, mut expected) = (Vec::new(), Vec::new());
            applied.package().write_to(&mut written).unwrap();
            effective.write_to(&mut expected).unwrap();
            assert_eq!(written, expected);
        }
        let mut sink = Vec::new();
        effective
            .write_text_to(&mut sink, litchi_core::TextOutputOptions::default())
            .unwrap();
        assert_eq!(
            sink,
            if changed {
                b"after".as_slice()
            } else {
                b"before".as_slice()
            }
        );
    }
    let foreign = SourceBackedPackage::from_vec(
        package_bytes(effective_xml("foreign"), None, false),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .prepare_topology(litchi_opc::SourceTopologyPlan::new())
    .unwrap();
    assert!(matches!(
        owner.with_prepared_topology(foreign),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn effective_owner_managed_document_retains_admissions_and_uses_fresh_indexes() {
    let owner = managed_raw_document(&effective_xml("before"));
    let physical = owner.source_backed_package().unwrap().clone();
    let old_view = owner.document().unwrap();
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_replace_part(
        PackURI::new("/word/document.xml").unwrap(),
        effective_xml("after"),
    )
    .unwrap();
    let changed = owner
        .with_prepared_topology(physical.prepare_topology(plan).unwrap())
        .unwrap();
    let view = changed.document().unwrap();
    assert_eq!(view.paragraph_text(0).unwrap().as_deref(), Some("after"));
    assert!(view.paragraphs().is_err());
    assert!(matches!(
        changed.to_owned_package(),
        Err(Error::Opc(OpcError::ManagedPackageMaterialization))
    ));
    drop(changed);
    drop(owner);
    assert_eq!(view.extract_text().unwrap(), "after");
    assert_eq!(old_view.extract_text().unwrap(), "before");
}

#[test]
fn effective_owner_checks_donor_fences_through_source_version() {
    let donor_source = Arc::new(ProbeSource::new(
        package_bytes(effective_xml("donor"), None, false),
        None,
    ));
    let donor = SourceBackedPackage::from_read_at(
        donor_source.clone(),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let xml = donor.main_document_part().unwrap().source_xml().unwrap();
    let physical = SourceBackedPackage::from_vec(
        package_bytes(effective_xml("base"), None, false),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_add_source_xml_part(PackURI::new("/word/copied.xml").unwrap(), xml)
        .unwrap();
    plan.try_replace_internal_relationship(
        PackURI::new("/").unwrap(),
        "rId1",
        rt::OFFICE_DOCUMENT,
        PackURI::new("/word/copied.xml").unwrap(),
    )
    .unwrap();
    let owner = owner
        .with_prepared_topology(physical.prepare_topology(plan).unwrap())
        .unwrap();
    let view = owner.document().unwrap();
    assert_eq!(view.extract_text().unwrap(), "donor");
    donor_source.changed();
    assert!(matches!(
        owner.source_version(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(owner.document().is_err());
    assert!(owner.to_owned_package().is_err());
    let mut sink = Vec::new();
    assert!(owner.write_to(&mut sink).is_err());
    assert!(sink.is_empty());
    assert_eq!(view.extract_text().unwrap(), "donor");
}

#[test]
fn effective_owner_pinned_document_checks_managed_donor_cancellation() {
    let memory = 16 * 1024 * 1024;
    let budget = litchi_core::Budget::root(
        "effective-donor",
        litchi_core::Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancel, cancellation) = litchi_core::CancellationSource::pair();
    let limits = litchi_core::ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let donor = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(package_bytes(
            effective_xml("donor"),
            None,
            false,
        ))),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        Some(litchi_core::ExecutionContext::new(
            budget.clone(),
            cancellation,
            limits,
        )),
    )
    .unwrap();
    let xml = donor.main_document_part().unwrap().source_xml().unwrap();
    let physical = SourceBackedPackage::from_vec(
        package_bytes(effective_xml("base"), None, false),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_add_source_xml_part(PackURI::new("/word/copied.xml").unwrap(), xml)
        .unwrap();
    plan.try_replace_internal_relationship(
        PackURI::new("/").unwrap(),
        "rId1",
        rt::OFFICE_DOCUMENT,
        PackURI::new("/word/copied.xml").unwrap(),
    )
    .unwrap();
    let owner = owner
        .with_prepared_topology(physical.prepare_topology(plan).unwrap())
        .unwrap();
    let view = owner.document().unwrap();
    assert_eq!(view.paragraph_text(0).unwrap().as_deref(), Some("donor"));
    assert!(matches!(view.body_blocks(), Err(Error::Refused(_))));
    assert!(matches!(view.paragraph(0), Err(Error::Refused(_))));
    assert!(matches!(view.blocks(), Err(Error::Refused(_))));
    assert!(matches!(view.elements(), Err(Error::Refused(_))));
    assert!(matches!(
        owner.document_snapshot(),
        Err(litchi_docx::document::Error::Document(Error::Opc(
            OpcError::ManagedPackageMaterialization
        )))
    ));
    assert!(
        owner
            .edit()
            .and_then(|edit| edit.commit().map(|_| ()))
            .is_err()
    );
    drop(donor);
    drop(owner);
    assert!(budget.used(litchi_core::Resource::Memory) > 0);
    cancel.cancel();
    assert!(matches!(
        view.extract_text(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        view.paragraph_count(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    drop(view);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}

#[test]
fn retained_semantic_patches_chain_and_inverse_without_reopening() {
    for managed in [false, true] {
        let xml = effective_xml("before");
        let source = if managed {
            managed_raw_document(&xml)
        } else {
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(package_bytes(
                xml, None, false,
            ))))
            .unwrap()
        };
        let mut edit = source.edit().unwrap();
        edit.replace_paragraph_text(litchi_core::Position::new(0), "first")
            .unwrap();
        let first = edit.commit().unwrap();
        let current = first.package().clone();
        assert_eq!(current.document().unwrap().extract_text().unwrap(), "first");
        let mut edit = current.edit().unwrap();
        edit.replace_paragraph_text(litchi_core::Position::new(0), "second")
            .unwrap();
        let second = edit.commit().unwrap();
        let next = second.clone().into_package();
        assert_eq!(next.document().unwrap().extract_text().unwrap(), "second");
        assert!(next.apply(first.patch()).is_err());
        let restored = next
            .apply(&second.patch().inverse())
            .unwrap()
            .into_package()
            .apply(&first.patch().inverse())
            .unwrap()
            .into_package();
        let mut before = Vec::new();
        source.write_to(&mut before).unwrap();
        let mut after = Vec::new();
        restored.write_to(&mut after).unwrap();
        assert_eq!(before, after);
        assert_eq!(current.document().unwrap().extract_text().unwrap(), "first");
        assert_eq!(source.document().unwrap().extract_text().unwrap(), "before");
    }
}

#[test]
fn retained_semantic_apply_preserves_graph_edits_and_effective_protection() {
    for policy in [
        "",
        "<w:documentProtection w:enforcement=\"1\"/>",
        "<w:trackRevisions/>",
    ] {
        let physical = SourceBackedPackage::from_vec(
            package_bytes(effective_xml("before"), None, false),
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            None,
        )
        .unwrap();
        let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
        let mut plan = litchi_opc::SourceTopologyPlan::new();
        let settings = PackURI::new("/word/settings.xml").unwrap();
        plan.try_add_part(
            settings.clone(),
            ct::WML_SETTINGS,
            format!("<w:settings xmlns:w=\"{W}\">{policy}</w:settings>").into_bytes(),
        )
        .unwrap();
        plan.try_add_internal_relationship(
            PackURI::new("/word/document.xml").unwrap(),
            "rSettings",
            rt::SETTINGS,
            settings.clone(),
        )
        .unwrap();
        plan.try_remove_part(PackURI::new("/word/unrelated.bin").unwrap())
            .unwrap();
        let owner = owner
            .with_prepared_topology(physical.prepare_topology(plan).unwrap())
            .unwrap();
        let noop = owner.edit().unwrap().commit().unwrap();
        owner.apply(noop.patch()).unwrap();
        let mut edit = owner.edit().unwrap();
        edit.replace_paragraph_text(litchi_core::Position::new(0), "after")
            .unwrap();
        let result = edit.commit();
        if !policy.is_empty() {
            assert!(matches!(
                result,
                Err(litchi_docx::document::Error::Document(Error::Refused(_)))
            ));
            let mut canonical = owner.edit().unwrap();
            canonical
                .replace_paragraph_text(litchi_core::Position::new(0), "after")
                .unwrap();
            assert!(matches!(
                canonical.commit(),
                Err(litchi_docx::document::Error::Document(Error::Refused(_)))
            ));
            assert_eq!(owner.document().unwrap().extract_text().unwrap(), "before");
            continue;
        }
        assert!(result.is_ok());
        let mut canonical = owner.edit().unwrap();
        canonical
            .replace_paragraph_text(litchi_core::Position::new(0), "after")
            .unwrap();
        let next = canonical.commit().unwrap().into_package();
        let mut bytes = Vec::new();
        next.write_to(&mut bytes).unwrap();
        let package = SourceBackedPackage::from_vec(
            bytes,
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            None,
        )
        .unwrap();
        assert!(package.part(&settings).is_ok());
        assert!(
            package
                .main_document_part()
                .unwrap()
                .rels()
                .get("rSettings")
                .is_some()
        );
        assert!(
            package
                .part(&PackURI::new("/word/unrelated.bin").unwrap())
                .is_err()
        );
        assert_eq!(next.document().unwrap().extract_text().unwrap(), "after");
        assert_eq!(owner.document().unwrap().extract_text().unwrap(), "before");
    }
}

#[test]
fn retained_semantic_apply_preserves_declaration_bom_and_revision_spelling() {
    use litchi_core::Position;
    use litchi_docx::document::{RevisionAction, RevisionKind, RevisionSelector};
    let xml = format!(
        "\u{feff}<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<w:document xmlns:w=\"{W}\"><w:body>\n<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>\n</w:body></w:document>"
    );
    let owner = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(package_bytes(
        xml.as_bytes().to_vec(),
        None,
        false,
    ))))
    .unwrap();
    let mut edit = owner.edit().unwrap();
    edit.apply_revision(
        RevisionSelector::new(Position::new(0), RevisionKind::Insertion, Position::new(0)),
        RevisionAction::Accept,
    )
    .unwrap();
    let committed = edit.commit().unwrap();
    let next = owner.apply(committed.patch()).unwrap().into_package();
    let snapshot = next.document_snapshot().unwrap();
    assert!(snapshot.xml_bytes().starts_with(b"\xef\xbb\xbf<?xml"));
    assert_eq!(snapshot.xml_bytes(), committed.patch().target().xml_bytes());
    let restored = next
        .apply(&committed.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(
        restored.document_snapshot().unwrap().xml_bytes(),
        xml.as_bytes()
    );
}

#[test]
fn canonical_source_edit_owns_package_and_isolates_staged_clones() {
    for managed in [false, true] {
        let xml = effective_xml("before");
        let source = if managed {
            managed_raw_document(&xml)
        } else {
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(package_bytes(
                xml, None, false,
            ))))
            .unwrap()
        };
        let mut edit: litchi_docx::document::Edit<source_backed::Package> = source.edit().unwrap();
        let view = source.document().unwrap();
        edit.replace_paragraph_text(litchi_core::Position::new(0), "first")
            .unwrap();
        let mut sibling = edit.try_clone().unwrap();
        sibling
            .replace_paragraph_text(litchi_core::Position::new(0), "second")
            .unwrap();
        drop(source);
        let commit: litchi_docx::document::Commit<source_backed::Package> = edit.commit().unwrap();
        assert_eq!(
            commit.package().document().unwrap().extract_text().unwrap(),
            "first"
        );
        assert_eq!(view.extract_text().unwrap(), "before");
        let (owner, patch) = commit.into_parts();
        assert!(!patch.is_noop());
        assert_eq!(
            owner
                .apply(&patch.inverse())
                .unwrap()
                .into_package()
                .document()
                .unwrap()
                .extract_text()
                .unwrap(),
            "before"
        );
        assert_eq!(
            sibling
                .commit()
                .unwrap()
                .into_package()
                .document()
                .unwrap()
                .extract_text()
                .unwrap(),
            "second"
        );
        let mut abandoned = owner.edit().unwrap();
        abandoned
            .replace_paragraph_text(litchi_core::Position::new(0), "abandoned")
            .unwrap();
        drop(abandoned);
        assert_eq!(owner.document().unwrap().extract_text().unwrap(), "first");
    }
}

#[test]
fn canonical_source_commit_refuses_a_changed_physical_source() {
    let physical = Arc::new(ProbeSource::new(
        package_bytes(effective_xml("before"), None, false),
        None,
    ));
    let owner = source_backed::Package::from_read_at(physical.clone()).unwrap();
    let mut edit = owner.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after")
        .unwrap();
    physical.changed();
    assert!(matches!(
        edit.commit(),
        Err(litchi_docx::document::Error::Document(Error::Opc(
            OpcError::SourceChanged { .. }
        )))
    ));
}

#[test]
fn effective_metadata_tracks_payload_relationships_and_owner_reports() {
    let bytes = package_bytes(
        effective_xml("before"),
        Some((&transitional_core("original"), rt::CORE_PROPERTIES)),
        false,
    );
    let physical = SourceBackedPackage::from_vec(
        bytes,
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
    assert_eq!(owner.metadata().unwrap().title.as_deref(), Some("original"));
    let mut edit = owner.edit().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after")
        .unwrap();
    let committed = edit.commit().unwrap().into_package();
    assert_eq!(
        committed.metadata().unwrap().title.as_deref(),
        Some("original")
    );
    let core = PackURI::new("/docProps/core.xml").unwrap();
    let root = PackURI::new("/").unwrap();
    let core_id = physical
        .rels()
        .iter()
        .find(|rel| rel.reltype() == rt::CORE_PROPERTIES)
        .unwrap()
        .r_id()
        .to_owned();
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_replace_part(core.clone(), transitional_core("changed").into_bytes())
        .unwrap();
    let changed = owner
        .with_prepared_topology(physical.clone().prepare_topology(plan).unwrap())
        .unwrap();
    assert_eq!(
        changed.metadata().unwrap().title.as_deref(),
        Some("changed")
    );
    assert_eq!(owner.metadata().unwrap().title.as_deref(), Some("original"));

    let mut plan = litchi_opc::SourceTopologyPlan::new();
    let target = PackURI::new("/custom/metadata.xml").unwrap();
    plan.try_add_part(
        target.clone(),
        ct::OPC_CORE_PROPERTIES,
        transitional_core("redirected").into_bytes(),
    )
    .unwrap();
    plan.try_remove_part(core.clone()).unwrap();
    plan.try_replace_internal_relationship(
        root.clone(),
        core_id.clone(),
        STRICT_CORE_RELATIONSHIP,
        target,
    )
    .unwrap();
    let redirected = owner
        .with_prepared_topology(physical.clone().prepare_topology(plan).unwrap())
        .unwrap();
    assert_eq!(
        redirected.metadata().unwrap().title.as_deref(),
        Some("redirected")
    );
    let repairs = redirected.tolerance_report().total();
    assert!(repairs > owner.tolerance_report().total());
    redirected.metadata().unwrap();
    assert_eq!(redirected.tolerance_report().total(), repairs);

    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_remove_relationship(root, core_id).unwrap();
    plan.try_remove_part(core).unwrap();
    let removed = owner
        .with_prepared_topology(physical.prepare_topology(plan).unwrap())
        .unwrap();
    assert_eq!(removed.metadata().unwrap().title, None);
    assert_eq!(
        changed.metadata().unwrap().title.as_deref(),
        Some("changed")
    );
}

#[test]
fn effective_metadata_reads_only_selected_payload_and_fences_source_changes() {
    for malformed in [false, true] {
        let core_xml = if malformed {
            "<broken/>".to_owned()
        } else {
            transitional_core("selected")
        };
        let bytes = package_bytes(
            effective_xml("before"),
            Some((&core_xml, rt::CORE_PROPERTIES)),
            false,
        );
        let range = payload_range(&bytes, "docProps/core.xml");
        let observed = Arc::new(ProbeSource::new(bytes, Some(range)));
        let physical = SourceBackedPackage::from_read_at(
            observed.clone(),
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
        let mut plan = litchi_opc::SourceTopologyPlan::new();
        plan.try_replace_part(
            PackURI::new("/word/document.xml").unwrap(),
            effective_xml("after"),
        )
        .unwrap();
        let owner = owner
            .with_prepared_topology(physical.prepare_topology(plan).unwrap())
            .unwrap();
        let main_reads = observed.main_reads.load(Ordering::SeqCst);
        let unrelated_reads = observed.unrelated_reads.load(Ordering::SeqCst);
        observed.arm_core_change();
        assert!(matches!(
            owner.metadata(),
            Err(Error::Opc(OpcError::SourceChanged { .. }))
        ));
        assert!(observed.core_reads.load(Ordering::SeqCst) > 0);
        assert_eq!(observed.main_reads.load(Ordering::SeqCst), main_reads);
        assert_eq!(
            observed.unrelated_reads.load(Ordering::SeqCst),
            unrelated_reads
        );
    }
}

#[test]
fn effective_metadata_checks_managed_donor_context_and_releases_admissions() {
    let memory = 16 * 1024 * 1024;
    let budget = litchi_core::Budget::root(
        "metadata-donor",
        litchi_core::Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancel, cancellation) = litchi_core::CancellationSource::pair();
    let limits = litchi_core::ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let donor = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(package_bytes(
            effective_xml("donor"),
            Some((&transitional_core("donor-title"), rt::CORE_PROPERTIES)),
            false,
        ))),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        Some(litchi_core::ExecutionContext::new(
            budget.clone(),
            cancellation,
            limits,
        )),
    )
    .unwrap();
    let core = PackURI::new("/docProps/core.xml").unwrap();
    let xml = donor.part(&core).unwrap().source_xml().unwrap();
    let physical = SourceBackedPackage::from_vec(
        package_bytes(effective_xml("base"), None, false),
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let owner = source_backed::Package::from_source_backed_package(physical.clone()).unwrap();
    let mut plan = litchi_opc::SourceTopologyPlan::new();
    plan.try_add_source_xml_part(core.clone(), xml).unwrap();
    plan.try_add_internal_relationship(
        PackURI::new("/").unwrap(),
        "rCore",
        rt::CORE_PROPERTIES,
        core,
    )
    .unwrap();
    let owner = owner
        .with_prepared_topology(physical.prepare_topology(plan).unwrap())
        .unwrap();
    drop(donor);
    assert_eq!(
        owner.metadata().unwrap().title.as_deref(),
        Some("donor-title")
    );
    assert!(budget.used(litchi_core::Resource::Memory) > 0);
    cancel.cancel();
    assert!(matches!(
        owner.metadata(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    drop(owner);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}
