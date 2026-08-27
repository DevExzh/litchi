use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_docx::section::{
    Column, Columns, Emu, Limits, Margins, Orientation, Ownership, PageSize, Property,
    PropertyValue, Selector, Start,
};
use litchi_docx::{Error, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MAIN: &str = "word/document.xml";
const HEADER: &str = "word/header1.xml";
const UNUSED: &str = "word/unused.bin";

fn fixture(document: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{MAIN}")).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document.to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{HEADER}")).unwrap(),
            ct::WML_HEADER.to_owned(),
            format!(
                r#"<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>inert header</w:t></w:r></w:p></w:hdr>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED}")).unwrap(),
            "application/octet-stream".to_owned(),
            vec![b'u'; 512 * 1024],
        )))
        .unwrap();
    package.relate_to(MAIN, rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn malformed_fixture(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/{MAIN}" ContentType="{}"/></Types>"#,
        ct::WML_DOCUMENT_MAIN
    );
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{}" Target="{MAIN}"/></Relationships>"#,
        rt::OFFICE_DOCUMENT
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .unwrap();
    writer.write_stored(MAIN, document).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn signed_section_fixture(document: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{MAIN}")).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document.to_vec(),
        )))
        .unwrap();
    package.relate_to(MAIN, rt::OFFICE_DOCUMENT);
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            b"<origin/>".to_vec(),
        )))
        .unwrap();
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    PackageWriter::to_bytes(&package).unwrap()
}

struct SectionFailingSink {
    accepted: usize,
    limit: usize,
}

impl Write for SectionFailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected section sink failure"));
        }
        let count = bytes.len().min(self.limit - self.accepted);
        self.accepted += count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn managed_document_fixture(document: &[u8]) -> (Budget, source_backed::Package) {
    let bytes = malformed_fixture(document);
    let memory = 16 * 1024 * 1024;
    let budget = Budget::root(
        "docx-managed-section-test",
        CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        litchi_opc::ReadLimits::default(),
        ExecutionContext::new(budget.clone(), cancellation, execution_limits),
    )
    .unwrap();
    (budget, package)
}

struct ObservedSource {
    bytes: Vec<u8>,
    main: std::ops::Range<usize>,
    header: std::ops::Range<usize>,
    unused: std::ops::Range<usize>,
    header_reads: AtomicUsize,
    unused_reads: AtomicUsize,
    revision: AtomicU64,
    change_after_read: AtomicBool,
}

impl ObservedSource {
    fn new(bytes: Vec<u8>, change_after_read: bool) -> Self {
        Self {
            main: payload_range(&bytes, MAIN),
            header: payload_range(&bytes, HEADER),
            unused: payload_range(&bytes, UNUSED),
            bytes,
            header_reads: AtomicUsize::new(0),
            unused_reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            change_after_read: AtomicBool::new(change_after_read),
        }
    }

    fn arm_change(&self) {
        self.change_after_read.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for ObservedSource {
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
        let requested = offset..end;
        if overlaps(&requested, &self.header) {
            self.header_reads.fetch_add(1, Ordering::SeqCst);
        }
        if overlaps(&requested, &self.unused) {
            self.unused_reads.fetch_add(1, Ordering::SeqCst);
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        if self.change_after_read.load(Ordering::SeqCst) && overlaps(&requested, &self.main) {
            self.revision.store(1, Ordering::SeqCst);
        }
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            8085,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn source_snapshot_reads_only_mandatory_main_and_retains_exact_version() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p/><w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgMar w:left="720"/></w:sectPr></w:body></w:document>"#
    );
    let observed = Arc::new(ObservedSource::new(fixture(document.as_bytes()), false));
    let read_at: Arc<dyn ReadAt> = observed.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let before = package.cache_diagnostics();
    let snapshot = package.section_inventory_snapshot().unwrap();
    let after = package.cache_diagnostics();

    assert_eq!(snapshot.source_version(), Some(SourceVersion::new(8085, 0)));
    assert_eq!(snapshot.inventory().sections().len(), 1);
    assert_eq!(
        snapshot.inventory().sections()[0].ownership(),
        Ownership::BodyFinal
    );
    assert_eq!(
        snapshot.inventory().sections()[0].headers()[0].relationship_id,
        "rHeader"
    );
    assert!(matches!(
        snapshot.property(0, Property::Margins),
        Some(PropertyValue::Margins(Some(_)))
    ));
    assert_eq!(after.cold_loads, before.cold_loads + 1);
    assert_eq!(observed.header_reads.load(Ordering::SeqCst), 0);
    assert_eq!(observed.unused_reads.load(Ordering::SeqCst), 0);

    let clone = snapshot.clone();
    assert!(snapshot.shares_allocation_with(&clone));
    let pinned = package.document().unwrap();
    assert_eq!(pinned.section_inventory().unwrap().sections().len(), 1);
}

#[test]
fn source_snapshot_applies_limits_and_rejects_a_changed_read_at() {
    let document = format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#);
    let bytes = fixture(document.as_bytes());

    let package =
        source_backed::Package::from_read_at(Arc::new(ObservedSource::new(bytes.clone(), false)))
            .unwrap();
    let limits = Limits {
        max_input_bytes: document.len() - 1,
        ..Limits::default()
    };
    assert!(matches!(
        package.section_inventory_snapshot_with_limits(&limits),
        Err(Error::SectionInventoryLimit { .. })
    ));

    let hostile = Arc::new(ObservedSource::new(bytes, false));
    let read_at: Arc<dyn ReadAt> = hostile.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    // The armed adapter advances its revision while the pinned main payload
    // is read, between the inventory's pre-capture and post-capture checks.
    hostile.arm_change();
    assert!(matches!(
        package.section_inventory_snapshot(),
        Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
    ));

    let event_limited = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        document.as_bytes(),
    ))))
    .unwrap();
    let limits = Limits {
        max_events: 1,
        ..Limits::default()
    };
    assert!(matches!(
        event_limited.section_inventory_snapshot_with_limits(&limits),
        Err(Error::SectionInventoryLimit {
            resource: "XML events",
            ..
        })
    ));

    let deep_document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>deep</w:t></w:r></w:p></w:body></w:document>"#
    );
    let depth_limited = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        malformed_fixture(deep_document.as_bytes()),
    )))
    .unwrap();
    let limits = Limits {
        max_depth: 2,
        ..Limits::default()
    };
    assert!(matches!(
        depth_limited.section_inventory_snapshot_with_limits(&limits),
        Err(Error::SectionInventoryLimit {
            resource: "XML depth",
            ..
        })
    ));
}

#[test]
fn source_snapshot_refuses_mce_dtd_and_entity_syntax() {
    let documents = [
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><mc:AlternateContent><mc:Fallback><w:sectPr/></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
        ),
        format!(
            r#"<!DOCTYPE w:document [<!ELEMENT document ANY>]><w:document xmlns:w="{W}"><w:body><w:sectPr/></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&custom;</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="http:&#x2f;&#x2f;schemas.openxmlformats.org&#x2f;markup-compatibility&#x2f;2006"><w:body><w:p/></w:body></w:document>"#
        ),
    ];

    for document in documents {
        let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
            malformed_fixture(document.as_bytes()),
        )))
        .unwrap();
        assert!(matches!(
            package.section_inventory_snapshot(),
            Err(Error::UnsafeEdit { .. })
        ));
    }

    let predefined = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>ok &amp; &#x41;</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        malformed_fixture(predefined.as_bytes()),
    )))
    .unwrap();
    assert!(package.section_inventory_snapshot().is_ok());

    let uri = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let harmless_uri = format!(
        r#"<w:document xmlns:w="{W}" xmlns:x="urn:test"><w:body><w:p x:note="{uri}"><w:r><w:t>{uri}</w:t></w:r></w:p><!-- {uri} --><w:sectPr/></w:body></w:document>"#
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        malformed_fixture(harmless_uri.as_bytes()),
    )))
    .unwrap();
    assert!(package.section_inventory_snapshot().is_ok());

    let strict = r#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p/><s:sectPr/></s:body></s:document>"#;
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        malformed_fixture(strict.as_bytes()),
    )))
    .unwrap();
    assert_eq!(
        package
            .section_inventory_snapshot()
            .unwrap()
            .inventory()
            .sections()
            .len(),
        1
    );

    let valid_references = format!(
        r#"<w:document xmlns:w="{W}" xmlns:x="urn:test"><w:body><w:p x:value="&amp; &apos; &gt; &lt; &quot; &#65; &#x1F600;"><w:r><w:t>&amp; &apos; &gt; &lt; &quot; &#65; &#x1F600;</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        malformed_fixture(valid_references.as_bytes()),
    )))
    .unwrap();
    assert!(package.section_inventory_snapshot().is_ok());

    for (index, invalid) in [
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&#x1;</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&#xZZ;</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&#;</w:t></w:r></w:p></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:x="urn:test"><w:body><w:p x:value="&custom;"/></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:x="urn:test"><w:body><w:p x:value="&#xZZ;"/></w:body></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:x="urn:test"><w:body><w:p x:value="&#1;"/></w:body></w:document>"#
        ),
    ]
    .into_iter()
    .enumerate()
    {
        let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
            malformed_fixture(invalid.as_bytes()),
        )))
        .unwrap();
        assert!(
            package.section_inventory_snapshot().is_err(),
            "invalid reference fixture {index} was accepted"
        );
    }
}

#[test]
fn managed_document_refuses_only_actual_mce_and_retains_part_data() {
    let uri = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let harmless = format!(
        r#"<w:document xmlns:w="{W}" xmlns:x="urn:test"><w:body><w:p x:note="{uri}"><w:r><w:t>{uri}</w:t></w:r></w:p><!-- {uri} --><w:sectPr/></w:body></w:document>"#
    );
    let (budget, package) = managed_document_fixture(harmless.as_bytes());
    let document = package.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), uri);
    assert_eq!(document.section_inventory().unwrap().sections().len(), 1);
    assert!(budget.used(Resource::Memory) > 0);
    drop(document);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);

    let encoded_mce = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="http:&#x2f;&#x2f;schemas.openxmlformats.org&#x2f;markup-compatibility&#x2f;2006"><w:body><w:p/></w:body></w:document>"#
    );
    let (_budget, package) = managed_document_fixture(encoded_mce.as_bytes());
    assert!(matches!(
        package.document(),
        Err(Error::UnsafeEdit {
            operation: "source-backed document read",
            ..
        })
    ));
}

#[test]
fn managed_document_bounds_namespace_scan_depth_and_events() {
    let mut exact_depth = format!(r#"<w:document xmlns:w="{W}"><w:body>"#);
    for _ in 0..254 {
        exact_depth.push_str("<w:sdt>");
    }
    for _ in 0..254 {
        exact_depth.push_str("</w:sdt>");
    }
    exact_depth.push_str("</w:body></w:document>");
    let (_budget, package) = managed_document_fixture(exact_depth.as_bytes());
    assert!(
        package.document().is_ok(),
        "the finite depth boundary is inclusive"
    );

    let mut deep = format!(r#"<w:document xmlns:w="{W}"><w:body>"#);
    for _ in 0..255 {
        deep.push_str("<w:sdt>");
    }
    for _ in 0..255 {
        deep.push_str("</w:sdt>");
    }
    deep.push_str("</w:body></w:document>");
    let (_budget, package) = managed_document_fixture(deep.as_bytes());
    assert!(matches!(
        package.document(),
        Err(Error::InvalidFormat(reason)) if reason.contains("depth limit")
    ));

    let mut events = format!(r#"<w:document xmlns:w="{W}"><w:body>"#);
    for _ in 0..1_000_001 {
        events.push_str("<w:p/>");
    }
    events.push_str("</w:body></w:document>");
    let (_budget, package) = managed_document_fixture(events.as_bytes());
    assert!(matches!(
        package.document(),
        Err(Error::InvalidFormat(reason)) if reason.contains("event limit")
    ));

    let mismatched =
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:p></w:r></w:body></w:document>"#);
    let (_budget, package) = managed_document_fixture(mismatched.as_bytes());
    assert!(matches!(
        package.document(),
        Err(Error::InvalidFormat(reason)) if reason.contains("mismatched")
    ));

    let unclosed = format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/>"#);
    let (_budget, package) = managed_document_fixture(unclosed.as_bytes());
    assert!(matches!(
        package.document(),
        Err(Error::InvalidFormat(reason)) if reason.contains("unclosed")
    ));

    let mut long_name = format!(r#"<w:document xmlns:w="{W}"><w:body><"#);
    long_name.extend(std::iter::repeat_n('x', 64 * 1024 + 1));
    long_name.push_str("/></w:body></w:document>");
    let (_budget, package) = managed_document_fixture(long_name.as_bytes());
    let result = package.document();
    assert!(matches!(
        result,
        Err(Error::InvalidFormat(reason)) if reason.contains("name-byte limit")
    ));
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
        let size = u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_len =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        if &zip[offset + 46..offset + 46 + name_len] == name {
            let local =
                u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
            let local_name =
                u16::from_le_bytes(zip[local + 26..local + 28].try_into().unwrap()) as usize;
            let local_extra =
                u16::from_le_bytes(zip[local + 28..local + 30].try_into().unwrap()) as usize;
            let start = local + 30 + local_name + local_extra;
            return start..start + size;
        }
    }
    panic!("missing ZIP member")
}

fn main_xml(zip: &[u8]) -> String {
    let package = OpcPackage::from_reader(io::Cursor::new(zip.to_vec())).unwrap();
    let main = package
        .get_part(&PackURI::new(format!("/{MAIN}")).unwrap())
        .unwrap()
        .blob();
    String::from_utf8(main.to_vec()).unwrap()
}

#[test]
fn source_backed_existing_section_layout_is_exact_and_cell_safe() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:layout"><w:body><w:p><w:pPr><w:sectPr><w:type w:val="continuous" x:opaque="type"/><w:pgSz w:w="12240" w:h="15840" x:opaque="size"/><w:pgMar w:left="720" x:opaque="margins"/><w:cols w:num="1" x:opaque="columns"/></w:sectPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr><w:type w:val="oddPage"/></w:sectPr></w:pPr></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
    );
    let source_bytes = fixture(document.as_bytes());
    let original_bytes = source_bytes.clone();
    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source_bytes.clone())))
            .unwrap();
    let source = package.section_layout_snapshot().unwrap();
    assert_eq!(source.inventory().sections().len(), 2);
    assert_eq!(
        source.inventory().sections()[0].ownership(),
        Ownership::Paragraph(litchi_core::Position::new(0))
    );
    assert_eq!(
        source.inventory().sections()[1].ownership(),
        Ownership::BodyFinal
    );
    assert_eq!(source.inventory().sections()[0].columns().unwrap().count, 1);
    assert!(
        source
            .edit(Selector::paragraph(litchi_core::Position::new(1)))
            .is_err()
    );

    let mut edit = source
        .edit(Selector::paragraph(litchi_core::Position::new(0)))
        .unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Landscape,
    }))
    .unwrap()
    .set_margins(Some(Margins {
        top: Some(Emu::from_twips(720)),
        right: Some(Emu::from_twips(1080)),
        bottom: Some(Emu::from_twips(720)),
        left: Some(Emu::from_twips(1080)),
        header: Some(Emu::from_twips(360)),
        footer: Some(Emu::from_twips(360)),
        gutter: None,
    }))
    .unwrap()
    .set_start(Some(Start::NewPage))
    .unwrap()
    .set_columns(Some(Columns {
        equal_width: true,
        count: 2,
        space: Some(Emu::from_twips(240)),
        separator: true,
        columns: Vec::new(),
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().changed());
    assert_eq!(commit.snapshot().inventory().sections().len(), 2);
    assert_eq!(
        commit.snapshot().inventory().sections()[0].ownership(),
        Ownership::Paragraph(litchi_core::Position::new(0))
    );
    assert_eq!(
        commit.snapshot().inventory().sections()[0]
            .page_size()
            .unwrap()
            .orientation,
        Orientation::Landscape
    );
    assert_eq!(
        commit.snapshot().inventory().sections()[0]
            .margins()
            .unwrap()
            .left,
        Some(Emu::from_twips(1080))
    );
    assert_eq!(
        commit.snapshot().inventory().sections()[0].start(),
        Some(Start::NewPage)
    );
    assert_eq!(
        commit.snapshot().inventory().sections()[0]
            .columns()
            .unwrap()
            .count,
        2
    );

    let changed = commit.patch().apply(&source).unwrap();
    let restored = commit.patch().inverse().apply(&changed).unwrap();
    assert_eq!(
        restored.inventory().sections(),
        source.inventory().sections()
    );
    assert!(commit.patch().apply(&source).is_ok());
    let mut stale_bytes = document.as_bytes().to_vec();
    stale_bytes.push(b' ');
    let stale = litchi_docx::section::layout::Snapshot::from_xml(stale_bytes).unwrap();
    assert!(commit.patch().apply(&stale).is_err());

    let mut output = Vec::new();
    let publication = package
        .publish_section_layout_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(
        publication
            .inverse_patch()
            .apply(publication.snapshot())
            .is_ok()
    );
    assert_ne!(output, source_bytes);
    assert_eq!(
        &output[payload_range(&output, UNUSED)],
        &source_bytes[payload_range(&source_bytes, UNUSED)]
    );
    assert_eq!(
        &output[payload_range(&output, HEADER)],
        &source_bytes[payload_range(&source_bytes, HEADER)]
    );
    assert_eq!(
        &output[payload_range(&output, "_rels/.rels")],
        &source_bytes[payload_range(&source_bytes, "_rels/.rels")]
    );
    let reopened =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    assert_eq!(
        reopened
            .section_layout_snapshot()
            .unwrap()
            .inventory()
            .sections()[0]
            .page_size()
            .unwrap()
            .orientation,
        Orientation::Landscape
    );

    let mut restored_output = Vec::new();
    reopened
        .publish_section_layout_inverse_to_stream(&mut restored_output, &publication)
        .unwrap();
    assert_eq!(restored_output, original_bytes);

    let foreign =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(original_bytes.clone())))
            .unwrap();
    let mut refused = Vec::new();
    assert!(matches!(
        foreign.publish_section_layout_commit_to_stream(&mut refused, &commit),
        Err(Error::SectionLayoutForeignSource)
    ));
    assert!(refused.is_empty());

    let stale =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(original_bytes))).unwrap();
    let mut stale_output = Vec::new();
    assert!(matches!(
        stale.publish_section_layout_inverse_to_stream(&mut stale_output, &publication),
        Err(Error::SectionLayoutStaleSource)
    ));
    assert!(stale_output.is_empty());
}

#[test]
fn source_backed_section_layout_noop_mce_strict_and_limits_are_typed() {
    let simple =
        format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/><w:sectPr/></w:body></w:document>"#);
    let source = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        simple.as_bytes(),
    ))))
    .unwrap();
    let snapshot = source.section_layout_snapshot().unwrap();
    let noop = snapshot.edit(0).unwrap().commit().unwrap();
    assert!(noop.patch().is_noop());
    let mut output = Vec::new();
    source
        .publish_section_layout_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, fixture(simple.as_bytes()));

    let mce = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><mc:AlternateContent><mc:Fallback><w:sectPr/></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
    );
    let mce_package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(mce.as_bytes()))))
            .unwrap();
    assert!(matches!(
        mce_package.section_layout_snapshot(),
        Err(Error::UnsafeEdit { .. })
    ));
    assert!(litchi_docx::section::layout::Snapshot::from_xml(mce.as_bytes().to_vec()).is_err());

    let dtd = format!(
        r#"<!DOCTYPE w:document [<!ENTITY custom "x">]><w:document xmlns:w="{W}"><w:body><w:sectPr/></w:body></w:document>"#
    );
    assert!(litchi_docx::section::layout::Snapshot::from_xml(dtd.into_bytes()).is_err());

    let strict = r#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p/><s:sectPr><s:pgSz s:w="12240" s:h="15840"/></s:sectPr></s:body></s:document>"#;
    let strict_package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        strict.as_bytes(),
    ))))
    .unwrap();
    assert!(strict_package.section_layout_snapshot().is_ok());

    let limited = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        simple.as_bytes(),
    ))))
    .unwrap();
    assert!(matches!(
        limited.section_layout_snapshot_with_limits(&Limits {
            max_section_bytes: 1,
            ..Limits::default()
        }),
        Err(Error::SectionInventoryLimit { .. })
    ));
}

#[test]
fn section_layout_clear_is_lossless_and_single_property_patch_preserves_siblings() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:x="urn:opaque"><w:body><w:p><w:pPr><w:sectPr><w:type w:val="continuous" x:type="keep"/><w:pgSz w:w="12240" w:h="15840" x:size="keep"/><w:pgMar w:left="720" x:margins="keep"/><w:cols w:num="1" x:columns="keep"/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let source_bytes = fixture(document.as_bytes());
    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source_bytes.clone())))
            .unwrap();
    let source = package.section_layout_snapshot().unwrap();

    let mut clear = source.edit(0).unwrap();
    clear.clear_page_size().unwrap();
    assert!(matches!(clear.commit(), Err(Error::UnsafeEdit { .. })));
    assert!(source.inventory().sections()[0].page_size().is_some());

    let mut edit = source.edit(0).unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Landscape,
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_section_layout_commit_to_stream(&mut output, &commit)
        .unwrap();
    let reopened = OpcPackage::from_reader(io::Cursor::new(output)).unwrap();
    let main = reopened
        .get_part(&PackURI::new(format!("/{MAIN}")).unwrap())
        .unwrap()
        .blob();
    let main = std::str::from_utf8(main).unwrap();
    assert!(main.contains("x:type=\"keep\""));
    assert!(main.contains("x:size=\"keep\""));
    assert!(main.contains("x:margins=\"keep\""));
    assert!(main.contains("x:columns=\"keep\""));
    assert!(main.contains("w:w=\"11906\""));
    assert!(main.contains("w:h=\"16838\""));

    let clean = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let clean_package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(clean.as_bytes()))))
            .unwrap();
    let clean_source = clean_package.section_layout_snapshot().unwrap();
    let mut clean_edit = clean_source.edit(0).unwrap();
    clean_edit.clear_page_size().unwrap();
    let clean_commit = clean_edit.commit().unwrap();
    assert!(
        clean_commit.snapshot().inventory().sections()[0]
            .page_size()
            .is_none()
    );
}

#[test]
fn section_layout_alias_foreign_children_and_column_comments_are_preserved() {
    let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    let document = format!(
        r#"<s:document xmlns:s="{strict}" xmlns:f="urn:foreign" xmlns:x="urn:opaque"><s:body><s:p><s:pPr><s:sectPr><f:pgSz f:w="foreign"/><s:pgSz s:w="12240" s:h="15840" x:opaque="keep"/><!-- preserve between known children --><s:pgMar s:left="720"/><s:cols><!-- preserve inside columns --><s:col s:w="240"/></s:cols></s:sectPr></s:pPr></s:p></s:body></s:document>"#
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        document.as_bytes(),
    ))))
    .unwrap();
    let source = package.section_layout_snapshot().unwrap();
    let mut edit = source.edit(0).unwrap();
    edit.set_margins(Some(Margins {
        top: Some(Emu::from_twips(720)),
        right: None,
        bottom: None,
        left: Some(Emu::from_twips(1080)),
        header: None,
        footer: None,
        gutter: None,
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_section_layout_commit_to_stream(&mut output, &commit)
        .unwrap();
    let reopened = OpcPackage::from_reader(io::Cursor::new(output)).unwrap();
    let main = reopened
        .get_part(&PackURI::new(format!("/{MAIN}")).unwrap())
        .unwrap()
        .blob();
    let main = std::str::from_utf8(main).unwrap();
    assert!(main.contains("<f:pgSz f:w=\"foreign\"/>"));
    assert!(main.contains("x:opaque=\"keep\""));
    assert!(main.contains("<!-- preserve between known children -->"));
    assert!(main.contains("<!-- preserve inside columns -->"));
    assert!(main.contains("<s:col s:w=\"240\"/>"));
}

#[test]
fn section_layout_namespace_shadowing_restores_section_root_alias() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:a="{W}"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840" xmlns:a="urn:foreign"/><a:pgMar a:left="720" a:right="720"/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        document.as_bytes(),
    ))))
    .unwrap();
    let source = package.section_layout_snapshot().unwrap();
    assert_eq!(
        source.inventory().sections()[0].margins().unwrap().left,
        Some(Emu::from_twips(720))
    );

    let mut edit = source.edit(0).unwrap();
    edit.set_margins(Some(Margins {
        top: None,
        right: Some(Emu::from_twips(720)),
        bottom: None,
        left: Some(Emu::from_twips(1080)),
        header: None,
        footer: None,
        gutter: None,
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_section_layout_commit_to_stream(&mut output, &commit)
        .unwrap();
    let reopened = OpcPackage::from_reader(io::Cursor::new(output)).unwrap();
    let main = reopened
        .get_part(&PackURI::new(format!("/{MAIN}")).unwrap())
        .unwrap()
        .blob();
    let main = std::str::from_utf8(main).unwrap();
    assert!(main.contains("xmlns:a=\"urn:foreign\""));
    assert!(main.contains("<a:pgMar a:left=\"1080\""));
}

#[test]
fn section_layout_default_word_namespace_rebinds_inserted_qnames() {
    for namespace in [W, "http://purl.oclc.org/ooxml/wordprocessingml/main"] {
        let document = format!(
            r#"<document xmlns="{namespace}"><body><p><pPr><sectPr/></pPr></p></body></document>"#
        );
        let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
            document.as_bytes(),
        ))))
        .unwrap();
        let source = package.section_layout_snapshot().unwrap();
        let mut edit = source.edit(0).unwrap();
        edit.set_page_size(Some(PageSize {
            width: Some(Emu::from_twips(11906)),
            height: Some(Emu::from_twips(16838)),
            orientation: Orientation::Portrait,
        }))
        .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            commit.snapshot().inventory().sections().len(),
            source.inventory().sections().len()
        );
        assert_eq!(
            commit.snapshot().inventory().sections()[0].ownership(),
            source.inventory().sections()[0].ownership()
        );
        assert_eq!(
            commit.snapshot().inventory().sections()[0]
                .page_size()
                .unwrap()
                .width,
            Some(Emu::from_twips(11906))
        );

        let mut output = Vec::new();
        package
            .publish_section_layout_commit_to_stream(&mut output, &commit)
            .unwrap();
        let reopened_opc = OpcPackage::from_reader(io::Cursor::new(output.clone())).unwrap();
        let main = reopened_opc
            .get_part(&PackURI::new(format!("/{MAIN}")).unwrap())
            .unwrap()
            .blob();
        let main = std::str::from_utf8(main).unwrap();
        assert!(!main.contains("<:"));
        assert!(!main.contains("</:"));

        let reopened =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
        assert_eq!(
            reopened
                .section_layout_snapshot()
                .unwrap()
                .inventory()
                .sections()[0]
                .page_size()
                .unwrap()
                .height,
            Some(Emu::from_twips(16838))
        );
    }
}

#[test]
fn section_layout_missing_children_precede_standard_barriers_losslessly() {
    let cases = [
        (
            "pgSz",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign"><w:body><w:p><w:pPr><w:sectPr><f:foreign f:keep="yes"/><!-- keep this comment --><w:pgBorders/><w:docGrid/><w:sectPrChange/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
            ),
            "pgBorders",
        ),
        (
            "pgMar",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><f:foreign f:keep="yes"/><!-- keep this comment --><w:docGrid/><w:sectPrChange/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
            ),
            "docGrid",
        ),
        (
            "cols",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="720"/><f:foreign f:keep="yes"/><!-- keep this comment --><w:sectPrChange/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
            ),
            "sectPrChange",
        ),
    ];

    for (field, document, barrier) in cases {
        let source_bytes = fixture(document.as_bytes());
        let package =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source_bytes.clone())))
                .unwrap();
        let source = package.section_layout_snapshot().unwrap();
        let mut edit = source.edit(0).unwrap();
        match field {
            "pgSz" => {
                edit.set_page_size(Some(PageSize {
                    width: Some(Emu::from_twips(11906)),
                    height: Some(Emu::from_twips(16838)),
                    orientation: Orientation::Landscape,
                }))
                .unwrap();
            },
            "pgMar" => {
                edit.set_margins(Some(Margins {
                    top: Some(Emu::from_twips(720)),
                    right: Some(Emu::from_twips(720)),
                    bottom: Some(Emu::from_twips(720)),
                    left: Some(Emu::from_twips(1080)),
                    header: None,
                    footer: None,
                    gutter: None,
                }))
                .unwrap();
            },
            "cols" => {
                edit.set_columns(Some(Columns {
                    equal_width: true,
                    count: 2,
                    space: Some(Emu::from_twips(240)),
                    separator: true,
                    columns: Vec::new(),
                }))
                .unwrap();
            },
            _ => unreachable!("unknown section-layout field"),
        }
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        package
            .publish_section_layout_commit_to_stream(&mut output, &commit)
            .unwrap();

        let main = main_xml(&output);
        let inserted = format!("<w:{field}");
        assert!(
            main.find(&inserted).unwrap() < main.find(&format!("<w:{barrier}")).unwrap(),
            "{field} was not inserted before {barrier}: {main}"
        );
        assert!(main.contains(r#"<f:foreign f:keep="yes"/>"#));
        assert!(main.contains("<!-- keep this comment -->"));
        assert_eq!(
            &output[payload_range(&output, HEADER)],
            &source_bytes[payload_range(&source_bytes, HEADER)]
        );
        assert_eq!(
            &output[payload_range(&output, UNUSED)],
            &source_bytes[payload_range(&source_bytes, UNUSED)]
        );

        let reopened =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
        let reopened_snapshot = reopened.section_layout_snapshot().unwrap();
        let section = &reopened_snapshot.inventory().sections()[0];
        match field {
            "pgSz" => assert_eq!(
                section.page_size().unwrap().orientation,
                Orientation::Landscape
            ),
            "pgMar" => assert_eq!(section.margins().unwrap().left, Some(Emu::from_twips(1080))),
            "cols" => assert_eq!(section.columns().unwrap().count, 2),
            _ => unreachable!("unknown section-layout field"),
        }
    }
}

#[test]
fn section_layout_local_prefix_shadowing_preserves_foreign_markup_and_reopens() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:x="{W}" xmlns:f="urn:foreign"><w:body><w:p><w:pPr><w:sectPr><x:pgSz xmlns:w="urn:page-shadow" f:marker="page" x:w="12240"/><x:pgMar xmlns:w="urn:margins-shadow" f:marker="margins" x:left="720"/><x:cols xmlns:w="urn:columns-shadow" f:marker="columns" x:num="1"><x:col xmlns:w="urn:column-shadow" f:marker="column" x:w="600"/></x:cols></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let source_bytes = fixture(document.as_bytes());
    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source_bytes.clone())))
            .unwrap();
    let source = package.section_layout_snapshot().unwrap();
    let mut edit = source.edit(0).unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Landscape,
    }))
    .unwrap()
    .set_margins(Some(Margins {
        top: Some(Emu::from_twips(720)),
        right: Some(Emu::from_twips(1080)),
        bottom: Some(Emu::from_twips(720)),
        left: Some(Emu::from_twips(1080)),
        header: Some(Emu::from_twips(360)),
        footer: Some(Emu::from_twips(360)),
        gutter: Some(Emu::from_twips(120)),
    }))
    .unwrap()
    .set_columns(Some(Columns {
        equal_width: false,
        count: 1,
        space: Some(Emu::from_twips(240)),
        separator: true,
        columns: vec![Column {
            width: Emu::from_twips(900),
            space: Some(Emu::from_twips(120)),
        }],
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_section_layout_commit_to_stream(&mut output, &commit)
        .unwrap();

    let main = main_xml(&output);
    for shadow in [
        r#"xmlns:w="urn:page-shadow""#,
        r#"xmlns:w="urn:margins-shadow""#,
        r#"xmlns:w="urn:columns-shadow""#,
        r#"xmlns:w="urn:column-shadow""#,
    ] {
        assert!(
            main.contains(shadow),
            "missing preserved declaration {shadow}"
        );
    }
    for marker in [
        r#"f:marker="page""#,
        r#"f:marker="margins""#,
        r#"f:marker="columns""#,
        r#"f:marker="column""#,
    ] {
        assert!(
            main.contains(marker),
            "missing preserved foreign attribute {marker}"
        );
    }
    for modeled in [
        r#"x:h="16838""#,
        r#"x:orient="landscape""#,
        r#"x:top="720""#,
        r#"x:right="1080""#,
        r#"x:bottom="720""#,
        r#"x:left="1080""#,
        r#"x:header="360""#,
        r#"x:footer="360""#,
        r#"x:gutter="120""#,
        r#"x:equalWidth="0""#,
        r#"x:space="240""#,
        r#"x:sep="1""#,
        r#"x:w="900""#,
        r#"x:space="120""#,
    ] {
        assert!(
            main.contains(modeled),
            "missing modeled attribute in section Word namespace: {modeled}"
        );
    }
    assert!(!main.contains(r#"w:h="16838""#));
    assert!(!main.contains(r#"w:top="720""#));

    let reopened =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    let reopened_snapshot = reopened.section_layout_snapshot().unwrap();
    let section = &reopened_snapshot.inventory().sections()[0];
    assert_eq!(
        section.page_size().unwrap().height,
        Some(Emu::from_twips(16838))
    );
    assert_eq!(section.margins().unwrap().left, Some(Emu::from_twips(1080)));
    assert_eq!(
        section.columns().unwrap(),
        Columns {
            equal_width: false,
            count: 1,
            space: Some(Emu::from_twips(240)),
            separator: true,
            columns: vec![Column {
                width: Emu::from_twips(900),
                space: Some(Emu::from_twips(120)),
            }],
        }
    );
}

#[test]
fn section_layout_missing_field_with_unknown_word_child_is_typed_unsafe_edit() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:sectPr><w:unknown w:opaque="keep"/><w:docGrid/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        document.as_bytes(),
    ))))
    .unwrap();
    let source = package.section_layout_snapshot().unwrap();
    let mut edit = source.edit(0).unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Portrait,
    }))
    .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            ..
        })
    ));
}

#[test]
fn section_layout_complex_noop_remains_byte_exact() {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign"><w:body><w:p><w:pPr><w:sectPr><f:foreign f:keep="yes"/><!-- keep this comment --><w:pgBorders/><w:docGrid/><w:sectPrChange/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let source_bytes = fixture(document.as_bytes());
    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source_bytes.clone())))
            .unwrap();
    let noop = package
        .section_layout_snapshot()
        .unwrap()
        .edit(0)
        .unwrap()
        .commit()
        .unwrap();
    assert!(noop.patch().is_noop());
    let mut output = Vec::new();
    package
        .publish_section_layout_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, source_bytes);
}

#[test]
fn section_layout_signed_noop_changed_and_partial_sink_are_typed() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let signed_bytes = signed_section_fixture(document.as_bytes());
    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(signed_bytes.clone())))
            .unwrap();
    let snapshot = package.section_layout_snapshot().unwrap();
    let noop = snapshot.edit(0).unwrap().commit().unwrap();
    let mut noop_output = Vec::new();
    package
        .publish_section_layout_commit_to_stream(&mut noop_output, &noop)
        .unwrap();
    assert_eq!(noop_output, signed_bytes);

    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(signed_bytes.clone())))
            .unwrap();
    let mut edit = package.section_layout_snapshot().unwrap().edit(0).unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Landscape,
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut changed_output = Vec::new();
    assert!(matches!(
        package.publish_section_layout_commit_to_stream(&mut changed_output, &commit),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(changed_output.is_empty());

    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        document.as_bytes(),
    ))))
    .unwrap();
    let mut edit = package.section_layout_snapshot().unwrap().edit(0).unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Landscape,
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = SectionFailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        package.publish_section_layout_commit_to_stream(&mut sink, &commit),
        Err(Error::Opc(litchi_opc::OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}

#[test]
fn section_layout_detached_patch_cannot_publish_to_source_package() {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p></w:body></w:document>"#
    );
    let detached =
        litchi_docx::section::layout::Snapshot::from_xml(document.as_bytes().to_vec()).unwrap();
    let mut edit = detached.edit(0).unwrap();
    edit.set_page_size(Some(PageSize {
        width: Some(Emu::from_twips(11906)),
        height: Some(Emu::from_twips(16838)),
        orientation: Orientation::Landscape,
    }))
    .unwrap();
    let commit = edit.commit().unwrap();
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture(
        document.as_bytes(),
    ))))
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_section_layout_commit_to_stream(&mut output, &commit),
        Err(Error::SectionLayoutAuthorizationConflict)
    ));
    assert!(output.is_empty());
}

#[test]
fn section_inventory_many_sections_stays_within_primary_event_budget() {
    let mut document = format!(r#"<w:document xmlns:w="{W}"><w:body>"#);
    for _ in 0..128 {
        document.push_str(r#"<w:p><w:pPr><w:sectPr/></w:pPr></w:p>"#);
    }
    document.push_str("</w:body></w:document>");
    let limits = Limits {
        max_events: 64,
        ..Limits::default()
    };
    assert!(matches!(
        litchi_docx::section::Inventory::parse_with_limits(document.as_bytes(), &limits),
        Err(Error::SectionInventoryLimit {
            resource: "XML events",
            ..
        })
    ));
}

#[test]
fn section_inventory_many_sections_consumes_one_primary_event_budget() {
    const SECTION_COUNT: usize = 128;
    let mut document = format!(r#"<w:document xmlns:w="{W}"><w:body>"#);
    for _ in 0..SECTION_COUNT {
        document.push_str(r#"<w:p><w:pPr><w:sectPr/></w:pPr></w:p>"#);
    }
    document.push_str("</w:body></w:document>");
    let limits = Limits {
        max_events: 5 + 7 * SECTION_COUNT,
        ..Limits::default()
    };
    let inventory =
        litchi_docx::section::Inventory::parse_with_limits(document.as_bytes(), &limits).unwrap();
    assert_eq!(inventory.sections().len(), SECTION_COUNT + 1);
}
