use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_docx::section::{Limits, Ownership, Property, PropertyValue};
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
