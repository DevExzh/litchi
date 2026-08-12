use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_docx::section::{Limits, Ownership, Property, PropertyValue};
use litchi_docx::{Error, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

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
    hostile.arm_change();
    assert!(package.section_inventory_snapshot().is_err());
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
