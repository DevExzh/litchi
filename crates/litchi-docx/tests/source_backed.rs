use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_docx::{Error, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter};

const MAIN_DOCUMENT: &str = "word/document.xml";
const UNUSED_PART: &str = "word/000-unused.bin";
const TRAILING_PART: &str = "word/zzz-trailing.bin";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

struct ObservedSource {
    bytes: Vec<u8>,
    main_payload: std::ops::Range<usize>,
    unused_payload: std::ops::Range<usize>,
    main_reads: AtomicUsize,
    unused_reads: AtomicUsize,
    revision: AtomicU64,
}

impl ObservedSource {
    fn new(
        bytes: Vec<u8>,
        main_payload: std::ops::Range<usize>,
        unused_payload: std::ops::Range<usize>,
    ) -> Self {
        Self {
            bytes,
            main_payload,
            unused_payload,
            main_reads: AtomicUsize::new(0),
            unused_reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for ObservedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let requested = offset..end;
        if overlaps(&requested, &self.main_payload) {
            self.main_reads.fetch_add(1, Ordering::SeqCst);
        }
        if overlaps(&requested, &self.unused_payload) {
            self.unused_reads.fetch_add(1, Ordering::SeqCst);
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(17, self.revision.load(Ordering::SeqCst)))
    }
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
        let compressed_size =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_length =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        let extra_length =
            u16::from_le_bytes(zip[offset + 30..offset + 32].try_into().unwrap()) as usize;
        if offset + 46 + name_length + extra_length > zip.len() {
            continue;
        }
        if &zip[offset + 46..offset + 46 + name_length] == name {
            let local_offset =
                u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
            let local_name_length = u16::from_le_bytes(
                zip[local_offset + 26..local_offset + 28]
                    .try_into()
                    .unwrap(),
            ) as usize;
            let local_extra_length = u16::from_le_bytes(
                zip[local_offset + 28..local_offset + 30]
                    .try_into()
                    .unwrap(),
            ) as usize;
            let data_start = local_offset + 30 + local_name_length + local_extra_length;
            assert!(data_start + compressed_size <= zip.len());
            return data_start..data_start + compressed_size;
        }
    }
    panic!("ZIP member {name:?} was not found")
}

fn incompressible_bytes() -> Vec<u8> {
    let mut state = 0x5EED_C0DE_u32;
    (0..1_048_576)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            (state >> 24) as u8
        })
        .collect()
}

fn fixture() -> (Vec<u8>, std::ops::Range<usize>, std::ops::Range<usize>) {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{MAIN_DOCUMENT}")).unwrap(),
            ct::WML_DOCUMENT_MAIN.to_string(),
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p></w:body></w:document>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED_PART}")).unwrap(),
            "application/octet-stream".to_string(),
            incompressible_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{TRAILING_PART}")).unwrap(),
            "application/octet-stream".to_string(),
            incompressible_bytes(),
        )))
        .unwrap();
    package.relate_to(MAIN_DOCUMENT, rt::OFFICE_DOCUMENT);
    let zip = PackageWriter::to_bytes(&package).unwrap();
    let main = payload_range(&zip, MAIN_DOCUMENT);
    let unused = payload_range(&zip, UNUSED_PART);
    (zip, main, unused)
}

#[test]
fn opening_stays_cold_and_document_query_loads_only_main_document() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));

    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).expect("open source-backed DOCX");
    assert_eq!(source.main_reads.load(Ordering::SeqCst), 0);
    assert_eq!(source.unused_reads.load(Ordering::SeqCst), 0);

    let document = package.document().expect("load main document");
    assert!(source.main_reads.load(Ordering::SeqCst) > 0);
    assert_eq!(source.unused_reads.load(Ordering::SeqCst), 0);
    assert_eq!(document.extract_text().unwrap(), "alphabeta");
    assert_eq!(document.paragraph_count().unwrap(), 2);
    assert_eq!(
        document.paragraph(1).unwrap().unwrap().text().unwrap(),
        "beta"
    );
    assert!(document.paragraph(2).unwrap().is_none());
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(paragraphs[0].text().unwrap(), "alpha");
    assert_eq!(paragraphs[1].text().unwrap(), "beta");
}

#[test]
fn source_changes_remain_visible_before_the_first_payload_read() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = source_backed::Package::from_read_at(read_at).expect("open source-backed DOCX");

    source.changed();
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
}

#[test]
fn read_limits_are_returned_as_the_original_opc_error() {
    let (bytes, main, unused) = fixture();
    let source = Arc::new(ObservedSource::new(bytes, main, unused));
    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let read_at: Arc<dyn ReadAt> = source;

    assert!(matches!(
        source_backed::Package::from_read_at_with_limits(read_at, limits),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));
}
