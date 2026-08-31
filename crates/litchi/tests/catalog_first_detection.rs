#![cfg(feature = "docx")]

use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};
use std::ops::Range;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};

use litchi::common::detection::{detect_file_format_from_bytes, detect_format_from_reader};
use litchi::common::{FileFormat, ReadAt, SourceVersion};
use litchi::detection_smart::{DetectedFormat, detect_format_smart};

const DOCX_MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const SENTINEL: &[u8] = b"ORDINARY_PAYLOAD_SENTINEL_0344_";

const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
const PACKAGE_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
const DOCUMENT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/0344" TargetMode="External"/></Relationships>"#;

struct Fixture {
    bytes: Vec<u8>,
    ordinary_payload: Range<usize>,
}

#[derive(Clone)]
struct ReadGuard {
    bytes: Arc<Vec<u8>>,
    forbidden: Range<usize>,
    reads: Arc<Mutex<Vec<Range<usize>>>>,
    violations: Arc<AtomicUsize>,
}

impl ReadGuard {
    fn new(bytes: Vec<u8>, forbidden: Range<usize>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            forbidden,
            reads: Arc::new(Mutex::new(Vec::new())),
            violations: Arc::new(AtomicUsize::new(0)),
        }
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        let range = start..start + count;
        if ranges_overlap(&range, &self.forbidden) {
            self.violations.fetch_add(1, Ordering::Relaxed);
            return Err(io::Error::other(
                "ordinary payload range read during catalog detection",
            ));
        }
        output[..count].copy_from_slice(&self.bytes[range.clone()]);
        self.reads
            .lock()
            .expect("read log is not poisoned")
            .push(range);
        Ok(count)
    }

    fn assert_no_forbidden_reads(&self) {
        assert_eq!(
            self.violations.load(Ordering::Relaxed),
            0,
            "catalog detection attempted an ordinary payload read"
        );
        let reads = self.reads.lock().expect("read log is not poisoned");
        assert!(
            reads
                .iter()
                .all(|range| !ranges_overlap(range, &self.forbidden)),
            "ordinary payload range appeared in the read log: {reads:?}"
        );
    }
}

struct GuardedReader {
    guard: ReadGuard,
    position: u64,
}

impl GuardedReader {
    fn new(bytes: Vec<u8>, forbidden: Range<usize>) -> Self {
        Self {
            guard: ReadGuard::new(bytes, forbidden),
            position: 0,
        }
    }

    fn guard(&self) -> &ReadGuard {
        &self.guard
    }
}

impl Read for GuardedReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let read = self.guard.read_at(self.position, output)?;
        self.position =
            self.position
                .checked_add(u64::try_from(read).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidData, "read count overflow")
                })?)
                .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "position overflow"))?;
        Ok(read)
    }
}

impl Seek for GuardedReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let target = match position {
            SeekFrom::Start(offset) => i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
            SeekFrom::End(offset) => {
                i128::try_from(self.guard.bytes.len())
                    .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "length overflow"))?
                    + i128::from(offset)
            },
        };
        let target = u64::try_from(target)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "invalid seek position"))?;
        self.position = target;
        Ok(target)
    }
}

#[derive(Clone)]
struct GuardedSource {
    guard: ReadGuard,
}

impl GuardedSource {
    fn new(bytes: Vec<u8>, forbidden: Range<usize>) -> Self {
        Self {
            guard: ReadGuard::new(bytes, forbidden),
        }
    }

    fn guard(&self) -> &ReadGuard {
        &self.guard
    }
}

impl ReadAt for GuardedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.guard.bytes.len())
            .map_err(|_| io::Error::other("source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.guard.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x0344_5445_5354, 0))
    }
}

fn ranges_overlap(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

fn zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for &(name, data) in entries {
        writer
            .start_file(name, options)
            .expect("ZIP member name is valid");
        writer.write_all(data).expect("ZIP member is writable");
    }
    writer.finish().expect("ZIP archive is writable");
    output.into_inner()
}

fn stored_member_payload(bytes: &[u8], member: &[u8]) -> Range<usize> {
    let mut offset = 0;
    while offset + 30 <= bytes.len() {
        if &bytes[offset..offset + 4] != b"PK\x03\x04" {
            offset += 1;
            continue;
        }
        let name_len = usize::from(u16::from_le_bytes(
            bytes[offset + 26..offset + 28]
                .try_into()
                .expect("ZIP name length is present"),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            bytes[offset + 28..offset + 30]
                .try_into()
                .expect("ZIP extra length is present"),
        ));
        let compressed_len = usize::try_from(u32::from_le_bytes(
            bytes[offset + 18..offset + 22]
                .try_into()
                .expect("ZIP compressed length is present"),
        ))
        .expect("ZIP compressed length fits usize");
        let name_start = offset + 30;
        let data_start = name_start + name_len + extra_len;
        if &bytes[name_start..data_start - extra_len] == member {
            return data_start..data_start + compressed_len;
        }
        offset = data_start + compressed_len;
    }
    panic!("stored ZIP member not found: {:?}", member);
}

fn docx_fixture() -> Fixture {
    let mut document = Vec::new();
    document.extend_from_slice(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>"#,
    );
    for _ in 0..32_768 {
        document.extend_from_slice(SENTINEL);
    }
    document.extend_from_slice(br#"</w:t></w:r></w:p></w:body></w:document>"#);

    let bytes = zip_bytes(&[
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", PACKAGE_RELATIONSHIPS),
        ("word/_rels/document.xml.rels", DOCUMENT_RELATIONSHIPS),
        ("word/document.xml", document.as_slice()),
    ]);
    let ordinary_payload = stored_member_payload(&bytes, b"word/document.xml");
    Fixture {
        bytes,
        ordinary_payload,
    }
}

fn malformed_opc_package() -> Vec<u8> {
    zip_bytes(&[
        ("_rels/.rels", PACKAGE_RELATIONSHIPS),
        ("word/document.xml", b"not admitted without content types"),
    ])
}

fn input_limit_below(fixture: &Fixture) -> litchi::opc::ReadLimits {
    litchi::opc::ReadLimits::builder()
        .max_input_bytes(u64::try_from(fixture.bytes.len() - 1).expect("fixture length fits u64"))
        .expect("positive input limit")
        .build()
        .expect("input limit is consistent")
}

#[test]
fn catalog_first_detection_matches_bytes_reader_and_typed_docx_without_payload_reads() {
    let fixture = docx_fixture();

    assert_eq!(
        detect_file_format_from_bytes(&fixture.bytes),
        Some(FileFormat::Docx)
    );

    let mut reader = GuardedReader::new(fixture.bytes.clone(), fixture.ordinary_payload.clone());
    reader
        .seek(SeekFrom::Start(41))
        .expect("initial reader position is valid");
    assert_eq!(
        detect_format_from_reader(&mut reader),
        Some(FileFormat::Docx)
    );
    assert_eq!(reader.stream_position().expect("position is readable"), 41);
    reader.guard().assert_no_forbidden_reads();

    let mut catalog_reader =
        GuardedReader::new(fixture.bytes.clone(), fixture.ordinary_payload.clone());
    catalog_reader
        .seek(SeekFrom::Start(17))
        .expect("initial reader position is valid");
    let catalog = litchi::opc::probe_package_catalog_from_reader(&mut catalog_reader)
        .expect("valid OPC catalog should be admitted");
    assert_eq!(catalog.part_count(), 1);
    assert!(
        catalog
            .part_content_types()
            .any(|content_type| content_type == DOCX_MAIN_CONTENT_TYPE)
    );
    assert_eq!(
        catalog_reader
            .stream_position()
            .expect("position is readable"),
        17
    );
    catalog_reader.guard().assert_no_forbidden_reads();

    let source = GuardedSource::new(fixture.bytes.clone(), fixture.ordinary_payload.clone());
    let source_for_open: Arc<dyn ReadAt> = Arc::new(source.clone());
    let typed = litchi::docx::source_backed::Package::from_read_at(source_for_open)
        .expect("typed DOCX source open should use the catalog only");
    assert_eq!(
        detect_file_format_from_bytes(&fixture.bytes),
        Some(FileFormat::Docx),
        "the typed source owner and both detector APIs identify DOCX"
    );
    let _ = typed;
    source.guard().assert_no_forbidden_reads();
}

#[test]
fn typed_catalog_and_owner_restore_reader_position_on_error_and_limit() {
    let fixture = docx_fixture();
    let malformed = malformed_opc_package();

    let mut malformed_reader = GuardedReader::new(malformed.clone(), 0..0);
    malformed_reader
        .seek(SeekFrom::Start(23))
        .expect("initial reader position is valid");
    let malformed_error = litchi::opc::probe_package_catalog_from_reader(&mut malformed_reader)
        .expect_err("malformed OPC catalog must return a typed error");
    assert!(matches!(
        malformed_error,
        litchi::opc::OpcError::PartNotFound(_)
    ));
    assert_eq!(
        malformed_reader
            .stream_position()
            .expect("position is readable"),
        23
    );

    let mut malformed_detection_reader = GuardedReader::new(malformed.clone(), 0..0);
    malformed_detection_reader
        .seek(SeekFrom::Start(31))
        .expect("initial reader position is valid");
    assert_eq!(
        detect_format_from_reader(&mut malformed_detection_reader),
        None
    );
    assert_eq!(
        malformed_detection_reader
            .stream_position()
            .expect("position is readable"),
        31
    );

    let mut limited_reader =
        GuardedReader::new(fixture.bytes.clone(), fixture.ordinary_payload.clone());
    limited_reader
        .seek(SeekFrom::Start(29))
        .expect("initial reader position is valid");
    let limited_error = litchi::opc::probe_package_catalog_from_reader_with_limits(
        &mut limited_reader,
        input_limit_below(&fixture),
    )
    .expect_err("input limit must be reported by the typed catalog probe");
    assert!(matches!(
        limited_error,
        litchi::opc::OpcError::ReadLimit {
            resource: litchi::opc::ReadResource::InputBytes,
            ..
        }
    ));
    assert_eq!(
        limited_reader
            .stream_position()
            .expect("position is readable"),
        29
    );
    limited_reader.guard().assert_no_forbidden_reads();

    let source = GuardedSource::new(fixture.bytes.clone(), fixture.ordinary_payload.clone());
    let owner_error = match litchi::docx::source_backed::Package::from_read_at_with_limits(
        Arc::new(source.clone()),
        input_limit_below(&fixture),
    ) {
        Ok(_) => panic!("typed DOCX owner unexpectedly accepted the input limit"),
        Err(error) => error,
    };
    assert!(matches!(
        owner_error,
        litchi::docx::Error::Opc(litchi::opc::OpcError::ReadLimit {
            resource: litchi::opc::ReadResource::InputBytes,
            ..
        })
    ));
    source.guard().assert_no_forbidden_reads();
}

#[test]
fn typed_malformed_package_error_is_distinct_from_fresh_no_match() {
    let malformed = malformed_opc_package();
    let fresh_no_match = b"this is not an Office package".to_vec();

    assert_eq!(detect_file_format_from_bytes(&malformed), None);
    assert_eq!(detect_file_format_from_bytes(&fresh_no_match), None);

    let malformed_source = GuardedSource::new(malformed, 0..0);
    let malformed_error =
        match litchi::docx::source_backed::Package::from_read_at(Arc::new(malformed_source)) {
            Ok(_) => panic!("malformed OPC unexpectedly opened as DOCX"),
            Err(error) => error,
        };
    assert!(matches!(
        malformed_error,
        litchi::docx::Error::Opc(litchi::opc::OpcError::PartNotFound(_))
    ));

    let fresh_source = GuardedSource::new(fresh_no_match, 0..0);
    let fresh_error =
        match litchi::docx::source_backed::Package::from_read_at(Arc::new(fresh_source)) {
            Ok(_) => panic!("fresh non-package input unexpectedly opened as DOCX"),
            Err(error) => error,
        };
    assert!(matches!(
        fresh_error,
        litchi::docx::Error::Opc(litchi::opc::OpcError::ZipError(_))
    ));
}

#[test]
fn smart_detection_still_returns_a_usable_eager_docx_variant() {
    let fixture = docx_fixture();

    match detect_format_smart(fixture.bytes) {
        Some(DetectedFormat::Docx(package)) => {
            let main = package
                .main_document_part()
                .expect("eager DOCX variant remains usable");
            assert_eq!(main.content_type(), DOCX_MAIN_CONTENT_TYPE);
        },
        #[cfg(any(
            feature = "doc",
            feature = "ppt",
            feature = "pptx",
            feature = "xls",
            feature = "xlsx",
            feature = "xlsb",
            feature = "pages",
            feature = "keynote",
            feature = "numbers",
            feature = "odt",
            feature = "ods",
            feature = "odp",
            feature = "rtf"
        ))]
        Some(other) => panic!("expected eager DOCX variant, got {other:?}"),
        None => panic!("valid DOCX fixture was not smart-detected"),
    }
}
