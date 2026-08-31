#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused probe assertions intentionally fail the test on fixture errors"
)]

//! Reader-level tests for the metadata-only OPC package catalog probe.

use std::io::{self, Cursor, Read, Seek, SeekFrom};
use std::ops::Range;

use litchi_opc::{
    OpcError, OpcPackage, ReadLimits, ReadResource, SourceBackedPackage,
    probe_package_catalog_from_reader, probe_package_catalog_from_reader_with_limits,
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DATA_REL: &str = "urn:litchi:test/data";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DATA_CONTENT_TYPE: &str = "application/octet-stream";
const DOCUMENT_PAYLOAD: &[u8] = b"deflated ordinary document payload sentinel";
const DATA_PAYLOAD: &[u8] = b"stored ordinary binary payload sentinel";
const ORPHAN_PAYLOAD: &[u8] = b"untyped unknown catalog member sentinel";

struct Fixture {
    bytes: Vec<u8>,
    content_types_bytes: usize,
    document_relationships_bytes: usize,
    structural_ranges: Vec<Range<usize>>,
    ordinary_ranges: Vec<Range<usize>>,
    central_directory: Range<usize>,
}

fn fixture() -> Fixture {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="{DOCUMENT_CONTENT_TYPE}"/><Override PartName="/custom/data.bin" ContentType="{DATA_CONTENT_TYPE}"/></Types>"#
    );
    let package_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let document_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rData" Type="{DATA_REL}" Target="../custom/data.bin"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", package_relationships.as_bytes())
        .unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            document_relationships.as_bytes(),
        )
        .unwrap();
    writer
        .write_deflated_sized("word/document.xml", DOCUMENT_PAYLOAD)
        .unwrap();
    writer
        .write_stored("custom/data.bin", DATA_PAYLOAD)
        .unwrap();
    writer
        .write_stored("custom/orphan.bin", ORPHAN_PAYLOAD)
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let structural_names = [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
    ];
    let ordinary_names = ["word/document.xml", "custom/data.bin", "custom/orphan.bin"];
    let structural_ranges = structural_names
        .into_iter()
        .map(|name| local_payload_range(&bytes, name))
        .collect();
    let ordinary_ranges = ordinary_names
        .into_iter()
        .map(|name| local_payload_range(&bytes, name))
        .collect();

    Fixture {
        content_types_bytes: content_types.len(),
        document_relationships_bytes: document_relationships.len(),
        central_directory: central_directory_range(&bytes),
        bytes,
        structural_ranges,
        ordinary_ranges,
    }
}

fn read_u16(bytes: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes(
        bytes[offset..offset + 2].try_into().unwrap(),
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> usize {
    usize::try_from(u32::from_le_bytes(
        bytes[offset..offset + 4].try_into().unwrap(),
    ))
    .unwrap()
}

fn local_payload_range(bytes: &[u8], wanted: &str) -> Range<usize> {
    let mut offset = 0;
    while offset + 30 <= bytes.len() {
        if &bytes[offset..offset + 4] != b"PK\x03\x04" {
            offset += 1;
            continue;
        }
        let name_len = read_u16(bytes, offset + 26);
        let extra_len = read_u16(bytes, offset + 28);
        let compressed_len = read_u32(bytes, offset + 18);
        let name_start = offset + 30;
        let data_start = name_start + name_len + extra_len;
        let data_end = data_start + compressed_len;
        assert!(data_end <= bytes.len(), "local ZIP member exceeds fixture");
        if bytes[name_start..data_start] == *wanted.as_bytes() {
            return data_start..data_end;
        }
        offset = data_end;
    }
    panic!("missing local ZIP member {wanted}");
}

fn central_directory_range(bytes: &[u8]) -> Range<usize> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == b"PK\x05\x06")
        .expect("fixture EOCD");
    let size = read_u32(bytes, eocd + 12);
    let offset = read_u32(bytes, eocd + 16);
    offset..offset + size
}

#[derive(Debug)]
struct ProbeReader {
    inner: Cursor<Vec<u8>>,
    reads: Vec<Range<usize>>,
    fail_reads: bool,
}

impl ProbeReader {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: Cursor::new(bytes),
            reads: Vec::new(),
            fail_reads: false,
        }
    }

    fn failing(bytes: Vec<u8>) -> Self {
        Self {
            fail_reads: true,
            ..Self::new(bytes)
        }
    }

    fn position(&self) -> u64 {
        self.inner.position()
    }

    fn read_ranges(&self) -> &[Range<usize>] {
        &self.reads
    }
}

impl Read for ProbeReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.fail_reads {
            return Err(io::Error::new(
                io::ErrorKind::PermissionDenied,
                "injected probe reader failure",
            ));
        }
        let start = usize::try_from(self.inner.position())
            .map_err(|_| io::Error::other("probe cursor does not fit usize"))?;
        let count = self.inner.read(output)?;
        self.reads.push(start..start + count);
        Ok(count)
    }
}

impl Seek for ProbeReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

fn intersects(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

fn assert_reads_touch(reads: &[Range<usize>], wanted: &Range<usize>, label: &str) {
    assert!(
        reads.iter().any(|read| intersects(read, wanted)),
        "probe never read {label} range {wanted:?}; reads were {reads:?}"
    );
}

fn assert_no_ordinary_payload_reads(fixture: &Fixture, reads: &[Range<usize>]) {
    for read in reads {
        assert!(
            fixture
                .ordinary_ranges
                .iter()
                .all(|ordinary| !intersects(read, ordinary)),
            "probe read ordinary payload range {read:?}; ordinary ranges were {:?}",
            fixture.ordinary_ranges
        );
    }
}

fn catalog_types(catalog: &litchi_opc::PackageCatalog) -> Vec<String> {
    let mut types = catalog
        .part_content_types()
        .map(|content_type| content_type.to_owned())
        .collect::<Vec<_>>();
    types.sort_unstable();
    types
}

fn expected_types() -> Vec<String> {
    let mut types = vec![
        DOCUMENT_CONTENT_TYPE.to_owned(),
        DATA_CONTENT_TYPE.to_owned(),
    ];
    types.sort_unstable();
    types
}

#[test]
fn default_probe_reads_zip_metadata_and_structural_members_only() {
    let fixture = fixture();
    let mut reader = ProbeReader::new(fixture.bytes.clone());
    let original_position = 17;
    reader.seek(SeekFrom::Start(original_position)).unwrap();

    let catalog = probe_package_catalog_from_reader(&mut reader).unwrap();

    assert_eq!(catalog.part_count(), 2);
    assert_eq!(catalog_types(&catalog), expected_types());
    assert_eq!(reader.position(), original_position);

    let reads = reader.read_ranges();
    assert!(!reads.is_empty());
    assert_reads_touch(
        reads,
        &fixture.central_directory,
        "ZIP central-directory metadata",
    );
    for (index, structural) in fixture.structural_ranges.iter().enumerate() {
        assert_reads_touch(reads, structural, &format!("structural member {index}"));
    }
    assert_no_ordinary_payload_reads(&fixture, reads);
}

#[test]
fn limited_probe_matches_eager_and_source_catalog_semantics() {
    let fixture = fixture();
    let limits = ReadLimits::builder()
        .max_input_bytes(fixture.bytes.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let mut reader = ProbeReader::new(fixture.bytes.clone());
    let original_position = 23;
    reader.seek(SeekFrom::Start(original_position)).unwrap();

    let catalog = probe_package_catalog_from_reader_with_limits(&mut reader, limits).unwrap();
    let eager = OpcPackage::from_bytes(&fixture.bytes).unwrap();
    let source = SourceBackedPackage::from_vec(fixture.bytes.clone()).unwrap();

    assert_eq!(catalog.part_count(), eager.part_count());
    assert_eq!(catalog.part_count(), source.iter_parts().count());
    assert_eq!(catalog_types(&catalog), expected_types());
    assert_eq!(catalog_types(&catalog), {
        let mut types = eager
            .iter_parts()
            .map(|part| part.content_type().to_owned())
            .collect::<Vec<_>>();
        types.sort_unstable();
        types
    });
    assert_eq!(catalog_types(&catalog), {
        let mut types = source
            .iter_parts()
            .map(|part| part.content_type().to_owned())
            .collect::<Vec<_>>();
        types.sort_unstable();
        types
    });
    assert_eq!(reader.position(), original_position);
    assert_no_ordinary_payload_reads(&fixture, reader.read_ranges());
}

#[test]
fn probe_limit_errors_are_typed_and_restore_the_cursor() {
    let fixture = fixture();
    let cases = [
        (
            ReadLimits::builder()
                .max_input_bytes((fixture.bytes.len() - 1) as u64)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::InputBytes,
        ),
        (
            ReadLimits::builder()
                .max_content_types_bytes(fixture.content_types_bytes - 1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::ContentTypesBytes,
        ),
        (
            ReadLimits::builder()
                .max_relationship_parts(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipParts,
        ),
        (
            ReadLimits::builder()
                .max_relationship_xml_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipXmlBytes,
        ),
    ];

    for (limits, expected_resource) in cases {
        let mut reader = ProbeReader::new(fixture.bytes.clone());
        let original_position = 29;
        reader.seek(SeekFrom::Start(original_position)).unwrap();
        let error = probe_package_catalog_from_reader_with_limits(&mut reader, limits)
            .expect_err("fixture must exceed the selected limit");

        assert!(matches!(
            error,
            OpcError::ReadLimit { resource, .. } if resource == expected_resource
        ));
        assert_eq!(reader.position(), original_position);
    }

    assert!(fixture.document_relationships_bytes > 1);
}

#[test]
fn malformed_probe_is_an_unknown_catalog_error_and_restores_the_cursor() {
    let fixture = fixture();
    let mut malformed = fixture.bytes;
    malformed.truncate(malformed.len() - 1);
    let mut reader = ProbeReader::new(malformed);
    let original_position = 31;
    reader.seek(SeekFrom::Start(original_position)).unwrap();

    let error = probe_package_catalog_from_reader(&mut reader)
        .expect_err("truncated ZIP must not produce a catalog");

    assert!(matches!(error, OpcError::ZipError(_)));
    assert_eq!(reader.position(), original_position);

    let mut non_zip = ProbeReader::new(b"not an OPC ZIP".to_vec());
    non_zip.seek(SeekFrom::Start(4)).unwrap();
    let error = probe_package_catalog_from_reader(&mut non_zip)
        .expect_err("non-ZIP input must remain unknown");
    assert!(matches!(error, OpcError::ZipError(_)));
    assert_eq!(non_zip.position(), 4);
}

#[test]
fn reader_failure_restores_the_cursor_without_fabricating_a_catalog() {
    let fixture = fixture();
    let original_position = 37;
    let mut reader = ProbeReader::failing(fixture.bytes);
    reader.seek(SeekFrom::Start(original_position)).unwrap();

    let error = probe_package_catalog_from_reader(&mut reader)
        .expect_err("injected reader failure must reject the catalog");

    assert!(matches!(
        error,
        OpcError::ZipError(_) | OpcError::IoError(_)
    ));
    assert_eq!(reader.position(), original_position);
}
