#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "This integration fixture is fixed-size and assertion-driven."
)]

//! Downstream coverage for source-backed ODF publication from a ZIP64 source.
//!
//! The fixture promotes a small ODF archive into a structurally valid ZIP64
//! archive without changing its local records.  Publication must keep every
//! untouched local span and central metadata record, retain the ZIP64 tail and
//! archive comment, and produce a package that can be reopened after writing
//! through a sequential sink.

use litchi_core::{OwnedSource, ReadAt};
use litchi_odf_common::core::SourceBackedPackage;
use soapberry_zip::{PreservationIndex, ZipArchive};
use std::collections::BTreeMap;
use std::io::{self, Cursor, Write};
use std::sync::Arc;
use zip::write::{ExtendedFileOptions, FileOptions, SimpleFileOptions};
use zip::{CompressionMethod as ZipCompressionMethod, ZipWriter};

const MIME: &str = "application/vnd.oasis.opendocument.text";
const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const SOURCE_CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text><office:p>source</office:p></office:text></office:body></office:document-content>"#;
const TARGET_CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text><office:p>target</office:p><office:p>second</office:p></office:text></office:body></office:document-content>"#;
const ARCHIVE_COMMENT: &[u8] = b"source ZIP64 archive comment";
const ZIP64_EXTENSIBLE_DATA: &[u8] = &[0xa1, 0xb2, 0xc3];

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

#[derive(Debug, Default)]
struct ChunkedSink {
    bytes: Vec<u8>,
    maximum_write: usize,
    write_calls: usize,
}

impl ChunkedSink {
    fn new(maximum_write: usize) -> Self {
        assert!(maximum_write > 0);
        Self {
            maximum_write,
            ..Self::default()
        }
    }
}

impl Write for ChunkedSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.write_calls += 1;
        let amount = input.len().min(self.maximum_write);
        self.bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn manifest() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Pictures/blob.bin" manifest:media-type="application/octet-stream"/><manifest:file-entry manifest:full-path="Padding/tail.bin" manifest:media-type="application/octet-stream"/></manifest:manifest>"#
    )
    .into_bytes()
}

fn file_options(method: ZipCompressionMethod) -> FileOptions<'static, ExtendedFileOptions> {
    let mut options = FileOptions::default()
        .compression_method(method)
        .with_file_comment("untouched member comment");
    options
        .add_extra_data(0x1234, b"untouched member extra", false)
        .unwrap();
    options
}

fn opaque_media() -> Vec<u8> {
    let mut state = 0x9e37_79b9_u32;
    (0..32 * 1024)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            state as u8
        })
        .collect()
}

fn zip32_odf_package() -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    writer
        .set_raw_comment(ARCHIVE_COMMENT.to_vec().into_boxed_slice())
        .unwrap();
    writer
        .start_file(
            "mimetype",
            SimpleFileOptions::default().compression_method(ZipCompressionMethod::Stored),
        )
        .unwrap();
    writer.write_all(MIME.as_bytes()).unwrap();
    writer
        .start_file("content.xml", file_options(ZipCompressionMethod::Deflated))
        .unwrap();
    writer.write_all(SOURCE_CONTENT).unwrap();
    writer
        .start_file(
            "Pictures/blob.bin",
            file_options(ZipCompressionMethod::Deflated),
        )
        .unwrap();
    writer.write_all(&opaque_media()).unwrap();
    writer
        .start_file(
            "Padding/tail.bin",
            file_options(ZipCompressionMethod::Stored),
        )
        .unwrap();
    writer
        .write_all(b"opaque tail bytes retained after the changed member")
        .unwrap();
    writer
        .start_file(
            "META-INF/manifest.xml",
            file_options(ZipCompressionMethod::Deflated),
        )
        .unwrap();
    writer.write_all(&manifest()).unwrap();
    writer.finish().unwrap().into_inner()
}

/// Promote ordinary central records and the end records to ZIP64 while preserving
/// the local-member bytes exactly.  The central ZIP64 extra stores the values
/// that were present in the original ZIP32 record, so no large allocation is
/// needed to exercise a real ZIP64 source. The fixed, first `mimetype` member
/// retains its canonical ODF central record without extra fields.
fn promote_to_zip64(mut source: Vec<u8>) -> Vec<u8> {
    let archive = ZipArchive::from_slice(&source).unwrap();
    assert!(!archive.is_zip64());
    let central_start = archive.directory_offset() as usize;
    let old_eocd = archive.eocd_offset() as usize;
    let original_classic_eocd = source[old_eocd..].to_vec();

    let mut central = Vec::new();
    let mut offset = central_start;
    while offset < old_eocd {
        assert_eq!(&source[offset..offset + 4], b"PK\x01\x02");
        let name_len =
            u16::from_le_bytes(source[offset + 28..offset + 30].try_into().unwrap()) as usize;
        let extra_len =
            u16::from_le_bytes(source[offset + 30..offset + 32].try_into().unwrap()) as usize;
        let comment_len =
            u16::from_le_bytes(source[offset + 32..offset + 34].try_into().unwrap()) as usize;
        let record_len = 46 + name_len + extra_len + comment_len;
        let record = &source[offset..offset + record_len];
        let variable_start = 46 + name_len;
        if &record[46..variable_start] == b"mimetype" {
            central.extend_from_slice(record);
            offset += record_len;
            continue;
        }

        let mut promoted = record[..variable_start].to_vec();
        promoted[6..8].copy_from_slice(&45_u16.to_le_bytes());
        let compressed_size = u32::from_le_bytes(record[20..24].try_into().unwrap());
        let uncompressed_size = u32::from_le_bytes(record[24..28].try_into().unwrap());
        let local_header_offset = u32::from_le_bytes(record[42..46].try_into().unwrap());
        promoted[20..24].copy_from_slice(&u32::MAX.to_le_bytes());
        promoted[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
        promoted[42..46].copy_from_slice(&u32::MAX.to_le_bytes());

        let mut zip64_extra = Vec::with_capacity(28);
        zip64_extra.extend_from_slice(&1_u16.to_le_bytes());
        zip64_extra.extend_from_slice(&24_u16.to_le_bytes());
        zip64_extra.extend_from_slice(&u64::from(uncompressed_size).to_le_bytes());
        zip64_extra.extend_from_slice(&u64::from(compressed_size).to_le_bytes());
        zip64_extra.extend_from_slice(&u64::from(local_header_offset).to_le_bytes());
        promoted[30..32].copy_from_slice(
            &(u16::try_from(extra_len + zip64_extra.len()).unwrap()).to_le_bytes(),
        );
        promoted.extend_from_slice(&record[variable_start..variable_start + extra_len]);
        promoted.extend_from_slice(&zip64_extra);
        promoted.extend_from_slice(&record[variable_start + extra_len..]);
        central.extend_from_slice(&promoted);
        offset += record_len;
    }

    source.truncate(central_start);
    source.extend_from_slice(&central);
    let zip64_eocd_offset = source.len();
    let entry_count = u64::from(u16::from_le_bytes(
        original_classic_eocd[8..10].try_into().unwrap(),
    ));

    let mut zip64_eocd = Vec::with_capacity(56 + ZIP64_EXTENSIBLE_DATA.len());
    zip64_eocd.extend_from_slice(&0x0606_4b50_u32.to_le_bytes());
    zip64_eocd.extend_from_slice(&(44_u64 + ZIP64_EXTENSIBLE_DATA.len() as u64).to_le_bytes());
    zip64_eocd.extend_from_slice(&45_u16.to_le_bytes());
    zip64_eocd.extend_from_slice(&45_u16.to_le_bytes());
    zip64_eocd.extend_from_slice(&0_u32.to_le_bytes());
    zip64_eocd.extend_from_slice(&0_u32.to_le_bytes());
    zip64_eocd.extend_from_slice(&entry_count.to_le_bytes());
    zip64_eocd.extend_from_slice(&entry_count.to_le_bytes());
    zip64_eocd.extend_from_slice(&(central.len() as u64).to_le_bytes());
    zip64_eocd.extend_from_slice(&(central_start as u64).to_le_bytes());
    zip64_eocd.extend_from_slice(ZIP64_EXTENSIBLE_DATA);
    source.extend_from_slice(&zip64_eocd);

    source.extend_from_slice(&0x0706_4b50_u32.to_le_bytes());
    source.extend_from_slice(&0_u32.to_le_bytes());
    source.extend_from_slice(&(zip64_eocd_offset as u64).to_le_bytes());
    source.extend_from_slice(&1_u32.to_le_bytes());

    let mut classic_eocd = original_classic_eocd;
    classic_eocd[8..10].copy_from_slice(&u16::MAX.to_le_bytes());
    classic_eocd[10..12].copy_from_slice(&u16::MAX.to_le_bytes());
    classic_eocd[12..16].copy_from_slice(&u32::MAX.to_le_bytes());
    classic_eocd[16..20].copy_from_slice(&u32::MAX.to_le_bytes());
    source.extend_from_slice(&classic_eocd);
    source
}

fn parse_archive(bytes: &[u8]) -> ZipArchive<Cursor<&[u8]>> {
    ZipArchive::from_slice(bytes).unwrap().into_zip_archive()
}

fn raw_members(bytes: &[u8]) -> BTreeMap<Vec<u8>, RawMember> {
    let parsed = parse_archive(bytes);
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&parsed, &mut buffer).unwrap();
    index
        .entries()
        .iter()
        .map(|entry| {
            let name = entry.raw_name_bytes().to_vec();
            let local =
                bytes[entry.local_span().start as usize..entry.local_span().end as usize].to_vec();
            let central_range = entry.central_record();
            let mut central =
                bytes[central_range.start as usize..central_range.end as usize].to_vec();
            let has_zip64_offset = central[42..46] == u32::MAX.to_le_bytes();
            central[42..46].fill(0);
            let name_len = u16::from_le_bytes(central[28..30].try_into().unwrap()) as usize;
            let extra_len = u16::from_le_bytes(central[30..32].try_into().unwrap()) as usize;
            let extra_start = 46 + name_len;
            let extra_end = extra_start + extra_len;
            let mut extra_cursor = extra_start;
            while extra_cursor + 4 <= extra_end {
                let field_id =
                    u16::from_le_bytes(central[extra_cursor..extra_cursor + 2].try_into().unwrap());
                let field_len = u16::from_le_bytes(
                    central[extra_cursor + 2..extra_cursor + 4]
                        .try_into()
                        .unwrap(),
                ) as usize;
                let field_end = extra_cursor + 4 + field_len;
                if field_end > extra_end {
                    break;
                }
                if field_id == 1 && field_len >= 24 && has_zip64_offset {
                    central[extra_cursor + 4 + 16..extra_cursor + 4 + 24].fill(0);
                }
                extra_cursor = field_end;
            }
            (
                name,
                RawMember {
                    local,
                    central_without_offset: central,
                },
            )
        })
        .collect()
}

fn zip64_extensible_data(bytes: &[u8]) -> Vec<u8> {
    let parsed = parse_archive(bytes);
    assert!(parsed.is_zip64());
    let start = parsed.head_eocd_offset() as usize + 56;
    let end = parsed.eocd_offset() as usize - 20;
    bytes[start..end].to_vec()
}

fn archive_comment(bytes: &[u8]) -> Vec<u8> {
    let parsed = parse_archive(bytes);
    let eocd = parsed.eocd_offset() as usize;
    let comment_len = u16::from_le_bytes(bytes[eocd + 20..eocd + 22].try_into().unwrap()) as usize;
    bytes[eocd + 22..eocd + 22 + comment_len].to_vec()
}

#[test]
fn public_zip64_source_content_replacement_preserves_untouched_spans_metadata_and_tail() {
    let source_bytes = promote_to_zip64(zip32_odf_package());
    assert!(parse_archive(&source_bytes).is_zip64());
    assert_eq!(zip64_extensible_data(&source_bytes), ZIP64_EXTENSIBLE_DATA);

    let before = raw_members(&source_bytes);
    let source = Arc::new(OwnedSource::new(source_bytes.clone()));
    let source_reader: Arc<dyn ReadAt> = source.clone();
    let package = SourceBackedPackage::from_read_at(source_reader).expect("valid ZIP64 ODF source");
    let mut sink = ChunkedSink::new(19);
    let report = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect("public source-backed ZIP64 publication");

    assert!(!report.is_no_op());
    assert_eq!(report.bytes(), sink.bytes.len() as u64);
    assert!(
        sink.write_calls > 1,
        "publication must use the sequential sink"
    );
    assert_eq!(archive_comment(&sink.bytes), ARCHIVE_COMMENT);
    assert_eq!(zip64_extensible_data(&sink.bytes), ZIP64_EXTENSIBLE_DATA);

    let after = raw_members(&sink.bytes);
    assert_ne!(
        after
            .get(b"content.xml".as_slice())
            .map(|member| &member.local),
        before
            .get(b"content.xml".as_slice())
            .map(|member| &member.local),
        "the selected content.xml member must change"
    );
    for (name, member) in before
        .iter()
        .filter(|(name, _)| name.as_slice() != b"content.xml")
    {
        assert_eq!(
            after.get(name.as_slice()),
            Some(member),
            "untouched ZIP64 member {name:?} changed"
        );
    }
    assert_eq!(source.as_slice(), source_bytes.as_slice());

    let output_archive = parse_archive(&sink.bytes);
    assert!(output_archive.is_zip64());
    assert_eq!(output_archive.entries_hint(), 5);
    let reopened = litchi_odf_common::core::OwnedPackage::from_bytes(sink.bytes.clone())
        .expect("published ZIP64 ODF must reopen");
    assert_eq!(reopened.get_file("content.xml").unwrap(), TARGET_CONTENT);
    assert_eq!(
        reopened.get_file("Pictures/blob.bin").unwrap(),
        opaque_media()
    );
    assert_eq!(
        reopened.get_file("Padding/tail.bin").unwrap(),
        b"opaque tail bytes retained after the changed member"
    );

    let reopened_source = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(sink.bytes)))
        .expect("published ZIP64 source-backed package must reopen");
    assert_eq!(
        reopened_source.get_file("content.xml").unwrap(),
        TARGET_CONTENT
    );
}
