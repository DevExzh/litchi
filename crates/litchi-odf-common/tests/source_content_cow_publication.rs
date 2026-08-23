#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "These integration fixtures are deliberately small, fixed, and assertion-driven."
)]

//! Adversarial coverage for the source-backed `content.xml` raw publisher.
//!
//! The fixture helpers intentionally retain ZIP framing rather than comparing
//! only decompressed payloads.  A content-only publication may change the
//! selected member, but every untouched member must retain its local span and
//! central record (apart from the central local-offset field, which necessarily
//! moves when the selected member changes length).

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as CoreLimits, OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_odf_common::core::{
    PackageWriter, Profile, SourceBackedPackage, SourceContentPublicationError,
    SourceContentPublicationOptions, SourceContentPublicationProgress,
};
use soapberry_zip::{PreservationIndex, ZipArchive};
use std::collections::BTreeMap;
use std::io::{self, Cursor, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};
use zip::write::{ExtendedFileOptions, FileOptions, SimpleFileOptions};
use zip::{CompressionMethod as ZipCompressionMethod, ZipWriter};

const MIME: &str = "application/vnd.oasis.opendocument.text";
const CONTENT_MEDIA: &str = "text/xml";
const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const SOURCE_CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text><office:p>source</office:p></office:text></office:body></office:document-content>"#;
const TARGET_CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text><office:p>target</office:p><office:p>second</office:p></office:text></office:body></office:document-content>"#;
const SIGNATURE_XML: &[u8] =
    br#"<document-signatures xmlns="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#;

static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(40_000);

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

#[derive(Debug, Clone, Default)]
struct ProbeState {
    bytes: Vec<u8>,
    revision: u64,
    reads: usize,
    bytes_read: usize,
    ranges: Vec<Range<u64>>,
    forbidden_until_output: Option<Range<u64>>,
    output_started: bool,
    fail_reads_after_output: bool,
    mutate_on_failed_read: bool,
    shorten_on_failed_read: bool,
    fail_version_after_output: bool,
}

/// A positional source that records physical ranges and can mutate between
/// any two source checks.  The publication contract is intentionally tested
/// against this rather than `OwnedSource`, whose immutable version cannot
/// exercise stale-source paths.
#[derive(Debug)]
struct ProbeSource {
    id: u64,
    state: Mutex<ProbeState>,
}

impl ProbeSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            state: Mutex::new(ProbeState {
                bytes,
                ..ProbeState::default()
            }),
        })
    }

    fn bump_revision(&self) {
        let mut state = self.state.lock().unwrap();
        state.revision = state.revision.saturating_add(1);
    }

    fn mutate(&self) {
        let mut state = self.state.lock().unwrap();
        if let Some(byte) = state.bytes.get_mut(1) {
            *byte ^= 0x01;
        }
        state.revision = state.revision.saturating_add(1);
    }

    fn bytes_read(&self) -> usize {
        self.state.lock().unwrap().bytes_read
    }

    fn ranges(&self) -> Vec<Range<u64>> {
        self.state.lock().unwrap().ranges.clone()
    }

    fn forbid_range_until_output(&self, range: Range<u64>) {
        self.state.lock().unwrap().forbidden_until_output = Some(range);
    }

    fn mark_output_started(&self) {
        self.state.lock().unwrap().output_started = true;
    }

    fn fail_reads_after_output(&self) {
        self.state.lock().unwrap().fail_reads_after_output = true;
    }

    fn mutate_and_fail_reads_after_output(&self) {
        let mut state = self.state.lock().unwrap();
        state.fail_reads_after_output = true;
        state.mutate_on_failed_read = true;
    }

    fn shorten_and_fail_reads_after_output(&self) {
        let mut state = self.state.lock().unwrap();
        state.fail_reads_after_output = true;
        state.shorten_on_failed_read = true;
    }

    fn fail_version_after_output(&self) {
        self.state.lock().unwrap().fail_version_after_output = true;
    }

    fn has_read_range(&self, wanted: Range<u64>) -> bool {
        self.ranges()
            .into_iter()
            .any(|range| overlaps(&range, &wanted))
    }
}

impl ReadAt for ProbeSource {
    fn len(&self) -> io::Result<u64> {
        let state = self.state.lock().unwrap();
        u64::try_from(state.bytes.len())
            .map_err(|_| io::Error::other("test source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let mut state = self.state.lock().unwrap();
        state.reads = state.reads.saturating_add(1);
        if state.output_started && state.fail_reads_after_output {
            if state.mutate_on_failed_read {
                state.revision = state.revision.saturating_add(1);
            }
            if state.shorten_on_failed_read {
                state.bytes.pop();
            }
            return Err(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "test source read failure after output",
            ));
        }
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        let Some(input) = state.bytes.get(start..) else {
            return Ok(0);
        };
        let amount = input.len().min(output.len());
        let end = offset
            .checked_add(amount as u64)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "range overflow"))?;
        if !state.output_started
            && state
                .forbidden_until_output
                .as_ref()
                .is_some_and(|range| overlaps(&(offset..end), range))
        {
            return Err(io::Error::other(format!(
                "opaque payload was read before publication output began: requested {offset}..{end}, forbidden {range:?}",
                range = state.forbidden_until_output.as_ref().unwrap()
            )));
        }
        output[..amount].copy_from_slice(&input[..amount]);
        state.bytes_read = state.bytes_read.saturating_add(amount);
        if amount != 0 {
            state.ranges.push(offset..end);
        }
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let state = self.state.lock().unwrap();
        if state.output_started && state.fail_version_after_output {
            return Err(io::Error::other("test source version failure after output"));
        }
        Ok(SourceVersion::new(self.id, state.revision))
    }
}

#[derive(Debug)]
struct ShortSink {
    bytes: Vec<u8>,
    maximum_write: usize,
}

impl Write for ShortSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        let amount = input.len().min(self.maximum_write);
        if amount == 0 {
            return Ok(0);
        }
        self.bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct TrackingSink {
    bytes: Vec<u8>,
    source: Arc<ProbeSource>,
}

#[derive(Debug)]
struct ReentrantReadSink<'a> {
    bytes: Vec<u8>,
    package: &'a SourceBackedPackage,
    read_once: bool,
}

impl Write for ReentrantReadSink<'_> {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if !self.read_once {
            self.read_once = true;
            self.package
                .get_file("content.xml")
                .map_err(io::Error::other)?;
        }
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for TrackingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.source.mark_output_started();
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct FlushFailSink {
    bytes: Vec<u8>,
}

impl Write for FlushFailSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Err(io::Error::new(
            io::ErrorKind::BrokenPipe,
            "test sink flush failure",
        ))
    }
}

#[derive(Debug)]
struct FlushCancellingSink {
    bytes: Vec<u8>,
    cancellation: CancellationSource,
}

impl Write for FlushCancellingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        self.cancellation.cancel();
        Ok(())
    }
}

#[derive(Debug)]
struct OverreportSink;

impl Write for OverreportSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        Ok(input.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct ErrorSink;

impl Write for ErrorSink {
    fn write(&mut self, _input: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::BrokenPipe,
            "test sink write failure",
        ))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct PartialErrorSink {
    bytes: Vec<u8>,
    first_acceptance: usize,
    wrote_once: bool,
}

impl Write for PartialErrorSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.wrote_once {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test sink failure after a known prefix",
            ));
        }
        self.wrote_once = true;
        let accepted = input.len().min(self.first_acceptance);
        self.bytes.extend_from_slice(&input[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct ZeroSink;

impl Write for ZeroSink {
    fn write(&mut self, _input: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct MutatingSink {
    bytes: Vec<u8>,
    source: Arc<ProbeSource>,
    mutate_after_first_write: bool,
    mutate_on_flush: bool,
}

impl Write for MutatingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.source.mark_output_started();
        self.bytes.extend_from_slice(input);
        if self.mutate_after_first_write {
            self.mutate_after_first_write = false;
            self.source.mutate();
        }
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.mutate_on_flush {
            self.mutate_on_flush = false;
            self.source.mutate();
        }
        Ok(())
    }
}

#[derive(Debug)]
struct CancellingSink {
    bytes: Vec<u8>,
    source: Arc<ProbeSource>,
    cancellation: CancellationSource,
    cancel_after_first_write: bool,
}

#[derive(Debug)]
struct PayloadObservationSink {
    bytes: Vec<u8>,
    source: Arc<ProbeSource>,
    payload_range: Range<u64>,
    payload_read_before_first_write: Option<bool>,
}

impl Write for PayloadObservationSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.payload_read_before_first_write.is_none() {
            self.payload_read_before_first_write =
                Some(self.source.has_read_range(self.payload_range.clone()));
            self.source.mark_output_started();
        }
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for CancellingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.source.mark_output_started();
        self.bytes.extend_from_slice(input);
        if self.cancel_after_first_write {
            self.cancel_after_first_write = false;
            self.cancellation.cancel();
        }
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn overlaps(left: &Range<u64>, right: &Range<u64>) -> bool {
    left.start < right.end && right.start < left.end
}

fn manifest(with_size: bool, aliases: bool, signed: bool, include_media: bool) -> Vec<u8> {
    let size = if with_size {
        format!(r#" manifest:size="{}""#, SOURCE_CONTENT.len())
    } else {
        String::new()
    };
    let alias = if aliases {
        r#"<manifest:file-entry manifest:full-path="/content.xml" manifest:media-type="text/xml"/>"#
    } else {
        ""
    };
    let media = if include_media {
        r#"<manifest:file-entry manifest:full-path="Pictures/blob.bin" manifest:media-type="application/octet-stream"/><manifest:file-entry manifest:full-path="Padding/tail.bin" manifest:media-type="application/octet-stream"/>"#
    } else {
        ""
    };
    let signature = if signed {
        r#"<manifest:file-entry manifest:full-path="META-INF/documentsignatures.xml" manifest:media-type="text/xml"/>"#
    } else {
        ""
    };
    format!(
        r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="{CONTENT_MEDIA}"{size}/>{alias}{media}{signature}</manifest:manifest>"#
    )
    .into_bytes()
}

fn file_options(method: ZipCompressionMethod) -> FileOptions<'static, ExtendedFileOptions> {
    let mut options = FileOptions::default()
        .compression_method(method)
        .with_file_comment("member comment");
    options
        .add_extra_data(0x1234, b"member-extra", false)
        .unwrap();
    options
}

fn zip_package(
    content_method: ZipCompressionMethod,
    media_method: ZipCompressionMethod,
    with_size: bool,
    aliases: bool,
    signed: bool,
    include_media: bool,
) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    writer
        .set_raw_comment(b"source archive comment".to_vec().into_boxed_slice())
        .unwrap();
    writer
        .start_file(
            "mimetype",
            SimpleFileOptions::default().compression_method(ZipCompressionMethod::Stored),
        )
        .unwrap();
    writer.write_all(MIME.as_bytes()).unwrap();
    writer
        .start_file("content.xml", file_options(content_method))
        .unwrap();
    writer.write_all(SOURCE_CONTENT).unwrap();
    if include_media {
        writer
            .start_file("Pictures/blob.bin", file_options(media_method))
            .unwrap();
        writer.write_all(&opaque_media()).unwrap();
        // Keep the selected opaque member outside the locator's terminal
        // search window. This stored tail is itself raw-preserved.
        writer
            .start_file(
                "Padding/tail.bin",
                file_options(ZipCompressionMethod::Stored),
            )
            .unwrap();
        writer.write_all(&vec![0_u8; 2 * 1024 * 1024]).unwrap();
    }
    if signed {
        writer
            .start_file(
                "META-INF/documentsignatures.xml",
                file_options(ZipCompressionMethod::Stored),
            )
            .unwrap();
        writer.write_all(SIGNATURE_XML).unwrap();
    }
    writer
        .start_file(
            "META-INF/manifest.xml",
            file_options(ZipCompressionMethod::Deflated),
        )
        .unwrap();
    writer
        .write_all(&manifest(with_size, aliases, signed, include_media))
        .unwrap();
    writer.finish().unwrap().into_inner()
}

fn streaming_package() -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer.write_stored("mimetype", MIME.as_bytes()).unwrap();
    writer
        .write_deflated("content.xml", SOURCE_CONTENT)
        .unwrap();
    writer
        .write_deflated("Pictures/blob.bin", &opaque_media())
        .unwrap();
    writer
        .write_deflated(
            "META-INF/manifest.xml",
            &manifest(false, false, false, true),
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn encrypted_package() -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption("test-password", Profile::compatible())
        .unwrap();
    writer.add_file("content.xml", SOURCE_CONTENT).unwrap();
    writer
        .add_file_with_media_type(
            "Pictures/blob.bin",
            &opaque_media(),
            "application/octet-stream",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn opaque_media() -> Vec<u8> {
    let mut state = 0x9e37_79b9_u32;
    // Keep the compressed payload beyond the ZIP locator's 1 MiB EOCD search
    // window so opening metadata does not itself overlap this range.
    (0..2 * 1024 * 1024)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            state as u8
        })
        .collect()
}

fn open(bytes: Vec<u8>) -> (Arc<ProbeSource>, SourceBackedPackage) {
    let source = ProbeSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(source_reader(&source))
        .expect("fixture must open as strict source-backed ODF");
    (source, package)
}

fn source_reader(source: &Arc<ProbeSource>) -> Arc<dyn ReadAt> {
    let source: Arc<dyn ReadAt> = source.clone();
    source
}

fn media_range(bytes: &[u8]) -> Range<u64> {
    let archive = ZipArchive::from_slice(bytes).unwrap();
    archive
        .entries()
        .filter_map(|entry| {
            let entry = entry.ok()?;
            if entry.file_path().as_ref() != b"Pictures/blob.bin" {
                return None;
            }
            {
                let (start, end) = archive
                    .get_entry(entry.wayfinder())
                    .unwrap()
                    .compressed_data_range();
                Some(start..end)
            }
        })
        .next()
        .unwrap()
}

fn raw_members(bytes: &[u8]) -> BTreeMap<Vec<u8>, RawMember> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
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
            central[42..46].fill(0);
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

fn physical_and_central_order(bytes: &[u8]) -> (Vec<Vec<u8>>, Vec<Vec<u8>>) {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let central = index
        .entries()
        .iter()
        .map(|entry| entry.raw_name_bytes().to_vec())
        .collect::<Vec<_>>();
    let mut physical = index.entries().to_vec();
    physical.sort_by_key(|entry| entry.local_span().start);
    (
        physical
            .into_iter()
            .map(|entry| entry.raw_name_bytes().to_vec())
            .collect(),
        central,
    )
}

fn archive_comment(bytes: &[u8]) -> Vec<u8> {
    let offset = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50u32.to_le_bytes())
        .unwrap();
    let length = u16::from_le_bytes([bytes[offset + 20], bytes[offset + 21]]) as usize;
    bytes[offset + 22..offset + 22 + length].to_vec()
}

fn reorder_central_directory(mut bytes: Vec<u8>) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50u32.to_le_bytes())
        .unwrap();
    let central_start =
        u32::from_le_bytes(bytes[eocd + 16..eocd + 20].try_into().unwrap()) as usize;
    let central_size = u32::from_le_bytes(bytes[eocd + 12..eocd + 16].try_into().unwrap()) as usize;
    let central_end = central_start + central_size;
    let mut records = Vec::new();
    let mut cursor = central_start;
    while cursor < central_end {
        assert_eq!(&bytes[cursor..cursor + 4], b"PK\x01\x02");
        let name_len =
            u16::from_le_bytes(bytes[cursor + 28..cursor + 30].try_into().unwrap()) as usize;
        let extra_len =
            u16::from_le_bytes(bytes[cursor + 30..cursor + 32].try_into().unwrap()) as usize;
        let comment_len =
            u16::from_le_bytes(bytes[cursor + 32..cursor + 34].try_into().unwrap()) as usize;
        let length = 46 + name_len + extra_len + comment_len;
        records.push(bytes[cursor..cursor + length].to_vec());
        cursor += length;
    }
    // ODF requires `mimetype` to remain the first central entry. Reorder the
    // remaining entries so physical and central order still differ.
    records[1..].reverse();
    bytes.splice(central_start..central_end, records.into_iter().flatten());
    bytes
}

fn patch_method(mut bytes: Vec<u8>, member: &[u8], method: u16) -> Vec<u8> {
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len =
            u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
        let extra_len =
            u16::from_le_bytes(bytes[local + 28..local + 30].try_into().unwrap()) as usize;
        let name_start = local + 30;
        let next = name_start + name_len + extra_len;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[local + 8..local + 10].copy_from_slice(&method.to_le_bytes());
        }
        if next >= bytes.len() {
            break;
        }
        cursor = next;
        if !bytes[cursor..]
            .windows(4)
            .any(|window| window == b"PK\x03\x04")
        {
            break;
        }
    }
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = cursor + relative;
        let name_len =
            u16::from_le_bytes(bytes[central + 28..central + 30].try_into().unwrap()) as usize;
        let extra_len =
            u16::from_le_bytes(bytes[central + 30..central + 32].try_into().unwrap()) as usize;
        let comment_len =
            u16::from_le_bytes(bytes[central + 32..central + 34].try_into().unwrap()) as usize;
        let name_start = central + 46;
        let next = name_start + name_len + extra_len + comment_len;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[central + 10..central + 12].copy_from_slice(&method.to_le_bytes());
        }
        if next >= bytes.len() {
            break;
        }
        cursor = next;
        if !bytes[cursor..]
            .windows(4)
            .any(|window| window == b"PK\x01\x02")
        {
            break;
        }
    }
    bytes
}

fn set_zip_encrypted_flag(mut bytes: Vec<u8>, member: &[u8]) -> Vec<u8> {
    let mut found_local = false;
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len =
            u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
        let name_start = local + 30;
        if &bytes[name_start..name_start + name_len] == member {
            let flags = u16::from_le_bytes(bytes[local + 6..local + 8].try_into().unwrap()) | 1;
            bytes[local + 6..local + 8].copy_from_slice(&flags.to_le_bytes());
            found_local = true;
            break;
        }
        cursor = name_start + name_len;
    }

    let mut found_central = false;
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = cursor + relative;
        let name_len =
            u16::from_le_bytes(bytes[central + 28..central + 30].try_into().unwrap()) as usize;
        let name_start = central + 46;
        if &bytes[name_start..name_start + name_len] == member {
            let flags =
                u16::from_le_bytes(bytes[central + 8..central + 10].try_into().unwrap()) | 1;
            bytes[central + 8..central + 10].copy_from_slice(&flags.to_le_bytes());
            found_central = true;
            break;
        }
        cursor = name_start + name_len;
    }
    assert!(found_local, "fixture local member not found");
    assert!(found_central, "fixture central member not found");
    bytes
}

fn patch_local_member_name(mut bytes: Vec<u8>, member: &[u8], replacement: &[u8]) -> Vec<u8> {
    assert_eq!(member.len(), replacement.len());
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len =
            u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
        let name_start = local + 30;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[name_start..name_start + name_len].copy_from_slice(replacement);
            return bytes;
        }
        cursor = name_start + name_len;
    }
    panic!("fixture local member not found");
}

fn member_layout(bytes: &[u8], member: &[u8]) -> (usize, usize, usize, usize) {
    let archive = ZipArchive::from_slice(bytes).unwrap();
    let payload_end = archive
        .entries()
        .find_map(|entry| {
            let entry = entry.ok()?;
            if entry.file_path().as_ref() != member {
                return None;
            }
            Some(
                archive
                    .get_entry(entry.wayfinder())
                    .unwrap()
                    .compressed_data_range()
                    .1 as usize,
            )
        })
        .unwrap();
    let archive = archive.into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let entry = index
        .entries()
        .iter()
        .find(|entry| entry.raw_name_bytes() == member)
        .unwrap();
    (
        entry.local_span().start as usize,
        entry.local_span().end as usize,
        entry.central_record().start as usize,
        payload_end,
    )
}

fn corrupt_descriptor_crc(mut bytes: Vec<u8>, member: &[u8]) -> Vec<u8> {
    let (_, local_end, _, descriptor_start) = member_layout(&bytes, member);
    assert!(descriptor_start < local_end);
    let signature = u32::from_le_bytes(
        bytes[descriptor_start..descriptor_start + 4]
            .try_into()
            .unwrap(),
    );
    let crc_offset = if signature == 0x0807_4b50 {
        descriptor_start + 4
    } else {
        descriptor_start
    };
    assert!(crc_offset + 4 <= local_end);
    bytes[crc_offset] ^= 1;
    bytes
}

/// Insert bytes between a valid streaming data descriptor and the following
/// local header while keeping the central-directory offsets internally
/// consistent. The normal ZIP indexer can still locate the package, but the
/// raw-preservation publisher must reject the now-padded descriptor before it
/// writes any sink bytes.
fn pad_data_descriptor(mut bytes: Vec<u8>, member: &[u8], padding: &[u8]) -> Vec<u8> {
    assert!(!padding.is_empty());
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50u32.to_le_bytes())
        .unwrap();
    let central_start =
        u32::from_le_bytes(bytes[eocd + 16..eocd + 20].try_into().unwrap()) as usize;
    let central_size = u32::from_le_bytes(bytes[eocd + 12..eocd + 16].try_into().unwrap()) as usize;
    let central_end = central_start + central_size;
    let mut cursor = central_start;
    let mut local_offset = None;
    let mut compressed_size = None;
    let mut flags = None;
    while cursor < central_end {
        assert_eq!(&bytes[cursor..cursor + 4], b"PK\x01\x02");
        let name_len =
            u16::from_le_bytes(bytes[cursor + 28..cursor + 30].try_into().unwrap()) as usize;
        let extra_len =
            u16::from_le_bytes(bytes[cursor + 30..cursor + 32].try_into().unwrap()) as usize;
        let comment_len =
            u16::from_le_bytes(bytes[cursor + 32..cursor + 34].try_into().unwrap()) as usize;
        let length = 46 + name_len + extra_len + comment_len;
        if &bytes[cursor + 46..cursor + 46 + name_len] == member {
            local_offset = Some(u32::from_le_bytes(
                bytes[cursor + 42..cursor + 46].try_into().unwrap(),
            ) as usize);
            compressed_size = Some(u32::from_le_bytes(
                bytes[cursor + 20..cursor + 24].try_into().unwrap(),
            ) as usize);
            flags = Some(u16::from_le_bytes(
                bytes[cursor + 8..cursor + 10].try_into().unwrap(),
            ));
        }
        cursor += length;
    }
    let local = local_offset.expect("descriptor member central record");
    assert_ne!(
        flags.unwrap() & 0x0008,
        0,
        "member must use a data descriptor"
    );
    let local_name_len =
        u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
    let local_extra_len =
        u16::from_le_bytes(bytes[local + 28..local + 30].try_into().unwrap()) as usize;
    let descriptor_start = local + 30 + local_name_len + local_extra_len + compressed_size.unwrap();
    assert_eq!(
        &bytes[descriptor_start..descriptor_start + 4],
        0x0807_4b50u32.to_le_bytes()
    );
    let insertion = descriptor_start + 16;
    let amount = padding.len();
    bytes.splice(insertion..insertion, padding.iter().copied());

    let new_eocd = eocd + amount;
    let new_central_start = central_start + amount;
    bytes[new_eocd + 16..new_eocd + 20].copy_from_slice(&(new_central_start as u32).to_le_bytes());
    let new_central_end = new_central_start + central_size;
    cursor = new_central_start;
    while cursor < new_central_end {
        assert_eq!(&bytes[cursor..cursor + 4], b"PK\x01\x02");
        let name_len =
            u16::from_le_bytes(bytes[cursor + 28..cursor + 30].try_into().unwrap()) as usize;
        let extra_len =
            u16::from_le_bytes(bytes[cursor + 30..cursor + 32].try_into().unwrap()) as usize;
        let comment_len =
            u16::from_le_bytes(bytes[cursor + 32..cursor + 34].try_into().unwrap()) as usize;
        let offset =
            u32::from_le_bytes(bytes[cursor + 42..cursor + 46].try_into().unwrap()) as usize;
        if offset >= insertion {
            bytes[cursor + 42..cursor + 46]
                .copy_from_slice(&((offset + amount) as u32).to_le_bytes());
        }
        cursor += 46 + name_len + extra_len + comment_len;
    }
    bytes
}

#[test]
fn changed_store_and_deflate_publications_preserve_untouched_raw_members_and_order() {
    for content_method in [ZipCompressionMethod::Stored, ZipCompressionMethod::Deflated] {
        let source_bytes = reorder_central_directory(zip_package(
            content_method,
            ZipCompressionMethod::Deflated,
            false,
            false,
            false,
            true,
        ));
        let (source, package) = open(source_bytes.clone());
        let before_raw = raw_members(&source_bytes);
        let (before_physical, before_central) = physical_and_central_order(&source_bytes);
        let before_comment = archive_comment(&source_bytes);
        let mut output = Vec::new();
        let report = package
            .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
            .expect("changed content publication");
        assert!(!report.is_no_op());
        assert_eq!(report.bytes(), output.len() as u64);
        assert_eq!(archive_comment(&output), before_comment);
        assert_eq!(physical_and_central_order(&output).0, before_physical);
        assert_eq!(physical_and_central_order(&output).1, before_central);

        let after_raw = raw_members(&output);
        for name in before_raw
            .keys()
            .filter(|name| name.as_slice() != b"content.xml")
        {
            assert_eq!(
                after_raw.get(name.as_slice()),
                before_raw.get(name.as_slice()),
                "raw member {name:?}"
            );
        }
        let reopened = litchi_odf_common::core::OwnedPackage::from_bytes(output).unwrap();
        assert_eq!(reopened.get_file("content.xml").unwrap(), TARGET_CONTENT);
        assert!(source.bytes_read() > 0);
    }
}

#[test]
fn exact_no_op_preserves_signed_source_byte_for_byte() {
    for source_bytes in [
        zip_package(
            ZipCompressionMethod::Deflated,
            ZipCompressionMethod::Stored,
            false,
            false,
            true,
            true,
        ),
        zip_package(
            ZipCompressionMethod::Deflated,
            ZipCompressionMethod::Stored,
            true,
            false,
            false,
            true,
        ),
    ] {
        let source = Arc::new(OwnedSource::new(source_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let mut output = Vec::new();
        let report = package
            .write_content_xml_to_stream(&mut output, SOURCE_CONTENT)
            .unwrap();
        assert!(report.is_no_op());
        assert_eq!(output, source_bytes);
        assert!(report.bytes() > 0);
    }
}

#[test]
fn exact_no_op_encrypted_source_is_refused_before_decryption_or_output() {
    let source = Arc::new(OwnedSource::new(encrypted_package()));
    let package = SourceBackedPackage::from_read_at_with_password(source, "test-password").unwrap();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream(&mut output, SOURCE_CONTENT)
        .expect_err("encrypted source is outside the raw publisher contract");
    assert!(matches!(
        error,
        SourceContentPublicationError::Unsupported { .. }
    ));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert!(output.is_empty());
}

#[test]
fn zip_encryption_flags_are_refused_before_member_reads_or_output() {
    for member in [b"content.xml".as_slice(), b"Pictures/blob.bin".as_slice()] {
        let bytes = set_zip_encrypted_flag(
            zip_package(
                ZipCompressionMethod::Deflated,
                ZipCompressionMethod::Stored,
                false,
                false,
                false,
                true,
            ),
            member,
        );
        let (source, package) = open(bytes);
        let reads_after_open = source.bytes_read();
        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
            .expect_err("ZIP encryption flags are outside the raw publisher contract");
        assert!(matches!(
            error,
            SourceContentPublicationError::Unsupported { .. }
        ));
        assert_eq!(
            error.progress(),
            SourceContentPublicationProgress::Untouched
        );
        assert_eq!(
            source.bytes_read(),
            reads_after_open,
            "publication read source bytes for {member:?}"
        );
        assert!(output.is_empty());
    }
}

#[test]
fn changed_publication_rejects_signed_encrypted_sized_and_aliased_sources() {
    for (bytes, label) in [
        (
            zip_package(
                ZipCompressionMethod::Deflated,
                ZipCompressionMethod::Stored,
                false,
                false,
                true,
                true,
            ),
            "signed",
        ),
        (encrypted_package(), "encrypted"),
        (
            zip_package(
                ZipCompressionMethod::Deflated,
                ZipCompressionMethod::Stored,
                true,
                false,
                false,
                true,
            ),
            "manifest-size",
        ),
    ] {
        let source = Arc::new(OwnedSource::new(bytes));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
            .expect_err(label);
        assert_eq!(
            error.progress(),
            SourceContentPublicationProgress::Untouched,
            "{label}"
        );
        assert!(output.is_empty(), "{label} must fail before sink output");
    }

    let aliased = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        true,
        false,
        true,
    );
    assert!(
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(aliased))).is_err(),
        "manifest aliases must be rejected while opening the source owner"
    );
}

#[test]
fn changed_publication_rejects_prefix_suffix_zip64_and_unsupported_compression() {
    let base = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    );

    let mut prefixed = b"prefix bytes".to_vec();
    prefixed.extend_from_slice(&base);
    assert!(SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(prefixed))).is_err());

    let mut suffixed = base.clone();
    suffixed.extend_from_slice(b"suffix bytes");
    let source = Arc::new(OwnedSource::new(suffixed));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect_err("suffix must not be silently discarded");
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert!(output.is_empty());

    let mut zip64 = base.clone();
    let eocd = zip64
        .windows(4)
        .rposition(|window| window == 0x0605_4b50u32.to_le_bytes())
        .unwrap();
    zip64[eocd + 8..eocd + 20].fill(0xff);
    assert!(SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(zip64))).is_err());

    let unsupported = patch_method(base.clone(), b"content.xml", 99);
    let source = Arc::new(OwnedSource::new(unsupported));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect_err("unsupported compression must be refused");
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert!(output.is_empty());

    let opaque_unknown = patch_method(base, b"Pictures/blob.bin", 99);
    let before_raw = raw_members(&opaque_unknown);
    let source = Arc::new(OwnedSource::new(opaque_unknown));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect("unknown untouched compression must remain raw-copyable");
    let after_raw = raw_members(&output);
    assert_eq!(
        after_raw.get(b"Pictures/blob.bin".as_slice()),
        before_raw.get(b"Pictures/blob.bin".as_slice())
    );
}

#[test]
fn changed_publication_checks_stale_source_before_during_and_after_output() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    );
    let clean_package =
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let mut candidate = Vec::new();
    clean_package
        .write_content_xml_to_stream(&mut candidate, TARGET_CONTENT)
        .unwrap();

    let source = ProbeSource::new(bytes.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    source.bump_revision();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect_err("stale-before publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged { .. }
    ));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert_eq!(error.written(), 0);
    assert!(output.is_empty());

    let source = ProbeSource::new(bytes.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    let mut sink = MutatingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
        mutate_after_first_write: true,
        mutate_on_flush: false,
    };
    let error = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect_err("stale-during publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged {
            progress: SourceContentPublicationProgress::Prefix { accepted },
            ..
        } if accepted == sink.bytes.len() as u64
    ));
    assert_eq!(error.written(), sink.bytes.len() as u64);
    assert!(!sink.bytes.is_empty());
    assert_eq!(sink.bytes, candidate[..sink.bytes.len()]);

    let source = ProbeSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    let mut sink = MutatingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
        mutate_after_first_write: false,
        mutate_on_flush: true,
    };
    let error = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect_err("stale-after publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged {
            progress: SourceContentPublicationProgress::Complete { bytes },
            ..
        } if bytes == sink.bytes.len() as u64
    ));
    assert_eq!(error.written(), sink.bytes.len() as u64);
    assert!(!sink.bytes.is_empty());
    assert_eq!(sink.bytes, candidate);
}

#[test]
fn no_op_source_read_failure_after_output_reports_the_exact_prefix() {
    let source_bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    );
    let source = ProbeSource::new(source_bytes.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    source.fail_reads_after_output();
    let mut sink = TrackingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
    };

    let error = package
        .write_content_xml_to_stream(&mut sink, SOURCE_CONTENT)
        .expect_err("source read failure after a prefix");
    assert!(matches!(
        error,
        SourceContentPublicationError::Source {
            progress: SourceContentPublicationProgress::Prefix { accepted },
            ..
        } if accepted == sink.bytes.len() as u64
    ));
    assert!(!sink.bytes.is_empty());
    assert_eq!(error.written(), sink.bytes.len() as u64);
    assert_eq!(&source_bytes[..sink.bytes.len()], sink.bytes.as_slice());
}

#[test]
fn failed_source_reads_still_prefer_revision_and_length_races() {
    let source_bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    );

    let source = ProbeSource::new(source_bytes.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    source.mutate_and_fail_reads_after_output();
    let mut sink = TrackingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
    };
    let error = package
        .write_content_xml_to_stream(&mut sink, SOURCE_CONTENT)
        .expect_err("revision change must win over the failed source read");
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged {
            progress: SourceContentPublicationProgress::Prefix { accepted },
            ..
        } if accepted == sink.bytes.len() as u64
    ));
    assert_eq!(&source_bytes[..sink.bytes.len()], sink.bytes.as_slice());

    let source = ProbeSource::new(source_bytes.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    source.shorten_and_fail_reads_after_output();
    let mut sink = TrackingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
    };
    let error = package
        .write_content_xml_to_stream(&mut sink, SOURCE_CONTENT)
        .expect_err("length change must win over the failed source read");
    assert!(matches!(
        error,
        SourceContentPublicationError::Source {
            progress: SourceContentPublicationProgress::Prefix { accepted },
            source: litchi_core::Error::InvalidFormat(ref message),
        } if accepted == sink.bytes.len() as u64
            && message.contains("length changed")
    ));
    assert_eq!(&source_bytes[..sink.bytes.len()], sink.bytes.as_slice());
}

#[test]
fn preservation_io_source_failures_report_source_and_exact_prefix_progress() {
    let source_bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    );
    let expected_package =
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(source_bytes.clone())))
            .unwrap();
    let mut expected = Vec::new();
    expected_package
        .write_content_xml_to_stream(&mut expected, TARGET_CONTENT)
        .unwrap();

    for fail_kind in ["read", "version"] {
        let source = ProbeSource::new(source_bytes.clone());
        let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
        if fail_kind == "read" {
            source.fail_reads_after_output();
        } else {
            source.fail_version_after_output();
        }
        let mut sink = TrackingSink {
            bytes: Vec::new(),
            source: Arc::clone(&source),
        };
        let error = package
            .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
            .expect_err(fail_kind);
        assert!(matches!(
            error,
            SourceContentPublicationError::Source {
                progress: SourceContentPublicationProgress::Prefix { accepted },
                ..
            } if accepted == sink.bytes.len() as u64
        ));
        assert!(!sink.bytes.is_empty(), "{fail_kind}");
        assert_eq!(error.written(), sink.bytes.len() as u64, "{fail_kind}");
        assert_eq!(
            &expected[..sink.bytes.len()],
            sink.bytes.as_slice(),
            "{fail_kind} preservation prefix"
        );
    }
}

#[test]
fn flush_failure_reports_complete_unflushed_progress() {
    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut sink = FlushFailSink { bytes: Vec::new() };
    let error = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect_err("flush failure");
    assert!(matches!(
        error,
        SourceContentPublicationError::Sink {
            progress: SourceContentPublicationProgress::CompleteUnflushed { bytes },
            ..
        } if bytes == sink.bytes.len() as u64
    ));
    assert!(!sink.bytes.is_empty());
    assert_eq!(error.written(), sink.bytes.len() as u64);
}

#[test]
fn cancellation_during_flush_reports_complete_progress() {
    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let (cancellation, token) = CancellationSource::pair();
    let mut sink = FlushCancellingSink {
        bytes: Vec::new(),
        cancellation,
    };
    let error = package
        .write_content_xml_to_stream_with_options(
            &mut sink,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .expect_err("cancellation during flush");
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled {
            progress: SourceContentPublicationProgress::Complete { bytes },
        } if bytes == sink.bytes.len() as u64
    ));
    assert!(!sink.bytes.is_empty());
    assert_eq!(error.written(), sink.bytes.len() as u64);
}

#[test]
fn overreporting_sink_is_indeterminate() {
    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut candidate = Vec::new();
    package
        .write_content_xml_to_stream(&mut candidate, TARGET_CONTENT)
        .unwrap();
    let error = package
        .write_content_xml_to_stream(OverreportSink, TARGET_CONTENT)
        .expect_err("overreporting sink");
    assert!(matches!(
        error,
        SourceContentPublicationError::Sink {
            progress: SourceContentPublicationProgress::Indeterminate { accepted_before: 0 },
            ..
        }
    ));
    assert_eq!(error.written(), 0);

    let error = package
        .write_content_xml_to_stream(ErrorSink, TARGET_CONTENT)
        .expect_err("a write error has unknowable per-call acceptance");
    assert!(matches!(
        error,
        SourceContentPublicationError::Sink {
            progress: SourceContentPublicationProgress::Indeterminate { accepted_before: 0 },
            ..
        }
    ));

    let mut partial = PartialErrorSink {
        bytes: Vec::new(),
        first_acceptance: 7,
        wrote_once: false,
    };
    let error = package
        .write_content_xml_to_stream(&mut partial, TARGET_CONTENT)
        .expect_err("a later write error has unknowable per-call acceptance");
    assert!(matches!(
        error,
        SourceContentPublicationError::Sink {
            progress: SourceContentPublicationProgress::Indeterminate { accepted_before: 7 },
            ..
        }
    ));
    assert_eq!(error.written(), 7);
    assert_eq!(partial.bytes, candidate[..7]);
}

#[test]
fn changed_publication_does_not_read_opaque_payload_before_output_begins() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Deflated,
        false,
        false,
        false,
        true,
    );
    let media = media_range(&bytes);
    let source = ProbeSource::new(bytes);
    source.forbid_range_until_output(media.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    let before = source.bytes_read();
    let mut sink = MutatingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
        mutate_after_first_write: false,
        mutate_on_flush: false,
    };
    let report = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect("opaque media is raw-copied only once output starts");
    assert_eq!(report.bytes(), sink.bytes.len() as u64);
    assert!(source.bytes_read() > before);
    assert!(source.has_read_range(media));
}

#[test]
fn changed_publication_roundtrips_changed_member() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Deflated,
        false,
        false,
        false,
        true,
    );
    let source = ProbeSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    let mut sink = MutatingSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
        mutate_after_first_write: false,
        mutate_on_flush: false,
    };
    let report = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect("changed publication should roundtrip the replacement");
    assert!(!report.is_no_op());
    assert_eq!(report.bytes(), sink.bytes.len() as u64);
    let reopened = litchi_odf_common::core::OwnedPackage::from_bytes(sink.bytes).unwrap();
    assert_eq!(reopened.get_file("content.xml").unwrap(), TARGET_CONTENT);
}

#[test]
fn changed_publication_keeps_signed_refusal_before_output() {
    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        true,
        false,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect_err("changed signed source must be refused");
    assert!(matches!(
        error,
        SourceContentPublicationError::Unsupported { .. }
    ));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert!(output.is_empty());
}

#[test]
fn changed_publication_keeps_unsupported_manifest_refusal_before_output() {
    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        true,
        false,
        false,
        false,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect_err("sized content manifest remains outside raw replacement");
    assert!(matches!(
        error,
        SourceContentPublicationError::Unsupported { ref reason }
            if reason.contains("manifest:size")
    ));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert!(output.is_empty());
}

#[test]
fn payload_verification_opt_in_reads_untouched_payload_before_output() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Deflated,
        false,
        false,
        false,
        true,
    );
    let media = media_range(&bytes);
    let source = ProbeSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    let mut sink = PayloadObservationSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
        payload_range: media,
        payload_read_before_first_write: None,
    };
    package
        .write_content_xml_to_stream_with_options(
            &mut sink,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_payload_verification(true),
        )
        .expect("payload verification publication");
    assert_eq!(sink.payload_read_before_first_write, Some(true));
}

#[test]
fn cancellation_and_output_limits_report_truthful_partial_progress() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    );
    let bounded_source = ProbeSource::new(bytes.clone());
    let bounded_package =
        SourceBackedPackage::from_read_at(source_reader(&bounded_source)).unwrap();
    let before_replacement_refusal = bounded_source.bytes_read();
    let mut bounded_output = Vec::new();
    let error = bounded_package
        .write_content_xml_to_stream_with_options(
            &mut bounded_output,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new()
                .with_max_replacement_bytes(TARGET_CONTENT.len() as u64 - 1),
        )
        .expect_err("replacement ceiling must be checked before source reads");
    assert!(matches!(
        error,
        SourceContentPublicationError::LimitExceeded {
            progress: SourceContentPublicationProgress::Untouched,
            actual,
            maximum,
        } if actual == TARGET_CONTENT.len() as u64 && maximum + 1 == actual
    ));
    assert!(bounded_output.is_empty());
    assert_eq!(bounded_source.bytes_read(), before_replacement_refusal);

    let source = Arc::new(OwnedSource::new(bytes));
    let package = SourceBackedPackage::from_read_at(source).unwrap();

    let (cancellation, token) = CancellationSource::pair();
    cancellation.cancel();
    let mut output = Vec::new();
    let error = package
        .write_content_xml_to_stream_with_options(
            &mut output,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .expect_err("pre-cancelled publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled { .. }
    ));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert!(output.is_empty());

    let (cancellation, token) = CancellationSource::pair();
    let source = ProbeSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    ));
    let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
    let mut sink = CancellingSink {
        bytes: Vec::new(),
        source,
        cancellation: cancellation.clone(),
        cancel_after_first_write: true,
    };
    let error = package
        .write_content_xml_to_stream_with_options(
            &mut sink,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .expect_err("cancellation during output");
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled { .. }
    ));
    assert!(error.written() > 0);
    assert_eq!(error.written(), sink.bytes.len() as u64);

    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut expected = Vec::new();
    let report = package
        .write_content_xml_to_stream(&mut expected, TARGET_CONTENT)
        .unwrap();
    let mut output = Vec::new();
    let maximum = report.bytes().saturating_sub(1);
    let error = package
        .write_content_xml_to_stream_with_options(
            &mut output,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_max_output_bytes(maximum),
        )
        .expect_err("one-byte-under output ceiling");
    assert!(matches!(
        error,
        SourceContentPublicationError::LimitExceeded {
            progress: SourceContentPublicationProgress::Prefix { accepted },
            ..
        } if accepted == output.len() as u64
    ));
    assert!(!output.is_empty());
    assert!(expected.starts_with(&output));
    assert_eq!(&expected[..output.len()], output.as_slice());
    assert_eq!(error.written(), output.len() as u64);
}

#[test]
fn short_and_zero_sinks_are_safe_and_report_progress() {
    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut expected = Vec::new();
    package
        .write_content_xml_to_stream(&mut expected, TARGET_CONTENT)
        .expect("reference publication");
    let mut sink = ShortSink {
        bytes: Vec::new(),
        maximum_write: 7,
    };
    let report = package
        .write_content_xml_to_stream(&mut sink, TARGET_CONTENT)
        .expect("short writes must be retried");
    assert_eq!(report.bytes(), sink.bytes.len() as u64);
    assert_eq!(
        sink.bytes, expected,
        "short writes must preserve exact output"
    );
    let reopened = litchi_odf_common::core::OwnedPackage::from_bytes(sink.bytes).unwrap();
    assert_eq!(reopened.get_file("content.xml").unwrap(), TARGET_CONTENT);
    assert_eq!(
        reopened.get_file("Pictures/blob.bin").unwrap(),
        opaque_media()
    );

    let source = Arc::new(OwnedSource::new(zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let error = package
        .write_content_xml_to_stream(ZeroSink, TARGET_CONTENT)
        .expect_err("zero sink must fail without looping");
    assert!(matches!(error, SourceContentPublicationError::Sink { .. }));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );
    assert_eq!(error.written(), 0);
}

#[test]
fn streaming_descriptor_source_reopens_after_content_replacement() {
    let bytes = streaming_package();
    let source = Arc::new(OwnedSource::new(bytes.clone()));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .expect("descriptor-bearing source publication");
    let reopened = litchi_odf_common::core::OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(reopened.get_file("content.xml").unwrap(), TARGET_CONTENT);
    assert_eq!(
        reopened.get_file("Pictures/blob.bin").unwrap(),
        opaque_media()
    );
}

#[test]
fn malformed_and_padded_data_descriptors_are_rejected_before_output() {
    for (bytes, label) in [
        (
            corrupt_descriptor_crc(streaming_package(), b"Pictures/blob.bin"),
            "malformed descriptor",
        ),
        (
            pad_data_descriptor(streaming_package(), b"Pictures/blob.bin", b"PAD!"),
            "padded descriptor",
        ),
        (
            patch_local_member_name(
                streaming_package(),
                b"Pictures/blob.bin",
                b"Pictures/blob.bad",
            ),
            "local/central name mismatch",
        ),
    ] {
        let source = Arc::new(OwnedSource::new(bytes));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let mut output = Vec::new();
        let error = package
            .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
            .expect_err(label);
        assert!(matches!(
            error,
            SourceContentPublicationError::Unsupported { .. }
        ));
        assert_eq!(
            error.progress(),
            SourceContentPublicationProgress::Untouched,
            "{label}"
        );
        assert!(output.is_empty(), "{label} must fail before output");
    }
}

#[test]
fn publication_report_and_reopen_payloads_match() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Deflated,
        false,
        false,
        false,
        true,
    );
    let source = Arc::new(OwnedSource::new(bytes.clone()));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let report = package
        .write_content_xml_to_stream(&mut output, TARGET_CONTENT)
        .unwrap();
    assert_eq!(report.bytes(), output.len() as u64);
    assert!(report.bytes() > 0);
    let reopened = litchi_odf_common::core::OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(reopened.get_file("content.xml").unwrap(), TARGET_CONTENT);
    assert_eq!(
        reopened.get_file("Pictures/blob.bin").unwrap(),
        opaque_media()
    );
}

#[test]
fn execution_context_refuses_typed_resources_and_releases_transient_memory() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Stored,
        false,
        false,
        false,
        false,
    );

    let source = Arc::new(OwnedSource::new(bytes.clone()));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let refusal_budget = Budget::root(
        "publication-refusal",
        CoreLimits::new(0, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, refusal_token) = CancellationSource::pair();
    let refusal_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(1).unwrap(),
        0,
    )
    .unwrap();
    let refusal_context =
        ExecutionContext::new(refusal_budget.clone(), refusal_token, refusal_limits);
    let mut refused_output = Vec::new();
    let error = package
        .write_content_xml_to_stream_with_options(
            &mut refused_output,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_execution_context(refusal_context),
        )
        .expect_err("memory budget refusal");
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            progress: SourceContentPublicationProgress::Untouched,
            source: ExecutionError::ResourceLimit(limit),
        } if limit.resource == Resource::Memory
    ));
    assert!(refused_output.is_empty());
    assert_eq!(refusal_budget.used(Resource::Memory), 0);

    let successful_budget = Budget::root(
        "publication-success",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, successful_token) = CancellationSource::pair();
    let successful_context =
        ExecutionContext::new(successful_budget.clone(), successful_token, refusal_limits);
    let successful_source = ProbeSource::new(bytes.clone());
    let package = SourceBackedPackage::from_read_at(source_reader(&successful_source)).unwrap();
    let source_reads_before = successful_source.bytes_read();
    let mut output = Vec::new();
    package
        .write_content_xml_to_stream_with_options(
            &mut output,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_execution_context(successful_context),
        )
        .expect("managed publication");
    let exact_physical_input = u64::try_from(
        successful_source
            .bytes_read()
            .saturating_sub(source_reads_before),
    )
    .unwrap();
    assert!(!output.is_empty());
    assert_eq!(successful_budget.used(Resource::Memory), 0);
    assert_eq!(
        successful_budget.used(Resource::InputBytes),
        exact_physical_input
    );
    assert_eq!(
        successful_budget.used(Resource::Work),
        package.len() + SOURCE_CONTENT.len() as u64 + 2 * TARGET_CONTENT.len() as u64
    );
    assert_eq!(successful_budget.used(Resource::Objects), 7);
    assert_eq!(
        successful_budget.used(Resource::OutputBytes),
        output.len() as u64
    );

    let reentrant_source = ProbeSource::new(bytes.clone());
    let reentrant_package =
        SourceBackedPackage::from_read_at(source_reader(&reentrant_source)).unwrap();
    let reentrant_before = reentrant_source.bytes_read();
    let reentrant_budget = Budget::root(
        "publication-reentrant",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, reentrant_token) = CancellationSource::pair();
    let reentrant_context =
        ExecutionContext::new(reentrant_budget.clone(), reentrant_token, refusal_limits);
    let mut reentrant_sink = ReentrantReadSink {
        bytes: Vec::new(),
        package: &reentrant_package,
        read_once: false,
    };
    reentrant_package
        .write_content_xml_to_stream_with_options(
            &mut reentrant_sink,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_execution_context(reentrant_context),
        )
        .expect("re-entrant caller read must remain outside publisher accounting");
    assert_eq!(reentrant_sink.bytes, output);
    assert_eq!(
        reentrant_budget.used(Resource::InputBytes),
        exact_physical_input
    );
    assert!(
        u64::try_from(reentrant_source.bytes_read() - reentrant_before).unwrap()
            > exact_physical_input
    );

    let concurrent_source = ProbeSource::new(bytes.clone());
    let concurrent_package =
        Arc::new(SourceBackedPackage::from_read_at(source_reader(&concurrent_source)).unwrap());
    let mut handles = Vec::new();
    for index in 0..2 {
        let package = Arc::clone(&concurrent_package);
        handles.push(std::thread::spawn(move || {
            let budget = Budget::root(
                format!("publication-concurrent-{index}"),
                CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
            );
            let (_, token) = CancellationSource::pair();
            let context = ExecutionContext::new(budget.clone(), token, refusal_limits);
            let mut bytes = Vec::new();
            package
                .write_content_xml_to_stream_with_options(
                    &mut bytes,
                    TARGET_CONTENT,
                    SourceContentPublicationOptions::new().with_execution_context(context),
                )
                .unwrap();
            (budget, bytes)
        }));
    }
    for handle in handles {
        let (budget, concurrent_output) = handle.join().unwrap();
        assert_eq!(concurrent_output, output);
        assert_eq!(budget.used(Resource::InputBytes), exact_physical_input);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    let input_limit = exact_physical_input - 1;
    let input_budget = Budget::root(
        "publication-input-refusal",
        CoreLimits::new(
            u64::MAX,
            input_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_, input_token) = CancellationSource::pair();
    let input_context = ExecutionContext::new(input_budget.clone(), input_token, refusal_limits);
    let input_source = ProbeSource::new(bytes.clone());
    let input_package = SourceBackedPackage::from_read_at(source_reader(&input_source)).unwrap();
    let input_before = input_source.bytes_read();
    let mut input_prefix = Vec::new();
    let error = input_package
        .write_content_xml_to_stream_with_options(
            &mut input_prefix,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_execution_context(input_context),
        )
        .expect_err("one-under physical input budget");
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            source: ExecutionError::ResourceLimit(limit),
            ..
        } if limit.resource == Resource::InputBytes
    ));
    assert_eq!(input_budget.used(Resource::InputBytes), input_limit);
    assert_eq!(
        u64::try_from(input_source.bytes_read() - input_before).unwrap(),
        input_limit
    );
    assert_eq!(input_prefix, output[..input_prefix.len()]);

    for resource in [Resource::Work, Resource::Objects] {
        let limits = match resource {
            Resource::Work => CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, 0),
            Resource::Objects => {
                CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, 0, u64::MAX, u64::MAX)
            },
            _ => unreachable!(),
        };
        let budget = Budget::root("publication-dimension-refusal", limits);
        let (_, token) = CancellationSource::pair();
        let context = ExecutionContext::new(budget, token, refusal_limits);
        let source = ProbeSource::new(bytes.clone());
        let package = SourceBackedPackage::from_read_at(source_reader(&source)).unwrap();
        let before = source.bytes_read();
        let mut refused = Vec::new();
        let error = package
            .write_content_xml_to_stream_with_options(
                &mut refused,
                TARGET_CONTENT,
                SourceContentPublicationOptions::new().with_execution_context(context),
            )
            .expect_err("work/object budget refusal");
        assert!(matches!(
            error,
            SourceContentPublicationError::Execution {
                progress: SourceContentPublicationProgress::Untouched,
                source: ExecutionError::ResourceLimit(limit),
            } if limit.resource == resource
        ));
        assert!(refused.is_empty());
        if resource == Resource::Work {
            assert_eq!(source.bytes_read(), before);
        }
    }

    let required_objects = successful_budget.used(Resource::Objects);
    let parent = Budget::root(
        "publication-parent",
        CoreLimits::new(
            u64::MAX,
            u64::MAX,
            u64::MAX,
            required_objects - 1,
            u64::MAX,
            u64::MAX,
        ),
    );
    let child = parent.child(
        "publication-child",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, hierarchy_token) = CancellationSource::pair();
    let hierarchy_context = ExecutionContext::new(child.clone(), hierarchy_token, refusal_limits);
    let hierarchy_source = Arc::new(OwnedSource::new(bytes.clone()));
    let hierarchy_package = SourceBackedPackage::from_read_at(hierarchy_source).unwrap();
    let mut hierarchy_output = Vec::new();
    let error = hierarchy_package
        .write_content_xml_to_stream_with_options(
            &mut hierarchy_output,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_execution_context(hierarchy_context),
        )
        .expect_err("ancestor object budget must reject the child operation");
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            progress: SourceContentPublicationProgress::Untouched,
            source: ExecutionError::ResourceLimit(limit),
        } if limit.resource == Resource::Objects && limit.scope.as_ref() == "publication-parent"
    ));
    assert!(hierarchy_output.is_empty());
    assert_eq!(parent.used(Resource::Objects), 1);
    assert_eq!(child.used(Resource::Objects), 1);

    let output_limit = output.len() as u64 - 1;
    let output_budget = Budget::root(
        "publication-output-refusal",
        CoreLimits::new(
            u64::MAX,
            u64::MAX,
            output_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_, output_token) = CancellationSource::pair();
    let output_context = ExecutionContext::new(output_budget.clone(), output_token, refusal_limits);
    let mut prefix = Vec::new();
    let error = package
        .write_content_xml_to_stream_with_options(
            &mut prefix,
            TARGET_CONTENT,
            SourceContentPublicationOptions::new().with_execution_context(output_context),
        )
        .expect_err("one-under managed output budget");
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            progress: SourceContentPublicationProgress::Prefix { accepted },
            source: ExecutionError::ResourceLimit(limit),
        } if limit.resource == Resource::OutputBytes && accepted == prefix.len() as u64
    ));
    assert!(!prefix.is_empty());
    assert_eq!(prefix, output[..prefix.len()]);
    assert_eq!(output_budget.used(Resource::Memory), 0);
    assert_eq!(
        output_budget.used(Resource::OutputBytes),
        prefix.len() as u64
    );
}
