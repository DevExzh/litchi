use super::{
    ExtraFieldId, ExtraFields, PreparedCentral, PreparedEntry, PreparedLocal, PreparedTail,
};
use crate::office::ArchiveReader;
use crate::{
    CompressionMethod, PreservationIndex, RECOMMENDED_BUFFER_SIZE, ReaderAt, ReplayLimits,
    ReplayPass, ReplayProgress, ReplayPublicationError, ReplayResource, ZipArchive,
    ZipArchiveWriter, ZipFileHeaderFixed, ZipLocalFileHeaderFixed, ZipOperationAccounting,
};
use std::fmt;
use std::io::{self, Write};
use std::panic::{AssertUnwindSafe, catch_unwind};

const TARGET_NAME: &str = "target.bin";

#[derive(Debug)]
struct ShortSink {
    bytes: Vec<u8>,
    maximum: usize,
}

impl ShortSink {
    fn new(maximum: usize) -> Self {
        assert!(maximum > 0);
        Self {
            bytes: Vec::new(),
            maximum,
        }
    }
}

impl Write for ShortSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let written = buffer.len().min(self.maximum);
        self.bytes.extend_from_slice(&buffer[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug, Default)]
struct ZeroSink {
    bytes: Vec<u8>,
}

impl Write for ZeroSink {
    fn write(&mut self, _buffer: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug, Default)]
struct OverreportSink;

impl Write for OverreportSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        Ok(buffer.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct FailAfterSink {
    bytes: Vec<u8>,
    fail_after: usize,
}

impl FailAfterSink {
    fn new(fail_after: usize) -> Self {
        Self {
            bytes: Vec::new(),
            fail_after,
        }
    }
}

impl Write for FailAfterSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.fail_after {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "replay sink failed",
            ));
        }
        let available = self.fail_after - self.bytes.len();
        let written = buffer.len().min(available);
        self.bytes.extend_from_slice(&buffer[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct FlushCountingFailSink {
    bytes: Vec<u8>,
    fail_after: usize,
    flushes: usize,
}

impl FlushCountingFailSink {
    fn new(fail_after: usize) -> Self {
        Self {
            bytes: Vec::new(),
            fail_after,
            flushes: 0,
        }
    }
}

impl Write for FlushCountingFailSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.fail_after {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "replay sink failed",
            ));
        }
        let available = self.fail_after - self.bytes.len();
        let written = buffer.len().min(available);
        self.bytes.extend_from_slice(&buffer[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.flushes += 1;
        Ok(())
    }
}

#[derive(Debug)]
struct SinkErrorSentinel;

impl fmt::Display for SinkErrorSentinel {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("typed replay sink failure")
    }
}

impl std::error::Error for SinkErrorSentinel {}

#[derive(Debug)]
struct TypedErrorSink {
    bytes: Vec<u8>,
    fail_after: usize,
}

impl TypedErrorSink {
    fn new(fail_after: usize) -> Self {
        Self {
            bytes: Vec::new(),
            fail_after,
        }
    }
}

impl Write for TypedErrorSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.fail_after {
            return Err(io::Error::new(io::ErrorKind::Other, SinkErrorSentinel));
        }
        let available = self.fail_after - self.bytes.len();
        let written = buffer.len().min(available);
        self.bytes.extend_from_slice(&buffer[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct InterruptOnceSink {
    bytes: Vec<u8>,
    interrupted: bool,
}

impl InterruptOnceSink {
    fn new() -> Self {
        Self {
            bytes: Vec::new(),
            interrupted: false,
        }
    }
}

impl Write for InterruptOnceSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::new(
                io::ErrorKind::Interrupted,
                "retry replay sink",
            ));
        }
        self.bytes.extend_from_slice(buffer);
        Ok(buffer.len())
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
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(buffer);
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Err(io::Error::new(io::ErrorKind::BrokenPipe, "flush failed"))
    }
}

fn source_archive() -> Vec<u8> {
    let mut writer = ZipArchiveWriter::new(Vec::new());
    writer
        .write_stored_file("keep.bin", b"untouched source bytes")
        .unwrap();
    writer
        .write_stored_file(TARGET_NAME, b"old replacement")
        .unwrap();
    writer
        .write_stored_file("tail.bin", b"tail source bytes")
        .unwrap();
    with_comment(writer.finish().unwrap(), b"replay archive comment")
}

fn with_comment(mut data: Vec<u8>, comment: &[u8]) -> Vec<u8> {
    let archive = ZipArchive::from_slice(&data).unwrap();
    let eocd = usize::try_from(archive.eocd_offset()).unwrap();
    data[eocd + 20..eocd + 22]
        .copy_from_slice(&u16::try_from(comment.len()).unwrap().to_le_bytes());
    data.extend_from_slice(comment);
    data
}

fn entry_id<R: ReaderAt>(
    index: &PreservationIndex<'_, R>,
    name: &str,
) -> crate::PreservationEntryId {
    index
        .entries()
        .iter()
        .find(|entry| entry.raw_name_bytes() == name.as_bytes())
        .unwrap()
        .id()
}

fn target_id<R: ReaderAt>(index: &PreservationIndex<'_, R>) -> crate::PreservationEntryId {
    entry_id(index, TARGET_NAME)
}

fn replay_to_vec(
    source: &[u8],
    method: CompressionMethod,
    payload: &[u8],
    chunk_size: usize,
) -> Vec<u8> {
    assert!(chunk_size > 0);
    let archive = ZipArchive::from_slice(source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);
    index
        .write_replacing_with_replay(
            target,
            method,
            ReplayLimits::default(),
            Vec::new(),
            |sink| {
                for chunk in payload.chunks(chunk_size) {
                    sink.write_all(chunk)?;
                }
                Ok::<(), io::Error>(())
            },
        )
        .unwrap()
}

fn replay_to_vec_with_accounting(
    source: &[u8],
    method: CompressionMethod,
    payload: &[u8],
) -> (Vec<u8>, ZipOperationAccounting) {
    let archive = ZipArchive::from_slice(source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);
    let mut accounting = ZipOperationAccounting::default();
    let output = index
        .write_replacing_with_replay_with_accounting(
            target,
            method,
            ReplayLimits::default(),
            Vec::new(),
            &mut accounting,
            |sink| {
                sink.write_all(payload).map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap();
    (output, accounting)
}

fn central_record(data: &[u8], name: &str) -> Vec<u8> {
    let archive = ZipArchive::from_slice(data).unwrap();
    let record = archive
        .entries()
        .map(|entry| entry.unwrap())
        .find(|entry| entry.file_path().as_ref() == name.as_bytes())
        .unwrap();
    let start = usize::try_from(record.central_directory_offset()).unwrap();
    let length =
        crate::ZipFileHeaderFixed::SIZE + usize::try_from(record.metadata_size_hint()).unwrap();
    data[start..start + length].to_vec()
}

fn central_record_without_offset(mut record: Vec<u8>) -> Vec<u8> {
    record[42..46].fill(0);
    record
}

fn local_span(data: &[u8], name: &str) -> Vec<u8> {
    let archive = ZipArchive::from_slice(data).unwrap();
    let records: Vec<_> = archive.entries().map(|entry| entry.unwrap()).collect();
    let record = records
        .iter()
        .find(|entry| entry.file_path().as_ref() == name.as_bytes())
        .unwrap();
    let start = record.local_header_offset();
    let end = records
        .iter()
        .map(|entry| entry.local_header_offset())
        .filter(|offset| *offset > start)
        .min()
        .unwrap_or(archive.directory_offset());
    data[usize::try_from(start).unwrap()..usize::try_from(end).unwrap()].to_vec()
}

fn archive_names(data: &[u8]) -> Vec<Vec<u8>> {
    ZipArchive::from_slice(data)
        .unwrap()
        .entries()
        .map(|entry| entry.unwrap().file_path().as_ref().to_vec())
        .collect()
}

fn target_payload_range(data: &[u8]) -> std::ops::Range<usize> {
    let archive = ZipArchive::from_slice(data).unwrap();
    let record = archive
        .entries()
        .map(|entry| entry.unwrap())
        .find(|entry| entry.file_path().as_ref() == TARGET_NAME.as_bytes())
        .unwrap();
    let entry = archive.get_entry(record.wayfinder()).unwrap();
    let (start, end) = entry.compressed_data_range();
    usize::try_from(start).unwrap()..usize::try_from(end).unwrap()
}

fn replay_layout_member() -> (Vec<u8>, Vec<u8>, u64) {
    let mut writer = ZipArchiveWriter::new(Vec::new());
    writer
        .write_stored_file(TARGET_NAME, b"synthetic replay payload")
        .unwrap();
    let data = writer.finish().unwrap();
    let archive = ZipArchive::from_slice(&data).unwrap();
    let local = ZipLocalFileHeaderFixed::parse(&data).unwrap();
    let framing_len = ZipLocalFileHeaderFixed::SIZE + local.variable_length();
    let record = archive
        .entries()
        .map(|entry| entry.unwrap())
        .next()
        .unwrap();
    let central_start = usize::try_from(record.central_directory_offset()).unwrap();
    let central_len =
        ZipFileHeaderFixed::SIZE + usize::try_from(record.metadata_size_hint()).unwrap();
    (
        data[..framing_len].to_vec(),
        data[central_start..central_start + central_len].to_vec(),
        record.compressed_size_hint(),
    )
}

fn name_mismatched_source() -> Vec<u8> {
    let mut data = source_archive();
    let archive = ZipArchive::from_slice(&data).unwrap();
    let record = archive
        .entries()
        .map(|entry| entry.unwrap())
        .find(|entry| entry.file_path().as_ref() == b"keep.bin")
        .unwrap();
    let start = usize::try_from(record.local_header_offset()).unwrap();
    let name_start = start + crate::ZipLocalFileHeaderFixed::SIZE;
    data[name_start..name_start + b"keep.bin".len()].copy_from_slice(b"kept.bin");
    data
}

#[test]
fn replay_replaces_large_and_empty_store_and_deflate_members() {
    let large = (0..(192 * 1024 + 17))
        .map(|index| u8::try_from(index % 251).unwrap())
        .collect::<Vec<_>>();
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        for payload in [Vec::new(), large.clone()] {
            let output = replay_to_vec(&source_archive(), method, &payload, 4096);
            let reader = ArchiveReader::new(&output).unwrap();
            assert_eq!(reader.read(TARGET_NAME).unwrap(), payload);
            assert_eq!(archive_names(&output), archive_names(&source_archive()));
        }
    }
}

#[test]
fn replay_same_name_preserves_untouched_records_order_and_comment() {
    let source = source_archive();
    let payload = b"replacement bytes with a different physical size";
    let output = replay_to_vec(&source, CompressionMethod::Deflate, payload, 7);
    let source_archive = ZipArchive::from_slice(&source).unwrap();
    let output_archive = ZipArchive::from_slice(&output).unwrap();

    assert_eq!(archive_names(&output), archive_names(&source));
    assert_eq!(
        output_archive.comment().as_bytes(),
        source_archive.comment().as_bytes()
    );
    for name in ["keep.bin", "tail.bin"] {
        assert_eq!(local_span(&output, name), local_span(&source, name));
        assert_eq!(
            central_record_without_offset(central_record(&output, name)),
            central_record_without_offset(central_record(&source, name))
        );
    }
    assert_eq!(
        ArchiveReader::new(&output)
            .unwrap()
            .read(TARGET_NAME)
            .unwrap(),
        payload
    );
}

#[test]
fn replay_output_is_independent_of_callback_chunk_boundaries() {
    let source = source_archive();
    let payload = (0..(96 * 1024 + 11))
        .map(|index| u8::try_from((index * 17) % 251).unwrap())
        .collect::<Vec<_>>();
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let one_chunk = replay_to_vec(&source, method, &payload, payload.len());
        let many_chunks = replay_to_vec(&source, method, &payload, 3);
        assert_eq!(one_chunk, many_chunks);
        assert_eq!(
            ArchiveReader::new(&many_chunks)
                .unwrap()
                .read(TARGET_NAME)
                .unwrap(),
            payload
        );
    }
}

#[test]
fn replay_measure_and_emit_allow_different_callback_chunk_boundaries() {
    let source = source_archive();
    let payload = (0..(128 * 1024 + 23))
        .map(|index| u8::try_from((index * 31) % 251).unwrap())
        .collect::<Vec<_>>();
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let expected = replay_to_vec(&source, method, &payload, payload.len());
        let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
        let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let target = target_id(&index);
        let mut calls = 0;
        let output = index
            .write_replacing_with_replay(
                target,
                method,
                ReplayLimits::default(),
                Vec::new(),
                |writer| {
                    calls += 1;
                    if calls == 1 {
                        writer.write_all(&payload)?;
                    } else {
                        for chunk in payload.chunks(3) {
                            writer.write_all(chunk)?;
                        }
                    }
                    Ok::<(), io::Error>(())
                },
            )
            .unwrap();
        assert_eq!(calls, 2);
        assert_eq!(output, expected);
    }
}

#[test]
fn replay_sink_short_writes_produce_exact_output() {
    let source = source_archive();
    let payload = b"short sink replacement payload repeated repeated";
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let expected = replay_to_vec(&source, method, payload, 5);
        let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
        let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let target = target_id(&index);
        let mut sink = ShortSink::new(3);
        let result = index.write_replacing_with_replay(
            target,
            method,
            ReplayLimits::default(),
            &mut sink,
            |writer| {
                for chunk in payload.chunks(5) {
                    writer.write_all(chunk)?;
                }
                Ok::<(), io::Error>(())
            },
        );
        result.unwrap();
        assert_eq!(sink.bytes, expected);
    }
}

#[test]
fn replay_accounting_charges_only_emitted_compressed_payload() {
    let source = source_archive();
    let payload = b"accounted replay payload repeated repeated bytes";
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let (output, accounting) = replay_to_vec_with_accounting(&source, method, payload);
        let compressed = target_payload_range(&output);
        let compressed_len = (compressed.end - compressed.start) as u64;
        assert_eq!(
            accounting.stored_payload_bytes_emitted(),
            if method == CompressionMethod::Store {
                compressed_len
            } else {
                0
            }
        );
        assert_eq!(
            accounting.generated_deflate_payload_bytes_emitted(),
            if method == CompressionMethod::Deflate {
                compressed_len
            } else {
                0
            }
        );
        assert_eq!(accounting.precompressed_payload_bytes_emitted(), 0);
        assert_eq!(
            ArchiveReader::new(&output)
                .unwrap()
                .read(TARGET_NAME)
                .unwrap(),
            payload
        );
    }
}

#[test]
fn replay_partial_sink_accounting_counts_accepted_generated_payload() {
    let source = source_archive();
    let payload = b"partial accounting payload repeated repeated bytes";
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let expected = replay_to_vec(&source, method, payload, payload.len());
        let payload_range = target_payload_range(&expected);
        let fail_after = payload_range.start + 3;
        let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
        let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
        let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
        let target = target_id(&index);
        let mut sink = FailAfterSink::new(fail_after);
        let mut accounting = ZipOperationAccounting::default();
        let error = index
            .write_replacing_with_replay_with_accounting(
                target,
                method,
                ReplayLimits::default(),
                &mut sink,
                &mut accounting,
                |writer| {
                    writer
                        .write_all(payload)
                        .map_err(|error| error.to_string())?;
                    Ok::<(), String>(())
                },
            )
            .unwrap_err();
        assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
        assert_eq!(sink.bytes, expected[..fail_after]);
        let accepted_payload = 3_u64;
        assert_eq!(
            accounting.stored_payload_bytes_emitted(),
            if method == CompressionMethod::Store {
                accepted_payload
            } else {
                0
            }
        );
        assert_eq!(
            accounting.generated_deflate_payload_bytes_emitted(),
            if method == CompressionMethod::Deflate {
                accepted_payload
            } else {
                0
            }
        );
        assert_eq!(accounting.precompressed_payload_bytes_emitted(), 0);
    }
}

#[test]
fn replay_sink_failures_report_zero_overreport_interrupt_and_flush_states() {
    let source = source_archive();
    let payload = b"sink failure payload repeated repeated";
    let expected = replay_to_vec(&source, CompressionMethod::Store, payload, payload.len());
    let payload_range = target_payload_range(&expected);

    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);

    let mut zero = ZeroSink::default();
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut zero,
            |writer| {
                writer
                    .write_all(payload)
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
    assert_eq!(error.pass(), ReplayPass::Emit);
    assert_eq!(error.progress(), ReplayProgress::Untouched);
    assert!(zero.bytes.is_empty());

    let mut overreport = OverreportSink;
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut overreport,
            |writer| {
                writer
                    .write_all(payload)
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
    assert_eq!(error.pass(), ReplayPass::Emit);
    assert!(matches!(
        error.progress(),
        ReplayProgress::Indeterminate { accepted_before: 0 }
    ));

    let central_start = usize::try_from(
        ZipArchive::from_slice(&expected)
            .unwrap()
            .directory_offset(),
    )
    .unwrap();
    let mut failure_points = vec![
        payload_range.start.saturating_sub(3),
        payload_range.start,
        payload_range.start + 3,
        central_start,
        central_start + 3,
    ];
    failure_points.sort_unstable();
    failure_points.dedup();
    for fail_after in failure_points {
        if fail_after >= expected.len() {
            continue;
        }
        let mut partial = FailAfterSink::new(fail_after);
        let error = index
            .write_replacing_with_replay(
                target,
                CompressionMethod::Store,
                ReplayLimits::default(),
                &mut partial,
                |writer| {
                    writer
                        .write_all(payload)
                        .map_err(|error| error.to_string())?;
                    Ok::<(), String>(())
                },
            )
            .unwrap_err();
        assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
        assert_eq!(error.pass(), ReplayPass::Emit);
        assert_eq!(
            error.progress(),
            ReplayProgress::Prefix {
                accepted: fail_after as u64
            }
        );
        assert_eq!(partial.bytes, expected[..fail_after]);
    }

    let mut interrupted = InterruptOnceSink::new();
    index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut interrupted,
            |writer| {
                writer
                    .write_all(payload)
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap();
    assert_eq!(interrupted.bytes, expected);

    let mut flush_failed = FlushFailSink { bytes: Vec::new() };
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut flush_failed,
            |writer| {
                writer
                    .write_all(payload)
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
    assert_eq!(error.pass(), ReplayPass::Emit);
    assert_eq!(
        error.progress(),
        ReplayProgress::CompleteUnflushed {
            bytes: flush_failed.bytes.len() as u64
        }
    );
    assert_eq!(flush_failed.bytes, expected);
}

#[test]
fn replay_early_callback_flush_does_not_make_late_failure_complete() {
    assert_eq!(
        super::replay::test_early_flush_then_late_write_failure_progress(),
        ReplayProgress::Prefix { accepted: 1 }
    );
}

#[test]
fn replay_callback_errors_are_not_swallowed_and_preserve_publication_progress() {
    let source = source_archive();
    let payload = b"callback failure payload repeated repeated";
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);

    let mut before = Vec::new();
    let mut calls = 0;
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut before,
            |_| {
                calls += 1;
                Err::<(), _>("measure failed")
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Callback { .. }));
    assert_eq!(error.pass(), ReplayPass::Measure);
    assert_eq!(error.progress(), ReplayProgress::Untouched);
    assert_eq!(calls, 1);
    assert!(before.is_empty());

    let mut after = Vec::new();
    let mut calls = 0;
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut after,
            |writer| {
                calls += 1;
                if calls == 1 {
                    writer
                        .write_all(payload)
                        .map_err(|error| error.to_string())?;
                    return Ok::<(), String>(());
                }
                writer
                    .write_all(&payload[..3])
                    .map_err(|error| error.to_string())?;
                Err("emit failed".to_string())
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Callback { .. }));
    assert_eq!(error.pass(), ReplayPass::Emit);
    assert_eq!(error.progress().accepted(), after.len() as u64);
    assert!(!after.is_empty());
    assert_eq!(calls, 2);

    let expected = replay_to_vec(&source, CompressionMethod::Store, payload, payload.len());
    let payload_start = target_payload_range(&expected).start;
    assert_eq!(after, expected[..payload_start + 3]);

    let mut swallowed = FailAfterSink::new(payload_start + 3);
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut swallowed,
            |writer| {
                let _ = writer.write_all(payload);
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
    assert_eq!(error.pass(), ReplayPass::Emit);
    assert!(error.callback().is_none());
    assert_eq!(swallowed.bytes, expected[..swallowed.bytes.len()]);
}

#[test]
fn replay_swallowed_sticky_failure_does_not_flush_sink_again() {
    let source = source_archive();
    let payload = b"sticky replay failure payload";
    let expected = replay_to_vec(&source, CompressionMethod::Store, payload, payload.len());
    let payload_start = target_payload_range(&expected).start;
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);

    let mut sink = FlushCountingFailSink::new(payload_start);
    let mut calls = 0;
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut sink,
            |writer| {
                calls += 1;
                if calls == 1 {
                    writer
                        .write_all(payload)
                        .map_err(|error| error.to_string())?;
                    return Ok::<(), String>(());
                }
                let _ = writer.write_all(payload);
                assert!(writer.flush().is_err());
                Ok::<(), String>(())
            },
        )
        .unwrap_err();

    assert!(matches!(&error, ReplayPublicationError::Sink { .. }));
    assert_eq!(error.pass(), ReplayPass::Emit);
    assert_eq!(
        error.progress(),
        ReplayProgress::Prefix {
            accepted: payload_start as u64,
        }
    );
    assert_eq!(calls, 2);
    assert_eq!(sink.flushes, 0);
    assert_eq!(sink.bytes, expected[..payload_start]);
}

#[test]
fn replay_sink_error_preserves_typed_io_source() {
    let source = source_archive();
    let payload = b"typed replay sink failure payload";
    let expected = replay_to_vec(&source, CompressionMethod::Store, payload, payload.len());
    let payload_start = target_payload_range(&expected).start;
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);

    let mut sink = TypedErrorSink::new(payload_start);
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut sink,
            |writer| {
                writer
                    .write_all(payload)
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    let ReplayPublicationError::Sink {
        source: sink_error,
        callback_error,
        ..
    } = error
    else {
        panic!("typed sink failure must remain a replay sink error");
    };
    assert!(callback_error.is_some());
    assert!(
        sink_error
            .get_ref()
            .and_then(|source| source.downcast_ref::<SinkErrorSentinel>())
            .is_some()
    );
    assert_eq!(sink.bytes, expected[..payload_start]);
}

#[test]
fn replay_callback_panic_does_not_flush_extra_output_on_drop() {
    let source = source_archive();
    let payload = b"panic payload repeated repeated";
    let expected = replay_to_vec(&source, CompressionMethod::Store, payload, payload.len());
    let payload_start = target_payload_range(&expected).start;
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);
    let mut sink = Vec::new();
    let mut calls = 0;
    let panic_result = catch_unwind(AssertUnwindSafe(|| {
        let _ = index.write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut sink,
            |writer| {
                calls += 1;
                if calls == 1 {
                    writer.write_all(payload).unwrap();
                    return Ok::<(), &'static str>(());
                }
                writer.write_all(&payload[..3]).unwrap();
                panic!("replay callback panic");
            },
        );
    }));
    assert!(panic_result.is_err());
    assert_eq!(calls, 2);
    assert_eq!(sink, expected[..payload_start + 3]);
}

#[test]
fn replay_detects_equal_length_crc_collision_with_sha256() {
    let source = source_archive();
    let first = b"\x1c\x45\x5e\xc1\x45\xf2";
    let second = b"\xfc\xd5\xf6\xa1\x34\x5c";
    assert_eq!(crate::crc32(first), crate::crc32(second));
    assert_ne!(first, second);

    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);
    let mut sink = Vec::new();
    let mut calls = 0;
    let error = index
        .write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut sink,
            |writer| {
                let payload = if calls == 0 { first } else { second };
                calls += 1;
                writer
                    .write_all(payload)
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    let ReplayPublicationError::NonDeterministic {
        expected,
        actual,
        progress,
    } = error
    else {
        panic!("CRC collision must be rejected by the SHA-256 fingerprints");
    };
    assert_eq!(expected.decoded_size(), actual.decoded_size());
    assert_eq!(expected.compressed_size(), actual.compressed_size());
    assert_eq!(expected.decoded_crc32(), actual.decoded_crc32());
    assert_ne!(expected.decoded_sha256(), actual.decoded_sha256());
    assert_ne!(expected.compressed_sha256(), actual.compressed_sha256());
    assert_eq!(progress.accepted(), sink.len() as u64);
}

#[test]
fn replay_layout_promotes_following_member_central_offset_to_zip64() {
    let source = source_archive();
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let synthetic = PreservationIndex {
        source: index.source,
        entries: index.entries.clone(),
        local_order: index.local_order.clone(),
        archive_comment: index.archive_comment.clone(),
        zip64_tail: index.zip64_tail.clone(),
        archive_end_offset: u64::MAX,
    };
    let (framing, central, compressed_len) = replay_layout_member();
    let mut prepared = vec![
        PreparedEntry {
            local: PreparedLocal::Copy(0..u64::from(u32::MAX)),
            central: PreparedCentral::Copy(0),
            generated_payload: None,
            omitted: false,
        },
        PreparedEntry {
            local: PreparedLocal::Replay {
                framing,
                compressed_len,
            },
            central: PreparedCentral::Generated(central),
            generated_payload: None,
            omitted: false,
        },
        PreparedEntry {
            local: PreparedLocal::Generated(Vec::new()),
            central: PreparedCentral::Generated(Vec::new()),
            generated_payload: None,
            omitted: true,
        },
    ];

    let layout = synthetic.validate_output_layout(&mut prepared).unwrap();
    let central = prepared[1]
        .central
        .checked_bytes(&synthetic.entries)
        .unwrap();
    let fixed = ZipFileHeaderFixed::parse(central).unwrap();
    assert_eq!(fixed.local_header_offset, u32::MAX);
    assert!(fixed.version_needed >= 45);
    let extra_start = ZipFileHeaderFixed::SIZE + usize::from(fixed.file_name_len);
    let extra_end = extra_start + usize::from(fixed.extra_field_len);
    let fields: Vec<_> = ExtraFields::new(&central[extra_start..extra_end]).collect();
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].0, ExtraFieldId::ZIP64);
    assert_eq!(fields[0].1, u64::from(u32::MAX).to_le_bytes().as_slice());
    assert!(matches!(layout.tail, PreparedTail::Zip64(_)));
}

#[test]
fn replay_layout_rejects_framing_and_compressed_length_overflow() {
    let source = source_archive();
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let synthetic = PreservationIndex {
        source: index.source,
        entries: index.entries.clone(),
        local_order: index.local_order.clone(),
        archive_comment: index.archive_comment.clone(),
        zip64_tail: index.zip64_tail.clone(),
        archive_end_offset: u64::MAX,
    };
    let mut prepared = vec![
        PreparedEntry {
            local: PreparedLocal::Generated(Vec::new()),
            central: PreparedCentral::Generated(Vec::new()),
            generated_payload: None,
            omitted: true,
        },
        PreparedEntry {
            local: PreparedLocal::Replay {
                framing: vec![0],
                compressed_len: u64::MAX,
            },
            central: PreparedCentral::Generated(Vec::new()),
            generated_payload: None,
            omitted: false,
        },
        PreparedEntry {
            local: PreparedLocal::Generated(Vec::new()),
            central: PreparedCentral::Generated(Vec::new()),
            generated_payload: None,
            omitted: true,
        },
    ];
    let error = synthetic.validate_output_layout(&mut prepared).unwrap_err();
    assert!(matches!(
        error.kind(),
        crate::ErrorKind::UnsupportedPreservation {
            reason: "generated replay local length overflow"
        }
    ));
}

#[test]
fn replay_limits_refuse_before_first_publication_byte() {
    let source = source_archive();
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let target = target_id(&index);

    for (limits, resource) in [
        (
            ReplayLimits::new(3, 1024, 1024 * 1024).unwrap(),
            ReplayResource::DecodedBytes,
        ),
        (
            ReplayLimits::new(1024, 3, 1024 * 1024).unwrap(),
            ReplayResource::CompressedBytes,
        ),
    ] {
        let mut sink = Vec::new();
        let result = index.write_replacing_with_replay(
            target,
            CompressionMethod::Store,
            limits,
            &mut sink,
            |writer| {
                writer
                    .write_all(b"1234")
                    .map_err(|error| error.to_string())?;
                Ok::<(), String>(())
            },
        );
        let error = result.expect_err("payload limits must fail before publication");
        assert!(matches!(&error, ReplayPublicationError::Limit { .. }));
        assert_eq!(error.pass(), ReplayPass::Measure);
        assert_eq!(error.progress(), ReplayProgress::Untouched);
        assert_eq!(sink.len(), 0);
        assert!(matches!(
            error,
            ReplayPublicationError::Limit {
                resource: actual,
                ..
            } if actual == resource
        ));
    }

    let mut sink = Vec::new();
    let result = index.write_replacing_with_replay(
        target,
        CompressionMethod::Store,
        ReplayLimits::new(1024, 1024, 1).unwrap(),
        &mut sink,
        |_| Ok::<(), String>(()),
    );
    let error = result.expect_err("archive limit must fail before publication");
    assert!(matches!(
        error,
        ReplayPublicationError::Limit {
            pass: ReplayPass::Preflight,
            resource: ReplayResource::ArchiveBytes,
            progress: ReplayProgress::Untouched,
            ..
        }
    ));
    assert!(sink.is_empty());
}

#[test]
fn replay_limits_require_nonzero_finite_ceilings() {
    for limits in [
        (0, 1, 1),
        (1, 0, 1),
        (1, 1, 0),
        (u64::MAX, 1, 1),
        (1, u64::MAX, 1),
        (1, 1, u64::MAX),
    ] {
        assert!(ReplayLimits::new(limits.0, limits.1, limits.2).is_err());
    }
}

#[test]
fn replay_wrong_index_and_name_mismatch_refuse_before_callback_and_output() {
    let source = source_archive();
    let archive = ZipArchive::from_slice(&source).unwrap().into_zip_archive();
    let mut buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let wrong_id = target_id(&index);

    let empty = ZipArchiveWriter::new(Vec::new()).finish().unwrap();
    let empty_archive = ZipArchive::from_slice(&empty).unwrap().into_zip_archive();
    let mut empty_buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let empty_index = PreservationIndex::new(&empty_archive, &mut empty_buffer).unwrap();
    let mut wrong_sink = Vec::new();
    let mut callback_calls = 0;
    let error = empty_index
        .write_replacing_with_replay(
            wrong_id,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut wrong_sink,
            |_| {
                callback_calls += 1;
                Ok::<(), String>(())
            },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Archive { .. }));
    assert_eq!(error.pass(), ReplayPass::Preflight);
    assert_eq!(error.progress(), ReplayProgress::Untouched);
    assert_eq!(callback_calls, 0);
    assert!(wrong_sink.is_empty());

    let mismatched = name_mismatched_source();
    let mismatched_archive = ZipArchive::from_slice(&mismatched)
        .unwrap()
        .into_zip_archive();
    let mut mismatched_buffer = vec![0; RECOMMENDED_BUFFER_SIZE];
    let mismatched_index =
        PreservationIndex::new(&mismatched_archive, &mut mismatched_buffer).unwrap();
    let mismatched_id = entry_id(&mismatched_index, "keep.bin");
    let mut mismatch_sink = Vec::new();
    let error = mismatched_index
        .write_replacing_with_replay(
            mismatched_id,
            CompressionMethod::Store,
            ReplayLimits::default(),
            &mut mismatch_sink,
            |_| -> Result<(), String> { panic!("name mismatch must be rejected before callback") },
        )
        .unwrap_err();
    assert!(matches!(&error, ReplayPublicationError::Archive { .. }));
    assert_eq!(error.pass(), ReplayPass::Preflight);
    assert_eq!(error.progress(), ReplayProgress::Untouched);
    assert!(mismatch_sink.is_empty());
}
