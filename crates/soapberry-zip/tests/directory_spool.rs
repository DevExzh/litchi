#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The directory-spool tests use small deterministic fixtures."
)]

//! Differential and failure coverage for the explicit central-directory spool.
//!
//! These tests deliberately keep the spool in a caller-owned `Read + Write +
//! Seek` object.  They therefore exercise the storage contract without
//! relying on a temporary path or an ambient filesystem.

use flate2::Compression;
use flate2::write::DeflateEncoder;
use soapberry_zip::extra_fields::ExtraFieldId;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveLimits, StreamingArchiveWriter};
use soapberry_zip::time::UtcDateTime;
use soapberry_zip::{
    CompressionMethod, DirectorySpoolLimits, ErrorKind, Header, ZipArchive, ZipArchiveWriter,
    ZipOperationAccounting,
};
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

const REPLAY_BUFFER: usize = 7;
const SPOOL_LIMIT: u64 = 1 << 20;
const STORED_PAYLOAD: &[u8] = b"stored payload with a stable byte layout";
const DEFLATED_PAYLOAD: &[u8] =
    b"deflated payload: repeated repeated repeated bytes for deterministic output";

fn spool_limits(max_bytes: u64) -> DirectorySpoolLimits {
    DirectorySpoolLimits::new(max_bytes, REPLAY_BUFFER)
}

fn timestamp() -> UtcDateTime {
    UtcDateTime::from_components(2024, 2, 3, 4, 5, 6, 0).expect("valid timestamp")
}

fn deflate(payload: &[u8]) -> Vec<u8> {
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(payload).expect("deflate input");
    encoder.finish().expect("deflate output")
}

fn baseline_archive(
    populate: for<'a, 'b> fn(
        &'a mut ZipArchiveWriter<&'b mut Cursor<Vec<u8>>>,
    ) -> Result<(), soapberry_zip::Error>,
) -> Result<Vec<u8>, soapberry_zip::Error> {
    let mut output = Cursor::new(Vec::<u8>::new());
    let mut archive = ZipArchiveWriter::new(&mut output);
    populate(&mut archive)?;
    archive.finish()?;
    Ok(output.into_inner())
}

fn spooled_archive(
    populate: for<'a, 'b> fn(
        &'a mut ZipArchiveWriter<&'b mut Cursor<Vec<u8>>>,
    ) -> Result<(), soapberry_zip::Error>,
) -> Result<Vec<u8>, soapberry_zip::Error> {
    let mut output = Cursor::new(Vec::new());
    let spool = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder().build_with_spool(
        &mut output,
        spool,
        spool_limits(SPOOL_LIMIT),
    )?;
    populate(&mut archive)?;
    archive.finish()?;
    Ok(output.into_inner())
}

fn populate_sized<W: Write>(archive: &mut ZipArchiveWriter<W>) -> Result<(), soapberry_zip::Error> {
    archive.write_stored_file("stored.bin", STORED_PAYLOAD)?;
    let compressed = deflate(DEFLATED_PAYLOAD);
    archive.write_precompressed_file(
        "precompressed.bin",
        CompressionMethod::Deflate,
        soapberry_zip::crc32(DEFLATED_PAYLOAD),
        DEFLATED_PAYLOAD.len() as u64,
        &compressed,
    )?;
    archive
        .new_dir("folder/")
        .last_modified(timestamp())
        .unix_permissions(0o755)
        .extra_field(ExtraFieldId::new(0xface), &[1, 2, 3, 4], Header::CENTRAL)?
        .create()?;
    Ok(())
}

fn populate_sized_concrete(
    archive: &mut ZipArchiveWriter<&mut Cursor<Vec<u8>>>,
) -> Result<(), soapberry_zip::Error> {
    populate_sized(archive)
}

fn write_borrowed<W: Write>(
    archive: &mut ZipArchiveWriter<W>,
    name: &str,
    method: CompressionMethod,
    payload: &[u8],
    zip64: bool,
) -> Result<(), soapberry_zip::Error> {
    let builder = archive
        .new_file(name)
        .compression_method(method)
        .zip64(zip64)
        .last_modified(timestamp())
        .unix_permissions(0o644)
        .extra_field(ExtraFieldId::new(0xbeef), &[9, 8, 7], Header::CENTRAL)?;
    let (mut entry, config) = builder.start()?;
    match method {
        CompressionMethod::Store => {
            let mut writer = config.wrap(&mut entry);
            writer.write_all(payload)?;
            let (_, descriptor) = writer.finish()?;
            entry.finish(descriptor)?;
        },
        CompressionMethod::Deflate => {
            let encoder = DeflateEncoder::new(&mut entry, Compression::default());
            let mut writer = config.wrap(encoder);
            writer.write_all(payload)?;
            let (encoder, descriptor) = writer.finish()?;
            encoder.finish()?;
            entry.finish(descriptor)?;
        },
        _ => unreachable!("the fixture uses only Store and Deflate"),
    }
    Ok(())
}

fn populate_borrowed<W: Write>(
    archive: &mut ZipArchiveWriter<W>,
) -> Result<(), soapberry_zip::Error> {
    write_borrowed(
        archive,
        "borrowed-store.bin",
        CompressionMethod::Store,
        STORED_PAYLOAD,
        false,
    )?;
    write_borrowed(
        archive,
        "borrowed-deflate.bin",
        CompressionMethod::Deflate,
        DEFLATED_PAYLOAD,
        false,
    )?;
    Ok(())
}

fn populate_borrowed_concrete(
    archive: &mut ZipArchiveWriter<&mut Cursor<Vec<u8>>>,
) -> Result<(), soapberry_zip::Error> {
    populate_borrowed(archive)
}

fn populate_borrowed_zip64<W: Write>(
    archive: &mut ZipArchiveWriter<W>,
) -> Result<(), soapberry_zip::Error> {
    write_borrowed(
        archive,
        "borrowed-zip64.bin",
        CompressionMethod::Store,
        STORED_PAYLOAD,
        true,
    )
}

fn populate_borrowed_zip64_concrete(
    archive: &mut ZipArchiveWriter<&mut Cursor<Vec<u8>>>,
) -> Result<(), soapberry_zip::Error> {
    populate_borrowed_zip64(archive)
}

fn populate_owned<W: Write>(
    archive: ZipArchiveWriter<W>,
) -> Result<ZipArchiveWriter<W>, soapberry_zip::Error> {
    let mut entry = archive.start_file_owned("owned-store.bin", CompressionMethod::Store)?;
    entry.write_all(STORED_PAYLOAD)?;
    let archive = entry.finish()?;

    let mut entry = archive.start_file_owned("owned-deflate.bin", CompressionMethod::Deflate)?;
    entry.write_all(DEFLATED_PAYLOAD)?;
    entry.finish()
}

fn baseline_owned() -> Result<Vec<u8>, soapberry_zip::Error> {
    let mut output = Cursor::new(Vec::new());
    let archive = ZipArchiveWriter::new(&mut output);
    let archive = populate_owned(archive)?;
    archive.finish()?;
    Ok(output.into_inner())
}

fn spooled_owned() -> Result<Vec<u8>, soapberry_zip::Error> {
    let mut output = Cursor::new(Vec::new());
    let archive = ZipArchiveWriter::builder().build_with_spool(
        &mut output,
        Cursor::new(Vec::new()),
        spool_limits(SPOOL_LIMIT),
    )?;
    let archive = populate_owned(archive)?;
    archive.finish()?;
    Ok(output.into_inner())
}

#[test]
fn explicit_spool_preserves_exact_bytes_for_sized_entries_and_directories() {
    let baseline = baseline_archive(populate_sized_concrete).expect("baseline archive");
    let spooled = spooled_archive(populate_sized_concrete).expect("spooled archive");
    assert_eq!(spooled, baseline);
    let archive = ZipArchive::from_slice(&spooled).expect("spooled ZIP");
    assert_eq!(archive.entries_hint(), 3);
}

#[test]
fn explicit_spool_preserves_exact_bytes_for_borrowed_descriptors_and_attributes() {
    let baseline = baseline_archive(populate_borrowed_concrete).expect("baseline archive");
    let spooled = spooled_archive(populate_borrowed_concrete).expect("spooled archive");
    assert_eq!(spooled, baseline);
    let reader = ArchiveReader::new(&spooled).expect("office reader");
    assert_eq!(
        reader.file_names().collect::<Vec<_>>(),
        ["borrowed-store.bin", "borrowed-deflate.bin",]
    );
}

#[test]
fn explicit_spool_preserves_exact_bytes_for_owned_store_and_deflate() {
    let baseline = baseline_owned().expect("baseline archive");
    let spooled = spooled_owned().expect("spooled archive");
    assert_eq!(spooled, baseline);
}

#[test]
fn explicit_spool_preserves_explicit_zip64_entry_layout() {
    let baseline = baseline_archive(populate_borrowed_zip64_concrete).expect("baseline archive");
    let spooled = spooled_archive(populate_borrowed_zip64_concrete).expect("spooled archive");
    assert_eq!(spooled, baseline);
    assert!(
        ZipArchive::from_slice(&spooled)
            .expect("spooled ZIP")
            .is_zip64()
    );
}

#[test]
fn exact_record_limit_succeeds_and_one_byte_below_refuses_before_output() {
    // A central record is 46 fixed bytes plus the one-byte member name.
    let exact = DirectorySpoolLimits::new(47, REPLAY_BUFFER);
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, Cursor::new(Vec::new()), exact)
        .expect("exact-limit construction");
    archive
        .write_stored_file("a", b"payload")
        .expect("exact central record fits");
    archive.finish().expect("exact-limit finish");
    assert!(!output.get_ref().is_empty());

    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, Cursor::new(Vec::new()), spool_limits(46))
        .expect("one-byte-below construction");
    let error = archive
        .write_stored_file("a", b"payload")
        .expect_err("central record should exceed the limit");
    assert!(matches!(
        error.kind(),
        ErrorKind::CentralDirectorySpoolLimitExceeded {
            actual: 47,
            maximum: 46
        }
    ));
    assert!(
        output.get_ref().is_empty(),
        "refusal must precede local output"
    );
}

#[test]
fn invalid_spool_initialization_writes_no_output() {
    let mut output = Cursor::new(Vec::<u8>::new());
    let error = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            Cursor::new(Vec::<u8>::new()),
            DirectorySpoolLimits::new(SPOOL_LIMIT, 0),
        )
        .expect_err("zero replay buffer is invalid");
    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    assert!(output.get_ref().is_empty());

    let mut output = Cursor::new(Vec::<u8>::new());
    let error = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, SeekFailureSpool, spool_limits(SPOOL_LIMIT))
        .expect_err("initial spool seek should fail");
    assert!(matches!(
        error.kind(),
        ErrorKind::CentralDirectorySpool {
            operation: "position",
            ..
        }
    ));
    assert!(output.get_ref().is_empty());
}

#[derive(Debug)]
struct SeekFailureSpool;

impl Read for SeekFailureSpool {
    fn read(&mut self, _output: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::Unsupported,
            "read unavailable",
        ))
    }
}

impl Write for SeekFailureSpool {
    fn write(&mut self, _input: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::Unsupported,
            "write unavailable",
        ))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for SeekFailureSpool {
    fn seek(&mut self, _position: SeekFrom) -> io::Result<u64> {
        Err(io::Error::new(
            io::ErrorKind::PermissionDenied,
            "seek unavailable",
        ))
    }
}

#[derive(Debug)]
struct WrongOffsetSpool {
    inner: Cursor<Vec<u8>>,
    seek_calls: usize,
}

impl WrongOffsetSpool {
    fn new() -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            seek_calls: 0,
        }
    }
}

impl Read for WrongOffsetSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read(output)
    }
}

impl Write for WrongOffsetSpool {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.inner.write(input)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for WrongOffsetSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.seek_calls += 1;
        let offset = self.inner.seek(position)?;
        if self.seek_calls == 2 {
            Ok(offset + 1)
        } else {
            Ok(offset)
        }
    }
}

#[derive(Debug, Clone)]
struct SharedBytes(Arc<Mutex<Vec<u8>>>);

impl SharedBytes {
    fn new(initial: &[u8]) -> (Self, Arc<Mutex<Vec<u8>>>) {
        let shared = Arc::new(Mutex::new(initial.to_vec()));
        (Self(Arc::clone(&shared)), shared)
    }
}

#[derive(Debug)]
struct SharedSpool {
    bytes: SharedBytes,
    position: u64,
    read_chunk: Option<usize>,
    write_chunk: Option<usize>,
}

impl SharedSpool {
    fn empty(read_chunk: Option<usize>, write_chunk: Option<usize>) -> (Self, Arc<Mutex<Vec<u8>>>) {
        let (bytes, shared) = SharedBytes::new(&[]);
        (
            Self {
                bytes,
                position: 0,
                read_chunk,
                write_chunk,
            },
            shared,
        )
    }

    fn with_prefix(prefix: &[u8]) -> (Self, Arc<Mutex<Vec<u8>>>) {
        let (bytes, shared) = SharedBytes::new(prefix);
        (
            Self {
                bytes,
                position: 0,
                read_chunk: None,
                write_chunk: None,
            },
            shared,
        )
    }
}

impl Read for SharedSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.0.lock().expect("spool lock");
        let start = usize::try_from(self.position).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "spool position overflows usize",
            )
        })?;
        if start >= bytes.len() || output.is_empty() {
            return Ok(0);
        }
        let amount = output
            .len()
            .min(bytes.len() - start)
            .min(self.read_chunk.unwrap_or(usize::MAX));
        output[..amount].copy_from_slice(&bytes[start..start + amount]);
        self.position += amount as u64;
        Ok(amount)
    }
}

impl Write for SharedSpool {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        let mut bytes = self.bytes.0.lock().expect("spool lock");
        let start = usize::try_from(self.position).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "spool position overflows usize",
            )
        })?;
        let amount = input.len().min(self.write_chunk.unwrap_or(usize::MAX));
        let end = start
            .checked_add(amount)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "spool length overflow"))?;
        if bytes.len() < end {
            bytes.resize(end, 0);
        }
        bytes[start..end].copy_from_slice(&input[..amount]);
        self.position = end as u64;
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for SharedSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let length = self.bytes.0.lock().expect("spool lock").len() as i128;
        let current = i128::from(self.position);
        let next = match position {
            SeekFrom::Start(value) => i128::from(value),
            SeekFrom::End(value) => length + i128::from(value),
            SeekFrom::Current(value) => current + i128::from(value),
        };
        if next < 0 || next > i128::from(u64::MAX) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "invalid spool seek",
            ));
        }
        self.position = next as u64;
        Ok(self.position)
    }
}

#[test]
fn empty_and_nonempty_supplied_spools_append_and_replay_only_the_new_extent() {
    let baseline = baseline_archive(populate_sized_concrete).expect("baseline archive");
    let central_start = ZipArchive::from_slice(&baseline)
        .expect("baseline ZIP")
        .directory_offset() as usize;
    let central = &baseline[central_start..baseline.len() - 22];

    let (empty, empty_bytes) = SharedSpool::empty(Some(2), Some(3));
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, empty, spool_limits(SPOOL_LIMIT))
        .expect("empty spool construction");
    populate_sized(&mut archive).expect("empty spool entries");
    archive.finish().expect("empty spool finish");
    assert_eq!(&*empty_bytes.lock().expect("empty spool lock"), central);
    assert_eq!(output.into_inner(), baseline);

    let prefix = b"caller-owned prefix";
    let (nonempty, nonempty_bytes) = SharedSpool::with_prefix(prefix);
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, nonempty, spool_limits(SPOOL_LIMIT))
        .expect("nonempty spool construction");
    populate_sized(&mut archive).expect("nonempty spool entries");
    archive.finish().expect("nonempty spool finish");
    let stored = nonempty_bytes.lock().expect("nonempty spool lock");
    assert_eq!(&stored[..prefix.len()], prefix);
    assert_eq!(&stored[prefix.len()..], central);
    assert_eq!(output.into_inner(), baseline);
}

#[test]
fn short_and_interrupted_spool_transfers_succeed() {
    let mut output = Cursor::new(Vec::new());
    let (spool, _) = SharedSpool::empty(Some(1), Some(1));
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, spool, spool_limits(SPOOL_LIMIT))
        .expect("short spool construction");
    archive
        .write_stored_file("short.bin", STORED_PAYLOAD)
        .expect("short spool write");
    archive.finish().expect("short spool replay");
    assert!(ZipArchive::from_slice(output.get_ref()).is_ok());

    let mut output = Cursor::new(Vec::new());
    let spool = InterruptedSpool::for_write();
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, spool, spool_limits(SPOOL_LIMIT))
        .expect("interrupted spool construction");
    archive
        .write_stored_file("interrupted.bin", STORED_PAYLOAD)
        .expect("Interrupted is retryable");
    archive.finish().expect("interrupted replay");
    assert!(ZipArchive::from_slice(output.get_ref()).is_ok());

    let mut output = Cursor::new(Vec::new());
    let spool = InterruptedSpool::for_read();
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(&mut output, spool, spool_limits(SPOOL_LIMIT))
        .expect("interrupted-read construction");
    archive
        .write_stored_file("interrupted-read.bin", STORED_PAYLOAD)
        .expect("interrupted-read entry");
    archive.finish().expect("Interrupted read is retryable");
    assert!(ZipArchive::from_slice(output.get_ref()).is_ok());
}

#[derive(Debug)]
struct InterruptedSpool {
    inner: Cursor<Vec<u8>>,
    interrupt_write: bool,
    interrupt_read: bool,
}

impl InterruptedSpool {
    fn for_write() -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            interrupt_write: true,
            interrupt_read: false,
        }
    }

    fn for_read() -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            interrupt_write: false,
            interrupt_read: true,
        }
    }
}

impl Read for InterruptedSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.interrupt_read {
            self.interrupt_read = false;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry read"));
        }
        self.inner.read(output)
    }
}

impl Write for InterruptedSpool {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.interrupt_write {
            self.interrupt_write = false;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry write"));
        }
        self.inner.write(input)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for InterruptedSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

#[derive(Debug)]
struct ReplayFailureSpool {
    inner: Cursor<Vec<u8>>,
    read_after: Option<usize>,
    read_total: usize,
    fail_replay_seek: bool,
    seek_calls: usize,
}

impl ReplayFailureSpool {
    fn zero_after() -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            read_after: Some(0),
            read_total: 0,
            fail_replay_seek: false,
            seek_calls: 0,
        }
    }

    fn truncated_after(bytes: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            read_after: Some(bytes),
            read_total: 0,
            fail_replay_seek: false,
            seek_calls: 0,
        }
    }

    fn seek_failure() -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            read_after: None,
            read_total: 0,
            fail_replay_seek: true,
            seek_calls: 0,
        }
    }
}

impl Read for ReplayFailureSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if let Some(limit) = self.read_after {
            if self.read_total >= limit {
                return Ok(0);
            }
            let amount = self
                .inner
                .get_ref()
                .len()
                .saturating_sub(self.inner.position() as usize)
                .min(output.len())
                .min(limit - self.read_total);
            if amount == 0 {
                return Ok(0);
            }
            self.inner.read_exact(&mut output[..amount])?;
            self.read_total += amount;
            return Ok(amount);
        }
        self.inner.read(output)
    }
}

impl Write for ReplayFailureSpool {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.inner.write(input)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for ReplayFailureSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.seek_calls += 1;
        if self.fail_replay_seek && self.seek_calls > 1 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "replay seek failed",
            ));
        }
        self.inner.seek(position)
    }
}

#[test]
fn zero_and_truncated_replay_poison_finish_after_partial_spool() {
    for spool in [
        ReplayFailureSpool::zero_after(),
        ReplayFailureSpool::truncated_after(3),
    ] {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::builder()
            .build_with_spool(&mut output, spool, spool_limits(SPOOL_LIMIT))
            .expect("replay failure construction");
        archive
            .write_stored_file("replay.bin", STORED_PAYLOAD)
            .expect("entry reaches spool");
        let error = archive.finish().expect_err("replay must fail");
        assert!(matches!(
            error.kind(),
            ErrorKind::CentralDirectorySpool { .. }
        ));
    }
}

#[test]
fn replay_seek_failure_is_typed() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            ReplayFailureSpool::seek_failure(),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("seek failure construction");
    archive
        .write_stored_file("seek.bin", STORED_PAYLOAD)
        .expect("seek entry");
    let error = archive.finish().expect_err("replay seek must fail");
    assert!(matches!(
        error.kind(),
        ErrorKind::CentralDirectorySpool {
            operation: "seek for replay",
            ..
        }
    ));
}

#[test]
fn replay_seek_wrong_offset_is_typed_refusal() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            WrongOffsetSpool::new(),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("wrong-offset construction");
    archive
        .write_stored_file("wrong-offset.bin", STORED_PAYLOAD)
        .expect("wrong-offset entry");
    let error = archive
        .finish()
        .expect_err("wrong replay seek offset must fail");
    match error.kind() {
        ErrorKind::CentralDirectorySpool { operation, source } => {
            assert_eq!(*operation, "seek for replay");
            assert_eq!(source.kind(), io::ErrorKind::InvalidData);
        },
        other => panic!("unexpected wrong-offset error: {other:?}"),
    }
}

#[derive(Debug)]
struct FailingWriteSpool {
    inner: Cursor<Vec<u8>>,
    fail_after: usize,
    written: usize,
}

impl FailingWriteSpool {
    fn new(fail_after: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            fail_after,
            written: 0,
        }
    }
}

impl Read for FailingWriteSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read(output)
    }
}

impl Write for FailingWriteSpool {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.written >= self.fail_after {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "spool write failed",
            ));
        }
        let amount = input.len().min(self.fail_after - self.written);
        self.inner.write_all(&input[..amount])?;
        self.written += amount;
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for FailingWriteSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

#[derive(Debug)]
struct SensitiveWriteSpool {
    inner: Cursor<Vec<u8>>,
    message: &'static str,
}

impl SensitiveWriteSpool {
    fn new(message: &'static str) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            message,
        }
    }
}

impl Read for SensitiveWriteSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read(output)
    }
}

impl Write for SensitiveWriteSpool {
    fn write(&mut self, _input: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(io::ErrorKind::BrokenPipe, self.message))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for SensitiveWriteSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

#[test]
fn spool_io_display_omits_provider_text_but_source_retains_it() {
    const SENSITIVE: &str = "injected sensitive provider text";
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            SensitiveWriteSpool::new(SENSITIVE),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("sensitive spool construction");
    let error = archive
        .write_stored_file("sensitive.bin", STORED_PAYLOAD)
        .expect_err("sensitive spool write must fail");
    assert!(matches!(
        error.kind(),
        ErrorKind::CentralDirectorySpool { .. }
    ));
    assert!(!error.to_string().contains(SENSITIVE));
    let source = std::error::Error::source(&error).expect("spool source retained");
    assert!(source.to_string().contains(SENSITIVE));
}

#[test]
fn partial_spool_write_poison_is_retained() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            FailingWriteSpool::new(4),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("partial-write construction");
    let error = archive
        .write_stored_file("partial.bin", STORED_PAYLOAD)
        .expect_err("partial spool must fail");
    assert!(matches!(
        error.kind(),
        ErrorKind::CentralDirectorySpool {
            operation: "write central-directory record",
            ..
        }
    ));
    let finish = archive
        .finish()
        .expect_err("poisoned spool must not replay");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
    assert!(
        !output.get_ref().is_empty(),
        "payload precedes spool failure"
    );
}

#[derive(Debug)]
struct OverreportSpool(Cursor<Vec<u8>>);

impl Read for OverreportSpool {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        self.0.read(output)
    }
}

impl Write for OverreportSpool {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        Ok(input.len() + 1)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for OverreportSpool {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.0.seek(position)
    }
}

#[test]
fn overreporting_spool_is_rejected_without_a_process_panic() {
    let mut output = Cursor::new(Vec::new());
    let result = catch_unwind(AssertUnwindSafe(|| {
        let mut archive = ZipArchiveWriter::builder()
            .build_with_spool(
                &mut output,
                OverreportSpool(Cursor::new(Vec::new())),
                spool_limits(SPOOL_LIMIT),
            )
            .expect("overreport construction");
        archive.write_stored_file("over.bin", STORED_PAYLOAD)
    }));
    let result = result.expect("overreporting must not panic");
    let error = result.expect_err("overreporting must fail");
    assert!(matches!(
        error.kind(),
        ErrorKind::CentralDirectorySpool { .. }
    ));
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
        Err(io::Error::new(io::ErrorKind::WriteZero, "flush failed"))
    }
}

#[test]
fn final_output_flush_failure_is_preserved() {
    let mut sink = FlushFailSink { bytes: Vec::new() };
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut sink,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("flush-failure construction");
    archive
        .write_stored_file("flush.bin", STORED_PAYLOAD)
        .expect("flush entry");
    let error = archive.finish().expect_err("sink flush must fail");
    assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
    assert!(!sink.bytes.is_empty());
}

#[test]
fn office_spool_progress_reports_accepted_output_bytes() {
    let mut sink = Cursor::new(Vec::new());
    let limits =
        StreamingArchiveLimits::new(8, 256, 1 << 20).with_byte_limits(1 << 20, 1 << 20, 1 << 20);
    let mut writer = StreamingArchiveWriter::with_writer_and_limits_and_spool(
        &mut sink,
        limits,
        Cursor::new(Vec::new()),
        spool_limits(SPOOL_LIMIT),
    )
    .expect("office spool construction");
    writer
        .write_stored("office.bin", STORED_PAYLOAD)
        .expect("office entry");
    let (returned, progress) = writer.finish_with_progress().expect("office finish");
    let returned_len = returned.get_ref().len();
    let _ = returned;
    assert_eq!(progress.output_bytes(), sink.get_ref().len() as u64);
    assert_eq!(returned_len, sink.get_ref().len());
    assert!(ArchiveReader::new(sink.get_ref()).is_ok());
}

#[test]
fn spooled_writer_remains_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<DirectorySpoolLimits>();
    assert_send_sync::<ZipArchiveWriter<Cursor<Vec<u8>>>>();
    assert_send_sync::<StreamingArchiveWriter<Cursor<Vec<u8>>>>();
}

#[test]
fn zip64_count_layout_is_identical_with_and_without_spool() {
    let mut baseline_output = Cursor::new(Vec::new());
    let mut baseline = ZipArchiveWriter::new(&mut baseline_output);
    for index in 0..=u16::MAX as usize {
        let name = format!("d{index:05}/");
        baseline
            .new_dir(&name)
            .create()
            .expect("baseline directory");
    }
    baseline.finish().expect("baseline count finish");

    let mut spooled_output = Cursor::new(Vec::new());
    let mut spooled = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut spooled_output,
            Cursor::new(Vec::new()),
            spool_limits(8 << 20),
        )
        .expect("spooled count construction");
    for index in 0..=u16::MAX as usize {
        let name = format!("d{index:05}/");
        spooled.new_dir(&name).create().expect("spooled directory");
    }
    spooled.finish().expect("spooled count finish");

    assert_eq!(spooled_output.get_ref(), baseline_output.get_ref());
    assert!(
        ZipArchive::from_slice(spooled_output.get_ref())
            .expect("ZIP64 count archive")
            .is_zip64()
    );
}

#[derive(Debug)]
struct VirtualOffsetSink {
    bytes: Vec<u8>,
}

impl Write for VirtualOffsetSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn zip64_offset_layout_is_identical_with_and_without_spool() {
    let offset = u32::MAX as u64;
    let mut baseline = VirtualOffsetSink { bytes: Vec::new() };
    let archive = ZipArchiveWriter::builder()
        .with_offset(offset)
        .build(&mut baseline);
    let mut entry = archive
        .start_file_owned("offset.bin", CompressionMethod::Store)
        .expect("baseline offset entry");
    entry.write_all(b"offset").expect("baseline offset payload");
    let archive = entry.finish().expect("baseline offset entry finish");
    archive.finish().expect("baseline offset finish");

    let mut spooled = VirtualOffsetSink { bytes: Vec::new() };
    let archive = ZipArchiveWriter::builder()
        .with_offset(offset)
        .build_with_spool(
            &mut spooled,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("spooled offset construction");
    let mut entry = archive
        .start_file_owned("offset.bin", CompressionMethod::Store)
        .expect("spooled offset entry");
    entry.write_all(b"offset").expect("spooled offset payload");
    let archive = entry.finish().expect("spooled offset entry finish");
    archive.finish().expect("spooled offset finish");

    assert_eq!(spooled.bytes, baseline.bytes);
    assert!(spooled.bytes.windows(4).any(|bytes| bytes == b"PK\x06\x06"));
}

fn offset_sized_attempt(
    maximum: u64,
    precompressed: bool,
) -> (Result<(), soapberry_zip::Error>, Vec<u8>) {
    let mut sink = VirtualOffsetSink { bytes: Vec::new() };
    let result = {
        let mut archive = ZipArchiveWriter::builder()
            .with_offset(u32::MAX as u64)
            .build_with_spool(&mut sink, Cursor::new(Vec::new()), spool_limits(maximum))
            .expect("offset-sized construction");
        let write_result = if precompressed {
            let compressed = deflate(DEFLATED_PAYLOAD);
            archive.write_precompressed_file(
                "a",
                CompressionMethod::Deflate,
                soapberry_zip::crc32(DEFLATED_PAYLOAD),
                DEFLATED_PAYLOAD.len() as u64,
                &compressed,
            )
        } else {
            archive.write_stored_file("a", STORED_PAYLOAD)
        };
        match write_result {
            Ok(()) => archive.finish().map(|_| ()),
            Err(error) => Err(error),
        }
    };
    (result, sink.bytes)
}

#[test]
fn offset_only_zip64_sized_records_use_the_exact_spool_budget() {
    // A one-byte name plus the 46-byte central fixed header and the
    // offset-only ZIP64 extra field (4-byte header plus 8-byte offset) totals
    // 59 bytes.
    for precompressed in [false, true] {
        let (result, bytes) = offset_sized_attempt(59, precompressed);
        result.expect("offset-only ZIP64 record fits exactly");
        let central_start = bytes
            .windows(4)
            .position(|window| window == b"PK\x01\x02")
            .expect("central directory signature");
        let record = &bytes[central_start..];
        let name_len = u16::from_le_bytes([record[28], record[29]]) as usize;
        let extra_len = u16::from_le_bytes([record[30], record[31]]) as usize;
        assert_eq!(&record[46..46 + name_len], b"a");
        assert_eq!(extra_len, 12);
        let extra = &record[46 + name_len..46 + name_len + extra_len];
        assert_eq!(
            u16::from_le_bytes([extra[0], extra[1]]),
            ExtraFieldId::ZIP64.as_u16()
        );
        assert_eq!(u16::from_le_bytes([extra[2], extra[3]]), 8);
        assert_eq!(
            u64::from_le_bytes(extra[4..12].try_into().expect("ZIP64 offset bytes")),
            u32::MAX as u64
        );

        let (result, bytes) = offset_sized_attempt(58, precompressed);
        let error = result.expect_err("one byte below offset-only ZIP64 budget");
        assert!(matches!(
            error.kind(),
            ErrorKind::CentralDirectorySpoolLimitExceeded {
                actual: 59,
                maximum: 58
            }
        ));
        assert!(bytes.is_empty(), "budget refusal must precede local output");
    }
}

#[test]
fn borrowed_entry_drop_does_not_publish_a_partial_central_record() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("drop construction");
    let (entry, _config) = archive
        .new_file("abandoned.bin")
        .start()
        .expect("entry start");
    drop(entry);
    let error = archive
        .finish()
        .expect_err("abandoned entry must poison archive");
    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
}

fn fresh_spooled_archive(output: &mut Cursor<Vec<u8>>) -> ZipArchiveWriter<&mut Cursor<Vec<u8>>> {
    ZipArchiveWriter::builder()
        .build_with_spool(output, Cursor::new(Vec::new()), spool_limits(SPOOL_LIMIT))
        .expect("spooled archive construction")
}

#[test]
fn dropped_borrowed_entry_rejects_all_reentrant_writer_operations() {
    macro_rules! assert_rejected_after_drop {
        ($archive:ident, $operation:expr) => {{
            let mut output = Cursor::new(Vec::new());
            let mut $archive = fresh_spooled_archive(&mut output);
            let (entry, _config) = $archive
                .new_file("abandoned.bin")
                .start()
                .expect("abandoned entry start");
            drop(entry);
            let before = $archive.stream_offset();
            let error = $operation.expect_err("unfinished borrowed entry must reject operation");
            assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
            assert_eq!(
                $archive.stream_offset(),
                before,
                "rejected operation wrote bytes"
            );
            let finish = $archive
                .finish()
                .expect_err("unfinished borrowed entry must reject finish");
            assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
        }};
    }

    assert_rejected_after_drop!(
        archive,
        archive.write_stored_file("next-stored.bin", b"next")
    );
    assert_rejected_after_drop!(
        archive,
        archive.write_precompressed_file(
            "next-precompressed.bin",
            CompressionMethod::Store,
            soapberry_zip::crc32(b"next"),
            4,
            b"next"
        )
    );
    assert_rejected_after_drop!(archive, archive.new_dir("next/").create());
    assert_rejected_after_drop!(archive, archive.new_file("next-borrowed.bin").start());
}

#[test]
fn dropped_borrowed_entry_rejects_owned_start_without_new_output() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = fresh_spooled_archive(&mut output);
    let (entry, _config) = archive
        .new_file("abandoned.bin")
        .start()
        .expect("abandoned entry start");
    drop(entry);
    let before = archive.stream_offset();
    let error = archive
        .start_file_owned("next-owned.bin", CompressionMethod::Store)
        .expect_err("unfinished borrowed entry must reject owned start");
    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    assert_eq!(
        before,
        output.position(),
        "rejected owned start wrote bytes"
    );
}

#[test]
fn borrowed_finish_collision_poisons_archive_before_finish() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .with_offset(u32::MAX as u64)
        .build_with_spool(
            &mut output,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("collision archive construction");
    let (mut entry, config) = archive
        .new_file("collision.bin")
        .extra_field(ExtraFieldId::ZIP64, &[0; 8], Header::CENTRAL)
        .expect("user ZIP64 field")
        .start()
        .expect("collision entry start");
    let mut writer = config.wrap(&mut entry);
    writer.write_all(b"collision").expect("collision payload");
    let (_, descriptor) = writer.finish().expect("collision descriptor");
    let error = entry
        .finish(descriptor)
        .expect_err("automatic ZIP64 field collision must fail at finish");
    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    let finish = archive
        .finish()
        .expect_err("failed borrowed finalization must poison archive");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
}

#[derive(Debug)]
struct RecoveringSink {
    bytes: Vec<u8>,
    fail_write: Arc<AtomicBool>,
    write_fault: Arc<Mutex<Option<io::ErrorKind>>>,
    fail_flush: Arc<AtomicBool>,
}

impl RecoveringSink {
    fn write_failure() -> (Self, Arc<AtomicBool>) {
        let fail_write = Arc::new(AtomicBool::new(false));
        let sink = Self {
            bytes: Vec::new(),
            fail_write: Arc::clone(&fail_write),
            write_fault: Arc::new(Mutex::new(None)),
            fail_flush: Arc::new(AtomicBool::new(false)),
        };
        (sink, fail_write)
    }

    fn flush_failure() -> Self {
        Self {
            bytes: Vec::new(),
            fail_write: Arc::new(AtomicBool::new(false)),
            write_fault: Arc::new(Mutex::new(None)),
            fail_flush: Arc::new(AtomicBool::new(true)),
        }
    }

    fn armed_write_fault() -> (Self, Arc<Mutex<Option<io::ErrorKind>>>) {
        let write_fault = Arc::new(Mutex::new(None));
        let sink = Self {
            bytes: Vec::new(),
            fail_write: Arc::new(AtomicBool::new(false)),
            write_fault: Arc::clone(&write_fault),
            fail_flush: Arc::new(AtomicBool::new(false)),
        };
        (sink, write_fault)
    }
}

impl Write for RecoveringSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if let Some(kind) = self.write_fault.lock().expect("write fault lock").take() {
            if kind == io::ErrorKind::WriteZero {
                return Ok(0);
            }
            return Err(io::Error::new(kind, "one scripted write fault"));
        }
        if self.fail_write.swap(false, Ordering::AcqRel) {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "one write failed",
            ));
        }
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.fail_flush.swap(false, Ordering::AcqRel) {
            return Err(io::Error::new(io::ErrorKind::WriteZero, "one flush failed"));
        }
        Ok(())
    }
}

#[derive(Debug)]
struct PartialRecoveringSink {
    bytes: Vec<u8>,
    remaining: usize,
    fail_next: bool,
}

impl PartialRecoveringSink {
    fn new(fail_after: usize) -> Self {
        assert!(fail_after > 0);
        Self {
            bytes: Vec::new(),
            remaining: fail_after,
            fail_next: false,
        }
    }
}

impl Write for PartialRecoveringSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.fail_next {
            self.fail_next = false;
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "one partial output failure",
            ));
        }
        if input.is_empty() {
            return Ok(0);
        }
        if self.remaining < input.len() {
            let accepted = self.remaining;
            self.bytes.extend_from_slice(&input[..accepted]);
            self.remaining = 0;
            self.fail_next = true;
            return Ok(accepted);
        }
        self.bytes.extend_from_slice(input);
        self.remaining -= input.len();
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn no_spool_partial_sized_write_poisons_future_writes_and_finish() {
    let mut sink = PartialRecoveringSink::new(40);
    let mut archive = ZipArchiveWriter::new(&mut sink);
    let error = archive
        .write_stored_file("sized.bin", STORED_PAYLOAD)
        .expect_err("partial sized output must fail");
    assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));

    let next = archive
        .write_stored_file("after.bin", b"after")
        .expect_err("partial sized failure must poison later writes");
    assert!(matches!(next.kind(), ErrorKind::InvalidInput { .. }));
    let finish = archive
        .finish()
        .expect_err("partial sized failure must poison finish");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
}

#[test]
fn no_spool_partial_directory_header_poisons_future_writes_and_finish() {
    let mut sink = PartialRecoveringSink::new(4);
    let mut archive = ZipArchiveWriter::new(&mut sink);
    let error = archive
        .new_dir("partial/")
        .create()
        .expect_err("partial directory header must fail");
    assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));

    let next = archive
        .write_stored_file("after.bin", b"after")
        .expect_err("partial directory failure must poison later writes");
    assert!(matches!(next.kind(), ErrorKind::InvalidInput { .. }));
    let finish = archive
        .finish()
        .expect_err("partial directory failure must poison finish");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
}

#[test]
fn ignored_borrowed_sink_write_error_cannot_finish_successfully() {
    let (mut sink, fail_write) = RecoveringSink::write_failure();
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut sink,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("write-failure archive construction");
    let (mut entry, config) = archive
        .new_file("write-failure.bin")
        .start()
        .expect("write-failure entry start");
    fail_write.store(true, Ordering::Release);
    let mut writer = config.wrap(&mut entry);
    let write_error = writer
        .write(STORED_PAYLOAD)
        .expect_err("sink write must fail once");
    assert_eq!(write_error.kind(), io::ErrorKind::BrokenPipe);
    let (_, descriptor) = writer
        .finish()
        .expect("recovered sink still supplies a descriptor");
    let finish = entry
        .finish(descriptor)
        .expect_err("ignored sink write failure must poison entry publication");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
    let finish = archive
        .finish()
        .expect_err("poisoned borrowed sink write must reject archive finish");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
}

#[test]
fn ignored_borrowed_sink_flush_error_cannot_finish_successfully() {
    let mut sink = RecoveringSink::flush_failure();
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut sink,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("flush-failure archive construction");
    let (mut entry, config) = archive
        .new_file("flush-failure.bin")
        .start()
        .expect("flush-failure entry start");
    let mut writer = config.wrap(&mut entry);
    writer.write_all(STORED_PAYLOAD).expect("payload write");
    let flush_error = writer.flush().expect_err("sink flush must fail once");
    assert_eq!(flush_error.kind(), io::ErrorKind::WriteZero);
    let (_, descriptor) = writer
        .finish()
        .expect("recovered sink still supplies a descriptor");
    let finish = entry
        .finish(descriptor)
        .expect_err("ignored sink flush failure must poison entry publication");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
    let finish = archive
        .finish()
        .expect_err("poisoned borrowed sink flush must reject archive finish");
    assert!(matches!(finish.kind(), ErrorKind::InvalidInput { .. }));
}

#[test]
fn borrowed_output_interrupted_once_then_write_all_and_finish_succeed() {
    let (mut sink, write_fault) = RecoveringSink::armed_write_fault();
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut sink,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("Interrupted archive construction");
    let (mut entry, config) = archive
        .new_file("interrupted-output.bin")
        .start()
        .expect("Interrupted entry start");
    *write_fault.lock().expect("write fault lock") = Some(io::ErrorKind::Interrupted);
    let mut writer = config.wrap(&mut entry);
    writer
        .write_all(STORED_PAYLOAD)
        .expect("write_all retries Interrupted");
    let (_, descriptor) = writer.finish().expect("Interrupted descriptor");
    entry.finish(descriptor).expect("Interrupted entry finish");
    archive.finish().expect("Interrupted archive finish");
    assert!(ZipArchive::from_slice(&sink.bytes).is_ok());
}

#[test]
fn borrowed_output_zero_write_poison_rejects_finish() {
    let (mut sink, write_fault) = RecoveringSink::armed_write_fault();
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut sink,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("zero-write archive construction");
    let (mut entry, config) = archive
        .new_file("zero-output.bin")
        .start()
        .expect("zero-write entry start");
    *write_fault.lock().expect("write fault lock") = Some(io::ErrorKind::WriteZero);
    let mut writer = config.wrap(&mut entry);
    let write_error = writer
        .write_all(STORED_PAYLOAD)
        .expect_err("write_all must report a non-progressing sink");
    assert_eq!(write_error.kind(), io::ErrorKind::WriteZero);
    let (_, descriptor) = writer.finish().expect("zero-write descriptor");
    let error = entry
        .finish(descriptor)
        .expect_err("ignored zero-write failure must reject publication");
    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    let error = archive
        .finish()
        .expect_err("zero-write poison must reject archive finish");
    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
}

#[test]
fn directory_at_zip64_offset_preserves_and_decodes_offset_extra() {
    let offset = u32::MAX as u64;
    let mut baseline_sink = VirtualOffsetSink { bytes: Vec::new() };
    let mut baseline = ZipArchiveWriter::builder()
        .with_offset(offset)
        .build(&mut baseline_sink);
    baseline
        .new_dir("virtual/")
        .create()
        .expect("baseline directory");
    baseline.finish().expect("baseline directory finish");

    let mut spooled_sink = VirtualOffsetSink { bytes: Vec::new() };
    let mut spooled = ZipArchiveWriter::builder()
        .with_offset(offset)
        .build_with_spool(
            &mut spooled_sink,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("spooled directory construction");
    spooled
        .new_dir("virtual/")
        .create()
        .expect("spooled directory");
    spooled.finish().expect("spooled directory finish");
    assert_eq!(spooled_sink.bytes, baseline_sink.bytes);

    let central_start = spooled_sink
        .bytes
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
        .expect("central directory signature");
    let record = &spooled_sink.bytes[central_start..];
    let name_len = u16::from_le_bytes([record[28], record[29]]) as usize;
    let extra_len = u16::from_le_bytes([record[30], record[31]]) as usize;
    assert_eq!(&record[46..46 + name_len], b"virtual/");
    let extra = &record[46 + name_len..46 + name_len + extra_len];
    assert_eq!(
        u16::from_le_bytes([extra[0], extra[1]]),
        ExtraFieldId::ZIP64.as_u16()
    );
    assert_eq!(u16::from_le_bytes([extra[2], extra[3]]), 8);
    assert_eq!(
        u64::from_le_bytes(extra[4..12].try_into().expect("ZIP64 offset bytes")),
        offset
    );
}

#[test]
fn local_only_extra_fields_do_not_consume_central_spool_quota() {
    let name = "local-only.bin";
    let exact_central_record = 46 + name.len() as u64;
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            Cursor::new(Vec::new()),
            spool_limits(exact_central_record),
        )
        .expect("local-only extra archive construction");
    let (mut entry, config) = archive
        .new_file(name)
        .extra_field(ExtraFieldId::new(0xcafe), &[1, 2, 3, 4], Header::LOCAL)
        .expect("local-only extra field")
        .start()
        .expect("local-only entry start");
    let mut writer = config.wrap(&mut entry);
    writer
        .write_all(b"local metadata")
        .expect("local-only payload");
    let (_, descriptor) = writer.finish().expect("local-only descriptor");
    entry.finish(descriptor).expect("exact central quota");
    archive.finish().expect("local-only finish");

    let bytes = output.into_inner();
    let archive = ZipArchive::from_slice(&bytes).expect("local-only ZIP");
    let record = archive
        .entries()
        .next_entry()
        .expect("local-only record result")
        .expect("local-only record");
    assert!(
        record
            .extra_fields()
            .all(|(id, _)| id != ExtraFieldId::new(0xcafe))
    );
    let local_extra_offset = 30 + name.len();
    assert_eq!(
        &bytes[local_extra_offset..local_extra_offset + 8],
        &[0xfe, 0xca, 4, 0, 1, 2, 3, 4],
        "the local header must preserve the complete little-endian extra field"
    );
}

#[test]
fn accounting_through_spooled_sized_writer_counts_accepted_payload() {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::builder()
        .build_with_spool(
            &mut output,
            Cursor::new(Vec::new()),
            spool_limits(SPOOL_LIMIT),
        )
        .expect("accounting construction");
    let mut accounting = ZipOperationAccounting::default();
    archive
        .write_stored_file_with_accounting("accounted.bin", STORED_PAYLOAD, &mut accounting)
        .expect("accounted write");
    assert_eq!(
        accounting.stored_payload_bytes_emitted(),
        STORED_PAYLOAD.len() as u64
    );
    archive.finish().expect("accounting finish");
}
