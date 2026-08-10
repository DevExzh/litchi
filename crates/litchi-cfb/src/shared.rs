//! Immutable, positional CFB reads backed by [`litchi_core::ReadAt`].

use crate::shared_bulk::SharedOleBulkRead;
use crate::{
    consts::{ENDOFCHAIN, MAXREGSECT, STGTY_STREAM},
    directory_name::directory_name_data,
    file::{DirectoryEntry, OleError, OleFile, ParsedOleIndex},
};
use litchi_core::{ExecutionContext, ReadAt, SourceVersion};
use std::{
    cmp::Ordering,
    io::{self, Read, Seek, SeekFrom},
    sync::{Arc, Mutex},
};

/// Bounded ingress settings for [`SharedOleFile`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SharedOleFileLimits {
    max_input_bytes: u64,
}

impl SharedOleFileLimits {
    /// Largest CFB input accepted by the default shared reader.
    pub const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

    /// Creates a finite input ceiling for one positional CFB source.
    ///
    /// # Errors
    ///
    /// Returns an error if the ceiling is zero or exceeds the CFB shared
    /// reader's hard ingress ceiling.
    pub fn new(max_input_bytes: u64) -> Result<Self, OleError> {
        if max_input_bytes == 0 || max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(OleError::InvalidData(format!(
                "shared CFB input limit must be between 1 and {} bytes",
                Self::MAX_INPUT_BYTES
            )));
        }
        Ok(Self { max_input_bytes })
    }

    /// Maximum source length accepted before parsing begins.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }
}

impl Default for SharedOleFileLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
        }
    }
}

/// One parsed, immutable CFB view over a thread-safe positional source.
///
/// Opening this type runs the existing [`OleFile`] validation pipeline once.
/// The validation cursor is then discarded; regular stream reads address the
/// [`ReadAt`] source directly and can run concurrently without a shared seek
/// cursor or reader lock. Mini-stream bytes remain lazy and are initialized at
/// most once, with failures left retryable.
pub struct SharedOleFile {
    source: Arc<dyn ReadAt>,
    expected_version: SourceVersion,
    index: Arc<ParsedOleIndex>,
    /// Serializes only lazy mini-stream initialization. Regular streams never
    /// acquire this lock or any shared cursor lock.
    ministream: Mutex<Option<Arc<[u8]>>>,
}

impl std::fmt::Debug for SharedOleFile {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SharedOleFile")
            .field("file_size", &self.index.file_size)
            .field("sector_size", &self.index.sector_size)
            .finish_non_exhaustive()
    }
}

impl SharedOleFile {
    /// Starts an explicit, bounded bulk-read session.
    ///
    /// This is the only shared-reader API that may schedule work. Normal
    /// [`Self::open_stream`] calls remain serial and create no runtime.
    #[must_use]
    pub fn bulk_read(&self, context: ExecutionContext) -> SharedOleBulkRead<'_> {
        SharedOleBulkRead::new(self, context)
    }

    /// Opens and validates a positional source under the default finite input
    /// ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds the default limit, changes
    /// during parsing, or is not a valid CFB file.
    pub fn open(source: Arc<dyn ReadAt>) -> Result<Self, OleError> {
        Self::open_with_limits(source, SharedOleFileLimits::default())
    }

    /// Opens and validates a positional source under caller-selected limits.
    ///
    /// The source version is captured before parsing and compared after the
    /// existing cursor-based parser has completely validated allocation chains,
    /// overlaps, directory trees, and final-sector truncation rules.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds the limit, changes during
    /// parsing, or fails the normal [`OleFile`] validation rules.
    pub fn open_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SharedOleFileLimits,
    ) -> Result<Self, OleError> {
        let expected_version = source.version()?;
        let source_length = source.len();
        let observed = source.version()?;
        if observed != expected_version {
            return Err(OleError::SourceChanged {
                expected: expected_version,
                observed,
            });
        }
        let source_length = source_length?;
        if source_length > limits.max_input_bytes() {
            return Err(OleError::InvalidData(format!(
                "shared CFB input length {source_length} exceeds configured limit {}",
                limits.max_input_bytes()
            )));
        }

        let parsed = OleFile::open(ReadAtCursor::new(source.clone(), source_length));
        let observed = source.version()?;
        if observed != expected_version {
            return Err(OleError::SourceChanged {
                expected: expected_version,
                observed,
            });
        }
        let index = parsed?.into_parsed_index();

        Ok(Self {
            source,
            expected_version,
            index: Arc::new(index),
            ministream: Mutex::new(None),
        })
    }

    /// Physical length captured while parsing this CFB file.
    #[must_use]
    pub fn file_size(&self) -> u64 {
        self.index.file_size
    }

    /// Returns the declared length of a stream without reading its contents.
    ///
    /// # Errors
    ///
    /// Returns an error if the path does not name a stream.
    pub fn stream_len(&self, path: &[&str]) -> Result<u64, OleError> {
        let entry = self.find_entry(path)?;
        if entry.entry_type != STGTY_STREAM {
            return Err(OleError::InvalidFormat("Not a stream".to_string()));
        }
        Ok(entry.size)
    }

    /// Returns whether an entry is present at `path`.
    #[must_use]
    pub fn exists(&self, path: &[&str]) -> bool {
        self.find_entry(path).is_ok()
    }

    /// Materializes one stream through immutable positional reads.
    ///
    /// Regular streams never acquire a shared lock or cursor. Small streams
    /// lazily initialize the root mini stream once; an initialization error is
    /// not retained, so a later read can retry.
    ///
    /// # Errors
    ///
    /// Returns an error if the path does not name a stream, source I/O fails,
    /// or the source version changes before or after the payload read.
    pub fn open_stream(&self, path: &[&str]) -> Result<Vec<u8>, OleError> {
        let (is_minifat, start_sector, size) = {
            let entry = self.find_entry(path)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()));
            }
            (entry.is_minifat, entry.start_sector, entry.size)
        };

        self.check_source_version()?;
        let result = if is_minifat {
            self.read_minifat_stream(start_sector, size)
        } else {
            self.read_fat_stream(start_sector, size)
        };
        self.check_source_version()?;
        result
    }

    fn check_source_version(&self) -> Result<(), OleError> {
        let observed = self.source.version()?;
        if observed == self.expected_version {
            Ok(())
        } else {
            Err(OleError::SourceChanged {
                expected: self.expected_version,
                observed,
            })
        }
    }

    fn read_fat_stream(&self, start_sector: u32, size: u64) -> Result<Vec<u8>, OleError> {
        let size = usize::try_from(size)
            .map_err(|_error| OleError::CorruptedFile("FAT stream is too large".to_string()))?;
        let required = size.div_ceil(self.index.sector_size);
        let mut data = try_zeroed_vec(size, "FAT stream data")?;
        self.read_chain_into(
            &self.index.fat,
            start_sector,
            required,
            self.index.sector_size,
            "FAT",
            &mut data,
        )?;
        Ok(data)
    }

    fn read_minifat_stream(&self, start_sector: u32, size: u64) -> Result<Vec<u8>, OleError> {
        let ministream = {
            let mut cached = self.ministream.lock().map_err(|_error| {
                OleError::InvalidData("shared mini-stream cache is poisoned".to_string())
            })?;
            if cached.is_none() {
                // Do not publish failed initialization: a transient source I/O
                // error must leave a subsequent mini-stream read free to retry.
                self.check_source_version()?;
                let loaded = self.load_ministream()?;
                self.check_source_version()?;
                *cached = Some(loaded);
            }
            cached.as_ref().cloned().ok_or_else(|| {
                OleError::CorruptedFile("shared mini-stream cache is missing".to_string())
            })?
        };
        let size = usize::try_from(size)
            .map_err(|_error| OleError::CorruptedFile("MiniFAT stream is too large".to_string()))?;
        let required = size.div_ceil(self.index.mini_sector_size);
        let mut data = try_zeroed_vec(size, "MiniFAT stream data")?;
        let mut sector = start_sector;

        for index in 0..required {
            let sector_index = usize::try_from(sector).map_err(|_error| {
                OleError::CorruptedFile("Mini sector index does not fit usize".to_string())
            })?;
            let position = sector_index
                .checked_mul(self.index.mini_sector_size)
                .ok_or_else(|| {
                    OleError::CorruptedFile("Mini sector offset overflow".to_string())
                })?;
            let end = position
                .checked_add(self.index.mini_sector_size)
                .ok_or_else(|| OleError::CorruptedFile("Mini sector end overflow".to_string()))?;
            if end > ministream.len() {
                return Err(OleError::CorruptedFile(
                    "Mini sector out of bounds".to_string(),
                ));
            }
            let output_start = index
                .checked_mul(self.index.mini_sector_size)
                .ok_or_else(|| {
                    OleError::CorruptedFile("MiniFAT stream offset overflow".to_string())
                })?;
            let output_end = output_start
                .checked_add(self.index.mini_sector_size)
                .unwrap_or(usize::MAX)
                .min(data.len());
            data[output_start..output_end].copy_from_slice(
                &ministream[position..position + output_end.saturating_sub(output_start)],
            );
            sector = next_chain_sector(&self.index.minifat, sector, "MiniFAT")?;
            if index + 1 == required {
                if sector != ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT chain exceeds its declared length".to_string(),
                    ));
                }
            } else if sector == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "MiniFAT chain ends before its declared length".to_string(),
                ));
            }
        }
        Ok(data)
    }

    fn load_ministream(&self) -> Result<Arc<[u8]>, OleError> {
        let root = self
            .index
            .root
            .as_ref()
            .ok_or_else(|| OleError::CorruptedFile("No root entry".to_string()))?;
        self.read_fat_stream(root.start_sector, root.size)
            .map(Vec::into)
    }

    fn read_chain_into(
        &self,
        table: &[u32],
        start_sector: u32,
        required: usize,
        sector_size: usize,
        table_name: &str,
        output: &mut [u8],
    ) -> Result<(), OleError> {
        if required == 0 {
            return Ok(());
        }
        if start_sector >= MAXREGSECT || required > table.len() {
            return Err(OleError::CorruptedFile(format!(
                "invalid {table_name} stream chain"
            )));
        }

        let mut sector = start_sector;
        let mut completed = 0usize;
        while completed < required {
            let run_start = sector;
            let mut run_count = 0usize;
            let next_after_run;
            loop {
                run_count = run_count.checked_add(1).ok_or_else(|| {
                    OleError::CorruptedFile("CFB sector run count overflow".to_string())
                })?;
                let next = next_chain_sector(table, sector, table_name)?;
                let total = completed.checked_add(run_count).ok_or_else(|| {
                    OleError::CorruptedFile("CFB sector count overflow".to_string())
                })?;
                if total == required {
                    next_after_run = next;
                    break;
                }
                if next == ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(format!(
                        "{table_name} chain ends before its declared length"
                    )));
                }
                if next
                    != sector.checked_add(1).ok_or_else(|| {
                        OleError::CorruptedFile("CFB sector index overflow".to_string())
                    })?
                {
                    next_after_run = next;
                    break;
                }
                sector = next;
            }
            if completed + run_count == required && next_after_run != ENDOFCHAIN {
                return Err(OleError::CorruptedFile(format!(
                    "{table_name} chain exceeds its declared length"
                )));
            }

            let output_start = completed
                .checked_mul(sector_size)
                .ok_or_else(|| OleError::CorruptedFile("CFB stream offset overflow".to_string()))?;
            let run_bytes = run_count.checked_mul(sector_size).ok_or_else(|| {
                OleError::CorruptedFile("CFB sector run size overflow".to_string())
            })?;
            let output_end = output_start
                .checked_add(run_bytes)
                .unwrap_or(usize::MAX)
                .min(output.len());
            self.read_sector_run(run_start, &mut output[output_start..output_end])?;

            completed += run_count;
            sector = next_after_run;
        }
        Ok(())
    }

    fn read_sector_run(&self, start_sector: u32, output: &mut [u8]) -> Result<(), OleError> {
        if output.is_empty() {
            return Ok(());
        }
        let position = (u64::from(start_sector) + 1)
            .checked_mul(self.index.sector_size as u64)
            .ok_or_else(|| OleError::CorruptedFile("Sector offset overflow".to_string()))?;
        if position >= self.index.file_size {
            return Err(OleError::CorruptedFile(format!(
                "Sector {start_sector} is outside the file"
            )));
        }
        let present = usize::try_from(
            self.index
                .file_size
                .saturating_sub(position)
                .min(output.len() as u64),
        )
        .map_err(|_error| OleError::CorruptedFile("sector read is too large".to_string()))?;
        self.source
            .read_exact_at(position, &mut output[..present])?;
        Ok(())
    }

    fn find_entry(&self, path: &[&str]) -> Result<&DirectoryEntry, OleError> {
        if path.is_empty() {
            return self.index.root.as_ref().ok_or(OleError::StreamNotFound);
        }
        let root = self.index.root.as_ref().ok_or(OleError::StreamNotFound)?;
        let mut current_sid = root.sid_child;
        for (index, name) in path.iter().enumerate() {
            let entry = self.find_child_by_name(current_sid, name)?;
            if index + 1 == path.len() {
                return Ok(entry);
            }
            current_sid = entry.sid_child;
        }
        Err(OleError::StreamNotFound)
    }

    fn find_child_by_name(&self, sid: u32, name: &str) -> Result<&DirectoryEntry, OleError> {
        let target = directory_name_data(name).map_err(|_error| OleError::StreamNotFound)?;
        let mut current_sid = sid;
        for _ in 0..self.index.dir_entries.len() {
            if current_sid == ENDOFCHAIN {
                return Err(OleError::StreamNotFound);
            }
            let index = usize::try_from(current_sid).map_err(|_error| OleError::StreamNotFound)?;
            let entry = self
                .index
                .dir_entries
                .get(index)
                .and_then(Option::as_ref)
                .ok_or(OleError::StreamNotFound)?;
            let entry_name = self
                .index
                .dir_name_data
                .get(index)
                .and_then(Option::as_ref)
                .ok_or(OleError::StreamNotFound)?;
            current_sid = match target.compare(entry_name) {
                Ordering::Less => entry.sid_left,
                Ordering::Equal => return Ok(entry),
                Ordering::Greater => entry.sid_right,
            };
        }
        Err(OleError::StreamNotFound)
    }
}

fn next_chain_sector(table: &[u32], sector: u32, table_name: &str) -> Result<u32, OleError> {
    if sector >= MAXREGSECT {
        return Err(OleError::CorruptedFile(format!(
            "invalid sector marker 0x{sector:08X} in {table_name} chain"
        )));
    }
    let index = usize::try_from(sector).map_err(|_error| {
        OleError::CorruptedFile(format!("invalid sector index {sector} in {table_name}"))
    })?;
    let next = *table.get(index).ok_or_else(|| {
        OleError::CorruptedFile(format!("invalid sector index {sector} in {table_name}"))
    })?;
    if next != ENDOFCHAIN && next >= MAXREGSECT {
        return Err(OleError::CorruptedFile(format!(
            "invalid sector marker 0x{next:08X} in {table_name} chain"
        )));
    }
    Ok(next)
}

fn try_zeroed_vec(length: usize, resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| OleError::allocation(resource, source))?;
    output.resize(length, 0);
    Ok(output)
}

/// A private local cursor used only to feed the existing validated parser.
struct ReadAtCursor {
    source: Arc<dyn ReadAt>,
    length: u64,
    position: u64,
}

impl ReadAtCursor {
    fn new(source: Arc<dyn ReadAt>, length: u64) -> Self {
        Self {
            source,
            length,
            position: 0,
        }
    }
}

impl Read for ReadAtCursor {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() || self.position >= self.length {
            return Ok(0);
        }
        let available = self.length - self.position;
        let count = usize::try_from(available.min(output.len() as u64))
            .map_err(|_error| io::Error::other("ReadAt cursor range is too large"))?;
        let read = self.source.read_at(self.position, &mut output[..count])?;
        if read > count {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional source reported more bytes than requested",
            ));
        }
        self.position = self
            .position
            .checked_add(read as u64)
            .ok_or_else(|| io::Error::other("ReadAt cursor position overflow"))?;
        Ok(read)
    }
}

impl Seek for ReadAtCursor {
    fn seek(&mut self, from: SeekFrom) -> io::Result<u64> {
        let base = match from {
            SeekFrom::Start(position) => {
                self.position = position;
                return Ok(position);
            },
            SeekFrom::End(offset) => i128::from(self.length) + i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
        };
        self.position = u64::try_from(base).map_err(|_error| {
            io::Error::new(io::ErrorKind::InvalidInput, "invalid ReadAt cursor seek")
        })?;
        Ok(self.position)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::{OleWriter, SharedOleBulkError};
    use litchi_core::{
        Budget, CancellationSource, CancellationToken, ExecutionContext, ExecutionError,
        ExecutionLimits, Limits, Resource,
    };
    use std::{
        io::Cursor,
        num::{NonZeroU64, NonZeroUsize},
        sync::{
            Barrier, Mutex,
            atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering as AtomicOrdering},
        },
        thread,
    };

    #[derive(Debug)]
    struct TestSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        reads: AtomicUsize,
        active_reads: AtomicUsize,
        max_active_reads: AtomicUsize,
        change_on_read: AtomicBool,
        fail_next_read: AtomicBool,
        cancel_on_read: AtomicBool,
        cancellation: Mutex<Option<CancellationSource>>,
        barrier: Mutex<Option<Arc<Barrier>>>,
        barrier_reads: AtomicUsize,
    }

    impl TestSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                reads: AtomicUsize::new(0),
                active_reads: AtomicUsize::new(0),
                max_active_reads: AtomicUsize::new(0),
                change_on_read: AtomicBool::new(false),
                fail_next_read: AtomicBool::new(false),
                cancel_on_read: AtomicBool::new(false),
                cancellation: Mutex::new(None),
                barrier: Mutex::new(None),
                barrier_reads: AtomicUsize::new(0),
            }
        }

        fn reset_read_count(&self) {
            self.reads.store(0, AtomicOrdering::SeqCst);
        }

        fn synchronize_next_two_reads(&self) {
            *self.barrier.lock().unwrap() = Some(Arc::new(Barrier::new(2)));
            self.barrier_reads.store(0, AtomicOrdering::SeqCst);
            self.max_active_reads.store(0, AtomicOrdering::SeqCst);
        }

        fn cancel_on_next_read(&self, source: CancellationSource) {
            *self.cancellation.lock().unwrap() = Some(source);
            self.cancel_on_read.store(true, AtomicOrdering::SeqCst);
        }
    }

    impl ReadAt for TestSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, AtomicOrdering::SeqCst);
            let active = self.active_reads.fetch_add(1, AtomicOrdering::SeqCst) + 1;
            self.max_active_reads
                .fetch_max(active, AtomicOrdering::SeqCst);
            let ticket = self.barrier_reads.fetch_add(1, AtomicOrdering::SeqCst);
            let barrier = self.barrier.lock().unwrap().clone();
            if ticket < 2 {
                if let Some(barrier) = barrier {
                    barrier.wait();
                }
            }
            if self.change_on_read.swap(false, AtomicOrdering::SeqCst) {
                self.revision.store(1, AtomicOrdering::SeqCst);
            }
            if self.cancel_on_read.swap(false, AtomicOrdering::SeqCst) {
                if let Some(source) = self.cancellation.lock().unwrap().as_ref() {
                    source.cancel();
                }
            }
            if self.fail_next_read.swap(false, AtomicOrdering::SeqCst) {
                self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
                return Err(io::Error::other("injected positional read failure"));
            }
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            let count = self.bytes.len().saturating_sub(start).min(output.len());
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                7,
                self.revision.load(AtomicOrdering::SeqCst),
            ))
        }
    }

    fn sample_bytes() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Small"], b"mini stream").unwrap();
        writer
            .create_stream_owned(&["Large"], vec![0xA5; 8192])
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn shared(source: Arc<TestSource>) -> SharedOleFile {
        SharedOleFile::open(source).unwrap()
    }

    fn context(
        workers: usize,
        max_tasks: usize,
        max_bytes: u64,
        min_parallel_bytes: u64,
        cancellation: CancellationToken,
        memory: u64,
    ) -> ExecutionContext {
        context_with_work(
            workers,
            max_tasks,
            max_bytes,
            min_parallel_bytes,
            cancellation,
            memory,
            1 << 20,
        )
    }

    fn context_with_work(
        workers: usize,
        max_tasks: usize,
        max_bytes: u64,
        min_parallel_bytes: u64,
        cancellation: CancellationToken,
        memory: u64,
        work: u64,
    ) -> ExecutionContext {
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(workers).unwrap(),
            NonZeroUsize::new(max_tasks).unwrap(),
            NonZeroU64::new(max_bytes).unwrap(),
            min_parallel_bytes,
        )
        .unwrap();
        ExecutionContext::new(
            Budget::root(
                "shared-cfb-test",
                Limits::new(memory, 1 << 20, 1 << 20, 100, 100, work),
            ),
            cancellation,
            limits,
        )
    }

    fn bulk_bytes() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer
            .create_stream_owned(&["First"], vec![0x11; 8192])
            .unwrap();
        writer
            .create_stream_owned(&["Second"], vec![0x22; 8192])
            .unwrap();
        writer.create_stream(&["Small"], b"mini stream").unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn preserves_existing_parser_corruption_checks() {
        let mut bytes = sample_bytes();
        bytes[0] ^= 0xFF;
        let source: Arc<dyn ReadAt> = Arc::new(TestSource::new(bytes));

        assert!(matches!(
            SharedOleFile::open(source),
            Err(OleError::NotOleFile)
        ));
    }

    #[test]
    fn rejects_input_before_parsing_when_over_the_finite_limit() {
        let source: Arc<dyn ReadAt> = Arc::new(TestSource::new(sample_bytes()));
        let limits = SharedOleFileLimits::new(512).unwrap();

        assert!(matches!(
            SharedOleFile::open_with_limits(source, limits),
            Err(OleError::InvalidData(message)) if message.contains("exceeds configured limit")
        ));
    }

    #[test]
    fn rejects_a_source_that_changes_during_structural_open() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(
            SharedOleFile::open(source),
            Err(OleError::SourceChanged { .. })
        ));
    }

    #[test]
    fn rejects_a_source_that_changes_during_payload_read() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(
            file.open_stream(&["Large"]),
            Err(OleError::SourceChanged { .. })
        ));
    }

    #[test]
    fn regular_stream_reads_are_concurrent_without_a_shared_cursor() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = Arc::new(shared(source.clone()));
        source.synchronize_next_two_reads();

        thread::scope(|scope| {
            let first = file.clone();
            let second = file.clone();
            let first_read = scope.spawn(move || first.open_stream(&["Large"]));
            let second_read = scope.spawn(move || second.open_stream(&["Large"]));
            assert_eq!(first_read.join().unwrap().unwrap().len(), 8192);
            assert_eq!(second_read.join().unwrap().unwrap().len(), 8192);
        });

        assert!(source.max_active_reads.load(AtomicOrdering::SeqCst) >= 2);
    }

    #[test]
    fn ministream_is_lazy_and_initialization_is_cached() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.reset_read_count();

        assert_eq!(file.open_stream(&["Large"]).unwrap().len(), 8192);
        let after_regular = source.reads.load(AtomicOrdering::SeqCst);
        assert!(after_regular > 0);

        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
        let after_first_small = source.reads.load(AtomicOrdering::SeqCst);
        assert!(after_first_small > after_regular);

        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
        assert_eq!(
            source.reads.load(AtomicOrdering::SeqCst),
            after_first_small,
            "the root mini stream must not be read twice"
        );
    }

    #[test]
    fn failed_ministream_initialization_can_retry() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(file.open_stream(&["Small"]), Err(OleError::Io(_))));
        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
    }

    #[test]
    fn bulk_reads_preserve_input_order_and_match_serial_reads() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let paths: &[&[&str]] = &[&["Second"], &["Small"], &["First"]];
        let (_cancel, token) = CancellationSource::pair();
        let context = context(2, 2, 16_384, 1, token, 32_768);

        let session = file.bulk_read(context);
        let bulk = session.read_streams(paths).unwrap();
        let repeated = session.read_streams(paths).unwrap();
        assert_eq!(repeated, bulk);
        assert_eq!(session.pool_build_count(), 1);
        let serial: Vec<_> = paths
            .iter()
            .map(|path| file.open_stream(path).unwrap())
            .collect();
        assert_eq!(bulk, serial);
        assert_eq!(bulk[0][0], 0x22);
        assert_eq!(bulk[2][0], 0x11);
    }

    #[test]
    fn bulk_reads_observe_preexisting_cancellation_atomically() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let (cancel, token) = CancellationSource::pair();
        cancel.cancel();

        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 8192))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::Cancelled))
        ));
    }

    #[test]
    fn bulk_reads_observe_mid_read_cancellation_atomically() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        let (cancel, token) = CancellationSource::pair();
        source.cancel_on_next_read(cancel);

        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 8192))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::Cancelled))
        ));
    }

    #[test]
    fn bulk_reads_enforce_byte_and_context_budget_caps() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let (_cancel, token) = CancellationSource::pair();
        assert!(matches!(
            file.bulk_read(context(1, 1, 4096, 0, token, 4096))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::StreamExceedsInFlightBytes { .. })
        ));

        let (_cancel, token) = CancellationSource::pair();
        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 4096))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(
                ExecutionError::ResourceLimit(_)
            ))
        ));
    }

    #[test]
    fn bulk_reads_charge_work_before_starting_payload_reads() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        source.reset_read_count();
        let (_cancel, token) = CancellationSource::pair();

        assert!(matches!(
            file.bulk_read(context_with_work(1, 1, 8192, 0, token, 8192, 4096))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Work
        ));
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), 0);
    }

    #[test]
    fn bulk_reads_honor_task_cap_and_use_a_local_pool_when_eligible() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        let (_cancel, token) = CancellationSource::pair();
        file.bulk_read(context(1, 1, 8192, 0, token, 16_384))
            .read_streams(&[&["First"], &["Second"]])
            .unwrap();
        assert_eq!(source.max_active_reads.load(AtomicOrdering::SeqCst), 1);

        source.synchronize_next_two_reads();
        let (_cancel, token) = CancellationSource::pair();
        file.bulk_read(context(2, 2, 16_384, 1, token, 16_384))
            .read_streams(&[&["First"], &["Second"]])
            .unwrap();
        assert!(source.max_active_reads.load(AtomicOrdering::SeqCst) >= 2);
    }

    #[test]
    fn bulk_session_skips_pool_when_only_bounded_batches_are_ineligible() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let (_cancel, token) = CancellationSource::pair();
        // Aggregate bytes exceed the threshold, but the 12 KiB batch cap
        // separates the two 8 KiB streams and neither one-stream batch can
        // qualify for parallel scheduling.
        let session = file.bulk_read(context(2, 2, 12_288, 12_288, token, 16_384));

        session.read_streams(&[&["First"], &["Second"]]).unwrap();
        assert_eq!(session.pool_build_count(), 0);
    }

    #[test]
    fn bulk_reads_reject_source_change_without_returning_partial_results() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        source.change_on_read.store(true, AtomicOrdering::SeqCst);
        let (_cancel, token) = CancellationSource::pair();

        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 8192))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Ole(OleError::SourceChanged { .. }))
        ));
    }
}
