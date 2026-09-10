//! A bounded, benchmark-private read-ahead adapter for positional sources.
//!
//! This module is intentionally owned by `litchi-perf-baseline`.  It is a
//! measurement candidate, rather than a production `ReadAt` implementation or
//! a proposal to change the resource-accounting contract in the core crates.
//! A single bounded window is retained for one adapter instance.  The window
//! starts at the caller's requested offset, so a positive short fill can be
//! returned as a valid prefix without being mistaken for end-of-file.

#![allow(clippy::module_name_repetitions)]

use std::{
    fmt, io,
    sync::{Arc, Mutex},
};

use litchi_core::{ReadAt, SourceVersion};
use serde::Serialize;

/// Largest window accepted by the benchmark candidate.
pub(crate) const MAX_READ_AHEAD_BYTES: usize = 64 * 1024;

/// Cumulative counters for one read-ahead adapter instance.
///
/// The counters are deliberately logical and bounded.  `requests` counts
/// every nonempty caller request.  A request is classified as exactly one
/// hit or miss; a miss at EOF can therefore have no fill.  A fill is one
/// physical delegated call, including a call that returns zero or an error;
/// `fill_returned_bytes` includes only successful returned bytes.  A short
/// fill is a successful delegated call whose returned count is smaller than
/// its requested count, including a zero-byte result.  `max_fill_bytes` is
/// the largest delegated request size, which is the physical bound relevant
/// to this candidate's window.
#[derive(Clone, Copy, Debug, Default, Serialize, PartialEq, Eq)]
pub(crate) struct ReadAheadSnapshot {
    /// Actual allocation capacity in bytes for the bounded window.
    pub(crate) window_capacity: usize,
    /// Number of nonempty caller requests.
    pub(crate) requests: u64,
    /// Number of requests served entirely from the retained window prefix.
    pub(crate) hits: u64,
    /// Number of requests that did not use a retained window.
    pub(crate) misses: u64,
    /// Number of delegated physical fills attempted.
    pub(crate) fills: u64,
    /// Sum of bytes requested from the wrapped source by fills.
    pub(crate) fill_requested_bytes: u64,
    /// Sum of bytes successfully returned by the wrapped source from fills.
    pub(crate) fill_returned_bytes: u64,
    /// Number of successful fills returning fewer bytes than requested.
    pub(crate) short_fills: u64,
    /// Largest delegated fill request in bytes.
    pub(crate) max_fill_bytes: u64,
    /// Number of nonempty `read_at` operations that returned an error.
    /// Lock-poison and counter-overflow errors make the entire snapshot
    /// unavailable and therefore are not represented by this field.
    pub(crate) failures: u64,
}

impl ReadAheadSnapshot {
    fn validate(self) -> io::Result<()> {
        let classified = self
            .hits
            .checked_add(self.misses)
            .ok_or_else(|| io::Error::other("read-ahead request classification overflows"))?;
        if classified != self.requests {
            return Err(io::Error::other(
                "read-ahead requests are not classified exactly once",
            ));
        }
        if self.fills > self.misses {
            return Err(io::Error::other("read-ahead fills exceed cache misses"));
        }
        if self.fill_returned_bytes > self.fill_requested_bytes {
            return Err(io::Error::other(
                "read-ahead returned fill bytes exceed requested fill bytes",
            ));
        }
        if self.short_fills > self.fills {
            return Err(io::Error::other("read-ahead short fills exceed fills"));
        }
        let capacity = u64::try_from(self.window_capacity)
            .map_err(|_| io::Error::other("read-ahead window does not fit u64"))?;
        if self.max_fill_bytes > capacity {
            return Err(io::Error::other(
                "read-ahead maximum fill exceeds configured window",
            ));
        }
        if self.fills == 0
            && (self.fill_requested_bytes != 0
                || self.fill_returned_bytes != 0
                || self.short_fills != 0
                || self.max_fill_bytes != 0)
        {
            return Err(io::Error::other(
                "read-ahead fill counters are nonzero without a fill",
            ));
        }
        Ok(())
    }
}

#[derive(Debug)]
struct State {
    /// The sole preallocated storage owned by this adapter.
    buffer: Vec<u8>,
    /// Offset at which the current valid prefix begins.
    cache_start: Option<u64>,
    /// Number of valid bytes in `buffer`.
    cache_len: usize,
    requests: u64,
    hits: u64,
    misses: u64,
    fills: u64,
    fill_requested_bytes: u64,
    fill_returned_bytes: u64,
    short_fills: u64,
    max_fill_bytes: u64,
    failures: u64,
    /// Counter overflow permanently makes diagnostic snapshots unavailable.
    metrics_failed: bool,
}

impl State {
    fn new(buffer: Vec<u8>) -> Self {
        Self {
            buffer,
            cache_start: None,
            cache_len: 0,
            requests: 0,
            hits: 0,
            misses: 0,
            fills: 0,
            fill_requested_bytes: 0,
            fill_returned_bytes: 0,
            short_fills: 0,
            max_fill_bytes: 0,
            failures: 0,
            metrics_failed: false,
        }
    }

    fn invalidate(&mut self) {
        self.cache_start = None;
        self.cache_len = 0;
    }

    fn snapshot(&self, window_capacity: usize) -> io::Result<ReadAheadSnapshot> {
        if self.metrics_failed {
            return Err(io::Error::other("read-ahead counters are unavailable"));
        }
        if window_capacity != self.buffer.capacity() {
            return Err(io::Error::other(
                "read-ahead reported capacity differs from its allocation",
            ));
        }
        if self.cache_len > self.buffer.len() {
            return Err(io::Error::other(
                "read-ahead cache length exceeds its buffer",
            ));
        }
        if self.cache_len != 0 && self.cache_start.is_none() {
            return Err(io::Error::other(
                "read-ahead cache has bytes without a start offset",
            ));
        }
        let snapshot = ReadAheadSnapshot {
            window_capacity,
            requests: self.requests,
            hits: self.hits,
            misses: self.misses,
            fills: self.fills,
            fill_requested_bytes: self.fill_requested_bytes,
            fill_returned_bytes: self.fill_returned_bytes,
            short_fills: self.short_fills,
            max_fill_bytes: self.max_fill_bytes,
            failures: self.failures,
        };
        snapshot.validate()?;
        Ok(snapshot)
    }
}

/// One-window read-ahead source used only by the performance benchmark.
pub(crate) struct ReadAheadReadAt {
    inner: Arc<dyn ReadAt>,
    expected_len: u64,
    expected_version: SourceVersion,
    /// Maximum bytes requested from the wrapped source by one fill.
    window_bytes: usize,
    /// Actual allocation capacity, retained separately from the configured
    /// fill bound because an allocator may return more than requested.
    window_capacity: usize,
    state: Mutex<State>,
}

impl fmt::Debug for ReadAheadReadAt {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReadAheadReadAt")
            .field("expected_len", &self.expected_len)
            .field("expected_version", &self.expected_version)
            .field("window_bytes", &self.window_bytes)
            .field("window_capacity", &self.window_capacity)
            .finish_non_exhaustive()
    }
}

impl ReadAheadReadAt {
    /// Constructs a fresh adapter and allocates its bounded window.
    ///
    /// The source length and version are fenced before and after the
    /// fallible preallocation.  A source that changes during construction is
    /// refused before the adapter becomes visible to a timed operation.
    pub(crate) fn new(inner: Arc<dyn ReadAt>, window_bytes: usize) -> io::Result<Self> {
        if window_bytes == 0 || window_bytes > MAX_READ_AHEAD_BYTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "read-ahead window must be in 1..=65536 bytes",
            ));
        }

        let before_len = inner.len()?;
        let before_version = inner.version()?;
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(window_bytes)
            .map_err(|error| io::Error::other(format!("read-ahead allocation failed: {error}")))?;
        // `try_reserve_exact` has established capacity before this resize, so
        // this initialization cannot trigger a second allocation.
        buffer.resize(window_bytes, 0);
        let window_capacity = buffer.capacity();
        if window_capacity > MAX_READ_AHEAD_BYTES {
            return Err(io::Error::other(
                "read-ahead allocation capacity exceeds its hard bound",
            ));
        }

        let after_len = inner.len()?;
        let after_version = inner.version()?;
        if before_len != after_len || before_version != after_version {
            return Err(source_changed(
                before_len,
                after_len,
                before_version,
                after_version,
            ));
        }

        Ok(Self {
            inner,
            expected_len: after_len,
            expected_version: after_version,
            window_bytes,
            window_capacity,
            state: Mutex::new(State::new(buffer)),
        })
    }

    /// Returns a checked cumulative counter snapshot.
    pub(crate) fn snapshot(&self) -> io::Result<ReadAheadSnapshot> {
        let state = self
            .state
            .lock()
            .map_err(|_| io::Error::other("read-ahead state lock is poisoned"))?;
        state.snapshot(self.window_capacity)
    }

    fn check_source(&self) -> io::Result<()> {
        let actual_len = self.inner.len()?;
        let actual_version = self.inner.version()?;
        if actual_len != self.expected_len || actual_version != self.expected_version {
            return Err(source_changed(
                self.expected_len,
                actual_len,
                self.expected_version,
                actual_version,
            ));
        }
        Ok(())
    }

    fn fail(state: &mut State, error: io::Error) -> io::Error {
        state.invalidate();
        if let Some(next) = state.failures.checked_add(1) {
            state.failures = next;
            error
        } else {
            state.metrics_failed = true;
            io::Error::other("read-ahead failure counter overflow")
        }
    }

    fn record_fill_failure(state: &mut State, error: io::Error) -> io::Error {
        Self::fail(state, error)
    }

    fn source_range_len(&self, offset: u64, output_len: usize) -> io::Result<usize> {
        let requested = u64::try_from(output_len)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "read length exceeds u64"))?;
        let available = self.expected_len.saturating_sub(offset);
        usize::try_from(available.min(requested))
            .map_err(|_| io::Error::other("bounded read length does not fit the platform usize"))
    }

    fn cache_contains(&self, state: &State, offset: u64) -> bool {
        let Some(start) = state.cache_start else {
            return false;
        };
        if state.cache_len == 0 {
            return false;
        }
        let Ok(cache_len) = u64::try_from(state.cache_len) else {
            return false;
        };
        let Some(end) = start.checked_add(cache_len) else {
            return false;
        };
        offset >= start && offset < end
    }

    fn add_requests(state: &mut State, amount: u64) -> io::Result<()> {
        let Some(next) = state.requests.checked_add(amount) else {
            state.metrics_failed = true;
            return Err(io::Error::other("read-ahead requests counter overflow"));
        };
        state.requests = next;
        Ok(())
    }

    fn add_hits(state: &mut State, amount: u64) -> io::Result<()> {
        let Some(next) = state.hits.checked_add(amount) else {
            state.metrics_failed = true;
            return Err(io::Error::other("read-ahead hits counter overflow"));
        };
        state.hits = next;
        Ok(())
    }

    fn add_misses(state: &mut State, amount: u64) -> io::Result<()> {
        let Some(next) = state.misses.checked_add(amount) else {
            state.metrics_failed = true;
            return Err(io::Error::other("read-ahead misses counter overflow"));
        };
        state.misses = next;
        Ok(())
    }

    fn add_fills(state: &mut State, requested: usize) -> io::Result<()> {
        let requested = u64::try_from(requested)
            .map_err(|_| io::Error::other("read-ahead fill length does not fit u64"))?;
        let Some(next_fills) = state.fills.checked_add(1) else {
            state.metrics_failed = true;
            return Err(io::Error::other("read-ahead fills counter overflow"));
        };
        let Some(next_requested) = state.fill_requested_bytes.checked_add(requested) else {
            state.metrics_failed = true;
            return Err(io::Error::other(
                "read-ahead requested fill bytes counter overflow",
            ));
        };
        state.fills = next_fills;
        state.fill_requested_bytes = next_requested;
        state.max_fill_bytes = state.max_fill_bytes.max(requested);
        Ok(())
    }

    fn add_returned_fill(state: &mut State, returned: usize, requested: usize) -> io::Result<()> {
        let returned = u64::try_from(returned)
            .map_err(|_| io::Error::other("read-ahead returned fill length does not fit u64"))?;
        let Some(next_returned) = state.fill_returned_bytes.checked_add(returned) else {
            state.metrics_failed = true;
            return Err(io::Error::other(
                "read-ahead returned fill bytes counter overflow",
            ));
        };
        state.fill_returned_bytes = next_returned;
        if returned
            < u64::try_from(requested).map_err(|_| {
                io::Error::other("read-ahead requested fill length does not fit u64")
            })?
        {
            let Some(next_short) = state.short_fills.checked_add(1) else {
                state.metrics_failed = true;
                return Err(io::Error::other("read-ahead short fills counter overflow"));
            };
            state.short_fills = next_short;
        }
        Ok(())
    }

    fn read_with_state(
        &self,
        state: &mut State,
        offset: u64,
        output: &mut [u8],
    ) -> io::Result<usize> {
        let requested_output = match self.source_range_len(offset, output.len()) {
            Ok(requested_output) => requested_output,
            Err(error) => {
                // A nonempty request that cannot be represented is a miss and
                // a typed failure; no caller bytes have been touched.
                Self::add_misses(state, 1)?;
                return Err(Self::fail(state, error));
            },
        };

        if self.cache_contains(state, offset) {
            // The source fence occurs before any caller bytes are modified.
            // There is no fallible work after the copy itself.
            if let Err(error) = self.check_source() {
                // A stale cached range is a miss from the caller's point of
                // view, even though a cache entry existed before the fence.
                if let Err(counter_error) = Self::add_misses(state, 1) {
                    return Err(Self::fail(state, counter_error));
                }
                return Err(Self::fail(state, error));
            }
            Self::add_hits(state, 1)?;
            let start = match state.cache_start {
                Some(start) => start,
                None => {
                    return Err(Self::fail(
                        state,
                        io::Error::other("read-ahead cache start disappeared"),
                    ));
                },
            };
            let relative = match offset
                .checked_sub(start)
                .and_then(|relative| usize::try_from(relative).ok())
            {
                Some(relative) => relative,
                None => {
                    return Err(Self::fail(
                        state,
                        io::Error::other("read-ahead cache offset is invalid"),
                    ));
                },
            };
            let available = match state.cache_len.checked_sub(relative) {
                Some(available) => available,
                None => {
                    return Err(Self::fail(
                        state,
                        io::Error::other("read-ahead cache offset exceeds valid bytes"),
                    ));
                },
            };
            let amount = available.min(requested_output).min(output.len());
            let end = match relative.checked_add(amount) {
                Some(end) => end,
                None => {
                    return Err(Self::fail(
                        state,
                        io::Error::other("read-ahead cache slice overflows usize"),
                    ));
                },
            };
            let cached = match state.buffer.get(relative..end) {
                Some(cached) => cached,
                None => {
                    return Err(Self::fail(
                        state,
                        io::Error::other("read-ahead cache slice exceeds its buffer"),
                    ));
                },
            };
            output[..amount].copy_from_slice(cached);
            return Ok(amount);
        }

        // Every nonempty request is classified exactly once, including EOF
        // requests and requests rejected by a source-version fence.
        Self::add_misses(state, 1)?;
        if let Err(error) = self.check_source() {
            return Err(Self::fail(state, error));
        }

        if requested_output == 0 {
            // Requests at or beyond the captured EOF are misses but do not
            // invoke the wrapped source or manufacture an UnexpectedEof.
            return Ok(0);
        }

        // The caller's requested offset is the cache start.  The fill is
        // bounded by both the configured window and the captured source EOF.
        let available = self.expected_len.saturating_sub(offset);
        let window_u64 = match u64::try_from(self.window_bytes) {
            Ok(window_u64) => window_u64,
            Err(error) => {
                return Err(Self::fail(
                    state,
                    io::Error::other(format!("read-ahead window does not fit u64: {error}")),
                ));
            },
        };
        let fill_len = match usize::try_from(available.min(window_u64)) {
            Ok(fill_len) => fill_len,
            Err(error) => {
                return Err(Self::fail(
                    state,
                    io::Error::other(format!(
                        "read-ahead fill length does not fit usize: {error}"
                    )),
                ));
            },
        };
        state.invalidate();
        if let Err(error) = Self::add_fills(state, fill_len) {
            return Err(Self::fail(state, error));
        }

        // A single delegated call is intentional.  In particular, an
        // Interrupted result is surfaced instead of silently retrying.
        let returned = match self.inner.read_at(offset, &mut state.buffer[..fill_len]) {
            Ok(returned) if returned <= fill_len => returned,
            Ok(returned) => {
                return Err(Self::record_fill_failure(
                    state,
                    io::Error::new(
                        io::ErrorKind::InvalidData,
                        format!(
                            "read-ahead source returned {returned} bytes for a {fill_len}-byte fill"
                        ),
                    ),
                ));
            },
            Err(error) => return Err(Self::record_fill_failure(state, error)),
        };

        // A source may mutate while the delegated call is in flight.  Do not
        // publish or copy the temporary bytes until this second fence passes.
        if let Err(error) = self.check_source() {
            return Err(Self::record_fill_failure(state, error));
        }
        if offset
            .checked_add(u64::try_from(returned).map_err(|_| {
                io::Error::other("read-ahead returned fill length does not fit u64")
            })?)
            .is_none()
        {
            return Err(Self::record_fill_failure(
                state,
                io::Error::new(
                    io::ErrorKind::InvalidData,
                    "read-ahead returned range overflows u64",
                ),
            ));
        }
        if let Err(error) = Self::add_returned_fill(state, returned, fill_len) {
            return Err(Self::fail(state, error));
        }

        if returned == 0 {
            // Zero before the captured EOF is the wrapped ReadAt contract's
            // normal short result.  It is not an invented UnexpectedEof.
            state.invalidate();
            return Ok(0);
        }

        state.cache_start = Some(offset);
        state.cache_len = returned;
        // All fallible work is complete before this copy.  A positive short
        // fill therefore remains a valid readable prefix on the next call.
        let amount = returned.min(requested_output).min(output.len());
        output[..amount].copy_from_slice(&state.buffer[..amount]);
        Ok(amount)
    }
}

impl ReadAt for ReadAheadReadAt {
    fn len(&self) -> io::Result<u64> {
        self.check_source()?;
        Ok(self.expected_len)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let mut state = self
            .state
            .lock()
            .map_err(|_| io::Error::other("read-ahead state lock is poisoned"))?;
        Self::add_requests(&mut state, 1)?;
        self.read_with_state(&mut state, offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.check_source()?;
        Ok(self.expected_version)
    }
}

fn source_changed(
    expected_len: u64,
    actual_len: u64,
    expected_version: SourceVersion,
    actual_version: SourceVersion,
) -> io::Error {
    io::Error::new(
        io::ErrorKind::InvalidData,
        format!(
            "read-ahead source changed (length {expected_len}->{actual_len}, version {:?}->{:?})",
            expected_version, actual_version
        ),
    )
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused adapter tests intentionally use direct assertions"
    )]

    use super::*;
    use std::sync::{
        Barrier,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    };
    use std::thread;

    fn source(bytes: Vec<u8>) -> Arc<dyn ReadAt> {
        Arc::new(litchi_core::OwnedSource::new(bytes))
    }

    #[test]
    fn constructor_rejects_unbounded_windows_and_stabilizes_source() {
        let normal = source(vec![1, 2, 3]);
        assert!(ReadAheadReadAt::new(Arc::clone(&normal), 0).is_err());
        assert!(ReadAheadReadAt::new(Arc::clone(&normal), MAX_READ_AHEAD_BYTES + 1).is_err());
        let adapter = ReadAheadReadAt::new(normal, 4).expect("bounded adapter");
        assert_eq!(adapter.len().unwrap(), 3);
        assert_eq!(adapter.version().unwrap().revision(), 0);
        let capacity = adapter.snapshot().unwrap().window_capacity;
        assert!((4..=MAX_READ_AHEAD_BYTES).contains(&capacity));
    }

    #[test]
    fn differential_reads_preserve_prefix_suffix_and_eof() {
        let bytes: Vec<u8> = (0..257).map(|value| value as u8).collect();
        let adapter = ReadAheadReadAt::new(source(bytes.clone()), 16).expect("adapter");
        let mut state = 0x9e37_79b9_u64;
        for _ in 0..1_000 {
            state = state
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            let offset = state % 300;
            let length = ((state >> 32) as usize % 24) + 1;
            let sentinel = 0xa5_u8;
            let mut output = vec![sentinel; length + 3];
            let returned = adapter.read_at(offset, &mut output[..length]).unwrap();
            let expected = usize::try_from(
                (bytes.len() as u64)
                    .saturating_sub(offset)
                    .min(length as u64)
                    .min(16),
            )
            .unwrap();
            assert!(returned <= expected, "offset={offset}, length={length}");
            if expected != 0 {
                assert!(returned != 0, "nonempty source prefix became empty");
            }
            let start = usize::try_from(offset).unwrap_or(bytes.len());
            if start < bytes.len() {
                assert_eq!(&output[..returned], &bytes[start..start + returned]);
            }
            assert!(output[returned..].iter().all(|&byte| byte == sentinel));
        }
        let mut whole = vec![0_u8; 64];
        adapter
            .read_exact_at(0, &mut whole)
            .expect("bounded prefixes can be retried to complete a read");
        assert_eq!(&whole, &bytes[..64]);
        let snapshot = adapter.snapshot().unwrap();
        assert_eq!(snapshot.requests, snapshot.hits + snapshot.misses);
        assert!(snapshot.fills <= snapshot.misses);
    }

    #[test]
    fn crossing_window_returns_prefix_without_an_extra_fill() {
        let adapter = ReadAheadReadAt::new(source((0..64).collect()), 8).expect("adapter");
        let mut first = [0xa5_u8; 24];
        assert_eq!(adapter.read_at(5, &mut first).unwrap(), 8);
        assert_eq!(&first[..8], &[5, 6, 7, 8, 9, 10, 11, 12]);
        assert!(first[8..].iter().all(|&byte| byte == 0xa5));
        let after_first = adapter.snapshot().unwrap();
        let mut second = [0xa5_u8; 24];
        assert_eq!(adapter.read_at(5, &mut second).unwrap(), 8);
        let after_second = adapter.snapshot().unwrap();
        assert_eq!(after_second.fills, after_first.fills);
        assert_eq!(after_second.hits, after_first.hits + 1);
    }

    #[derive(Debug)]
    struct OneByteSource {
        bytes: Vec<u8>,
        version: SourceVersion,
    }

    impl ReadAt for OneByteSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            if start >= self.bytes.len() || output.is_empty() {
                return Ok(0);
            }
            output[0] = self.bytes[start];
            Ok(1)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(self.version)
        }
    }

    #[test]
    fn positive_short_fill_is_a_valid_prefix_and_zero_is_not_eof_error() {
        let inner: Arc<dyn ReadAt> = Arc::new(OneByteSource {
            bytes: b"abcdef".to_vec(),
            version: SourceVersion::new(7, 0),
        });
        let adapter = ReadAheadReadAt::new(inner, 4).expect("adapter");
        let mut output = [0xa5_u8; 4];
        assert_eq!(adapter.read_at(0, &mut output).unwrap(), 1);
        assert_eq!(&output[..1], b"a");
        assert!(output[1..].iter().all(|&byte| byte == 0xa5));
        let mut second = [0xa5_u8; 4];
        assert_eq!(adapter.read_at(1, &mut second).unwrap(), 1);
        assert_eq!(&second[..1], b"b");
        let snapshot = adapter.snapshot().unwrap();
        assert_eq!(snapshot.fills, 2);
        assert_eq!(snapshot.short_fills, 2);
        assert_eq!(snapshot.fill_requested_bytes, 8);
        assert_eq!(snapshot.fill_returned_bytes, 2);

        #[derive(Debug)]
        struct ZeroSource;
        impl ReadAt for ZeroSource {
            fn len(&self) -> io::Result<u64> {
                Ok(10)
            }
            fn read_at(&self, _: u64, _: &mut [u8]) -> io::Result<usize> {
                Ok(0)
            }
            fn version(&self) -> io::Result<SourceVersion> {
                Ok(SourceVersion::new(8, 0))
            }
        }
        let zero = ReadAheadReadAt::new(Arc::new(ZeroSource), 4).expect("zero adapter");
        let mut untouched = [0xa5_u8; 4];
        assert_eq!(zero.read_at(3, &mut untouched).unwrap(), 0);
        assert_eq!(untouched, [0xa5; 4]);
        assert_eq!(zero.snapshot().unwrap().failures, 0);
    }

    #[derive(Clone, Copy, Debug)]
    enum ReadBehavior {
        Interrupted,
        Overcount,
        MutateDuringRead,
    }

    #[derive(Debug)]
    struct ScriptedSource {
        bytes: Vec<u8>,
        behavior: ReadBehavior,
        revision: AtomicU64,
    }

    impl ReadAt for ScriptedSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            match self.behavior {
                ReadBehavior::Interrupted => Err(io::Error::new(
                    io::ErrorKind::Interrupted,
                    "injected interruption",
                )),
                ReadBehavior::Overcount => Ok(output.len() + 1),
                ReadBehavior::MutateDuringRead => {
                    let start = usize::try_from(offset).unwrap_or(self.bytes.len());
                    let count = self.bytes.len().saturating_sub(start).min(output.len());
                    if count != 0 {
                        output[..count].copy_from_slice(&self.bytes[start..start + count]);
                    }
                    if matches!(self.behavior, ReadBehavior::MutateDuringRead) {
                        self.revision.fetch_add(1, Ordering::SeqCst);
                    }
                    Ok(count)
                },
            }
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(9, self.revision.load(Ordering::SeqCst)))
        }
    }

    #[test]
    fn interruption_and_overcount_invalidate_without_touching_output() {
        for behavior in [ReadBehavior::Interrupted, ReadBehavior::Overcount] {
            let inner: Arc<dyn ReadAt> = Arc::new(ScriptedSource {
                bytes: b"source".to_vec(),
                behavior,
                revision: AtomicU64::new(0),
            });
            let adapter = ReadAheadReadAt::new(inner, 4).expect("adapter");
            let mut output = [0xa5_u8; 4];
            let error = adapter
                .read_at(0, &mut output)
                .expect_err("failure expected");
            assert_eq!(output, [0xa5; 4]);
            assert_eq!(
                error.kind(),
                if matches!(behavior, ReadBehavior::Interrupted) {
                    io::ErrorKind::Interrupted
                } else {
                    io::ErrorKind::InvalidData
                }
            );
            let snapshot = adapter.snapshot().unwrap();
            assert_eq!(snapshot.fills, 1);
            assert_eq!(snapshot.failures, 1);
        }
    }

    #[test]
    fn mutation_across_fill_is_refused_and_does_not_publish_bytes() {
        let inner: Arc<dyn ReadAt> = Arc::new(ScriptedSource {
            bytes: b"source".to_vec(),
            behavior: ReadBehavior::MutateDuringRead,
            revision: AtomicU64::new(0),
        });
        let adapter = ReadAheadReadAt::new(inner, 4).expect("adapter");
        let mut output = [0xa5_u8; 4];
        let error = adapter
            .read_at(0, &mut output)
            .expect_err("mutation expected");
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);
        assert_eq!(output, [0xa5; 4]);
        let snapshot = adapter.snapshot().unwrap();
        assert_eq!(snapshot.fills, 1);
        assert_eq!(snapshot.failures, 1);

        // A source mutation observed before a would-be hit must also refuse
        // before copying the retained bytes.
        let mutable: Arc<MutableSource> = Arc::new(MutableSource::new(b"source"));
        let adapter =
            ReadAheadReadAt::new(Arc::clone(&mutable) as Arc<dyn ReadAt>, 4).expect("adapter");
        let mut first = [0_u8; 4];
        assert_eq!(adapter.read_at(0, &mut first).unwrap(), 4);
        mutable.bump();
        let mut untouched = [0xa5_u8; 4];
        let error = adapter
            .read_at(0, &mut untouched)
            .expect_err("stale hit expected");
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);
        assert_eq!(untouched, [0xa5; 4]);
    }

    #[derive(Debug)]
    struct MutableSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
    }

    impl MutableSource {
        fn new(bytes: &[u8]) -> Self {
            Self {
                bytes: bytes.to_vec(),
                revision: AtomicU64::new(0),
            }
        }

        fn bump(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for MutableSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            let count = self.bytes.len().saturating_sub(start).min(output.len());
            if count != 0 {
                output[..count].copy_from_slice(&self.bytes[start..start + count]);
            }
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(10, self.revision.load(Ordering::SeqCst)))
        }
    }

    #[derive(Debug)]
    struct HugeSource;

    impl ReadAt for HugeSource {
        fn len(&self) -> io::Result<u64> {
            Ok(u64::MAX)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            if output.is_empty() || offset >= u64::MAX - 1 {
                return Ok(0);
            }
            output[0] = 0x5a;
            Ok(1)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(11, 0))
        }
    }

    #[test]
    fn near_u64_max_is_bounded_without_offset_overflow() {
        let adapter = ReadAheadReadAt::new(Arc::new(HugeSource), 4).expect("adapter");
        let mut output = [0xa5_u8; 4];
        assert_eq!(adapter.read_at(u64::MAX - 2, &mut output).unwrap(), 1);
        assert_eq!(output[0], 0x5a);
        assert!(output[1..].iter().all(|&byte| byte == 0xa5));
        assert_eq!(adapter.read_at(u64::MAX, &mut output).unwrap(), 0);
        assert!(adapter.snapshot().is_ok());
    }

    #[derive(Debug)]
    struct CountingSource {
        bytes: Vec<u8>,
        reads: AtomicUsize,
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            let count = self.bytes.len().saturating_sub(start).min(output.len());
            if count != 0 {
                output[..count].copy_from_slice(&self.bytes[start..start + count]);
            }
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(12, 0))
        }
    }

    #[test]
    fn concurrent_same_offset_reads_singleflight_and_bound_memory() {
        let inner = Arc::new(CountingSource {
            bytes: (0..128).collect(),
            reads: AtomicUsize::new(0),
        });
        let adapter = Arc::new(
            ReadAheadReadAt::new(Arc::clone(&inner) as Arc<dyn ReadAt>, 32).expect("adapter"),
        );
        let barrier = Arc::new(Barrier::new(9));
        let mut workers = Vec::new();
        for _ in 0..8 {
            let adapter = Arc::clone(&adapter);
            let barrier = Arc::clone(&barrier);
            workers.push(thread::spawn(move || {
                barrier.wait();
                let mut output = [0xa5_u8; 8];
                let returned = adapter.read_at(16, &mut output).expect("concurrent read");
                assert_eq!(returned, 8);
                assert_eq!(output, [16, 17, 18, 19, 20, 21, 22, 23]);
                returned
            }));
        }
        barrier.wait();
        for worker in workers {
            assert_eq!(worker.join().expect("worker"), 8);
        }
        assert_eq!(inner.reads.load(Ordering::SeqCst), 1);
        let snapshot = adapter.snapshot().unwrap();
        assert!((32..=MAX_READ_AHEAD_BYTES).contains(&snapshot.window_capacity));
        assert_eq!(snapshot.fills, 1);
        assert_eq!(snapshot.max_fill_bytes, 32);
        assert_eq!(snapshot.requests, 8);
        assert_eq!(snapshot.hits, 7);
        assert_eq!(snapshot.misses, 1);
    }

    #[test]
    fn poisoned_state_is_a_typed_io_failure() {
        let adapter = ReadAheadReadAt::new(source(b"abc".to_vec()), 2).expect("adapter");
        let poison = Arc::new(adapter);
        let thread_adapter = Arc::clone(&poison);
        let _ = thread::spawn(move || {
            let _guard = thread_adapter.state.lock().expect("state lock");
            panic!("poison test");
        })
        .join();
        let error = poison.snapshot().expect_err("poison should be reported");
        assert_eq!(error.kind(), io::ErrorKind::Other);
    }

    #[test]
    fn source_zero_before_eof_is_counted_as_a_short_fill_without_cache() {
        let adapter = ReadAheadReadAt::new(Arc::new(ZeroLengthSource), 2).expect("adapter");
        let mut output = [0xa5_u8; 2];
        assert_eq!(adapter.read_at(0, &mut output).unwrap(), 0);
        let before_retry = adapter.snapshot().unwrap();
        assert_eq!((before_retry.fills, before_retry.short_fills), (1, 1));
        assert_eq!(adapter.read_at(0, &mut output).unwrap(), 0);
        let after_retry = adapter.snapshot().unwrap();
        assert_eq!(after_retry.fills, 2);
        assert_eq!(after_retry.hits, 0);
    }

    #[derive(Debug)]
    struct ZeroLengthSource;

    impl ReadAt for ZeroLengthSource {
        fn len(&self) -> io::Result<u64> {
            Ok(5)
        }
        fn read_at(&self, _: u64, _: &mut [u8]) -> io::Result<usize> {
            Ok(0)
        }
        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(13, 0))
        }
    }

    #[test]
    fn empty_requests_are_noops_without_source_or_counter_work() {
        let changed = Arc::new(AtomicBool::new(false));
        let inner: Arc<dyn ReadAt> = Arc::new(EmptyProbe {
            touched: Arc::clone(&changed),
        });
        let adapter = ReadAheadReadAt::new(inner, 2).expect("adapter");
        changed.store(false, Ordering::SeqCst);
        let mut output: [u8; 0] = [];
        assert_eq!(adapter.read_at(0, &mut output).unwrap(), 0);
        assert!(!changed.load(Ordering::SeqCst));
        assert_eq!(adapter.snapshot().unwrap().requests, 0);
    }

    #[derive(Debug)]
    struct EmptyProbe {
        touched: Arc<AtomicBool>,
    }

    impl ReadAt for EmptyProbe {
        fn len(&self) -> io::Result<u64> {
            self.touched.store(true, Ordering::SeqCst);
            Ok(0)
        }
        fn read_at(&self, _: u64, _: &mut [u8]) -> io::Result<usize> {
            self.touched.store(true, Ordering::SeqCst);
            Ok(0)
        }
        fn version(&self) -> io::Result<SourceVersion> {
            self.touched.store(true, Ordering::SeqCst);
            Ok(SourceVersion::new(14, 0))
        }
    }
}
