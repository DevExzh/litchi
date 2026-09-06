//! A deterministic, caller-owned `ReadAt` adapter for range-source evidence.
//!
//! The adapter is deliberately kept in the performance tool.  It adds a
//! bounded short-read model and a fixed delay around an explicit source while
//! leaving the core and format crates unaware of simulated remote behavior.
//! The counters describe logical adapter calls; they are not physical I/O
//! observations.

use std::{
    fmt, io,
    sync::{
        Arc,
        atomic::{AtomicBool, AtomicU64, Ordering},
    },
    thread,
    time::Duration,
};

use litchi_core::{ReadAt, SourceVersion};
use serde::{Deserialize, Serialize};

/// Number of fixed request-size histogram buckets in a range-source snapshot.
pub const PPTX_RANGE_REQUEST_SIZE_BUCKETS: usize = 18;

/// Returns the fixed histogram bucket for one nonempty caller request.
///
/// Bucket `0` contains one byte, bucket `1` contains two bytes, buckets `2`
/// through `16` contain the inclusive ranges `3..=4` through
/// `32769..=65536`, and bucket `17` contains every larger request.  A zero
/// input is mapped to bucket `0`; empty `ReadAt` calls are never counted.
#[must_use]
pub fn request_size_bucket(requested_bytes: u64) -> usize {
    if requested_bytes <= 1 {
        0
    } else {
        let bucket = (u64::BITS - (requested_bytes - 1).leading_zeros()) as usize;
        bucket.min(PPTX_RANGE_REQUEST_SIZE_BUCKETS - 1)
    }
}

/// Configuration for [`PptxRangeSource`].
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct PptxRangeSourceConfig {
    /// Maximum number of bytes that one nonempty logical read may return.
    /// `None` preserves the wrapped source's normal read length.  `Some(0)`
    /// produces a zero-byte short read without invoking the wrapped source.
    pub max_returned_bytes: Option<usize>,
    /// Fixed delay applied once per nonempty delegated range call.  Empty
    /// caller buffers and zero-byte capped calls do not sleep.
    pub fixed_delay: Option<Duration>,
}

impl PptxRangeSourceConfig {
    /// Creates a range-source configuration.
    #[must_use]
    pub const fn new(max_returned_bytes: Option<usize>, fixed_delay: Option<Duration>) -> Self {
        Self {
            max_returned_bytes,
            fixed_delay,
        }
    }
}

/// Cumulative logical range-read counters.
///
/// The snapshot is immutable after construction and can be serialized into a
/// result record.  `min_request_bytes` and `max_request_bytes` are cumulative
/// bounds over the calls represented by this snapshot.  The checked interval
/// method carries those bounds from its after-snapshot, while checking every
/// monotonic event counter arithmetically.
#[derive(Clone, Copy, Debug, Default, Deserialize, Serialize, PartialEq, Eq)]
pub struct PptxRangeSourceSnapshot {
    /// Number of nonempty logical `read_at` calls accepted by the adapter.
    pub logical_calls: u64,
    /// Sum of caller-requested output lengths.
    pub requested_bytes: u64,
    /// Sum of bytes returned to callers.
    pub returned_bytes: u64,
    /// Smallest nonempty caller-requested output length observed so far.
    pub min_request_bytes: Option<u64>,
    /// Largest nonempty caller-requested output length observed so far.
    pub max_request_bytes: Option<u64>,
    /// Number of nonempty calls that returned fewer bytes than requested.
    pub short_reads: u64,
    /// Number of nonempty delegated calls for which the configured delay was
    /// recorded.  A zero-duration configured delay is still recorded.
    pub delayed_calls: u64,
    /// Number of nonempty caller requests in each fixed request-size bucket.
    pub request_size_counts: [u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS],
}

impl PptxRangeSourceSnapshot {
    /// Computes a checked interval from `before` to this after-snapshot.
    ///
    /// A counter regression is returned as an I/O error instead of being
    /// interpreted as unsigned wraparound.  Request bounds are point-in-time
    /// observations and therefore are carried from the after-snapshot rather
    /// than subtracted.
    pub fn checked_delta(self, before: Self) -> io::Result<PptxRangeSourceDelta> {
        let logical_calls = checked_difference(
            self.logical_calls,
            before.logical_calls,
            "logical range calls",
        )?;
        let request_size_counts = checked_request_size_delta(self, before)?;
        let histogram_calls = request_size_counts.iter().try_fold(0_u64, |total, count| {
            total
                .checked_add(*count)
                .ok_or_else(|| io::Error::other("request-size histogram interval overflows"))
        })?;
        if histogram_calls != logical_calls {
            return Err(io::Error::other(
                "request-size histogram interval differs from logical range calls",
            ));
        }
        Ok(PptxRangeSourceDelta {
            logical_calls,
            requested_bytes: checked_difference(
                self.requested_bytes,
                before.requested_bytes,
                "requested range bytes",
            )?,
            returned_bytes: checked_difference(
                self.returned_bytes,
                before.returned_bytes,
                "returned range bytes",
            )?,
            min_request_bytes: self.min_request_bytes,
            max_request_bytes: self.max_request_bytes,
            short_reads: checked_difference(
                self.short_reads,
                before.short_reads,
                "short range reads",
            )?,
            delayed_calls: checked_difference(
                self.delayed_calls,
                before.delayed_calls,
                "delayed range calls",
            )?,
            request_size_counts,
        })
    }

    /// Returns the number of calls represented by this snapshot.
    #[must_use]
    pub const fn read_calls(self) -> u64 {
        self.logical_calls
    }

    /// Returns the smallest request bound recorded by this snapshot.
    #[must_use]
    pub const fn minimum_requested_bytes(self) -> Option<u64> {
        self.min_request_bytes
    }

    /// Returns the largest request bound recorded by this snapshot.
    #[must_use]
    pub const fn maximum_requested_bytes(self) -> Option<u64> {
        self.max_request_bytes
    }
}

/// Checked interval counters for [`PptxRangeSource`].
#[derive(Clone, Copy, Debug, Default, Deserialize, Serialize, PartialEq, Eq)]
pub struct PptxRangeSourceDelta {
    /// Nonempty logical calls in the interval.
    pub logical_calls: u64,
    /// Caller-requested bytes in the interval.
    pub requested_bytes: u64,
    /// Bytes returned in the interval.
    pub returned_bytes: u64,
    /// Cumulative minimum request bound from the after-snapshot.
    pub min_request_bytes: Option<u64>,
    /// Cumulative maximum request bound from the after-snapshot.
    pub max_request_bytes: Option<u64>,
    /// Short reads in the interval.
    pub short_reads: u64,
    /// Recorded delayed calls in the interval.
    pub delayed_calls: u64,
    /// Checked request-size histogram counters for the interval.  Snapshot
    /// fields are cumulative; each bucket here is the checked after-minus-
    /// before count.
    pub request_size_counts: [u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS],
}

impl PptxRangeSourceDelta {
    /// Returns the number of calls represented by this interval.
    #[must_use]
    pub const fn read_calls(self) -> u64 {
        self.logical_calls
    }

    /// Returns the smallest cumulative request bound from the after-snapshot.
    #[must_use]
    pub const fn minimum_requested_bytes(self) -> Option<u64> {
        self.min_request_bytes
    }

    /// Returns the largest cumulative request bound from the after-snapshot.
    #[must_use]
    pub const fn maximum_requested_bytes(self) -> Option<u64> {
        self.max_request_bytes
    }
}

/// Caller-supplied positional source wrapped with deterministic range
/// behavior and logical counters.
pub struct PptxRangeSource {
    inner: Arc<dyn ReadAt>,
    config: PptxRangeSourceConfig,
    logical_calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    min_request_bytes: AtomicU64,
    max_request_bytes: AtomicU64,
    short_reads: AtomicU64,
    delayed_calls: AtomicU64,
    request_size_counts: [AtomicU64; PPTX_RANGE_REQUEST_SIZE_BUCKETS],
    metrics_failed: AtomicBool,
}

impl fmt::Debug for PptxRangeSource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PptxRangeSource")
            .field("config", &self.config)
            .field(
                "metrics_failed",
                &self.metrics_failed.load(Ordering::Acquire),
            )
            .finish_non_exhaustive()
    }
}

impl PptxRangeSource {
    /// Wraps an explicit caller-owned positional source.
    #[must_use]
    pub fn new(inner: Arc<dyn ReadAt>, config: PptxRangeSourceConfig) -> Self {
        Self {
            inner,
            config,
            logical_calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            min_request_bytes: AtomicU64::new(u64::MAX),
            max_request_bytes: AtomicU64::new(0),
            short_reads: AtomicU64::new(0),
            delayed_calls: AtomicU64::new(0),
            request_size_counts: std::array::from_fn(|_| AtomicU64::new(0)),
            metrics_failed: AtomicBool::new(false),
        }
    }

    /// Wraps a source with the supplied short-read cap and fixed delay.
    #[must_use]
    pub fn with_limits(
        inner: Arc<dyn ReadAt>,
        max_returned_bytes: Option<usize>,
        fixed_delay: Option<Duration>,
    ) -> Self {
        Self::new(
            inner,
            PptxRangeSourceConfig::new(max_returned_bytes, fixed_delay),
        )
    }

    /// Returns the configured range behavior.
    #[must_use]
    pub const fn config(&self) -> PptxRangeSourceConfig {
        self.config
    }

    /// Returns a serializable immutable cumulative counter snapshot.
    pub fn snapshot(&self) -> io::Result<PptxRangeSourceSnapshot> {
        self.ensure_metrics_available()?;
        let snapshot = self.snapshot_unchecked();
        self.ensure_metrics_available()?;
        let observed_calls =
            snapshot
                .request_size_counts
                .iter()
                .try_fold(0_u64, |total, count| {
                    total
                        .checked_add(*count)
                        .ok_or_else(|| io::Error::other("request-size histogram total overflow"))
                })?;
        if observed_calls != snapshot.logical_calls {
            return Err(io::Error::other(
                "request-size histogram total differs from logical range calls",
            ));
        }
        Ok(snapshot)
    }

    /// Takes an after-snapshot and computes its checked interval from
    /// `before`.
    pub fn checked_delta(
        &self,
        before: PptxRangeSourceSnapshot,
    ) -> io::Result<PptxRangeSourceDelta> {
        self.snapshot()?.checked_delta(before)
    }

    fn snapshot_unchecked(&self) -> PptxRangeSourceSnapshot {
        let logical_calls = self.logical_calls.load(Ordering::Acquire);
        let minimum = self.min_request_bytes.load(Ordering::Acquire);
        PptxRangeSourceSnapshot {
            logical_calls,
            requested_bytes: self.requested_bytes.load(Ordering::Acquire),
            returned_bytes: self.returned_bytes.load(Ordering::Acquire),
            min_request_bytes: (minimum != u64::MAX).then_some(minimum),
            max_request_bytes: (logical_calls != 0)
                .then(|| self.max_request_bytes.load(Ordering::Acquire)),
            short_reads: self.short_reads.load(Ordering::Acquire),
            delayed_calls: self.delayed_calls.load(Ordering::Acquire),
            request_size_counts: std::array::from_fn(|index| {
                self.request_size_counts[index].load(Ordering::Acquire)
            }),
        }
    }

    fn ensure_metrics_available(&self) -> io::Result<()> {
        if self.metrics_failed.load(Ordering::Acquire) {
            Err(metrics_unavailable())
        } else {
            Ok(())
        }
    }

    fn fail_metrics(&self, error: io::Error) -> io::Error {
        self.metrics_failed.store(true, Ordering::Release);
        error
    }

    fn checked_add(&self, counter: &AtomicU64, amount: u64, label: &'static str) -> io::Result<()> {
        self.ensure_metrics_available()?;
        counter
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |current| {
                current.checked_add(amount)
            })
            .map(|_| ())
            .map_err(|_| self.fail_metrics(io::Error::other(format!("{label} overflow"))))
    }

    fn update_min(&self, candidate: u64) {
        let mut current = self.min_request_bytes.load(Ordering::Acquire);
        while candidate < current {
            match self.min_request_bytes.compare_exchange_weak(
                current,
                candidate,
                Ordering::AcqRel,
                Ordering::Acquire,
            ) {
                Ok(_) => return,
                Err(observed) => current = observed,
            }
        }
    }

    fn update_max(&self, candidate: u64) {
        let mut current = self.max_request_bytes.load(Ordering::Acquire);
        while candidate > current {
            match self.max_request_bytes.compare_exchange_weak(
                current,
                candidate,
                Ordering::AcqRel,
                Ordering::Acquire,
            ) {
                Ok(_) => return,
                Err(observed) => current = observed,
            }
        }
    }
}

impl ReadAt for PptxRangeSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        self.ensure_metrics_available()?;
        let requested = u64::try_from(output.len())
            .map_err(|error| self.fail_metrics(io::Error::other(error.to_string())))?;
        self.checked_add(&self.logical_calls, 1, "logical range calls")?;
        self.checked_add(&self.requested_bytes, requested, "requested range bytes")?;
        self.checked_add(
            &self.request_size_counts[request_size_bucket(requested)],
            1,
            "request-size histogram count",
        )?;
        self.update_min(requested);
        self.update_max(requested);

        let delegated_len = self
            .config
            .max_returned_bytes
            .map_or(output.len(), |maximum| output.len().min(maximum));
        if delegated_len != 0 {
            if self.config.fixed_delay.is_some() {
                self.checked_add(&self.delayed_calls, 1, "delayed range calls")?;
            }
            if let Some(delay) = self.config.fixed_delay {
                thread::sleep(delay);
            }
        }

        let returned = if delegated_len == 0 {
            0
        } else {
            self.inner.read_at(offset, &mut output[..delegated_len])?
        };
        if returned > delegated_len {
            return Err(self.fail_metrics(io::Error::new(
                io::ErrorKind::InvalidData,
                "wrapped source returned more bytes than its output range",
            )));
        }
        let returned_u64 = u64::try_from(returned)
            .map_err(|error| self.fail_metrics(io::Error::other(error.to_string())))?;
        self.checked_add(&self.returned_bytes, returned_u64, "returned range bytes")?;
        if returned < output.len() {
            self.checked_add(&self.short_reads, 1, "short range reads")?;
        }
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn checked_difference(after: u64, before: u64, label: &'static str) -> io::Result<u64> {
    after
        .checked_sub(before)
        .ok_or_else(|| io::Error::other(format!("{label} counter moved backwards")))
}

fn checked_request_size_delta(
    after: PptxRangeSourceSnapshot,
    before: PptxRangeSourceSnapshot,
) -> io::Result<[u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS]> {
    let mut delta = [0_u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS];
    for (index, value) in delta.iter_mut().enumerate() {
        *value = checked_difference(
            after.request_size_counts[index],
            before.request_size_counts[index],
            "request-size histogram count",
        )?;
    }
    Ok(delta)
}

fn metrics_unavailable() -> io::Error {
    io::Error::other("PPTX range source metrics are unavailable after counter overflow")
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::OwnedSource;

    #[derive(Debug)]
    struct VersionedSource {
        inner: OwnedSource,
        version: SourceVersion,
    }

    impl ReadAt for VersionedSource {
        fn len(&self) -> io::Result<u64> {
            self.inner.len()
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.inner.read_at(offset, output)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(self.version)
        }
    }

    fn source(bytes: &[u8]) -> Arc<dyn ReadAt> {
        Arc::new(OwnedSource::new(bytes.to_vec()))
    }

    #[test]
    fn short_reads_reconstruct_the_exact_source() {
        let adapter = PptxRangeSource::with_limits(source(b"abcdefgh"), Some(2), None);
        let mut reconstructed = Vec::new();
        let mut offset = 0_u64;
        while offset < 8 {
            let mut output = [0_u8; 4];
            let returned = adapter.read_at(offset, &mut output).expect("short read");
            assert!(returned > 0);
            reconstructed.extend_from_slice(&output[..returned]);
            offset = offset
                .checked_add(u64::try_from(returned).expect("usize fits u64"))
                .expect("test offset does not overflow");
        }
        assert_eq!(reconstructed, b"abcdefgh");
        assert_eq!(
            adapter.snapshot().expect("metrics snapshot"),
            PptxRangeSourceSnapshot {
                logical_calls: 4,
                requested_bytes: 16,
                returned_bytes: 8,
                min_request_bytes: Some(4),
                max_request_bytes: Some(4),
                short_reads: 4,
                delayed_calls: 0,
                request_size_counts: {
                    let mut counts = [0_u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS];
                    counts[2] = 4;
                    counts
                },
            }
        );
    }

    #[test]
    fn empty_buffer_and_eof_have_distinct_call_semantics() {
        let adapter = PptxRangeSource::with_limits(source(b"abc"), Some(2), Some(Duration::ZERO));
        let mut empty = [];
        assert_eq!(adapter.read_at(0, &mut empty).expect("empty read"), 0);
        let mut eof = [0_u8; 2];
        assert_eq!(adapter.read_at(99, &mut eof).expect("EOF read"), 0);
        let snapshot = adapter.snapshot().expect("metrics snapshot");
        assert_eq!(snapshot.logical_calls, 1);
        assert_eq!(snapshot.requested_bytes, 2);
        assert_eq!(snapshot.returned_bytes, 0);
        assert_eq!(snapshot.short_reads, 1);
        assert_eq!(snapshot.delayed_calls, 1);
        assert_eq!(snapshot.request_size_counts[1], 1);
        assert_eq!(
            snapshot.request_size_counts.iter().sum::<u64>(),
            snapshot.logical_calls
        );
    }

    #[test]
    fn forwards_length_and_version() {
        let expected = SourceVersion::new(0x1234, 9);
        let source: Arc<dyn ReadAt> = Arc::new(VersionedSource {
            inner: OwnedSource::new(b"source".to_vec()),
            version: expected,
        });
        let adapter = PptxRangeSource::new(source, PptxRangeSourceConfig::default());
        assert_eq!(adapter.len().expect("length"), 6);
        assert_eq!(adapter.version().expect("version"), expected);
    }

    #[test]
    fn checked_delta_rejects_counter_regression() {
        let before = PptxRangeSourceSnapshot {
            logical_calls: 4,
            requested_bytes: 8,
            ..PptxRangeSourceSnapshot::default()
        };
        let after = PptxRangeSourceSnapshot {
            logical_calls: 3,
            requested_bytes: 9,
            ..PptxRangeSourceSnapshot::default()
        };
        let error = after
            .checked_delta(before)
            .expect_err("counter regression must fail");
        assert_eq!(error.kind(), io::ErrorKind::Other);
    }

    #[test]
    fn mixed_request_sizes_fill_the_fixed_histogram() {
        let requests = [
            1_u64, 2, 3, 4, 5, 8, 9, 16, 17, 32, 33, 64, 65, 128, 129, 256, 257, 512, 513, 1024,
            1025, 2048, 2049, 4096, 4097, 8192, 8193, 16_384, 16_385, 32_768, 32_769, 65_536,
            65_537,
        ];
        let adapter = PptxRangeSource::new(
            source(&vec![0_u8; 65_537]),
            PptxRangeSourceConfig::default(),
        );
        for request in requests {
            let mut output = vec![0_u8; usize::try_from(request).expect("test size")];
            assert_eq!(
                adapter.read_at(0, &mut output).expect("range read"),
                output.len()
            );
        }
        let snapshot = adapter.snapshot().expect("metrics snapshot");
        assert_eq!(snapshot.logical_calls, requests.len() as u64);
        assert_eq!(snapshot.requested_bytes, requests.iter().sum::<u64>());
        assert_eq!(snapshot.returned_bytes, requests.iter().sum::<u64>());
        assert_eq!(snapshot.min_request_bytes, Some(1));
        assert_eq!(snapshot.max_request_bytes, Some(65_537));
        assert_eq!(snapshot.short_reads, 0);
        assert_eq!(
            snapshot.request_size_counts.iter().sum::<u64>(),
            snapshot.logical_calls
        );
        for (request, count) in snapshot.request_size_counts.iter().enumerate() {
            let expected = requests
                .iter()
                .filter(|&&size| request_size_bucket(size) == request)
                .count() as u64;
            assert_eq!(*count, expected, "bucket {request}");
        }
        let delta = snapshot
            .checked_delta(PptxRangeSourceSnapshot::default())
            .expect("histogram delta");
        assert_eq!(delta.request_size_counts, snapshot.request_size_counts);
        assert_eq!(
            delta.request_size_counts.iter().sum::<u64>(),
            delta.logical_calls
        );
    }

    #[test]
    fn request_size_histogram_overflow_fails_closed_without_wrapping() {
        let adapter = PptxRangeSource::new(source(b"abc"), PptxRangeSourceConfig::default());
        adapter.request_size_counts[0].store(u64::MAX, Ordering::Relaxed);
        let mut output = [0_u8; 1];
        let error = adapter
            .read_at(0, &mut output)
            .expect_err("counter overflow must fail");
        assert_eq!(error.kind(), io::ErrorKind::Other);
        assert_eq!(
            adapter.request_size_counts[0].load(Ordering::Relaxed),
            u64::MAX
        );
        assert!(adapter.snapshot().is_err());
    }

    #[test]
    fn logical_counter_overflow_fails_closed_without_wrapping() {
        let adapter = PptxRangeSource::new(source(b"abc"), PptxRangeSourceConfig::default());
        adapter.logical_calls.store(u64::MAX, Ordering::Relaxed);
        let mut output = [0_u8; 1];
        let error = adapter
            .read_at(0, &mut output)
            .expect_err("counter overflow must fail");
        assert_eq!(error.kind(), io::ErrorKind::Other);
        assert_eq!(adapter.logical_calls.load(Ordering::Relaxed), u64::MAX);
        assert!(adapter.snapshot().is_err());
    }
}
