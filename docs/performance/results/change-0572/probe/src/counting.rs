//! A counting `litchi_core::ReadAt` wrapper with an optional simulated transport.
//!
//! The wrapper sits exactly where the library's caller-supplied source sits, so
//! every `(offset, length)` it records is a request the library issued -- there
//! is no inference. `litchi-opc`'s `SourceReader` forwards `soapberry_zip`'s
//! `ReaderAt::read_at` straight to this source under the `exact` policy, and
//! through `ArchiveReadAhead` under a `forward_start` policy; in both cases the
//! physical call lands here.
//!
//! The transport model reproduces the `0493` configuration used elsewhere in
//! this program: a fixed service time per physical request, a nominal transfer
//! rate, a minimum-service combination policy (the fixed sleep and the wrapped
//! source's own work count toward the transfer target rather than adding to
//! it), and a maximum physical range that is served as a *short read* -- the
//! caller loops, exactly as a real byte-range transport would force it to.

use std::{
    io,
    sync::{Mutex, atomic::{AtomicU64, Ordering}},
    thread,
    time::{Duration, Instant},
};

use litchi_core::{ReadAt, SourceVersion};

/// One recorded logical request, in issue order.
#[derive(Clone, Copy, Debug)]
pub struct Request {
    /// Byte offset the library asked for.
    pub offset: u64,
    /// Output length the library offered.
    pub length: u64,
    /// Bytes actually returned (a short read under a capped transport).
    pub returned: u64,
}

/// Transport configuration. `None` fields disable that part of the model.
#[derive(Clone, Copy, Debug, Default)]
pub struct Transport {
    /// Fixed service time charged once per nonempty physical request.
    pub fixed_service: Option<Duration>,
    /// Nominal transfer rate in bytes per second.
    pub bytes_per_second: Option<u64>,
    /// Maximum bytes one physical request may return; larger asks are short.
    pub max_range_bytes: Option<usize>,
}

impl Transport {
    /// The zero-delay control: no service time, no rate, no range cap.
    pub const fn control() -> Self {
        Self { fixed_service: None, bytes_per_second: None, max_range_bytes: None }
    }

    /// The 0493 configuration: 1 ms fixed service, 100 MiB/s, 64 KiB maximum
    /// physical range, minimum-service combination.
    pub const fn delayed_0493() -> Self {
        Self {
            fixed_service: Some(Duration::from_millis(1)),
            bytes_per_second: Some(104_857_600),
            max_range_bytes: Some(65_536),
        }
    }

    /// The range cap with no delay, used only to separate a sequence change
    /// caused by the cap from one caused by timing.
    pub const fn capped_control() -> Self {
        Self { fixed_service: None, bytes_per_second: None, max_range_bytes: Some(65_536) }
    }
}

/// A `ReadAt` that records every request in order and optionally paces it.
pub struct CountingSource {
    inner: Vec<u8>,
    version: SourceVersion,
    transport: Transport,
    log: Mutex<Vec<Request>>,
    calls: AtomicU64,
}

impl CountingSource {
    /// Wraps owned bytes. The bytes are held in memory so the control arm
    /// measures the library and the model, never the page cache.
    pub fn new(bytes: Vec<u8>, transport: Transport) -> Self {
        Self {
            inner: bytes,
            version: SourceVersion::new(0x0572, 1),
            transport,
            log: Mutex::new(Vec::new()),
            calls: AtomicU64::new(0),
        }
    }

    /// Drains the recorded request list.
    pub fn take_log(&self) -> Vec<Request> {
        let mut guard = self.log.lock().expect("request log poisoned");
        std::mem::take(&mut *guard)
    }

    /// Number of nonempty requests recorded so far.
    pub fn calls(&self) -> u64 {
        self.calls.load(Ordering::SeqCst)
    }
}

/// Rounds a nominal transfer target up to whole nanoseconds.
fn transfer_delay_ns(bytes: u64, rate: u64) -> u64 {
    if rate == 0 {
        return 0;
    }
    let nanos = u128::from(bytes) * 1_000_000_000_u128;
    let target = nanos.div_ceil(u128::from(rate));
    u64::try_from(target).unwrap_or(u64::MAX)
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.inner.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        // An empty caller buffer is not a request and is never recorded; it
        // costs no service time either.
        if output.is_empty() {
            return Ok(0);
        }
        let requested = output.len() as u64;
        self.calls.fetch_add(1, Ordering::SeqCst);

        let delegated = self
            .transport
            .max_range_bytes
            .map_or(output.len(), |cap| output.len().min(cap));

        let started = (delegated != 0 && self.transport.bytes_per_second.is_some())
            .then(Instant::now);
        if delegated != 0 && let Some(service) = self.transport.fixed_service {
            thread::sleep(service);
        }

        let returned = if delegated == 0 {
            0
        } else {
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.inner.len() {
                0
            } else {
                let take = delegated.min(self.inner.len() - start);
                output[..take].copy_from_slice(&self.inner[start..start + take]);
                take
            }
        };

        if returned != 0 && let Some(rate) = self.transport.bytes_per_second {
            let nominal = Duration::from_nanos(transfer_delay_ns(returned as u64, rate));
            // Minimum-service: wait only until fixed + nominal has elapsed.
            let target = self.transport.fixed_service.unwrap_or_default() + nominal;
            let elapsed = started.map(|s| s.elapsed()).unwrap_or_default();
            let remaining = target.saturating_sub(elapsed);
            if !remaining.is_zero() {
                thread::sleep(remaining);
            }
        }

        self.log
            .lock()
            .expect("request log poisoned")
            .push(Request { offset, length: requested, returned: returned as u64 });
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}
