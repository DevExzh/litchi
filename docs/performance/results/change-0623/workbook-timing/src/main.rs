//! Time a source-backed OOXML open on change 0572's simulated transport.
//!
//! The transport is the one change 0493 and change 0572 fixed and change 0611
//! reused: 1 ms of fixed service per physical request, 100 MiB/s of bandwidth,
//! and a 64 KiB maximum physical range, so one logical read of more than
//! 64 KiB costs several physical requests. Requests are counted as well as
//! timed, so the median can be read against a deterministic request count.
use std::{
    fs, io,
    sync::{Arc, atomic::{AtomicU64, Ordering}},
    time::{Duration, Instant},
};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

const FIXED_LATENCY: Duration = Duration::from_micros(1000);
const BANDWIDTH_BYTES_PER_SEC: u64 = 100 * 1024 * 1024;
const MAX_PHYSICAL_BYTES: usize = 64 * 1024;

struct Delayed {
    bytes: Vec<u8>,
    requests: AtomicU64,
    physical: AtomicU64,
    read_bytes: AtomicU64,
}

impl Delayed {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            requests: AtomicU64::new(0),
            physical: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
        }
    }
}

impl ReadAt for Delayed {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        self.requests.fetch_add(1, Ordering::Relaxed);
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let taken = if start >= self.bytes.len() {
            0
        } else {
            output.len().min(self.bytes.len() - start)
        };
        let physical = output.len().div_ceil(MAX_PHYSICAL_BYTES).max(1);
        self.physical.fetch_add(physical as u64, Ordering::Relaxed);
        self.read_bytes.fetch_add(taken as u64, Ordering::Relaxed);
        let transfer = Duration::from_nanos(
            (taken as u64).saturating_mul(1_000_000_000) / BANDWIDTH_BYTES_PER_SEC,
        );
        std::thread::sleep(FIXED_LATENCY * physical as u32 + transfer);
        if taken > 0 {
            output[..taken].copy_from_slice(&self.bytes[start..start + taken]);
        }
        Ok(taken)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x0623, 1))
    }
}

fn main() {
    let mut args = std::env::args().skip(1);
    let path = args.next().expect("usage: <fixture> <warmup> <samples>");
    let warmup: usize = args.next().expect("warmup").parse().expect("warmup");
    let samples: usize = args.next().expect("samples").parse().expect("samples");
    let bytes = fs::read(&path).expect("fixture");
    let mut timings: Vec<u128> = Vec::with_capacity(samples);
    let mut requests = Vec::with_capacity(samples);
    let mut physical = Vec::with_capacity(samples);
    let mut read_bytes = Vec::with_capacity(samples);
    for iteration in 0..warmup + samples {
        let source = Arc::new(Delayed::new(bytes.clone()));
        let dynamic: Arc<dyn ReadAt> = Arc::clone(&source) as Arc<dyn ReadAt>;
        let started = Instant::now();
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
                dynamic,
                ReadLimits::default(),
                SourceCacheLimits::default(),
                SourceReadPolicy::exact(),
            )
            .expect("open");
        let elapsed = started.elapsed();
        std::hint::black_box(package.iter_parts().count());
        if iteration >= warmup {
            timings.push(elapsed.as_nanos());
            requests.push(source.requests.load(Ordering::Relaxed));
            physical.push(source.physical.load(Ordering::Relaxed));
            read_bytes.push(source.read_bytes.load(Ordering::Relaxed));
        }
    }
    timings.sort_unstable();
    let quantile = |q: f64| timings[((timings.len() - 1) as f64 * q).round() as usize];
    let mean: f64 = timings.iter().map(|value| *value as f64).sum::<f64>() / timings.len() as f64;
    requests.sort_unstable();
    physical.sort_unstable();
    read_bytes.sort_unstable();
    println!(
        "{{\"fixture\":\"{path}\",\"samples\":{},\"p50_us\":{:.1},\"mean_us\":{:.1},\"p95_us\":{:.1},\"p99_us\":{:.1},\"requests_min\":{},\"requests_max\":{},\"physical_min\":{},\"physical_max\":{},\"bytes_min\":{},\"bytes_max\":{}}}",
        timings.len(),
        quantile(0.50) as f64 / 1000.0,
        mean / 1000.0,
        quantile(0.95) as f64 / 1000.0,
        quantile(0.99) as f64 / 1000.0,
        requests[0],
        requests[requests.len() - 1],
        physical[0],
        physical[physical.len() - 1],
        read_bytes[0],
        read_bytes[read_bytes.len() - 1],
    );
}
