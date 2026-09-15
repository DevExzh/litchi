//! Coexistence probe for change 0615.
//!
//! The workspace has exactly three explicitly scheduled parallel sessions, and
//! each of them owns its own worker set:
//!
//! 1. `soapberry_zip::office::ParallelReadSession` (a private Rayon pool built
//!    in `ParallelReadSession::new`), reached through
//!    `litchi_opc::OpenSession`;
//! 2. `litchi_cfb::SharedOleBulkRead` (a private Rayon pool built lazily on the
//!    first eligible batch);
//! 3. `litchi_opc::SourceBackedPackage::read_parts_ordered` (`std::thread::scope`
//!    workers spawned per operation).
//!
//! Each is configured by an `ExecutionLimits` whose `workers` field is a
//! *per-session* ceiling. This probe drives all three from **one** hierarchical
//! `Budget` root with `workers = W` and counts what the process actually holds:
//! live OS threads at each stage, and the maximum number of `ReadAt::read_at`
//! calls in flight at once.
//!
//! Nothing here is a latency measurement. Every number is a count.

use std::fs;
use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_cfb::{OleWriter, SharedOleFile, SharedOleFileLimits};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, SourceVersion,
};
use litchi_opc::{OpenSession, PackURI, ReadLimits, SourceBackedPackage};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/// Members are one MiB each so that every batch clears `min_parallel_bytes`
/// and every session takes its parallel branch.
const MEMBER_BYTES: usize = 1024 * 1024;
const MEMBERS: usize = 4;

/// Process-wide in-flight positional reads, across every source in this
/// process. Nothing in `ExecutionLimits` bounds this number; the probe counts
/// what the three sessions produce together.
static WORKERS: AtomicUsize = AtomicUsize::new(4);
static GLOBAL_IN_FLIGHT: AtomicUsize = AtomicUsize::new(0);
static GLOBAL_MAX_IN_FLIGHT: AtomicUsize = AtomicUsize::new(0);

/// Positional source that counts concurrent `read_at` calls.
///
/// `max_in_flight` is the exact maximum number of simultaneous positional
/// reads the source observed; `max_threads_at_peak` is the live OS thread
/// count sampled at the moment a new maximum was recorded.
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    version: SourceVersion,
    reads: AtomicU64,
    in_flight: AtomicUsize,
    max_in_flight: AtomicUsize,
    max_threads_at_peak: AtomicUsize,
    /// Fixed per-read delay, in microseconds. A deterministic instrument that
    /// widens the window in which concurrent reads can be observed; it models
    /// no production storage. Zero in the default counting mode.
    delay_us: u64,
}

impl CountingSource {
    fn new(bytes: Vec<u8>, delay_us: u64) -> Self {
        Self {
            bytes: Arc::new(bytes),
            version: SourceVersion::new(1, 1),
            reads: AtomicU64::new(0),
            in_flight: AtomicUsize::new(0),
            max_in_flight: AtomicUsize::new(0),
            max_threads_at_peak: AtomicUsize::new(0),
            delay_us,
        }
    }

    fn reset(&self) {
        self.reads.store(0, Ordering::SeqCst);
        self.max_in_flight.store(0, Ordering::SeqCst);
        self.max_threads_at_peak.store(0, Ordering::SeqCst);
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source length exceeds u64"))
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let depth = self.in_flight.fetch_add(1, Ordering::SeqCst) + 1;
        let global = GLOBAL_IN_FLIGHT.fetch_add(1, Ordering::SeqCst) + 1;
        GLOBAL_MAX_IN_FLIGHT.fetch_max(global, Ordering::SeqCst);
        // Sample the live thread count only when this read sets a new
        // concurrency record, so the /proc walk cannot dominate the run.
        if depth > self.max_in_flight.fetch_max(depth, Ordering::SeqCst) {
            self.max_threads_at_peak
                .fetch_max(live_threads(), Ordering::SeqCst);
        }
        self.reads.fetch_add(1, Ordering::SeqCst);
        if self.delay_us > 0 {
            std::thread::sleep(std::time::Duration::from_micros(self.delay_us));
        }
        let start = usize::try_from(offset)
            .unwrap_or(usize::MAX)
            .min(self.bytes.len());
        let available = &self.bytes[start..];
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.in_flight.fetch_sub(1, Ordering::SeqCst);
        GLOBAL_IN_FLIGHT.fetch_sub(1, Ordering::SeqCst);
        Ok(count)
    }
}

/// Live OS threads in this process, from `/proc/self/task`.
fn live_threads() -> usize {
    fs::read_dir("/proc/self/task")
        .map(|entries| entries.count())
        .unwrap_or(0)
}

fn payload(seed: u8) -> Vec<u8> {
    // Incompressible enough that a stored member is the honest shape, and
    // distinct per member so the verification below is real.
    let mut state = u64::from(seed).wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1;
    (0..MEMBER_BYTES)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            u8::try_from(state & 0xFF).unwrap_or(0)
        })
        .collect()
}

fn member_name(index: usize) -> String {
    format!("custom/member{index}.bin")
}

fn member_uri(index: usize) -> String {
    format!("/custom/member{index}.bin")
}

fn cfb_stream_name(index: usize) -> String {
    format!("Stream{index}")
}

fn opc_archive(payloads: &[Vec<u8>]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", content_types.as_bytes())?;
    writer.write_stored("_rels/.rels", root_relationships.as_bytes())?;
    writer.write_stored("word/document.xml", b"<w:document xmlns:w=\"x\"/>")?;
    for (index, bytes) in payloads.iter().enumerate() {
        writer.write_stored(&member_name(index), bytes)?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn cfb_archive(payloads: &[Vec<u8>]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut writer = OleWriter::new();
    for (index, bytes) in payloads.iter().enumerate() {
        let name = cfb_stream_name(index);
        writer.create_stream_owned(&[name.as_str()], bytes.clone())?;
    }
    let mut output = io::Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

/// One `ExecutionContext` per session, all charging children of `root`.
fn context(
    root: &Budget,
    scope: &'static str,
) -> Result<ExecutionContext, Box<dyn std::error::Error>> {
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(WORKERS.load(Ordering::SeqCst)).ok_or("worker count must be nonzero")?,
        NonZeroUsize::new(32).ok_or("task cap must be nonzero")?,
        NonZeroU64::new(64 * 1024 * 1024).ok_or("byte cap must be nonzero")?,
        64 * 1024,
    )?;
    // Nothing in this probe cancels; the token stays valid after the source
    // is dropped because both hold the same shared flag.
    let (_source, token) = CancellationSource::pair();
    let child = root.child(
        scope,
        Limits::new(
            512 * 1024 * 1024,
            512 * 1024 * 1024,
            64 * 1024 * 1024,
            1_000_000,
            1_024,
            512 * 1024 * 1024,
        ),
    );
    Ok(ExecutionContext::new(child, token, limits))
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let delay_us: u64 = std::env::args()
        .nth(1)
        .as_deref()
        .unwrap_or("0")
        .parse()
        .map_err(|_| "argument 1 must be a per-read delay in microseconds")?;
    let workers: usize = std::env::args()
        .nth(2)
        .as_deref()
        .unwrap_or("4")
        .parse()
        .map_err(|_| "argument 2 must be a worker count")?;
    WORKERS.store(workers, Ordering::SeqCst);
    let payloads: Vec<Vec<u8>> = (0..MEMBERS)
        .map(|index| payload(u8::try_from(index).unwrap_or(0).wrapping_add(0x11)))
        .collect();
    let opc_bytes = opc_archive(&payloads)?;
    let cfb_bytes = cfb_archive(&payloads)?;

    // One hierarchical root for all three sessions: the Budget IS shareable.
    let root = Budget::root(
        "probe0615",
        Limits::new(
            2 * 1024 * 1024 * 1024,
            2 * 1024 * 1024 * 1024,
            256 * 1024 * 1024,
            4_000_000,
            4_096,
            2 * 1024 * 1024 * 1024,
        ),
    );

    println!("# probe 0615: three parallel sessions in one process");
    println!("workers_per_session       {workers}");
    println!("opc_archive_bytes         {}", opc_bytes.len());
    println!("cfb_archive_bytes         {}", cfb_bytes.len());
    println!("member_bytes              {MEMBER_BYTES}");
    println!("members                   {MEMBERS}");
    println!("budget_roots              1");
    println!("read_delay_us             {delay_us}");

    let baseline = live_threads();
    println!("threads_baseline          {baseline}");

    // --- Session 1: soapberry-zip ParallelReadSession, through OpenSession ---
    let zip_session = OpenSession::new(context(&root, "zip-open-session")?)?;
    let after_zip_new = live_threads();
    println!("threads_after_zip_session_new   {after_zip_new}");
    let package = zip_session.from_bytes(&opc_bytes, ReadLimits::default())?;
    let after_zip_open = live_threads();
    println!("threads_after_eager_opc_open    {after_zip_open}");
    println!(
        "eager_opc_part_count            {}",
        package.iter_parts().count()
    );

    // --- Session 2: litchi-cfb SharedOleBulkRead ---
    let cfb_source = Arc::new(CountingSource::new(cfb_bytes.clone(), delay_us));
    let ole = SharedOleFile::open_with_limits(
        Arc::clone(&cfb_source) as Arc<dyn ReadAt>,
        SharedOleFileLimits::new(u64::try_from(cfb_bytes.len())?)?,
    )?;
    let after_cfb_open = live_threads();
    println!("threads_after_cfb_open          {after_cfb_open}");
    cfb_source.reset();
    let cfb_names: Vec<String> = (0..MEMBERS).map(cfb_stream_name).collect();
    let cfb_path_storage: Vec<Vec<&str>> = cfb_names.iter().map(|n| vec![n.as_str()]).collect();
    let cfb_paths: Vec<&[&str]> = cfb_path_storage.iter().map(Vec::as_slice).collect();
    let bulk = ole.bulk_read(context(&root, "cfb-bulk-read")?);
    let streams = bulk.read_streams(&cfb_paths)?;
    let after_cfb_bulk = live_threads();
    for (index, stream) in streams.iter().enumerate() {
        assert_eq!(stream, &payloads[index], "CFB stream {index} differs");
    }
    println!("threads_after_cfb_bulk_read     {after_cfb_bulk}");
    println!(
        "cfb_max_concurrent_read_at      {}",
        cfb_source.max_in_flight.load(Ordering::SeqCst)
    );
    println!(
        "cfb_read_at_calls               {}",
        cfb_source.reads.load(Ordering::SeqCst)
    );
    println!(
        "cfb_threads_at_read_peak        {}",
        cfb_source.max_threads_at_peak.load(Ordering::SeqCst)
    );

    // --- Session 3: SourceBackedPackage::read_parts_ordered scoped workers ---
    let opc_source = Arc::new(CountingSource::new(opc_bytes.clone(), delay_us));
    let source_package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::clone(&opc_source) as Arc<dyn ReadAt>,
        ReadLimits::default(),
        context(&root, "opc-part-batch")?,
    )?;
    let after_source_open = live_threads();
    println!("threads_after_source_open       {after_source_open}");
    opc_source.reset();
    let uris: Vec<PackURI> = (0..MEMBERS)
        .map(|index| PackURI::new(member_uri(index)))
        .collect::<Result<_, _>>()?;
    let batch = source_package.read_parts_ordered(&uris)?;
    let after_batch = live_threads();
    assert_eq!(batch.len(), MEMBERS, "batch returned the wrong part count");
    for (index, expected) in payloads.iter().enumerate() {
        let part = batch.get(index).ok_or("missing batch part")?;
        assert_eq!(part.as_bytes(), expected.as_slice(), "part {index} differs");
    }
    println!("threads_after_part_batch        {after_batch}");
    println!(
        "opc_max_concurrent_read_at      {}",
        opc_source.max_in_flight.load(Ordering::SeqCst)
    );
    println!(
        "opc_read_at_calls               {}",
        opc_source.reads.load(Ordering::SeqCst)
    );
    println!(
        "opc_threads_at_batch_peak       {}",
        opc_source.max_threads_at_peak.load(Ordering::SeqCst)
    );

    // --- All three sessions live at once ---
    // Re-run each parallel operation with every session still held, and sample
    // the live thread count from inside the batch workers' own reads.
    // A second source-backed package so the joint batch reads cold rather than
    // from the first package's part cache.
    let joint_source = Arc::new(CountingSource::new(opc_bytes.clone(), delay_us));
    let joint_package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::clone(&joint_source) as Arc<dyn ReadAt>,
        ReadLimits::default(),
        context(&root, "opc-part-batch-joint")?,
    )?;
    joint_source.reset();
    cfb_source.reset();
    GLOBAL_MAX_IN_FLIGHT.store(0, Ordering::SeqCst);
    let joint = std::thread::scope(|scope| -> Result<usize, String> {
        let cfb_handle = scope.spawn(|| bulk.read_streams(&cfb_paths).map_err(|e| e.to_string()));
        let opc_handle = scope.spawn(|| {
            joint_package
                .read_parts_ordered(&uris)
                .map(|b| b.len())
                .map_err(|e| e.to_string())
        });
        let zip_handle = scope.spawn(|| {
            zip_session
                .from_bytes(&opc_bytes, ReadLimits::default())
                .map(|p| p.iter_parts().count())
                .map_err(|e| e.to_string())
        });
        let mut peak = 0usize;
        while !(cfb_handle.is_finished() && opc_handle.is_finished() && zip_handle.is_finished()) {
            peak = peak.max(live_threads());
            std::hint::spin_loop();
        }
        peak = peak.max(live_threads());
        cfb_handle
            .join()
            .map_err(|_| "cfb worker panicked".to_string())??;
        opc_handle
            .join()
            .map_err(|_| "opc worker panicked".to_string())??;
        zip_handle
            .join()
            .map_err(|_| "zip worker panicked".to_string())??;
        Ok(peak)
    })?;
    // Three probe-owned driver threads plus the sampling main thread are part
    // of this figure and are reported so the session share can be derived.
    println!("threads_peak_all_sessions_live  {joint}");
    println!("probe_driver_threads            3");
    println!(
        "joint_cfb_max_concurrent_read_at {}",
        cfb_source.max_in_flight.load(Ordering::SeqCst)
    );
    println!(
        "joint_opc_max_concurrent_read_at {}",
        joint_source.max_in_flight.load(Ordering::SeqCst)
    );
    println!(
        "joint_opc_read_at_calls          {}",
        joint_source.reads.load(Ordering::SeqCst)
    );
    println!(
        "joint_threads_at_opc_read_peak   {}",
        joint_source.max_threads_at_peak.load(Ordering::SeqCst)
    );
    println!(
        "joint_process_max_concurrent_read_at {}",
        GLOBAL_MAX_IN_FLIGHT.load(Ordering::SeqCst)
    );

    drop(batch);
    drop(joint_package);
    let after_all = live_threads();
    println!("threads_after_operations        {after_all}");
    drop(source_package);
    drop(bulk);
    drop(ole);
    drop(package);
    drop(zip_session);
    println!("threads_after_sessions_dropped  {}", live_threads());
    // Rayon terminates a dropped pool's workers asynchronously, so sample
    // again after a settle window before saying anything about persistence.
    std::thread::sleep(std::time::Duration::from_millis(250));
    println!("threads_after_250ms_settle      {}", live_threads());
    Ok(())
}
