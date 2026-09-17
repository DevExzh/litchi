//! Scratch probe for change 0662: parallel deflate of a publication's changed
//! member set, on the real `SourceBackedPackage` publication route.
//!
//! `tools/perf-baseline`'s managed XLSX selectors construct their execution
//! context with exactly one worker (`lib.rs:44497`), so no harness selector can
//! drive a wave wider than one. This probe opens each fixture through the
//! documented managed constructor
//! `SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context`
//! and publishes through the documented `write_part_overlays_shared_to_stream`.
//!
//! A "changed set" is built by replacing the first `k` ordinary XML Parts whose
//! source member is Deflate-compressed with their own bytes plus a trailing XML
//! comment — a minimal, well-formed change that keeps realistic payload sizes.
//!
//! Modes:
//!   corpus <root> [k]            one line per fixture: digests at widths 1/2/4/8
//!                                against the unmanaged sequential publication
//!   census <root> [k]            per fixture: tasks, total, largest, remainder
//!   time <file> <k> <workers> <warmup> <samples> [pool]
//!                                per-sample open+publish nanoseconds; `pool`
//!                                attaches a caller-owned worker facility
//!   open <file> <warmup> <samples>
//!                                per-sample open-only nanoseconds
//!   sweep <members> <bytes> <workers> <warmup> <samples> [pool]
//!                                synthetic package crossover, OPC route
//!   zipsweep <members> <bytes> <workers> <warmup> <samples>
//!                                the same crossover at the ZIP publication
//!                                boundary with both byte thresholds at zero,
//!                                which is what the balance rule is chosen from

use std::env;
use std::fs;
use std::io::Write;
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Instant;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    ScopedWorkers,
};
use litchi_opc::{PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits};
use sha2::{Digest, Sha256};

const EXTENSIONS: &[&str] = &[
    "xlsx", "xlsm", "xltx", "xltm", "docx", "docm", "dotx", "dotm", "pptx", "pptm", "potx", "ppsx",
];

const MARKER: &[u8] = b"<!--0662-->";

/// A caller-owned Rayon pool offered to the library as an execution facility.
///
/// The library never creates this; the caller does, exactly as ADR 0031 §4
/// describes. It is built once and reused by every publication.
#[derive(Debug)]
struct PoolWorkers {
    pool: rayon::ThreadPool,
}

impl ScopedWorkers for PoolWorkers {
    fn run_all(&self, tasks: &mut [&mut (dyn FnMut() + Send)]) {
        use rayon::prelude::*;
        self.pool.install(|| tasks.par_iter_mut().for_each(|task| task()));
    }
}

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe0662 corpus <root> [k] | census <root> [k] | time <file> <k> <workers> <warmup> <samples> [pool] | open <file> <warmup> <samples> | sweep <members> <bytes> <workers> <warmup> <samples> [pool]";
    match arguments.first().map(String::as_str) {
        Some("corpus") => {
            let root = arguments.get(1).expect(usage);
            let parts = arguments.get(2).map_or(64, |value| value.parse().expect(usage));
            corpus(Path::new(root), parts);
        },
        Some("census") => {
            let root = arguments.get(1).expect(usage);
            let parts = arguments.get(2).map_or(64, |value| value.parse().expect(usage));
            census(Path::new(root), parts);
        },
        Some("time") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let parts: usize = arguments.get(2).expect(usage).parse().expect(usage);
            let workers: usize = arguments.get(3).expect(usage).parse().expect(usage);
            let warmup: usize = arguments.get(4).expect(usage).parse().expect(usage);
            let samples: usize = arguments.get(5).expect(usage).parse().expect(usage);
            let pooled = arguments.get(6).is_some_and(|value| value == "pool");
            time(&file, parts, workers, warmup, samples, pooled);
        },
        Some("open") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let warmup: usize = arguments.get(2).expect(usage).parse().expect(usage);
            let samples: usize = arguments.get(3).expect(usage).parse().expect(usage);
            time_open(&file, warmup, samples);
        },
        Some("zipsweep") => {
            let members: usize = arguments.get(1).expect(usage).parse().expect(usage);
            let bytes: usize = arguments.get(2).expect(usage).parse().expect(usage);
            let workers: usize = arguments.get(3).expect(usage).parse().expect(usage);
            let warmup: usize = arguments.get(4).expect(usage).parse().expect(usage);
            let samples: usize = arguments.get(5).expect(usage).parse().expect(usage);
            zip_sweep(members, bytes, workers, warmup, samples);
        },
        Some("sweep") => {
            let members: usize = arguments.get(1).expect(usage).parse().expect(usage);
            let bytes: usize = arguments.get(2).expect(usage).parse().expect(usage);
            let workers: usize = arguments.get(3).expect(usage).parse().expect(usage);
            let warmup: usize = arguments.get(4).expect(usage).parse().expect(usage);
            let samples: usize = arguments.get(5).expect(usage).parse().expect(usage);
            let pooled = arguments.get(6).is_some_and(|value| value == "pool");
            sweep(members, bytes, workers, warmup, samples, pooled);
        },
        _ => {
            eprintln!("{usage}");
            std::process::exit(2);
        },
    }
}

fn context_for(workers: usize, facility: Option<Arc<dyn ScopedWorkers>>) -> ExecutionContext {
    let budget = Budget::root(
        "probe0662",
        Limits::new(
            2 * 1024 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (source, cancellation) = CancellationSource::pair();
    // The probe never cancels; leak the handle so the token stays live.
    std::mem::forget(source);
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(workers).expect("worker count"),
        NonZeroUsize::new(workers.max(64)).expect("task count"),
        NonZeroU64::new(1024 * 1024 * 1024).expect("byte bound"),
        0,
    )
    .expect("execution limits");
    let context = ExecutionContext::new(budget, cancellation, limits);
    match facility {
        Some(workers) => context.with_scoped_workers(workers),
        None => context,
    }
}

/// Opens a package and returns the replacement set for its first `parts`
/// Deflate-compressed XML Parts.
fn replacements(
    source: &Arc<dyn ReadAt>,
    parts: usize,
) -> Result<Vec<(PackURI, Arc<Vec<u8>>)>, String> {
    let package = SourceBackedPackage::from_read_at_with_limits(Arc::clone(source), read_limits())
        .map_err(|error| error.to_string())?;
    let mut selected = Vec::new();
    for view in package.iter_parts() {
        if selected.len() >= parts.min(64) {
            break;
        }
        let name = view.partname().clone();
        if !name.as_str().ends_with(".xml") && !name.as_str().ends_with(".rels") {
            continue;
        }
        let Ok(data) = package.part(&name).and_then(|part| part.data()) else {
            continue;
        };
        if data.as_bytes().is_empty() {
            continue;
        }
        // The publication audits every *overlay* payload as authored XML, so
        // a replacement must be compact even when the producer's own bytes
        // are not (which, for real packages, they usually are not: change
        // 0602 found 94 of 95 real packages non-compact). The probe therefore
        // re-serializes the part's own content compactly — which is what a
        // real editor's writer emits — and appends a marker comment so the
        // member is genuinely changed rather than copied.
        let Some(payload) = compact_overlay(data.as_bytes()) else {
            continue;
        };
        selected.push((name, Arc::new(payload)));
    }
    if selected.is_empty() {
        return Err("no eligible parts".to_string());
    }
    Ok(selected)
}

/// Re-serializes `input` as compact XML and appends the change marker.
///
/// Returns `None` when the result would not pass the publication's own
/// authored-XML audit, so the probe never asks the library to accept a payload
/// its documented contract refuses.
fn compact_overlay(input: &[u8]) -> Option<Vec<u8>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(input);
    let mut writer = quick_xml::Writer::new(Vec::with_capacity(input.len()));
    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Text(text)) => {
                if text.iter().any(|byte| !byte.is_ascii_whitespace()) {
                    writer.write_event(Event::Text(text)).ok()?;
                }
            },
            Ok(event) => writer.write_event(event).ok()?,
            Err(_) => return None,
        }
    }
    let mut payload = writer.into_inner();
    payload.extend_from_slice(MARKER);
    xml_minifier::audit::verify_authored(&payload, xml_minifier::audit::Limits::default()).ok()?;
    Some(payload)
}

fn read_limits() -> ReadLimits {
    ReadLimits::default()
}

fn publish(
    source: &Arc<dyn ReadAt>,
    context: Option<ExecutionContext>,
    replacements: Vec<(PackURI, Arc<Vec<u8>>)>,
) -> Result<Vec<u8>, String> {
    let package = match context {
        Some(context) => {
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                Arc::clone(source),
                read_limits(),
                SourceCacheLimits::default(),
                context,
            )
        },
        None => SourceBackedPackage::from_read_at_with_limits(Arc::clone(source), read_limits()),
    }
    .map_err(|error| error.to_string())?;
    let mut output = Vec::new();
    package
        .write_part_overlays_shared_to_stream(&mut output, replacements)
        .map_err(|error| error.to_string())?;
    Ok(output)
}

fn digest(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    format!("{:x}", hasher.finalize())[..32].to_string()
}

fn corpus(root: &Path, parts: usize) {
    let mut files = Vec::new();
    collect(root, &mut files);
    files.sort();
    for file in files {
        let Ok(bytes) = fs::read(&file) else { continue };
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
        let set = match replacements(&source, parts) {
            Ok(set) => set,
            Err(error) => {
                println!("{} skip {error}", file.display());
                continue;
            },
        };
        let sequential = match publish(&source, None, set.clone()) {
            Ok(bytes) => bytes,
            Err(error) => {
                // A refused publication must be refused identically at every
                // width, with the same typed error text.
                let mut verdict = "same-refusal";
                for workers in [1_usize, 2, 4, 8] {
                    let scheduled = publish(&source, Some(context_for(workers, None)), set.clone());
                    match scheduled {
                        Err(scheduled) if scheduled == error => {},
                        _ => verdict = "DIFFERENT",
                    }
                }
                println!("{} refused {} {verdict} :: {error}", file.display(), set.len());
                continue;
            },
        };
        let reference = digest(&sequential);
        let mut line = format!("{} ok {} {reference}", file.display(), set.len());
        let mut verdict = "identical";
        for workers in [1_usize, 2, 4, 8] {
            let scheduled = publish(&source, Some(context_for(workers, None)), set.clone());
            let observed = match scheduled {
                Ok(bytes) => digest(&bytes),
                Err(error) => format!("error:{error}"),
            };
            if observed != reference {
                verdict = "DIFFERENT";
            }
            line.push_str(&format!(" w{workers}={observed}"));
        }
        let facility: Arc<dyn ScopedWorkers> = Arc::new(PoolWorkers {
            pool: rayon::ThreadPoolBuilder::new()
                .num_threads(4)
                .build()
                .expect("probe pool"),
        });
        let pooled = publish(&source, Some(context_for(4, Some(facility))), set.clone());
        let observed = match pooled {
            Ok(bytes) => digest(&bytes),
            Err(error) => format!("error:{error}"),
        };
        if observed != reference {
            verdict = "DIFFERENT";
        }
        line.push_str(&format!(" facility={observed} {verdict}"));
        println!("{line}");
    }
}

fn census(root: &Path, parts: usize) {
    let mut files = Vec::new();
    collect(root, &mut files);
    files.sort();
    println!("fixture parts total_bytes largest_bytes remainder_bytes");
    for file in files {
        let Ok(bytes) = fs::read(&file) else { continue };
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
        let Ok(set) = replacements(&source, parts) else {
            continue;
        };
        let mut sizes: Vec<u64> = set.iter().map(|(_, data)| data.len() as u64).collect();
        sizes.sort_unstable_by(|left, right| right.cmp(left));
        let total: u64 = sizes.iter().sum();
        let largest = sizes.first().copied().unwrap_or_default();
        println!(
            "{} {} {total} {largest} {}",
            file.display(),
            sizes.len(),
            total - largest
        );
    }
}

fn facility_for(workers: usize) -> Arc<dyn ScopedWorkers> {
    Arc::new(PoolWorkers {
        pool: rayon::ThreadPoolBuilder::new()
            .num_threads(workers)
            .build()
            .expect("probe pool"),
    })
}

fn time(file: &Path, parts: usize, workers: usize, warmup: usize, samples: usize, pooled: bool) {
    let bytes = fs::read(file).expect("fixture");
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    let set = replacements(&source, parts).expect("replacement set");
    let facility = pooled.then(|| facility_for(workers));
    for _ in 0..warmup {
        let context = context_for(workers, facility.clone());
        publish(&source, Some(context), set.clone()).expect("warmup publication");
    }
    let mut out = std::io::stdout().lock();
    for _ in 0..samples {
        let context = context_for(workers, facility.clone());
        let replacements = set.clone();
        let started = Instant::now();
        let published = publish(&source, Some(context), replacements).expect("publication");
        let elapsed = started.elapsed().as_nanos();
        std::hint::black_box(&published);
        writeln!(out, "{elapsed}").expect("sample");
    }
}

fn time_open(file: &Path, warmup: usize, samples: usize) {
    let bytes = fs::read(file).expect("fixture");
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    for _ in 0..warmup {
        let package =
            SourceBackedPackage::from_read_at_with_limits(Arc::clone(&source), read_limits())
                .expect("warmup open");
        std::hint::black_box(&package);
    }
    let mut out = std::io::stdout().lock();
    for _ in 0..samples {
        let started = Instant::now();
        let package =
            SourceBackedPackage::from_read_at_with_limits(Arc::clone(&source), read_limits())
                .expect("open");
        let elapsed = started.elapsed().as_nanos();
        std::hint::black_box(&package);
        writeln!(out, "{elapsed}").expect("sample");
    }
}

/// A synthetic package of `members` Deflate members of `bytes` each.
fn synthetic_package(members: usize, bytes: usize) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored(
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
        )
        .expect("content types");
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        )
        .expect("package relationships");
    writer
        .write_deflated("word/document.xml", &synthetic_payload(1, bytes))
        .expect("document");
    for index in 0..members {
        writer
            .write_deflated(
                &format!("custom/part-{index:04}.xml"),
                &synthetic_payload(index + 2, bytes),
            )
            .expect("member");
    }
    writer.finish_to_bytes().expect("archive")
}

fn synthetic_payload(seed: usize, bytes: usize) -> Vec<u8> {
    let mut data = Vec::with_capacity(bytes + 32);
    data.extend_from_slice(b"<p>");
    let mut state = (seed as u64).wrapping_mul(0x9e37_79b9_7f4a_7c15).max(1);
    while data.len() < bytes {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        data.extend_from_slice(format!("<c r=\"{}\"/>", state % 8192).as_bytes());
    }
    data.extend_from_slice(b"</p>");
    data
}

fn sweep(members: usize, bytes: usize, workers: usize, warmup: usize, samples: usize, pooled: bool) {
    let archive = synthetic_package(members, bytes);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(archive));
    let set: Vec<(PackURI, Arc<Vec<u8>>)> = (0..members.min(64))
        .map(|index| {
            (
                PackURI::new(format!("/custom/part-{index:04}.xml")).expect("part name"),
                Arc::new(synthetic_payload(index + 500, bytes)),
            )
        })
        .collect();
    let facility = pooled.then(|| facility_for(workers));
    for _ in 0..warmup {
        let context = context_for(workers, facility.clone());
        publish(&source, Some(context), set.clone()).expect("warmup publication");
    }
    let mut out = std::io::stdout().lock();
    for _ in 0..samples {
        let context = context_for(workers, facility.clone());
        let replacements = set.clone();
        let started = Instant::now();
        let published = publish(&source, Some(context), replacements).expect("publication");
        let elapsed = started.elapsed().as_nanos();
        std::hint::black_box(&published);
        writeln!(out, "{elapsed}").expect("sample");
    }
}

#[derive(Debug)]
struct NeverCancelled;

impl soapberry_zip::office::CancellationProbe for NeverCancelled {
    fn is_cancelled(&self) -> bool {
        false
    }
}

/// Times the ZIP publication boundary itself, with both balance thresholds at
/// zero so the crossover can be found rather than assumed.
///
/// A width of zero times the sequential `write_to` this crate has always had.
fn zip_sweep(members: usize, bytes: usize, workers: usize, warmup: usize, samples: usize) {
    use soapberry_zip::{
        CompressionMethod, PreservationAction, PreservationIndex, PreservationPlan,
        RegeneratedEntry, ZipArchive,
    };

    let archive_bytes = synthetic_package(members, bytes);
    let archive = ZipArchive::from_slice(archive_bytes.as_slice())
        .expect("synthetic archive")
        .into_zip_archive();
    let mut scratch = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut scratch).expect("preservation index");
    let mut plan = PreservationPlan::new();
    plan.try_reserve_exact(index.entries().len())
        .expect("plan reservation");
    let mut replaced = 0_usize;
    for (ordinal, entry) in index.entries().iter().enumerate() {
        let name = String::from_utf8(entry.raw_name_bytes().to_vec()).expect("member name");
        if name.starts_with("custom/part-") {
            plan.push(PreservationAction::Regenerate {
                id: entry.id(),
                entry: RegeneratedEntry::new(name, synthetic_payload(ordinal + 700, bytes))
                    .compression_method(CompressionMethod::Deflate),
            });
            replaced += 1;
        } else {
            plan.push(PreservationAction::Copy(entry.id()));
        }
    }
    assert_eq!(replaced, members, "every synthetic member is regenerated");

    let limits = soapberry_zip::office::ParallelWriteLimits::new(
        NonZeroUsize::new(workers.max(1)).expect("worker count"),
        NonZeroUsize::new(workers.max(64)).expect("task count"),
        0,
        0,
        0,
    )
    .expect("write limits");
    let session = soapberry_zip::office::ParallelWriteSession::new(limits);
    let wave = plan.deflate_wave(limits);
    let cancellation = NeverCancelled;
    let mut run = |sink: Vec<u8>| -> Vec<u8> {
        if workers == 0 {
            index.write_to(&plan, sink).expect("sequential publication")
        } else {
            index
                .write_to_with_session(&plan, sink, &session, wave, &cancellation)
                .expect("scheduled publication")
        }
    };
    for _ in 0..warmup {
        let published = run(Vec::with_capacity(archive_bytes.len() * 2));
        std::hint::black_box(&published);
    }
    let mut out = std::io::stdout().lock();
    for _ in 0..samples {
        let sink = Vec::with_capacity(archive_bytes.len() * 2);
        let started = Instant::now();
        let published = run(sink);
        let elapsed = started.elapsed().as_nanos();
        std::hint::black_box(&published);
        writeln!(out, "{elapsed}").expect("sample");
    }
    eprintln!(
        "tasks={} width={} total={} largest={} remainder={}",
        wave.tasks(),
        wave.width(),
        wave.total_bytes(),
        wave.largest_bytes(),
        wave.remainder_bytes()
    );
}

fn collect(root: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, files);
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| EXTENSIONS.contains(&extension))
        {
            files.push(path);
        }
    }
}
