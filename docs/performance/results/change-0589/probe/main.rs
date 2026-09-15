//! Scratch probe for change 0589 (source-identity fingerprint passes).
//!
//! Derived from the change-0587 `docppt_survey` driver; the path dependencies
//! are rewritten per leg so the same source measures both checkouts.
//!
//! census-doc <root>            : which *.doc the DOC SourceSnapshot admits
//! census-ppt <root>            : which *.ppt the PPT text-edit snapshot admits
//! counts <mode> <path>         : deterministic source-read counts for one op
//! profile <mode> <path> <w> <n>: loop, for callgrind / perf stat isolation
//! bench <mode> <path> <w> <n>  : per-sample nanoseconds, one line each
//!
//! modes: doc-snapshot-open ppt-textedit-open doc-snapshot-resolve

use std::hint::black_box;
use std::io;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

use litchi_core::{FileSource, ReadAt, SourceVersion};

type BoxError = Box<dyn std::error::Error>;

/// Positional adapter that counts every call the fence makes.
struct CountingSource {
    inner: FileSource,
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    len_calls: AtomicU64,
    version_calls: AtomicU64,
}

impl CountingSource {
    fn open(path: &Path) -> io::Result<Self> {
        Ok(Self {
            inner: FileSource::open(path)?,
            read_calls: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
            len_calls: AtomicU64::new(0),
            version_calls: AtomicU64::new(0),
        })
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        self.len_calls.fetch_add(1, Ordering::Relaxed);
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        let read = self.inner.read_at(offset, output)?;
        self.read_bytes.fetch_add(read as u64, Ordering::Relaxed);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.version_calls.fetch_add(1, Ordering::Relaxed);
        self.inner.version()
    }
}

fn walk(root: &Path, ext: &str, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, ext, out);
            } else if path
                .extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| e.eq_ignore_ascii_case(ext))
            {
                out.push(path);
            }
        }
    }
}

fn short(err: impl std::fmt::Debug) -> String {
    let s = format!("{err:?}").replace('\n', " ");
    if s.len() > 160 { format!("{}...", &s[..160]) } else { s }
}

fn file_source(path: &Path) -> Result<Arc<dyn ReadAt>, BoxError> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn run(mode: &str, source: Arc<dyn ReadAt>) -> Result<usize, BoxError> {
    match mode {
        "doc-snapshot-open" => {
            let snapshot = litchi_doc::body_text::source::SourceSnapshot::open(source)?;
            black_box(&snapshot);
            Ok(0)
        }
        "doc-snapshot-resolve" => {
            let snapshot = litchi_doc::body_text::source::SourceSnapshot::open(source)?;
            let paragraph = snapshot.paragraph(litchi_core::Position::new(0))?;
            let n = paragraph.text().len();
            black_box(&paragraph);
            Ok(n)
        }
        "cfb-index-open" => {
            // One CFB directory/FAT index parse, the unit the DOC ReadAt path
            // pays twice. Unchanged by change 0589; measured for its design
            // section only.
            let shared = litchi_cfb::SharedOleFile::open(source)?;
            let n = shared.file_size() as usize;
            black_box(&shared);
            Ok(n)
        }
        "ppt-textedit-open" => {
            let snapshot = litchi_ppt::text_edit::SourceSnapshot::open(source)?;
            black_box(&snapshot);
            Ok(0)
        }
        other => Err(format!("unknown mode {other}").into()),
    }
}

fn census(root: &Path, ext: &str, mode: &str) {
    let mut files = Vec::new();
    walk(root, ext, &mut files);
    files.sort();
    println!("file\tbytes\tresult");
    for path in files {
        let bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let result = match file_source(&path) {
            Ok(source) => match run(mode, source) {
                Ok(_) => "OK".to_string(),
                Err(e) => format!("ERR {}", short(e)),
            },
            Err(e) => format!("IO {}", short(e)),
        };
        println!("{}\t{bytes}\t{result}", path.display());
    }
}

fn counts(mode: &str, path: &Path) -> Result<(), BoxError> {
    let counting = Arc::new(CountingSource::open(path)?);
    let length = counting.inner.len()?;
    let source: Arc<dyn ReadAt> = counting.clone();
    run(mode, source)?;
    let read_bytes = counting.read_bytes.load(Ordering::Relaxed);
    println!(
        "{{\"mode\":\"{mode}\",\"file\":\"{}\",\"bytes\":{length},\"read_calls\":{},\"read_bytes\":{read_bytes},\"full_artifact_reads\":{:.4},\"len_calls\":{},\"version_calls\":{}}}",
        path.file_name().and_then(|n| n.to_str()).unwrap_or("?"),
        counting.read_calls.load(Ordering::Relaxed),
        read_bytes as f64 / length as f64,
        counting.len_calls.load(Ordering::Relaxed),
        counting.version_calls.load(Ordering::Relaxed),
    );
    Ok(())
}

fn profile(mode: &str, path: &Path, warmups: usize, samples: usize) -> Result<(), BoxError> {
    let mut checksum = 0u64;
    for _ in 0..(warmups + samples) {
        let source = file_source(path)?;
        checksum = checksum.wrapping_add(black_box(run(mode, source)?) as u64);
    }
    println!(
        "{{\"mode\":\"{mode}\",\"iterations\":{},\"checksum\":{checksum}}}",
        warmups + samples
    );
    Ok(())
}

fn bench(mode: &str, path: &Path, warmups: usize, samples: usize) -> Result<(), BoxError> {
    let mut checksum = 0u64;
    for _ in 0..warmups {
        let source = file_source(path)?;
        checksum = checksum.wrapping_add(black_box(run(mode, source)?) as u64);
    }
    for _ in 0..samples {
        let source = file_source(path)?;
        let start = Instant::now();
        let value = run(mode, source)?;
        let elapsed = start.elapsed().as_nanos();
        checksum = checksum.wrapping_add(black_box(value) as u64);
        println!("{elapsed}");
    }
    eprintln!("checksum={checksum}");
    Ok(())
}


fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn sha256_hex(bytes: &[u8]) -> String {
    use sha2::{Digest as _, Sha256};
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hex(&hasher.finalize())
}

/// Differential line for one artifact: every identity this change can reach.
///
/// Printed by both the before and the after build; the two reports must be
/// byte-identical, which is the whole value-identity claim.
fn digest(path: &Path) -> Result<(), BoxError> {
    let name = path.display().to_string();
    let bytes = std::fs::read(path)?;
    let mut fields: Vec<String> = vec![format!("bytes={}", bytes.len())];

    // 1. The empty-splice identity plan, generic ReadAt and owned.
    match litchi_cfb::SharedOleFile::open(file_source(path)?) {
        Ok(shared) => {
            let limits = litchi_cfb::StreamSpliceLimits::default();
            match shared.plan_same_length_stream_splices(Vec::new(), limits) {
                Ok(plan) => {
                    fields.push(format!("identity_src={}", hex(plan.source_fingerprint().as_bytes())));
                    fields.push(format!("identity_tgt={}", hex(plan.target_fingerprint().as_bytes())));
                    fields.push(format!("identity_noop={}", plan.is_noop()));
                    let mut out = Vec::new();
                    match plan.write_to(&mut out) {
                        Ok(report) => fields.push(format!(
                            "identity_out={} report_src={} report_tgt={}",
                            sha256_hex(&out),
                            hex(report.source_fingerprint().as_bytes()),
                            hex(report.target_fingerprint().as_bytes())
                        )),
                        Err(e) => fields.push(format!("identity_out=ERR {}", short(e))),
                    }
                },
                Err(e) => fields.push(format!("identity=ERR {}", short(e))),
            }

            // 2. An exact byte no-op splice and an effective one-byte splice
            //    on the first non-empty stream, so both span shapes are covered.
            let selected = shared
                .directory_entries()
                .find(|entry| entry.entry_type == 2 && entry.size > 0)
                .map(|entry| (entry.name.clone(), entry.size));
            if let Some((stream, size)) = selected {
                fields.push(format!("stream={stream} stream_len={size}"));
                let mut first = [0u8; 1];
                if shared.read_stream_range(&[stream.as_str()], 0, &mut first).is_ok() {
                    let limits = litchi_cfb::StreamSpliceLimits::default();
                    let expected: Arc<[u8]> = Arc::from(vec![first[0]]);
                    let noop = litchi_cfb::SameLengthStreamSplice::new(
                        vec![stream.clone()], 0, Arc::clone(&expected), Arc::clone(&expected));
                    match shared.plan_same_length_stream_splices(vec![noop], limits) {
                        Ok(plan) => {
                            let mut out = Vec::new();
                            let published = plan.write_to(&mut out).map(|_| sha256_hex(&out))
                                .unwrap_or_else(|e| format!("ERR {}", short(e)));
                            fields.push(format!(
                                "noop_src={} noop_tgt={} noop_spans={} noop_out={}",
                                hex(plan.source_fingerprint().as_bytes()),
                                hex(plan.target_fingerprint().as_bytes()),
                                plan.changed_spans(), published));
                        },
                        Err(e) => fields.push(format!("noop=ERR {}", short(e))),
                    }
                    let replacement: Arc<[u8]> = Arc::from(vec![first[0] ^ 0xff]);
                    let effective = litchi_cfb::SameLengthStreamSplice::new(
                        vec![stream.clone()], 0, expected, replacement);
                    match shared.plan_same_length_stream_splices(vec![effective], limits) {
                        Ok(plan) => {
                            let mut out = Vec::new();
                            let published = plan.write_to(&mut out).map(|_| sha256_hex(&out))
                                .unwrap_or_else(|e| format!("ERR {}", short(e)));
                            fields.push(format!(
                                "eff_src={} eff_tgt={} eff_spans={} eff_out={}",
                                hex(plan.source_fingerprint().as_bytes()),
                                hex(plan.target_fingerprint().as_bytes()),
                                plan.changed_spans(), published));
                        },
                        Err(e) => fields.push(format!("eff=ERR {}", short(e))),
                    }
                }
            }
        },
        Err(e) => fields.push(format!("cfb=ERR {}", short(e))),
    }

    // 3. The two format-owned snapshots this change is measured on.
    match litchi_doc::body_text::source::SourceSnapshot::open(file_source(path)?) {
        Ok(snapshot) => fields.push(format!("doc_snapshot={}", hex(snapshot.fingerprint().as_bytes()))),
        Err(e) => fields.push(format!("doc_snapshot=ERR {}", short(e))),
    }
    match litchi_ppt::text_edit::SourceSnapshot::open(file_source(path)?) {
        Ok(snapshot) => fields.push(format!("ppt_snapshot={}", hex(snapshot.fingerprint().as_bytes()))),
        Err(e) => fields.push(format!("ppt_snapshot=ERR {}", short(e))),
    }

    println!("{name}\t{}", fields.join("\t"));
    Ok(())
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("census-doc") => census(Path::new(&args[1]), "doc", "doc-snapshot-open"),
        Some("census-ppt") => census(Path::new(&args[1]), "ppt", "ppt-textedit-open"),
        Some("counts") => counts(&args[1], Path::new(&args[2]))?,
        Some("digest") => {
            let mut files = Vec::new();
            walk(Path::new(&args[1]), "doc", &mut files);
            walk(Path::new(&args[1]), "ppt", &mut files);
            files.sort();
            for file in &files {
                digest(file)?;
            }
        },
        Some("profile") => profile(&args[1], Path::new(&args[2]), args[3].parse()?, args[4].parse()?)?,
        Some("bench") => bench(&args[1], Path::new(&args[2]), args[3].parse()?, args[4].parse()?)?,
        _ => return Err("usage: census-doc|census-ppt <root> | counts <mode> <path> | profile|bench <mode> <path> <warmups> <samples>".into()),
    }
    Ok(())
}
