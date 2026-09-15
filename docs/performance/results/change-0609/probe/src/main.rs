//! Scratch probe for change 0609 (facade `.doc` route sizing). Not part of the tree.
//!
//! Modes:
//!   census   <root>                      TSV: both routes' admission over every *.doc
//!   identity <root>                      TSV: value comparison on fixtures both routes admit
//!   alloc    <mode> <path>               counting-allocator peak and retained bytes
//!   readat   <mode> <path>               counting `ReadAt`: calls and bytes (source routes)
//!   loop     <mode> <path> <warm> <n>    isolation loop for callgrind / perf stat
//!
//! Modes for alloc/readat/loop:
//!   facade-open facade-text facade-count facade-para0
//!   snap-open   snap-para0

use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use litchi_core::{FileSource, Position, ReadAt, SourceVersion};
use litchi_doc::body_text::source::{Error as SnapError, SourceSnapshot};

// ---------------------------------------------------------------- allocator

static LIVE: AtomicU64 = AtomicU64::new(0);
static PEAK: AtomicU64 = AtomicU64::new(0);
static TOTAL: AtomicU64 = AtomicU64::new(0);
static COUNT: AtomicU64 = AtomicU64::new(0);
static ARMED: AtomicBool = AtomicBool::new(false);

struct Counting;

#[inline]
fn on_alloc(size: usize) {
    if !ARMED.load(Ordering::Relaxed) {
        return;
    }
    COUNT.fetch_add(1, Ordering::Relaxed);
    TOTAL.fetch_add(size as u64, Ordering::Relaxed);
    let live = LIVE.fetch_add(size as u64, Ordering::Relaxed) + size as u64;
    PEAK.fetch_max(live, Ordering::Relaxed);
}

#[inline]
fn on_free(size: usize) {
    if !ARMED.load(Ordering::Relaxed) {
        return;
    }
    LIVE.fetch_sub(size as u64, Ordering::Relaxed);
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let p = unsafe { System.alloc(layout) };
        if !p.is_null() {
            on_alloc(layout.size());
        }
        p
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        on_free(layout.size());
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let p = unsafe { System.alloc_zeroed(layout) };
        if !p.is_null() {
            on_alloc(layout.size());
        }
        p
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let p = unsafe { System.realloc(ptr, layout, new_size) };
        if !p.is_null() {
            on_free(layout.size());
            on_alloc(new_size);
        }
        p
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn arm() {
    LIVE.store(0, Ordering::SeqCst);
    PEAK.store(0, Ordering::SeqCst);
    TOTAL.store(0, Ordering::SeqCst);
    COUNT.store(0, Ordering::SeqCst);
    ARMED.store(true, Ordering::SeqCst);
}

fn disarm() -> (u64, u64, u64, u64) {
    ARMED.store(false, Ordering::SeqCst);
    (
        PEAK.load(Ordering::SeqCst),
        LIVE.load(Ordering::SeqCst),
        TOTAL.load(Ordering::SeqCst),
        COUNT.load(Ordering::SeqCst),
    )
}

// ------------------------------------------------------------ counting ReadAt

struct CountingSource {
    inner: FileSource,
    calls: AtomicU64,
    bytes: AtomicU64,
    versions: AtomicU64,
    lens: AtomicU64,
}

impl std::fmt::Debug for CountingSource {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("CountingSource").finish_non_exhaustive()
    }
}

impl ReadAt for CountingSource {
    fn read_at(&self, offset: u64, buf: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read_at(offset, buf)?;
        self.calls.fetch_add(1, Ordering::Relaxed);
        self.bytes.fetch_add(n as u64, Ordering::Relaxed);
        Ok(n)
    }
    fn len(&self) -> std::io::Result<u64> {
        self.lens.fetch_add(1, Ordering::Relaxed);
        self.inner.len()
    }
    fn version(&self) -> std::io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::Relaxed);
        self.inner.version()
    }
}

// ---------------------------------------------------------------- utilities

type BoxError = Box<dyn std::error::Error>;

fn walk(root: &Path, ext: &str, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, ext, out);
            } else if path
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.eq_ignore_ascii_case(ext))
                .unwrap_or(false)
            {
                out.push(path);
            }
        }
    }
}

fn short(err: impl std::fmt::Debug) -> String {
    let s = format!("{err:?}").replace(['\n', '\t'], " ");
    if s.len() > 160 {
        format!("{}...", &s[..160])
    } else {
        s
    }
}

/// Collapse a litchi-doc error to the refusal/limit family the design records.
fn snapshot_kind(err: &SnapError) -> String {
    match err {
        SnapError::Refused(reason) => format!("Refused::{reason:?}"),
        SnapError::Limit(_) => "Limit".to_string(),
        SnapError::Overlay(_) => "Overlay".to_string(),
        other => {
            let s = format!("{other:?}");
            let head = s.split(['(', ' ', '{']).next().unwrap_or("Other");
            head.to_string()
        },
    }
}

fn facade_kind(err: &litchi::common::Error) -> String {
    let s = format!("{err:?}");
    let head = s.split(['(', ' ', '{']).next().unwrap_or("Other");
    head.to_string()
}

fn file_source(path: &Path) -> Result<Arc<dyn ReadAt>, BoxError> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn digest(text: &str) -> u64 {
    // FNV-1a over the UTF-8 bytes; only used to compare two strings cheaply.
    let mut h: u64 = 0xcbf2_9ce4_8422_2325;
    for b in text.as_bytes() {
        h ^= u64::from(*b);
        h = h.wrapping_mul(0x1000_0000_01b3);
    }
    h
}

// ------------------------------------------------------------------- census

fn census(root: &Path) {
    let mut files = Vec::new();
    walk(root, "doc", &mut files);
    files.sort();
    println!("file\tbytes\tfacade\tfacade_kind\tsnapshot\tsnapshot_kind");
    for path in files {
        let bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let (facade, fkind) = match litchi::Document::open(&path) {
            Ok(doc) => {
                black_box(&doc);
                ("OK".to_string(), String::new())
            },
            Err(e) => (format!("ERR {}", short(&e)), facade_kind(&e)),
        };
        let (snapshot, skind) = match file_source(&path) {
            Ok(source) => match SourceSnapshot::open(source) {
                Ok(snap) => {
                    black_box(&snap);
                    ("OK".to_string(), String::new())
                },
                Err(e) => (format!("ERR {}", short(&e)), snapshot_kind(&e)),
            },
            Err(e) => (format!("IO {}", short(&e)), "Io".to_string()),
        };
        println!(
            "{}\t{bytes}\t{facade}\t{fkind}\t{snapshot}\t{skind}",
            path.display()
        );
    }
}

// ----------------------------------------------------------------- identity

/// Compare, on every fixture, what each route can answer for the three facade
/// queries the brief names: `text()`, `paragraph_count()`, `paragraph_text(0)`.
fn identity(root: &Path) {
    let mut files = Vec::new();
    walk(root, "doc", &mut files);
    files.sort();
    println!(
        "file\tbytes\tfacade_text_len\tfacade_text_digest\tfacade_paras\tfacade_para0_len\tfacade_para0_digest\tsnap_para0\tsnap_para0_len\tsnap_para0_digest\tsnap_walk_paras\tsnap_walk_text_len\tsnap_walk_digest"
    );
    for path in files {
        let bytes = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        let mut ftl = String::from("-");
        let mut ftd = String::from("-");
        let mut fpc = String::from("-");
        let mut fp0l = String::from("-");
        let mut fp0d = String::from("-");
        if let Ok(doc) = litchi::Document::open(&path) {
            if let Ok(text) = doc.text() {
                ftl = text.len().to_string();
                ftd = format!("{:016x}", digest(&text));
            }
            if let Ok(count) = doc.paragraph_count() {
                fpc = count.to_string();
            }
            if let Ok(Some(p)) = doc.paragraph_text(0) {
                fp0l = p.len().to_string();
                fp0d = format!("{:016x}", digest(&p));
            }
        }

        let mut s0 = String::from("-");
        let mut s0l = String::from("-");
        let mut s0d = String::from("-");
        let mut walk_n = String::from("-");
        let mut walk_len = String::from("-");
        let mut walk_d = String::from("-");
        if let Ok(source) = file_source(&path)
            && let Ok(snap) = SourceSnapshot::open(source)
        {
            match snap.paragraph(Position::new(0)) {
                Ok(p) => {
                    s0 = "OK".to_string();
                    s0l = p.text().len().to_string();
                    s0d = format!("{:016x}", digest(p.text()));
                },
                Err(e) => s0 = snapshot_kind(&e),
            }
            // Walk positions until the snapshot refuses, concatenating text the
            // way a `text()` built on this reader would have to.
            let mut joined = String::new();
            let mut n = 0usize;
            while let Ok(p) = snap.paragraph(Position::new(n)) {
                joined.push_str(p.text());
                n += 1;
                if n > 100_000 {
                    break;
                }
            }
            walk_n = n.to_string();
            walk_len = joined.len().to_string();
            walk_d = format!("{:016x}", digest(&joined));
        }
        println!(
            "{}\t{bytes}\t{ftl}\t{ftd}\t{fpc}\t{fp0l}\t{fp0d}\t{s0}\t{s0l}\t{s0d}\t{walk_n}\t{walk_len}\t{walk_d}",
            path.display()
        );
    }
}

// ------------------------------------------------------------------ workloads

fn run_facade(mode: &str, path: &Path) -> Result<u64, BoxError> {
    let doc = litchi::Document::open(path)?;
    let acc = match mode {
        "facade-open" => 1,
        "facade-text" => doc.text().map(|t| t.len() as u64).unwrap_or(0),
        "facade-count" => doc.paragraph_count().map(|c| c as u64).unwrap_or(0),
        "facade-para0" => doc
            .paragraph_text(0)
            .ok()
            .flatten()
            .map(|t| t.len() as u64)
            .unwrap_or(0),
        other => return Err(format!("unknown facade mode {other}").into()),
    };
    black_box(&doc);
    Ok(acc)
}

fn run_snapshot(mode: &str, source: Arc<dyn ReadAt>) -> Result<u64, BoxError> {
    if mode == "snap-open-try" {
        // The probe a source-first facade route would pay before falling back:
        // an open that is expected to be refused, with the refusal swallowed.
        return Ok(match SourceSnapshot::open(source) {
            Ok(snap) => {
                black_box(&snap);
                1
            },
            Err(error) => {
                black_box(&error);
                0
            },
        });
    }
    let snap = SourceSnapshot::open(source)?;
    let acc = match mode {
        "snap-open" => 1,
        "snap-para0" => snap
            .paragraph(Position::new(0))
            .map(|p| p.text().len() as u64)
            .unwrap_or(0),
        other => return Err(format!("unknown snapshot mode {other}").into()),
    };
    black_box(&snap);
    Ok(acc)
}

fn run_once(mode: &str, path: &Path) -> Result<u64, BoxError> {
    if mode.starts_with("facade-") {
        run_facade(mode, path)
    } else {
        run_snapshot(mode, file_source(path)?)
    }
}

// --------------------------------------------------------------------- alloc

fn alloc_mode(mode: &str, path: &Path) -> Result<(), BoxError> {
    // One untimed warm-up so lazily initialised statics are not charged.
    let _ = run_once(mode, path);
    if mode.starts_with("facade-") {
        arm();
        let doc = litchi::Document::open(path)?;
        let peak_open = PEAK.load(Ordering::SeqCst);
        let live_open = LIVE.load(Ordering::SeqCst);
        let acc = match mode {
            "facade-open" => 1,
            "facade-text" => doc.text().map(|t| t.len() as u64).unwrap_or(0),
            "facade-count" => doc.paragraph_count().map(|c| c as u64).unwrap_or(0),
            "facade-para0" => doc
                .paragraph_text(0)
                .ok()
                .flatten()
                .map(|t| t.len() as u64)
                .unwrap_or(0),
            other => return Err(format!("unknown mode {other}").into()),
        };
        let (peak, live, total, count) = disarm();
        println!(
            "mode\t{mode}\nfile\t{}\nacc\t{acc}\npeak_open_bytes\t{peak_open}\nretained_open_bytes\t{live_open}\npeak_bytes\t{peak}\nretained_bytes\t{live}\ntotal_alloc_bytes\t{total}\nalloc_count\t{count}",
            path.display()
        );
        drop(doc);
    } else {
        let source = file_source(path)?;
        arm();
        let snap = SourceSnapshot::open(source)?;
        let peak_open = PEAK.load(Ordering::SeqCst);
        let live_open = LIVE.load(Ordering::SeqCst);
        let acc = match mode {
            "snap-open" => 1,
            "snap-para0" => snap
                .paragraph(Position::new(0))
                .map(|p| p.text().len() as u64)
                .unwrap_or(0),
            other => return Err(format!("unknown mode {other}").into()),
        };
        let (peak, live, total, count) = disarm();
        println!(
            "mode\t{mode}\nfile\t{}\nacc\t{acc}\npeak_open_bytes\t{peak_open}\nretained_open_bytes\t{live_open}\npeak_bytes\t{peak}\nretained_bytes\t{live}\ntotal_alloc_bytes\t{total}\nalloc_count\t{count}",
            path.display()
        );
        drop(snap);
    }
    Ok(())
}

// -------------------------------------------------------------------- readat

fn readat_mode(mode: &str, path: &Path) -> Result<(), BoxError> {
    let counting = Arc::new(CountingSource {
        inner: FileSource::open(path)?,
        calls: AtomicU64::new(0),
        bytes: AtomicU64::new(0),
        versions: AtomicU64::new(0),
        lens: AtomicU64::new(0),
    });
    let handle = Arc::clone(&counting);
    let source: Arc<dyn ReadAt> = counting;
    let result = run_snapshot(mode, source);
    println!(
        "mode\t{mode}\nfile\t{}\nok\t{}\nread_calls\t{}\nread_bytes\t{}\nversion_calls\t{}\nlen_calls\t{}\nfile_bytes\t{}",
        path.display(),
        result.is_ok(),
        handle.calls.load(Ordering::SeqCst),
        handle.bytes.load(Ordering::SeqCst),
        handle.versions.load(Ordering::SeqCst),
        handle.lens.load(Ordering::SeqCst),
        std::fs::metadata(path).map(|m| m.len()).unwrap_or(0),
    );
    Ok(())
}

// ---------------------------------------------------------------------- loop

fn loop_mode(mode: &str, path: &Path, warm: usize, samples: usize) -> Result<(), BoxError> {
    for _ in 0..warm {
        let _ = black_box(run_once(mode, path));
    }
    let mut acc: u64 = 0;
    for _ in 0..samples {
        acc = acc.wrapping_add(black_box(run_once(mode, path))?);
    }
    println!("{mode} {} samples={samples} acc={acc}", path.display());
    Ok(())
}

// --------------------------------------------------------------------- bench

/// Paired wall-clock timing. Prints one nanoseconds-per-operation line per
/// sample, each sample being `ops` operations of `mode`.
fn bench_mode(
    mode: &str,
    path: &Path,
    warm: usize,
    samples: usize,
    ops: usize,
) -> Result<(), BoxError> {
    for _ in 0..warm {
        let _ = black_box(run_once(mode, path));
    }
    for _ in 0..samples {
        let start = std::time::Instant::now();
        let mut acc: u64 = 0;
        for _ in 0..ops {
            acc = acc.wrapping_add(black_box(run_once(mode, path))?);
        }
        let elapsed = start.elapsed();
        black_box(acc);
        println!("{}", elapsed.as_nanos() as u64 / ops as u64);
    }
    Ok(())
}

// ---------------------------------------------------------------------- main

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("census") => census(Path::new(&args[2])),
        Some("identity") => identity(Path::new(&args[2])),
        Some("alloc") => alloc_mode(&args[2], Path::new(&args[3]))?,
        Some("readat") => readat_mode(&args[2], Path::new(&args[3]))?,
        Some("bench") => bench_mode(
            &args[2],
            Path::new(&args[3]),
            args[4].parse()?,
            args[5].parse()?,
            args[6].parse()?,
        )?,
        Some("loop") => loop_mode(
            &args[2],
            Path::new(&args[3]),
            args[4].parse()?,
            args[5].parse()?,
        )?,
        _ => {
            eprintln!(
                "usage: census <root> | identity <root> | alloc <mode> <path> | readat <mode> <path> | loop <mode> <path> <warm> <n> | bench <mode> <path> <warm> <samples> <ops>"
            );
            std::process::exit(2);
        },
    }
    Ok(())
}
