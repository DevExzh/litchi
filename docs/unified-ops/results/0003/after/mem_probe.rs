//! M0a step S10 allocation probe (plan 0003 risk R5).
//!
//! Counts the heap a facade open holds and allocates, through the facade API
//! of one revision:
//! - feature `before`: the pre-M0a facade types (`litchi::Document`,
//!   `litchi::Presentation`, `litchi::Workbook`), `open(path)` /
//!   `from_bytes(Vec<u8>)` and `text()`;
//! - feature `after`: `litchi::open(path)` / `litchi::open_bytes(Vec<u8>)` and
//!   `File::text()`.
//!
//! The counting allocator wraps `std::alloc::System` and counts requested
//! layout sizes, like the harness's `litchi-perf-baseline-alloc` wrapper: it
//! excludes allocator overhead and realloc overlap, and it is not RSS.
//!
//! Per fixture and route, after two warm-up rounds, every measured round
//! records (bytes relative to the live heap just before the route started):
//! - `held_after_open`: heap still live after the open, the handle alive. On
//!   the byte route the window starts before `fs::read`, so retained input
//!   bytes count and released ones do not;
//! - `peak_open`: the high-water mark during the open;
//! - `open_alloc_calls`: allocation and reallocation calls during the open;
//! - `peak_text`: the high-water mark during `text()`, the handle alive;
//! - `held_after_text`: heap live after `text()` returned and its `String`
//!   was dropped (lazy caches the read filled stay here);
//! - `live_after_drop`: heap live after the handle was dropped (0 = no leak).
//!
//! Output: one JSON object per line on stdout.

use std::alloc::{GlobalAlloc, Layout, System};
use std::path::Path;
use std::sync::atomic::{AtomicU64, Ordering::SeqCst};

struct Counting;

static LIVE: AtomicU64 = AtomicU64::new(0);
static PEAK: AtomicU64 = AtomicU64::new(0);
static CALLS: AtomicU64 = AtomicU64::new(0);

fn grow(bytes: usize) {
    let live = LIVE.fetch_add(bytes as u64, SeqCst) + bytes as u64;
    PEAK.fetch_max(live, SeqCst);
}

fn shrink(bytes: usize) {
    LIVE.fetch_sub(bytes as u64, SeqCst);
}

// SAFETY: every call forwards to `System` with the caller's arguments; the
// counters only observe successful results.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            CALLS.fetch_add(1, SeqCst);
            grow(layout.size());
        }
        ptr
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc_zeroed(layout) };
        if !ptr.is_null() {
            CALLS.fetch_add(1, SeqCst);
            grow(layout.size());
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        shrink(layout.size());
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let new = unsafe { System.realloc(ptr, layout, new_size) };
        if !new.is_null() {
            CALLS.fetch_add(1, SeqCst);
            if new_size >= layout.size() {
                grow(new_size - layout.size());
            } else {
                shrink(layout.size() - new_size);
            }
        }
        new
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

#[cfg(all(feature = "before", feature = "after"))]
compile_error!("select exactly one of the features `before` and `after`");

#[cfg(feature = "after")]
mod api {
    use std::path::Path;

    pub const REVISION_API: &str = "after: litchi::open_with(path, &opts) / litchi::open_bytes_with(Vec<u8>, &opts) -> litchi::File; File::text(); opts = OpenOptions::default() (retain_source(false) when S10_RETAIN_SOURCE=0)";

    pub type Handle = litchi::File;
    pub type OpenError = litchi::Error;

    /// Set once by `main` before any measured round: `S10_RETAIN_SOURCE=0`
    /// opens with `OpenOptions::retain_source(false)`, the R5 opt-out (plan
    /// 0003 §2.3). Read inside the window without allocating.
    pub static RETAIN: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(true);

    fn options() -> litchi::OpenOptions {
        litchi::OpenOptions::default()
            .retain_source(RETAIN.load(std::sync::atomic::Ordering::SeqCst))
    }

    pub fn open_path(_format: &str, path: &Path) -> Result<Handle, OpenError> {
        litchi::open_with(path, &options())
    }

    pub fn open_bytes(_format: &str, bytes: Vec<u8>) -> Result<Handle, OpenError> {
        litchi::open_bytes_with(bytes, &options())
    }

    pub fn text(handle: &Handle) -> Result<String, OpenError> {
        handle.text()
    }
}

#[cfg(feature = "before")]
mod api {
    use std::path::Path;

    pub const REVISION_API: &str = "before: litchi::{Document, Presentation, Workbook}::open(path) / from_bytes(Vec<u8>); text()";

    pub enum Handle {
        Document(litchi::Document),
        Presentation(litchi::Presentation),
        Workbook(litchi::Workbook),
    }

    /// `litchi::Workbook` reports a boxed error; the other two `litchi::Error`.
    pub type OpenError = Box<dyn std::error::Error + Send + Sync>;

    enum Family {
        Document,
        Presentation,
        Workbook,
    }

    fn family(format: &str) -> Family {
        match format {
            "doc" | "docx" | "rtf" | "odt" => Family::Document,
            "ppt" | "pptx" | "odp" => Family::Presentation,
            "xlsx" | "ods" => Family::Workbook,
            other => panic!("no facade family for {other}"),
        }
    }

    pub fn open_path(format: &str, path: &Path) -> Result<Handle, OpenError> {
        Ok(match family(format) {
            Family::Document => Handle::Document(litchi::Document::open(path)?),
            Family::Presentation => Handle::Presentation(litchi::Presentation::open(path)?),
            Family::Workbook => Handle::Workbook(litchi::Workbook::open(path)?),
        })
    }

    pub fn open_bytes(format: &str, bytes: Vec<u8>) -> Result<Handle, OpenError> {
        Ok(match family(format) {
            Family::Document => Handle::Document(litchi::Document::from_bytes(bytes)?),
            Family::Presentation => Handle::Presentation(litchi::Presentation::from_bytes(bytes)?),
            Family::Workbook => Handle::Workbook(litchi::Workbook::from_bytes(bytes)?),
        })
    }

    pub fn text(handle: &Handle) -> Result<String, OpenError> {
        match handle {
            Handle::Document(document) => document.text().map_err(Into::into),
            Handle::Presentation(presentation) => presentation.text().map_err(Into::into),
            Handle::Workbook(workbook) => workbook.text(),
        }
    }
}

#[derive(Clone, PartialEq, Eq)]
struct Sample {
    open_error: Option<String>,
    text_error: Option<String>,
    held_after_open: i64,
    peak_open: u64,
    open_alloc_calls: u64,
    text_bytes: Option<usize>,
    peak_text: Option<u64>,
    held_after_text: Option<i64>,
    live_after_drop: i64,
}

fn live() -> u64 {
    LIVE.load(SeqCst)
}

fn sample(format: &str, path: &Path, bytes_route: bool) -> Sample {
    let base = live();
    PEAK.store(base, SeqCst);
    let calls = CALLS.load(SeqCst);
    let opened = if bytes_route {
        let bytes = std::fs::read(path).expect("read fixture");
        api::open_bytes(format, bytes)
    } else {
        api::open_path(format, path)
    };
    let held_after_open = live() as i64 - base as i64;
    let peak_open = PEAK.load(SeqCst) - base;
    let open_alloc_calls = CALLS.load(SeqCst) - calls;
    let mut text_bytes = None;
    let mut peak_text = None;
    let mut held_after_text = None;
    let mut text_error = None;
    if let Ok(handle) = &opened {
        PEAK.store(live(), SeqCst);
        let text = api::text(handle);
        peak_text = Some(PEAK.load(SeqCst) - base);
        match text {
            Ok(text) => {
                text_bytes = Some(text.len());
                drop(text);
            },
            Err(error) => {
                // Formatted after the read's peak was taken.
                text_error = Some(error.to_string());
            },
        }
        held_after_text = Some(live() as i64 - base as i64 - text_error.as_ref().map_or(0, |e| e.capacity() as i64));
    }
    let open_error = opened.as_ref().err().map(ToString::to_string);
    drop(opened);
    let pending = open_error.as_ref().map_or(0, |e| e.capacity()) + text_error.as_ref().map_or(0, |e| e.capacity());
    let live_after_drop = live() as i64 - base as i64 - pending as i64;
    Sample {
        open_error,
        text_error,
        held_after_open,
        peak_open,
        open_alloc_calls,
        text_bytes,
        peak_text,
        held_after_text,
        live_after_drop,
    }
}

fn json_str(value: &str) -> String {
    let mut out = String::with_capacity(value.len() + 2);
    out.push('"');
    for ch in value.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

fn opt<T: std::fmt::Display>(value: &Option<T>) -> String {
    value.as_ref().map_or_else(|| "null".to_owned(), ToString::to_string)
}

fn opt_str(value: &Option<String>) -> String {
    value.as_deref().map_or_else(|| "null".to_owned(), json_str)
}

fn main() {
    #[cfg(feature = "after")]
    api::RETAIN.store(
        std::env::var("S10_RETAIN_SOURCE").map_or(true, |value| value != "0"),
        SeqCst,
    );
    let mut args = std::env::args().skip(1);
    let rounds: usize = args.next().expect("rounds").parse().expect("rounds");
    let warmup: usize = args.next().expect("warmup").parse().expect("warmup");
    let fixtures: Vec<(String, String)> = args
        .map(|arg| {
            let (format, path) = arg.split_once(':').expect("format:path");
            (format.to_owned(), path.to_owned())
        })
        .collect();
    for (format, path) in &fixtures {
        let path = Path::new(path);
        let input_bytes = std::fs::metadata(path).expect("fixture metadata").len();
        for bytes_route in [false, true] {
            for _ in 0..warmup {
                sample(format, path, bytes_route);
            }
            let samples: Vec<Sample> = (0..rounds).map(|_| sample(format, path, bytes_route)).collect();
            let deterministic = samples.windows(2).all(|pair| pair[0] == pair[1]);
            let first = &samples[0];
            let list = |f: &dyn Fn(&Sample) -> String| {
                samples.iter().map(f).collect::<Vec<_>>().join(",")
            };
            println!(
                "{{\"api\":{},\"format\":{},\"fixture\":{},\"input_bytes\":{},\"route\":{},\"rounds\":{},\"warmup\":{},\"deterministic\":{},\"open_error\":{},\"text_error\":{},\"held_after_open\":[{}],\"peak_open\":[{}],\"open_alloc_calls\":[{}],\"text_bytes\":{},\"peak_text\":[{}],\"held_after_text\":[{}],\"live_after_drop\":[{}]}}",
                json_str(api::REVISION_API),
                json_str(format),
                json_str(&path.display().to_string()),
                input_bytes,
                json_str(if bytes_route { "bytes" } else { "path" }),
                rounds,
                warmup,
                deterministic,
                opt_str(&first.open_error),
                opt_str(&first.text_error),
                list(&|s| s.held_after_open.to_string()),
                list(&|s| s.peak_open.to_string()),
                list(&|s| s.open_alloc_calls.to_string()),
                opt(&first.text_bytes),
                list(&|s| opt(&s.peak_text)),
                list(&|s| opt(&s.held_after_text)),
                list(&|s| s.live_after_drop.to_string()),
            );
        }
    }
}
