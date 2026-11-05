use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use litchi_odf_common::flat::Parts;
use sha2::{Digest, Sha256};

#[cfg(feature = "final")]
type ProbeLimits = litchi_odf_common::flat::Limits;
#[cfg(not(feature = "final"))]
#[derive(Clone, Copy)]
struct ProbeLimits;

#[cfg(feature = "final")]
fn default_limits() -> ProbeLimits {
    ProbeLimits::default()
}
#[cfg(not(feature = "final"))]
fn default_limits() -> ProbeLimits {
    ProbeLimits
}

#[cfg(feature = "final")]
fn strict_depth_limits() -> ProbeLimits {
    use litchi_odf_common::{SemanticLimits, SourceLimits};
    let semantic = SemanticLimits::builder().max_depth(2).unwrap().build().unwrap();
    ProbeLimits::new(SourceLimits::default(), semantic)
}
#[cfg(not(feature = "final"))]
fn strict_depth_limits() -> ProbeLimits {
    ProbeLimits
}

#[cfg(feature = "final")]
fn split_parts(source: &[u8], limits: &ProbeLimits) -> Result<Parts, String> {
    Parts::split(source, limits).map_err(|error| format!("{error:?}"))
}
#[cfg(not(feature = "final"))]
fn split_parts(source: &[u8], _limits: &ProbeLimits) -> Result<Parts, String> {
    Parts::split(source).map_err(|error| format!("{error:?}"))
}

#[cfg(feature = "final")]
fn to_package(parts: &Parts, limits: &ProbeLimits) -> Result<Vec<u8>, String> {
    parts.to_bytes(limits).map_err(|error| format!("{error:?}"))
}
#[cfg(not(feature = "final"))]
fn to_package(parts: &Parts, _limits: &ProbeLimits) -> Result<Vec<u8>, String> {
    parts.to_bytes().map_err(|error| format!("{error:?}"))
}

#[cfg(feature = "final")]
fn merge_parts(
    parts: &Parts,
    flat: &[u8],
    content: Option<&str>,
    styles: Option<&str>,
    meta: Option<&str>,
    settings: Option<&str>,
    limits: &ProbeLimits,
) -> Result<Vec<u8>, String> {
    parts
        .merge(flat, content, styles, meta, settings, limits)
        .map_err(|error| format!("{error:?}"))
}
#[cfg(not(feature = "final"))]
fn merge_parts(
    parts: &Parts,
    flat: &[u8],
    content: Option<&str>,
    styles: Option<&str>,
    meta: Option<&str>,
    settings: Option<&str>,
    _limits: &ProbeLimits,
) -> Result<Vec<u8>, String> {
    parts
        .merge(flat, content, styles, meta, settings)
        .map_err(|error| format!("{error:?}"))
}

#[cfg(feature = "final")]
fn merge_package(
    parts: &Parts,
    flat: &[u8],
    package: &[u8],
    limits: &ProbeLimits,
) -> Result<Vec<u8>, String> {
    parts
        .merge_package(flat, package, limits)
        .map_err(|error| format!("{error:?}"))
}
#[cfg(not(feature = "final"))]
fn merge_package(
    parts: &Parts,
    flat: &[u8],
    package: &[u8],
    _limits: &ProbeLimits,
) -> Result<Vec<u8>, String> {
    parts
        .merge_package(flat, package)
        .map_err(|error| format!("{error:?}"))
}

#[derive(Default)]
struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_REQUESTED: AtomicU64 = AtomicU64::new(0);
static ALLOC_LIVE: AtomicU64 = AtomicU64::new(0);
static ALLOC_MEASURING: AtomicBool = AtomicBool::new(false);
static ALLOC_INTERVAL_PEAK: AtomicU64 = AtomicU64::new(0);

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn observe_live(live: u64) {
    if ALLOC_MEASURING.load(Ordering::Relaxed) {
        let mut peak = ALLOC_INTERVAL_PEAK.load(Ordering::Relaxed);
        while live > peak {
            match ALLOC_INTERVAL_PEAK.compare_exchange_weak(
                peak,
                live,
                Ordering::Relaxed,
                Ordering::Relaxed,
            ) {
                Ok(_) => break,
                Err(current) => peak = current,
            }
        }
    }
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOC_REQUESTED.fetch_add(layout.size() as u64, Ordering::Relaxed);
            let live = ALLOC_LIVE.fetch_add(layout.size() as u64, Ordering::Relaxed) + layout.size() as u64;
            observe_live(live);
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOC_REQUESTED.fetch_add(layout.size() as u64, Ordering::Relaxed);
            let live = ALLOC_LIVE.fetch_add(layout.size() as u64, Ordering::Relaxed) + layout.size() as u64;
            observe_live(live);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        ALLOC_LIVE.fetch_sub(layout.size() as u64, Ordering::Relaxed);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if !result.is_null() {
            ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOC_REQUESTED.fetch_add(new_size as u64, Ordering::Relaxed);
            let old_size = layout.size() as u64;
            let new_size = new_size as u64;
            let live = if new_size >= old_size {
                ALLOC_LIVE.fetch_add(new_size - old_size, Ordering::Relaxed) + new_size - old_size
            } else {
                ALLOC_LIVE.fetch_sub(old_size - new_size, Ordering::Relaxed) - (old_size - new_size)
            };
            observe_live(live);
        }
        result
    }
}

#[derive(Clone, Copy)]
struct Baseline {
    calls: u64,
    requested: u64,
    live: u64,
}

#[derive(Clone, Copy)]
struct Metrics {
    calls: u64,
    requested: u64,
    retained: i64,
    peak: u64,
}

fn measure_begin() -> Baseline {
    let live = ALLOC_LIVE.load(Ordering::Relaxed);
    ALLOC_INTERVAL_PEAK.store(live, Ordering::Relaxed);
    ALLOC_MEASURING.store(true, Ordering::Relaxed);
    Baseline {
        calls: ALLOC_CALLS.load(Ordering::Relaxed),
        requested: ALLOC_REQUESTED.load(Ordering::Relaxed),
        live,
    }
}

fn measure_end(before: Baseline) -> Metrics {
    ALLOC_MEASURING.store(false, Ordering::Relaxed);
    let live = ALLOC_LIVE.load(Ordering::Relaxed);
    Metrics {
        calls: ALLOC_CALLS.load(Ordering::Relaxed).saturating_sub(before.calls),
        requested: ALLOC_REQUESTED
            .load(Ordering::Relaxed)
            .saturating_sub(before.requested),
        retained: live as i64 - before.live as i64,
        peak: ALLOC_INTERVAL_PEAK
            .load(Ordering::Relaxed)
            .saturating_sub(before.live),
    }
}

fn digest(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut text = String::with_capacity(64);
    for byte in digest {
        use std::fmt::Write;
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn error_report(route: &str, error: String, metrics: Metrics) -> serde_json::Value {
    serde_json::json!({
        "schema": "odf_flat_probe.v1",
        "route": route,
        "status": "error",
        "error": error,
        "alloc": {
            "calls": metrics.calls,
            "requested_bytes": metrics.requested,
            "retained_bytes": metrics.retained,
            "peak_bytes": metrics.peak,
        },
    })
}

fn output_report(
    route: &str,
    output: &[u8],
    source: &[u8],
    metrics: Metrics,
    extra: serde_json::Value,
) -> serde_json::Value {
    let mut value = serde_json::json!({
        "schema": "odf_flat_probe.v1",
        "route": route,
        "status": "ok",
        "output_bytes": output.len(),
        "output_sha256": digest(output),
        "source_sha256": digest(source),
        "output_equals_source": output == source,
        "alloc": {
            "calls": metrics.calls,
            "requested_bytes": metrics.requested,
            "retained_bytes": metrics.retained,
            "peak_bytes": metrics.peak,
        },
    });
    if let (Some(target), Some(fields)) = (value.as_object_mut(), extra.as_object()) {
        for (key, field) in fields {
            target.insert(key.clone(), field.clone());
        }
    }
    value
}

fn flat_source(path: &std::path::Path) -> Vec<u8> {
    std::fs::read(path).expect("flat fixture")
}

fn synthetic_depth_source() -> Vec<u8> {
    br#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text"><office:body><office:text/></office:body></office:document>"#.to_vec()
}

fn main() {
    let route = std::env::args().nth(1).unwrap_or_else(|| "help".to_string());
    let fixture = std::env::args()
        .nth(2)
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| {
            panic!("usage: probe ROUTE FIXTURE_PATH")
        });
    let source = flat_source(&fixture);
    let source = source.as_slice();
    let limits = default_limits();
    let report = match route.as_str() {
        "split" => {
            let before = measure_begin();
            let result = split_parts(source, &limits);
            let metrics = measure_end(before);
            match result {
                Ok(parts) => output_report(
                    "split",
                    parts.content_xml().as_bytes(),
                    source,
                    metrics,
                    serde_json::json!({
                        "content_bytes": parts.content_xml().len(),
                        "content_sha256": digest(parts.content_xml().as_bytes()),
                        "styles_bytes": parts.styles_xml().len(),
                        "styles_sha256": digest(parts.styles_xml().as_bytes()),
                        "meta_bytes": parts.meta_xml().map(str::len),
                        "settings_bytes": parts.settings_xml().map(str::len),
                    }),
                ),
                Err(error) => error_report("split", error, metrics),
            }
        }
        "to_package" => {
            let parts = split_parts(source, &limits).expect("fixture split");
            let before = measure_begin();
            let result = to_package(&parts, &limits);
            let metrics = measure_end(before);
            match result {
                Ok(package) => output_report(
                    "to_package",
                    &package,
                    source,
                    metrics,
                    serde_json::json!({ "package_bytes": package.len() }),
                ),
                Err(error) => error_report("to_package", error, metrics),
            }
        }
        "merge_noop" => {
            let parts = split_parts(source, &limits).expect("fixture split");
            let independent = source.to_vec();
            let before = measure_begin();
            let result = merge_parts(&parts, &independent, None, None, None, None, &limits);
            let metrics = measure_end(before);
            match result {
                Ok(output) => output_report(
                    "merge_noop",
                    &output,
                    source,
                    metrics,
                    serde_json::json!({
                        "independent_source": true,
                        "source_content_proof": digest(&independent),
                    }),
                ),
                Err(error) => error_report("merge_noop", error, metrics),
            }
        }
        "merge_edit" => {
            let parts = split_parts(source, &limits).expect("fixture split");
            let content = parts.content_xml().replace("Sheet", "Skeet");
            assert_ne!(content, parts.content_xml(), "fixture edit marker missing");
            let before = measure_begin();
            let result = merge_parts(&parts, source, Some(&content), None, None, None, &limits);
            let metrics = measure_end(before);
            match result {
                Ok(output) => output_report(
                    "merge_edit",
                    &output,
                    source,
                    metrics,
                    serde_json::json!({
                        "changed": output != source,
                        "edited_content_sha256": digest(content.as_bytes()),
                    }),
                ),
                Err(error) => error_report("merge_edit", error, metrics),
            }
        }
        "merge_foreign" => {
            let parts = split_parts(source, &limits).expect("fixture split");
            let mut foreign = source.to_vec();
            let marker = foreign.iter().position(|byte| *byte == b'S').expect("ASCII marker");
            foreign[marker] = b'T';
            assert_eq!(foreign.len(), source.len());
            let before = measure_begin();
            let result = merge_parts(&parts, &foreign, None, None, None, None, &limits);
            let metrics = measure_end(before);
            match result {
                Ok(output) => output_report(
                    "merge_foreign",
                    &output,
                    source,
                    metrics,
                    serde_json::json!({
                        "same_length_foreign": true,
                        "foreign_sha256": digest(&foreign),
                        "source_content_proof": digest(&foreign),
                    }),
                ),
                Err(error) => error_report("merge_foreign", error, metrics),
            }
        }
        "merge_tight" => {
            let synthetic = synthetic_depth_source();
            let loose = default_limits();
            let parts = split_parts(&synthetic, &loose).expect("synthetic split");
            let external = synthetic.to_vec();
            let tight = strict_depth_limits();
            let before = measure_begin();
            let result = merge_parts(&parts, &external, None, None, None, None, &tight);
            let metrics = measure_end(before);
            match result {
                Ok(output) => output_report(
                    "merge_tight",
                    &output,
                    &synthetic,
                    metrics,
                    serde_json::json!({
                        "independent_source": true,
                        "empty_element_depth": 3,
                        "tight_depth": 2,
                        "expected_tight_depth_refusal": cfg!(feature = "final"),
                    }),
                ),
                Err(error) => error_report("merge_tight", error, metrics),
            }
        }
        "split_tight" => {
            let synthetic = synthetic_depth_source();
            let tight = strict_depth_limits();
            let before = measure_begin();
            let result = split_parts(&synthetic, &tight);
            let metrics = measure_end(before);
            match result {
                Ok(parts) => output_report(
                    "split_tight",
                    parts.content_xml().as_bytes(),
                    &synthetic,
                    metrics,
                    serde_json::json!({
                        "empty_element_depth": 3,
                        "tight_depth": 2,
                        "expected_tight_depth_refusal": cfg!(feature = "final"),
                    }),
                ),
                Err(error) => error_report("split_tight", error, metrics),
            }
        }
        "merge_package" => {
            let parts = split_parts(source, &limits).expect("fixture split");
            let package = to_package(&parts, &limits).expect("package projection");
            let before = measure_begin();
            let result = merge_package(&parts, source, &package, &limits);
            let metrics = measure_end(before);
            match result {
                Ok(output) => output_report(
                    "merge_package",
                    &output,
                    source,
                    metrics,
                    serde_json::json!({
                        "package_bytes": package.len(),
                        "projection_source": true,
                    }),
                ),
                Err(error) => error_report("merge_package", error, metrics),
            }
        }
        "owner_roundtrip" => {
            let before = measure_begin();
            let result = litchi_ods::flat::Workbook::from_bytes(source)
                .map_err(|error| format!("{error:?}"))
                .and_then(|workbook| {
                    let sheets = workbook.sheets().len();
                    let output = workbook.to_bytes();
                    let text = workbook.text().map_err(|error| format!("{error:?}"))?;
                    Ok((output, text, sheets))
                });
            let metrics = measure_end(before);
            match result {
                Ok((output, text, sheets)) => output_report(
                    "owner_roundtrip",
                    output.as_bytes(),
                    source,
                    metrics,
                    serde_json::json!({
                        "sheet_count": sheets,
                        "text_bytes": text.len(),
                        "text_sha256": digest(text.as_bytes()),
                    }),
                ),
                Err(error) => error_report("owner_roundtrip", error, metrics),
            }
        }
        _ => serde_json::json!({
            "schema": "odf_flat_probe.v1",
            "status": "error",
            "error": "route must be split, to_package, merge_noop, merge_edit, merge_foreign, merge_tight, split_tight, merge_package, or owner_roundtrip",
        }),
    };
    println!("{}", serde_json::to_string(&report).expect("report JSON"));
}
