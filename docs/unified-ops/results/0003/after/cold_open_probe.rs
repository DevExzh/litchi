//! M0a S10 follow-up probe: the first open in a fresh process.
//!
//! The harness's `*_file_*` routes time one open per fresh child process, so
//! one-time costs land inside the timer. This probe times, in a fresh
//! process, the first and then a second open of one file, through either the
//! revision's facade or the leaf calls the S8b leaf twin makes, with the
//! process's minor-fault count around each open.
//!
//! - feature `before`: facade = `litchi::Workbook::open` (xlsx) /
//!   `litchi::Document::open` (docx) / `litchi::Presentation::open` (pptx);
//! - feature `after`: facade = `litchi::open`;
//! - leaf (both): xlsx = `FileSource::open` + `SourceBackedPackage::
//!   from_read_at_with_limits` + `read_source_backed_with_report` +
//!   `SourceBackedWorkbook::from_source_backed_package`, each step timed;
//!   docx = `litchi_docx::source_backed::Package::from_path`;
//!   pptx = `litchi_pptx::SourceBackedPresentation::from_path`.
//!
//! - mode `slide-count` (pptx): open through the facade outside the timer,
//!   then time three `slide_count` calls (the first one cold).
//!
//! usage: cold-open-probe <xlsx|docx|pptx> <facade|leaf|slide-count> <path>
//! Output: one JSON line.

use std::path::Path;
use std::sync::Arc;
use std::time::Instant;

fn minor_faults() -> u64 {
    let stat = std::fs::read_to_string("/proc/self/stat").expect("procfs");
    let after = &stat[stat.rfind(')').expect("comm") + 2..];
    after.split(' ').nth(7).expect("minflt").parse().expect("minflt")
}

#[cfg(feature = "after")]
fn facade(format: &str, path: &Path) {
    let _ = format;
    std::hint::black_box(litchi::open(path).expect("facade open"));
}

#[cfg(feature = "before")]
fn facade(format: &str, path: &Path) {
    match format {
        "xlsx" => {
            std::hint::black_box(litchi::Workbook::open(path).expect("facade open"));
        },
        "docx" => {
            std::hint::black_box(litchi::Document::open(path).expect("facade open"));
        },
        "pptx" => {
            std::hint::black_box(litchi::Presentation::open(path).expect("facade open"));
        },
        other => panic!("unknown format {other}"),
    }
}

/// Opens a PPTX through the revision's facade (untimed) and returns a closure
/// that asks it for the slide count, the query `pptx_file_*_slide_count`
/// times.
#[cfg(feature = "after")]
fn slide_count_query(path: &Path) -> impl Fn() -> usize {
    let file = litchi::open(path).expect("facade open");
    move || file.presentation().expect("view").slide_count().expect("count")
}

#[cfg(feature = "before")]
fn slide_count_query(path: &Path) -> impl Fn() -> usize {
    let presentation = litchi::Presentation::open(path).expect("facade open");
    move || presentation.slide_count().expect("count")
}

/// The leaf calls the harness's prepared query routes time, at both
/// revisions: an eager DOCX package (`OpcPackage::from_bytes_with_limits` +
/// `litchi_docx::Package::from_opc_package`) or a source-backed PPTX
/// presentation, opened outside the timer; the returned closure is the
/// timed query.
fn leaf_query(mode: &str, path: &Path) -> Box<dyn Fn()> {
    match mode {
        "docx-eager-paragraph-count" | "docx-eager-text" | "docx-eager-separated-text" => {
            let bytes = std::fs::read(path).expect("read");
            let opc = litchi_opc::OpcPackage::from_bytes_with_limits(
                &bytes,
                litchi_opc::ReadLimits::default(),
            )
            .expect("opc");
            let package = litchi_docx::Package::from_opc_package(opc).expect("docx package");
            let mode = mode.to_owned();
            Box::new(move || {
                let document = package.document().expect("document");
                match mode.as_str() {
                    "docx-eager-paragraph-count" => {
                        std::hint::black_box(document.paragraph_count().expect("count"));
                    },
                    "docx-eager-text" => {
                        std::hint::black_box(document.text().expect("text"));
                    },
                    _ => separated_text(&document),
                }
            })
        },
        "pptx-source-selected-slide" => {
            let presentation =
                litchi_pptx::SourceBackedPresentation::from_path(path).expect("pptx open");
            Box::new(move || {
                let slide = presentation.slide(100).expect("slide 100");
                std::hint::black_box(slide.text_and_name().expect("text and name"));
            })
        },
        other => panic!("unknown query mode {other}"),
    }
}

/// The separated text `File::text()` builds since S8a; the pre-M0a leaf has
/// no separator mode, so the before build times plain `text()` here.
#[cfg(feature = "after")]
fn separated_text(document: &litchi_docx::Document<'_>) {
    std::hint::black_box(document.text_with_paragraph_separator("\n").expect("text"));
}

#[cfg(feature = "before")]
fn separated_text(document: &litchi_docx::Document<'_>) {
    std::hint::black_box(document.text().expect("text"));
}

/// Returns the per-step nanoseconds of the leaf open.
fn leaf(format: &str, path: &Path) -> Vec<(&'static str, u64)> {
    let mut steps = Vec::new();
    let mut mark = Instant::now();
    let mut step = |name: &'static str, steps: &mut Vec<(&'static str, u64)>| {
        let now = Instant::now();
        steps.push((name, u64::try_from((now - mark).as_nanos()).expect("ns")));
        mark = now;
    };
    match format {
        "xlsx" => {
            let source: Arc<dyn litchi_core::ReadAt> =
                Arc::new(litchi_core::FileSource::open(path).expect("file source"));
            step("file_source", &mut steps);
            let package = litchi_opc::SourceBackedPackage::from_read_at_with_limits(
                source,
                litchi_opc::ReadLimits::default(),
            )
            .expect("opc index");
            step("opc_index", &mut steps);
            let mut repairs = litchi_ooxml_common::leniency::ToleranceReport::default();
            let metadata =
                litchi_ooxml_common::properties::read_source_backed_with_report(&package, &mut repairs)
                    .expect("core properties");
            std::hint::black_box(&metadata);
            step("core_properties", &mut steps);
            let workbook = litchi_xlsx::SourceBackedWorkbook::from_source_backed_package(package)
                .expect("workbook");
            step("workbook_catalog", &mut steps);
            std::hint::black_box(workbook);
        },
        "docx" => {
            std::hint::black_box(
                litchi_docx::source_backed::Package::from_path(path).expect("docx source open"),
            );
            step("open", &mut steps);
        },
        "pptx" => {
            std::hint::black_box(
                litchi_pptx::SourceBackedPresentation::from_path(path).expect("pptx source open"),
            );
            step("open", &mut steps);
        },
        other => panic!("unknown format {other}"),
    }
    steps
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let format = args[1].as_str();
    let mode = args[2].as_str();
    let path = Path::new(&args[3]);
    // S10_ROUNDS (default 3) sets how many timed rounds run; callgrind runs
    // use 0, 1 and 2 so that I(1) - I(0) is the first round's instructions.
    let count: usize = std::env::var("S10_ROUNDS").ok().and_then(|v| v.parse().ok()).unwrap_or(3);
    let mut rounds = Vec::new();
    if mode == "slide-count" || mode.starts_with("docx-eager-") || mode == "pptx-source-selected-slide" {
        // The open stays outside the timer, as in the harness's prepared
        // query routes; then the timed calls, the first one cold.
        let query: Box<dyn Fn()> = if mode == "slide-count" {
            let count_query = slide_count_query(path);
            Box::new(move || {
                std::hint::black_box(count_query());
            })
        } else {
            leaf_query(mode, path)
        };
        for _ in 0..count {
            let faults = minor_faults();
            let started = Instant::now();
            query();
            let elapsed = u64::try_from(started.elapsed().as_nanos()).expect("ns");
            let faults = minor_faults() - faults;
            rounds.push(format!("{{\"elapsed_ns\":{elapsed},\"minor_faults\":{faults},\"steps\":{{}}}}"));
        }
        println!(
            "{{\"revision\":\"{}\",\"format\":\"{format}\",\"mode\":\"{mode}\",\"rounds\":[{}]}}",
            if cfg!(feature = "after") { "after" } else { "before" },
            rounds.join(",")
        );
        return;
    }
    for _ in 0..count {
        let faults = minor_faults();
        let started = Instant::now();
        let steps = if mode == "facade" {
            facade(format, path);
            Vec::new()
        } else {
            leaf(format, path)
        };
        let elapsed = u64::try_from(started.elapsed().as_nanos()).expect("ns");
        let faults = minor_faults() - faults;
        let steps = steps
            .iter()
            .map(|(name, ns)| format!("\"{name}\":{ns}"))
            .collect::<Vec<_>>()
            .join(",");
        rounds.push(format!("{{\"elapsed_ns\":{elapsed},\"minor_faults\":{faults},\"steps\":{{{steps}}}}}"));
    }
    println!(
        "{{\"revision\":\"{}\",\"format\":\"{format}\",\"mode\":\"{mode}\",\"rounds\":[{}]}}",
        if cfg!(feature = "after") { "after" } else { "before" },
        rounds.join(",")
    );
}
