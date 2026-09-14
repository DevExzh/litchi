//! For a caller that must dispatch across the three facade types, compare:
//!   A) detect_file_format(path) then open the right facade
//!   B) try each facade opener in turn until one succeeds
use std::hint::black_box;
use std::time::Instant;

use litchi::common::detection::FileFormat;

const ROOT: &str = "/home/zhuhe/code/litchi/";

fn stats(mut v: Vec<u128>) -> (f64, f64, f64) {
    v.sort_unstable();
    let n = v.len();
    let pick = |q: f64| v[((n as f64 - 1.0) * q).round() as usize] as f64 / 1000.0;
    (pick(0.5), pick(0.25), pick(0.75))
}

/// Pattern A: one detect, then exactly one open.
fn dispatch_detect(p: &str) {
    match litchi::detect_file_format(p) {
        Some(FileFormat::Doc | FileFormat::Docx | FileFormat::Odt | FileFormat::Rtf) => {
            black_box(litchi::Document::open(p).unwrap());
        },
        Some(FileFormat::Ppt | FileFormat::Pptx | FileFormat::Odp) => {
            black_box(litchi::Presentation::open(p).unwrap());
        },
        Some(FileFormat::Xls | FileFormat::Xlsx | FileFormat::Xlsb | FileFormat::Ods) => {
            black_box(litchi::sheet::Workbook::open(p).unwrap());
        },
        other => panic!("unexpected {other:?}"),
    }
}

/// Pattern B: no pre-detect; try each facade in declaration order.
fn dispatch_cascade(p: &str) {
    if let Ok(d) = litchi::Document::open(p) {
        black_box(d);
        return;
    }
    if let Ok(d) = litchi::Presentation::open(p) {
        black_box(d);
        return;
    }
    if let Ok(d) = litchi::sheet::Workbook::open(p) {
        black_box(d);
        return;
    }
    panic!("no facade accepted {p}");
}

fn run(label: &str, path: &str, iters: usize) {
    for _ in 0..200 {
        dispatch_detect(path);
        dispatch_cascade(path);
    }
    let mut ta = Vec::with_capacity(iters);
    let mut tb = Vec::with_capacity(iters);
    for _ in 0..iters {
        let t = Instant::now();
        dispatch_detect(path);
        ta.push(t.elapsed().as_nanos());
        let t = Instant::now();
        dispatch_cascade(path);
        tb.push(t.elapsed().as_nanos());
    }
    let (am, a25, a75) = stats(ta);
    let (bm, b25, b75) = stats(tb);
    println!(
        "{label:<12} detect-then-open {am:8.2} us (p25 {a25:7.2} p75 {a75:7.2})   \
         try-each-facade {bm:8.2} us (p25 {b25:7.2} p75 {b75:7.2})   \
         cascade is {:+7.2} us ({:+6.1}%)",
        bm - am,
        (bm - am) / am * 100.0
    );
}

fn main() {
    let iters: usize = std::env::args()
        .nth(1)
        .and_then(|s| s.parse().ok())
        .unwrap_or(2000);
    let cases = [
        ("DOCX", "test-data/ooxml/docx/Hyperlink.docx"),
        ("PPTX", "test-data/ooxml/pptx/shapes.pptx"),
        ("XLSX", "test-data/ooxml/xlsx/sheet-names.xlsx"),
    ];
    for (l, r) in cases {
        let p = format!("{ROOT}{r}");
        let _ = black_box(std::fs::read(&p).unwrap());
    }
    println!("iters={iters}\n");
    for (l, r) in cases {
        run(l, &format!("{ROOT}{r}"), iters);
    }
}
