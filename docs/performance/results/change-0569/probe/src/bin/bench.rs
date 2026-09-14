//! Paired A/B microbenchmark: two-call (detect then open) vs single-call open.
//! A and B alternate within one loop so host drift affects both equally.
use std::hint::black_box;
use std::time::Instant;

const ROOT: &str = "/home/zhuhe/code/litchi/";

fn stats(mut v: Vec<u128>) -> (f64, f64, f64, f64) {
    v.sort_unstable();
    let n = v.len();
    let pick = |q: f64| v[((n as f64 - 1.0) * q).round() as usize] as f64 / 1000.0;
    (pick(0.5), pick(0.25), pick(0.75), v[0] as f64 / 1000.0)
}

fn bench(label: &str, path: &str, iters: usize, warm: usize, a: &dyn Fn(&str), b: &dyn Fn(&str)) {
    for _ in 0..warm {
        a(path);
        b(path);
    }
    let mut ta = Vec::with_capacity(iters);
    let mut tb = Vec::with_capacity(iters);
    for _ in 0..iters {
        let t = Instant::now();
        a(path);
        ta.push(t.elapsed().as_nanos());
        let t = Instant::now();
        b(path);
        tb.push(t.elapsed().as_nanos());
    }
    let (am, a25, a75, amin) = stats(ta);
    let (bm, b25, b75, bmin) = stats(tb);
    println!(
        "{label:<34} two-call median {am:8.2} us (p25 {a25:7.2} p75 {a75:7.2} min {amin:7.2})  \
         single-call median {bm:8.2} us (p25 {b25:7.2} p75 {b75:7.2} min {bmin:7.2})  \
         delta {:7.2} us  ({:5.1}% of two-call)",
        am - bm,
        (am - bm) / am * 100.0
    );
}

fn bench_one(label: &str, path: &str, iters: usize, warm: usize, f: &dyn Fn(&str)) {
    for _ in 0..warm {
        f(path);
    }
    let mut t = Vec::with_capacity(iters);
    for _ in 0..iters {
        let s = Instant::now();
        f(path);
        t.push(s.elapsed().as_nanos());
    }
    let (m, p25, p75, min) = stats(t);
    println!("{label:<34} median {m:8.2} us (p25 {p25:7.2} p75 {p75:7.2} min {min:7.2})");
}

fn main() {
    let iters: usize = std::env::args()
        .nth(1)
        .and_then(|s| s.parse().ok())
        .unwrap_or(2000);
    let warm = 300usize;
    let docx = format!("{ROOT}test-data/ooxml/docx/Hyperlink.docx");
    let pptx = format!("{ROOT}test-data/ooxml/pptx/shapes.pptx");
    let xlsx = format!("{ROOT}test-data/ooxml/xlsx/sheet-names.xlsx");

    // Warm the page cache.
    for p in [&docx, &pptx, &xlsx] {
        let _ = black_box(std::fs::read(p).unwrap());
    }

    println!("iters={iters} warm={warm}\n");
    println!("-- isolated cost of the extra call --");
    bench_one("detect_file_format DOCX", &docx, iters, warm, &|p| {
        black_box(litchi::detect_file_format(black_box(p)));
    });
    bench_one("detect_file_format PPTX", &pptx, iters, warm, &|p| {
        black_box(litchi::detect_file_format(black_box(p)));
    });
    bench_one("detect_file_format XLSX", &xlsx, iters, warm, &|p| {
        black_box(litchi::detect_file_format(black_box(p)));
    });
    println!("\n-- isolated cost of the single-call open --");
    bench_one("Document::open DOCX", &docx, iters, warm, &|p| {
        black_box(litchi::Document::open(black_box(p)).unwrap());
    });
    bench_one("Presentation::open PPTX", &pptx, iters, warm, &|p| {
        black_box(litchi::Presentation::open(black_box(p)).unwrap());
    });
    bench_one("sheet::Workbook::open XLSX", &xlsx, iters, warm, &|p| {
        black_box(litchi::sheet::Workbook::open(black_box(p)).unwrap());
    });
    bench_one("sheet::open_workbook XLSX", &xlsx, iters, warm, &|p| {
        black_box(litchi::sheet::open_workbook(black_box(p)).unwrap());
    });

    println!("\n-- paired A/B: two-call vs single-call --");
    bench(
        "DOCX Document::open",
        &docx,
        iters,
        warm,
        &|p| {
            let f = black_box(litchi::detect_file_format(black_box(p)));
            assert_eq!(f, Some(litchi::common::detection::FileFormat::Docx));
            black_box(litchi::Document::open(black_box(p)).unwrap());
        },
        &|p| {
            black_box(litchi::Document::open(black_box(p)).unwrap());
        },
    );
    bench(
        "PPTX Presentation::open",
        &pptx,
        iters,
        warm,
        &|p| {
            let f = black_box(litchi::detect_file_format(black_box(p)));
            assert_eq!(f, Some(litchi::common::detection::FileFormat::Pptx));
            black_box(litchi::Presentation::open(black_box(p)).unwrap());
        },
        &|p| {
            black_box(litchi::Presentation::open(black_box(p)).unwrap());
        },
    );
    bench(
        "XLSX sheet::Workbook::open",
        &xlsx,
        iters,
        warm,
        &|p| {
            let f = black_box(litchi::detect_file_format(black_box(p)));
            assert_eq!(f, Some(litchi::common::detection::FileFormat::Xlsx));
            black_box(litchi::sheet::Workbook::open(black_box(p)).unwrap());
        },
        &|p| {
            black_box(litchi::sheet::Workbook::open(black_box(p)).unwrap());
        },
    );
    bench(
        "XLSX sheet::open_workbook",
        &xlsx,
        iters,
        warm,
        &|p| {
            let f = black_box(litchi::detect_file_format(black_box(p)));
            assert_eq!(f, Some(litchi::common::detection::FileFormat::Xlsx));
            black_box(litchi::sheet::open_workbook(black_box(p)).unwrap());
        },
        &|p| {
            black_box(litchi::sheet::open_workbook(black_box(p)).unwrap());
        },
    );
}
