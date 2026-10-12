//! Paired native/unified PPT selected-shape and early-query measurements.
//!
//! Run with `cargo run --release -p litchi-ops --no-default-features --features ppt
//! --example ppt_read_bench`. Opening, fixture generation, path/query compilation,
//! and correctness validation are outside timing. Samples alternate AB/BA order.
//! This exploratory probe does not replace the frozen performance harness gates.

#[cfg(feature = "ppt")]
mod enabled {
    use litchi_core::FileFormat;
    use litchi_ops::{
        path::Path,
        query::{Plan, Query},
    };
    use serde::Serialize;
    use std::{hint::black_box, time::Instant};

    #[derive(Serialize)]
    struct Case {
        slides: usize,
        source: &'static str,
        repeats: usize,
        native_ns: Vec<u128>,
        unified_get_ns: Vec<u128>,
        unified_first_ns: Vec<u128>,
    }

    fn measure(repeats: usize, mut action: impl FnMut() -> usize) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            black_box(action());
        }
        start.elapsed().as_nanos()
    }

    pub(super) fn run() {
        let samples = std::env::var("LITCHI_PPT_BENCH_SAMPLES")
            .ok()
            .and_then(|v| v.parse::<usize>().ok())
            .filter(|&v| v > 0)
            .unwrap_or(7);
        let repeats = 100;
        let mut cases = Vec::new();
        for count in [1, 100, 1_000] {
            let mut builder = litchi_ppt::Builder::new();
            for _ in 0..count {
                let slide = builder.add_slide().expect("slide");
                builder
                    .add_textbox(slide, 10, 10, 240, 40, "selected text")
                    .expect("shape");
            }
            let bytes = builder.build().expect("package").to_bytes().expect("bytes");
            let temp = tempfile::NamedTempFile::new().expect("file");
            std::fs::write(temp.path(), &bytes).expect("fixture");
            for positional in [false, true] {
                let package = if positional {
                    litchi_ppt::Package::open(temp.path()).expect("native source")
                } else {
                    litchi_ppt::Package::from_bytes(&bytes).expect("native bytes")
                };
                let presentation = package.presentation().expect("presentation");
                let file = if positional {
                    litchi_ops::open(temp.path()).expect("unified source")
                } else {
                    litchi_ops::open_bytes(bytes.clone()).expect("unified bytes")
                };
                let path = Path::parse("/slide[1]/shape[1]").expect("path");
                let query = Query::parse("shape").expect("query");
                let plan = Plan::compile(&query, FileFormat::Ppt).expect("plan");
                let native = || {
                    let slide = presentation
                        .slide_at(0)
                        .expect("slide")
                        .expect("first slide");
                    slide.shapes().expect("shapes")[0].text().expect("text")
                };
                let unified = || {
                    file.get(&path)
                        .expect("get")
                        .view()
                        .text()
                        .expect("text")
                        .expect("some text")
                        .into_owned()
                };
                let first = || {
                    file.query(&plan)
                        .expect("query")
                        .first()
                        .expect("first")
                        .expect("match")
                        .view()
                        .text()
                        .expect("text")
                        .expect("some text")
                        .into_owned()
                };
                assert_eq!(native(), "selected text");
                assert_eq!(native(), unified());
                assert_eq!(native(), first());
                let mut case = Case {
                    slides: count,
                    source: if positional { "positional" } else { "memory" },
                    repeats,
                    native_ns: Vec::new(),
                    unified_get_ns: Vec::new(),
                    unified_first_ns: Vec::new(),
                };
                for sample in 0..samples {
                    let mut n = || case.native_ns.push(measure(repeats, || native().len()));
                    let mut u = || {
                        case.unified_get_ns
                            .push(measure(repeats, || unified().len()))
                    };
                    let mut q = || {
                        case.unified_first_ns
                            .push(measure(repeats, || first().len()))
                    };
                    if sample % 2 == 0 {
                        n();
                        u();
                        q();
                    } else {
                        q();
                        u();
                        n();
                    }
                }
                cases.push(case);
            }
        }
        serde_json::to_writer_pretty(std::io::stdout().lock(), &cases).expect("JSON");
        println!();
    }
}

fn main() {
    #[cfg(feature = "ppt")]
    enabled::run();
    #[cfg(not(feature = "ppt"))]
    eprintln!("enable the ppt feature");
}
