//! Quick open-route loop for record 0012 (not a registered benchmark).
//!
//! Usage: uo0012-open-bench <route> <path> <warmup> <samples>
//! Routes:
//!   opc-open        litchi_opc::SourceBackedPackage::from_path
//!   docx-open       litchi_docx::source_backed::Package::open
//!   docx-open-text  ... then document()?.extract_text()
//!   pptx-open       litchi_pptx::SourceBackedPresentation::open
//!   pptx-open-text  ... then text() of every slide
//!   opc-clone       (after only) one open, then $CLONES clone/drop pairs
//!   opc-clone-ns    (after only) ns per clone/drop pair, batches of 1,000
//!   opc-clone-ns-managed  (after only) the same on a managed package
//!
//! Every sample opens the path afresh (warm page cache) and drops the owner
//! inside the timed region. Prints one JSON line with the sorted samples' p10,
//! p50, p90 and mean in nanoseconds, plus a checksum of the work.

use std::hint::black_box;
use std::time::Instant;

fn run(route: &str, path: &str) -> usize {
    match route {
        "opc-open" => {
            let package = litchi_opc::SourceBackedPackage::from_path(path).expect("opc open");
            let count = package.iter_parts().count();
            drop(black_box(package));
            count
        },
        "docx-open" => {
            let package = litchi_docx::source_backed::Package::open(path).expect("docx open");
            drop(black_box(package));
            1
        },
        "docx-open-text" => {
            let package = litchi_docx::source_backed::Package::open(path).expect("docx open");
            let text = package
                .document()
                .expect("docx document")
                .extract_text()
                .expect("docx text");
            let len = text.len();
            drop(black_box(text));
            drop(black_box(package));
            len
        },
        "pptx-open" => {
            let presentation =
                litchi_pptx::SourceBackedPresentation::open(path).expect("pptx open");
            let count = presentation.slides().len();
            drop(black_box(presentation));
            count
        },
        "pptx-open-text" => {
            let presentation =
                litchi_pptx::SourceBackedPresentation::open(path).expect("pptx open");
            let mut len = 0;
            for slide in presentation.slides() {
                len += slide.text().expect("pptx slide text").len();
            }
            drop(black_box(presentation));
            len
        },
        // After-only route: one open, then `CLONES` clone-and-drop pairs.
        // Run under an allocation counter with two values of `CLONES` to
        // isolate the cost of `clone()`.
        "opc-clone" => {
            let clones: usize = std::env::var("CLONES")
                .ok()
                .and_then(|value| value.parse().ok())
                .unwrap_or(0);
            let package = litchi_opc::SourceBackedPackage::from_path(path).expect("opc open");
            for _ in 0..clones {
                let clone = black_box(package.clone());
                drop(black_box(clone));
            }
            clones
        },
        other => panic!("unknown route {other}"),
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let route = &args[1];
    let path = &args[2];
    let warmup: usize = args[3].parse().expect("warmup");
    let samples: usize = args[4].parse().expect("samples");
    if route == "opc-clone-ns" || route == "opc-clone-ns-managed" {
        // After only: open once, then time `samples` batches of 1,000
        // clone-and-drop pairs; report nanoseconds per pair. The managed
        // variant opens under an execution context with a budget.
        let package = if route == "opc-clone-ns" {
            litchi_opc::SourceBackedPackage::from_path(path).expect("opc open")
        } else {
            use std::num::{NonZeroU64, NonZeroUsize};
            let budget = litchi_core::Budget::root(
                "uo0012-bench",
                litchi_core::Limits::new(1 << 32, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
            );
            let (_source, cancellation) = litchi_core::CancellationSource::pair();
            let limits = litchi_core::ExecutionLimits::new(
                NonZeroUsize::new(1).unwrap(),
                NonZeroUsize::new(1).unwrap(),
                NonZeroU64::new(1 << 30).unwrap(),
                0,
            )
            .unwrap();
            let context = litchi_core::ExecutionContext::new(budget, cancellation, limits);
            let bytes = std::fs::read(path).expect("read fixture");
            litchi_opc::SourceBackedPackage::from_read_at_with_execution_context(
                std::sync::Arc::new(litchi_core::OwnedSource::new(bytes)),
                litchi_opc::ReadLimits::default(),
                context,
            )
            .expect("managed open")
        };
        let mut per_pair = Vec::with_capacity(samples);
        for batch in 0..warmup + samples {
            let start = Instant::now();
            for _ in 0..1_000 {
                let clone = black_box(package.clone());
                drop(black_box(clone));
            }
            if batch >= warmup {
                per_pair.push(start.elapsed().as_nanos() as f64 / 1_000.0);
            }
        }
        per_pair.sort_by(f64::total_cmp);
        println!(
            "{{\"route\":\"{route}\",\"path\":\"{path}\",\"batches\":{samples},\"p10_ns\":{:.1},\"p50_ns\":{:.1},\"p90_ns\":{:.1}}}",
            per_pair[per_pair.len() / 10],
            per_pair[per_pair.len() / 2],
            per_pair[per_pair.len() * 9 / 10]
        );
        return;
    }
    let mut checksum = 0;
    for _ in 0..warmup {
        checksum = run(route, path);
    }
    let mut times = Vec::with_capacity(samples);
    for _ in 0..samples {
        let start = Instant::now();
        let value = run(route, path);
        let elapsed = start.elapsed().as_nanos() as u64;
        assert_eq!(value, checksum, "route output changed between samples");
        times.push(elapsed);
    }
    times.sort_unstable();
    let pick = |q: f64| times[((times.len() - 1) as f64 * q).round() as usize];
    let mean = times.iter().sum::<u64>() / times.len() as u64;
    println!(
        "{{\"route\":\"{route}\",\"path\":\"{path}\",\"samples\":{samples},\"checksum\":{checksum},\"p10_ns\":{},\"p50_ns\":{},\"p90_ns\":{},\"mean_ns\":{mean}}}",
        pick(0.10),
        pick(0.50),
        pick(0.90)
    );
}
