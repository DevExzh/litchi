//! One-off cost of the registry export (record 0263).
//!
//! Run with `cargo run --release -p litchi-ops --all-features --example
//! schema_export_bench`, ideally pinned (`taskset -c 2`). For each format it
//! times `litchi_ops::schema(format)` alone, a walk over every element and
//! property through the typed accessors, and `serde_json::to_string` of the
//! export, and prints the median of 31 samples per leg as one JSON line.
//! The export is not a hot path; this records its order of magnitude.

use std::hint::black_box;
use std::time::Instant;

use litchi_core::FileFormat;

const SAMPLES: usize = 31;

fn median(repeats: u32, mut operation: impl FnMut() -> usize) -> f64 {
    let mut samples: Vec<f64> = (0..SAMPLES)
        .map(|_| {
            let start = Instant::now();
            for _ in 0..repeats {
                black_box(operation());
            }
            start.elapsed().as_nanos() as f64 / f64::from(repeats)
        })
        .collect();
    samples.sort_by(f64::total_cmp);
    samples[SAMPLES / 2]
}

fn main() {
    for (name, format) in [
        ("doc", FileFormat::Doc),
        ("docx", FileFormat::Docx),
        ("ppt", FileFormat::Ppt),
        ("pptx", FileFormat::Pptx),
        ("rtf", FileFormat::Rtf),
        ("xls", FileFormat::Xls),
        ("xlsb", FileFormat::Xlsb),
        ("xlsx", FileFormat::Xlsx),
    ] {
        let build = median(10_000, || {
            usize::from(black_box(litchi_ops::schema(black_box(format))).registered())
        });
        let walk = median(100, || {
            let schema = litchi_ops::schema(black_box(format));
            let mut count = 0;
            for element in schema.elements() {
                for property in element.properties() {
                    count += property.aliases().count() + usize::from(property.clearable());
                }
            }
            count
        });
        let mut bytes = 0;
        let json = median(20, || {
            let text = serde_json::to_string(&litchi_ops::schema(black_box(format))).unwrap();
            bytes = text.len();
            bytes
        });
        let properties: usize = litchi_ops::schema(format)
            .elements()
            .map(|element| element.properties().count())
            .sum();
        println!(
            "{{\"format\":\"{name}\",\"properties\":{properties},\"jsonBytes\":{bytes},\
             \"buildNs\":{build:.1},\"walkUs\":{:.2},\"jsonUs\":{:.1}}}",
            walk / 1000.0,
            json / 1000.0
        );
    }
}
