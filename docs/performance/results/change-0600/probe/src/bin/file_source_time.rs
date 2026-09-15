//! Paired timing for the scenarios no perf-baseline selector covers: a cold
//! Part read on a *file* source, and the cold reads that follow a `stream_to`
//! on the same package. `version()` costs a syscall only on `FileSource`, so
//! the in-memory selectors cannot show this path at all.
//!
//! `file_source_time <fixture> <partname> <scenario> <warmup> <samples>`
//! prints `timing_samples_ns=` in the shape change 0594's probe used.
use std::hint::black_box;
use std::sync::Arc;
use std::time::Instant;

use litchi_core::{FileSource, ReadAt};
use litchi_opc::{PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

fn open(path: &str) -> SourceBackedPackage {
    let source: Arc<dyn ReadAt> = Arc::new(FileSource::open(path).expect("file source"));
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        source,
        ReadLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
    )
    .expect("open")
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let path = args[0].clone();
    let part = PackURI::new(args[1].as_str()).expect("partname");
    let scenario = args[2].clone();
    let warmup: usize = args[3].parse().expect("warmup");
    let samples: usize = args[4].parse().expect("samples");

    // Every scenario opens its own package so each read is genuinely cold.
    let once = |scenario: &str| -> u64 {
        match scenario {
            // Open only.
            "open" => {
                let start = Instant::now();
                let package = open(&path);
                let elapsed = start.elapsed().as_nanos() as u64;
                black_box(&package);
                elapsed
            },
            // Open, then one cold Part read.
            "cold_read" => {
                let package = open(&path);
                let start = Instant::now();
                let data = package.part(&part).expect("part").data().expect("data");
                let elapsed = start.elapsed().as_nanos() as u64;
                black_box(data.as_bytes().len());
                elapsed
            },
            // The sticky-monitor scenario: stream one Part, then cold-read
            // every other Part on the same package. Only the reads are timed.
            "reads_after_stream" => {
                let package = open(&path);
                let mut sink = Vec::new();
                package
                    .part(&part)
                    .expect("part")
                    .stream_to(&mut sink)
                    .expect("stream");
                let names: Vec<String> = package
                    .iter_parts()
                    .map(|view| view.partname().to_string())
                    .collect();
                let start = Instant::now();
                let mut total = 0usize;
                for name in &names {
                    let uri = PackURI::new(name.as_str()).expect("uri");
                    total += package
                        .part(&uri)
                        .expect("part")
                        .data()
                        .expect("data")
                        .as_bytes()
                        .len();
                }
                let elapsed = start.elapsed().as_nanos() as u64;
                black_box(total);
                elapsed
            },
            // The same reads on a package that never streamed (control).
            "reads_no_stream" => {
                let package = open(&path);
                let names: Vec<String> = package
                    .iter_parts()
                    .map(|view| view.partname().to_string())
                    .collect();
                let start = Instant::now();
                let mut total = 0usize;
                for name in &names {
                    let uri = PackURI::new(name.as_str()).expect("uri");
                    total += package
                        .part(&uri)
                        .expect("part")
                        .data()
                        .expect("data")
                        .as_bytes()
                        .len();
                }
                let elapsed = start.elapsed().as_nanos() as u64;
                black_box(total);
                elapsed
            },
            other => panic!("unknown scenario {other}"),
        }
    };

    for _ in 0..warmup {
        black_box(once(&scenario));
    }
    let mut measured = Vec::with_capacity(samples);
    for _ in 0..samples {
        measured.push(once(&scenario));
    }
    println!("scenario={scenario} fixture={path}");
    println!(
        "timing_samples_ns={}",
        measured
            .iter()
            .map(u64::to_string)
            .collect::<Vec<_>>()
            .join(",")
    );
}
