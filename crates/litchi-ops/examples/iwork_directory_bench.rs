//! Paired leaf-vs-unified measurements of the iWork open routes record 0306
//! added or changed.
//!
//! For each format, three scenarios run as leaf/unified pairs in one
//! process:
//!
//! - `dir_open`: open the app-authored package directory fixture (leaf
//!   `Package::open(dir)` vs `litchi_ops::open(dir)`): the bundle capture and
//!   conversion, plus the unified layer's directory classification;
//! - `file_open`: open the single-file fixture from its path (leaf
//!   `Package::open(file)` vs `litchi_ops::open(file)`);
//! - `bytes_open`: open in-memory bytes (leaf
//!   `Package::from_shared_bytes_with_options` with the defaults vs
//!   `litchi_ops::open_bytes`), the reference for the next scenario;
//! - `bytes_open_options`: open in-memory bytes with the format's leaf
//!   options set to their defaults (leaf `Package::from_shared_bytes_with_options`
//!   vs `litchi_ops::open_bytes_with` carrying `{pages,keynote,numbers}_options`),
//!   the plumbing that must keep the single detection pass.
//!
//! Each pair runs in ABBA order for `--rounds` rounds of `--iterations`
//! timed calls per leg; every leg's per-call median is one JSON line, and a
//! summary line per scenario gives the median unified/leaf ratio over the
//! rounds (the shape of `iwork_bench`). Pin the process with `taskset`.
//!
//! ```text
//! cargo build --release -p litchi-ops --features pages,keynote,numbers \
//!     --example iwork_directory_bench
//! taskset -c 13 target/release/examples/iwork_directory_bench --rounds 8 --iterations 30
//! ```

use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::{Duration, Instant};

const FIXTURES: &[(&str, &str, &str)] = &[
    (
        "pages",
        "iwork/pages/basic.pages",
        "iwork/directory/pages/basic.pages",
    ),
    (
        "keynote",
        "iwork/keynote/basic.key",
        "iwork/directory/keynote/basic.key",
    ),
    (
        "numbers",
        "iwork/numbers/basic.numbers",
        "iwork/directory/numbers/basic.numbers",
    ),
];

fn test_data() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
}

fn median(mut samples: Vec<Duration>) -> Duration {
    samples.sort_unstable();
    samples[samples.len() / 2]
}

/// Times `iterations` calls of `call`, each on a fresh input from `input`
/// prepared outside the timed region, and returns the per-call median.
fn leg<I>(iterations: usize, mut input: impl FnMut() -> I, mut call: impl FnMut(I)) -> Duration {
    let mut samples = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let value = input();
        let start = Instant::now();
        call(value);
        samples.push(start.elapsed());
    }
    median(samples)
}

fn leaf_path_open(kind: &str, path: &Path) {
    match kind {
        "pages" => {
            black_box(litchi_pages::Package::open(path).unwrap());
        },
        "keynote" => {
            black_box(litchi_keynote::Package::open(path).unwrap());
        },
        _ => {
            black_box(litchi_numbers::Package::open(path).unwrap());
        },
    }
}

fn leaf_bytes_open(kind: &str, bytes: Arc<[u8]>) {
    match kind {
        "pages" => {
            black_box(
                litchi_pages::Package::from_shared_bytes_with_options(
                    bytes,
                    litchi_pages::OpenOptions::new(),
                )
                .unwrap(),
            );
        },
        "keynote" => {
            black_box(
                litchi_keynote::Package::__from_shared_source_with_options(
                    bytes,
                    litchi_keynote::OpenOptions::new(),
                )
                .unwrap(),
            );
        },
        _ => {
            black_box(
                litchi_numbers::Package::from_shared_bytes_with_options(
                    bytes,
                    litchi_numbers::OpenOptions::new(),
                )
                .unwrap(),
            );
        },
    }
}

fn unified_options(kind: &str) -> litchi_ops::OpenOptions {
    let options = litchi_ops::OpenOptions::new();
    match kind {
        "pages" => options.pages_options(litchi_pages::OpenOptions::new()),
        "keynote" => options.keynote_options(litchi_keynote::OpenOptions::new()),
        _ => options.numbers_options(litchi_numbers::OpenOptions::new()),
    }
}

fn argument(name: &str, default: usize) -> usize {
    let arguments: Vec<String> = std::env::args().collect();
    arguments
        .iter()
        .position(|argument| argument == name)
        .and_then(|index| arguments.get(index + 1))
        .and_then(|value| value.parse().ok())
        .unwrap_or(default)
}

fn main() {
    let rounds = argument("--rounds", 8);
    let iterations = argument("--iterations", 30);
    for (kind, single, bundle) in FIXTURES {
        let single = test_data().join(single);
        let bundle = test_data().join(bundle);
        let bytes = std::fs::read(&single).unwrap();
        let options = unified_options(kind);
        for scenario in ["dir_open", "file_open", "bytes_open", "bytes_open_options"] {
            let mut ratios = Vec::with_capacity(rounds);
            let mut leaf_medians = Vec::new();
            let mut unified_medians = Vec::new();
            for round in 0..rounds {
                // ABBA: leaf, unified, unified, leaf.
                for (position, leg_name) in
                    ["leaf", "unified", "unified", "leaf"].iter().enumerate()
                {
                    let median = match (scenario, *leg_name) {
                        ("dir_open", "leaf") => {
                            leg(iterations, || (), |()| leaf_path_open(kind, &bundle))
                        },
                        ("dir_open", _) => leg(
                            iterations,
                            || (),
                            |()| {
                                black_box(litchi_ops::open(&bundle).unwrap());
                            },
                        ),
                        ("file_open", "leaf") => {
                            leg(iterations, || (), |()| leaf_path_open(kind, &single))
                        },
                        ("file_open", _) => leg(
                            iterations,
                            || (),
                            |()| {
                                black_box(litchi_ops::open(&single).unwrap());
                            },
                        ),
                        ("bytes_open", "leaf") | ("bytes_open_options", "leaf") => leg(
                            iterations,
                            || Arc::<[u8]>::from(bytes.as_slice()),
                            |shared| leaf_bytes_open(kind, shared),
                        ),
                        ("bytes_open", _) => leg(
                            iterations,
                            || bytes.clone(),
                            |owned| {
                                black_box(litchi_ops::open_bytes(owned).unwrap());
                            },
                        ),
                        _ => leg(
                            iterations,
                            || bytes.clone(),
                            |owned| {
                                black_box(litchi_ops::open_bytes_with(owned, &options).unwrap());
                            },
                        ),
                    };
                    println!(
                        "{{\"format\":\"{kind}\",\"scenario\":\"{scenario}\",\"round\":{round},\"position\":{position},\"leg\":\"{leg_name}\",\"median_ns\":{}}}",
                        median.as_nanos()
                    );
                    if *leg_name == "leaf" {
                        leaf_medians.push(median);
                    } else {
                        unified_medians.push(median);
                    }
                }
                let leaf = leaf_medians[leaf_medians.len() - 2..]
                    .iter()
                    .sum::<Duration>()
                    .as_secs_f64();
                let unified = unified_medians[unified_medians.len() - 2..]
                    .iter()
                    .sum::<Duration>()
                    .as_secs_f64();
                ratios.push(unified / leaf);
            }
            ratios.sort_by(f64::total_cmp);
            let leaf = median(leaf_medians);
            let unified = median(unified_medians);
            println!(
                "{{\"summary\":true,\"format\":\"{kind}\",\"scenario\":\"{scenario}\",\"leaf_median_ns\":{},\"unified_median_ns\":{},\"ratio_median\":{:.4},\"ratio_min\":{:.4},\"ratio_max\":{:.4},\"delta_ns\":{}}}",
                leaf.as_nanos(),
                unified.as_nanos(),
                ratios[ratios.len() / 2],
                ratios[0],
                ratios[ratios.len() - 1],
                unified.as_nanos() as i128 - leaf.as_nanos() as i128,
            );
        }
    }
}
