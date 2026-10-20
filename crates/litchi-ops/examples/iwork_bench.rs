//! Paired leaf-vs-unified measurements of the iWork formats (record 0296).
//!
//! For each fixture, three scenarios run as leaf/unified pairs over the same
//! bytes in one process:
//!
//! - `open_get`: open an in-memory package and read one element's text
//!   (Pages: the first body paragraph; Keynote: the first slide's title;
//!   Numbers: cell B2 of the first table);
//! - `set`: change that element's text (or a number cell) on an open
//!   package and commit;
//! - `open`: open only.
//!
//! Each pair runs in ABBA order for `--rounds` rounds of `--iterations`
//! timed calls per leg; the per-call median of every leg run is printed as
//! one JSON line, and a summary line gives the median ratio
//! unified/leaf over the rounds. Pin the process with `taskset`.
//!
//! ```text
//! cargo run --release -p litchi-ops --features pages,keynote,numbers \
//!     --example iwork_bench -- [--rounds 8] [--iterations 40]
//! ```

use std::hint::black_box;
use std::time::{Duration, Instant};

use litchi_ops::props;

const FIXTURES: &[(&str, &str)] = &[
    ("pages", "iwork/pages/body-sections-unicode.pages"),
    ("keynote", "iwork/keynote/basic.key"),
    ("numbers", "iwork/numbers/basic.numbers"),
];

fn test_data() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
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

fn leaf_open(kind: &str, bytes: &[u8]) {
    match kind {
        "pages" => {
            black_box(litchi_pages::Package::from_bytes(bytes).unwrap());
        },
        "keynote" => {
            black_box(litchi_keynote::Package::from_bytes(bytes).unwrap());
        },
        _ => {
            black_box(litchi_numbers::Package::from_bytes(bytes).unwrap());
        },
    }
}

fn leaf_open_get(kind: &str, bytes: &[u8]) {
    match kind {
        "pages" => {
            let package = litchi_pages::Package::from_bytes(bytes).unwrap();
            let text = package
                .section_text(litchi_pages::SectionSelector::index(0))
                .unwrap();
            black_box(text.split('\n').next().map(str::to_owned));
        },
        "keynote" => {
            let package = litchi_keynote::Package::from_bytes(bytes).unwrap();
            black_box(
                package
                    .slide_title(litchi_keynote::SlideSelector::index(0))
                    .unwrap(),
            );
        },
        _ => {
            let package = litchi_numbers::Package::from_bytes(bytes).unwrap();
            let table = package.sheets()[0].tables().next().unwrap();
            black_box(table.get_a1("B2").unwrap().cloned());
        },
    }
}

fn unified_open_get(kind: &str, bytes: Vec<u8>) {
    let file = litchi_ops::open_bytes(bytes).unwrap();
    let path = match kind {
        "pages" => "/body/p[1]",
        "keynote" => "/slide[1]/placeholder[title]",
        _ => "/sheet[1]/table[1]/B2",
    };
    let node = file.get(path).unwrap();
    let node = node.as_iwork().unwrap();
    black_box(node.text().unwrap().map(std::borrow::Cow::into_owned));
}

/// The leaf packages a set starts from: opened outside the timed region.
enum Leaf {
    Pages(litchi_pages::Package),
    Keynote(litchi_keynote::Package),
    Numbers(litchi_numbers::Package),
}

fn leaf_package(kind: &str, bytes: &[u8]) -> Leaf {
    match kind {
        "pages" => Leaf::Pages(litchi_pages::Package::from_bytes(bytes).unwrap()),
        "keynote" => Leaf::Keynote(litchi_keynote::Package::from_bytes(bytes).unwrap()),
        _ => Leaf::Numbers(litchi_numbers::Package::from_bytes(bytes).unwrap()),
    }
}

fn leaf_set(package: Leaf) {
    match package {
        Leaf::Pages(package) => {
            let selector = litchi_pages::SectionSelector::index(0);
            let text = package.section_text(selector).unwrap();
            let length = text.split('\n').next().unwrap().encode_utf16().count();
            let span = litchi_pages::TextSpan::from_utf16_indexes(0, length).unwrap();
            let mut edit = package.edit_section_text(selector).unwrap();
            edit.replace(span, "Benchmark paragraph").unwrap();
            drop(black_box(edit.commit().unwrap()));
        },
        Leaf::Keynote(package) => {
            let mut edit = package
                .edit_slide_text(
                    litchi_keynote::SlideSelector::index(0),
                    litchi_keynote::slide::placeholder::Kind::Title,
                )
                .unwrap();
            edit.set("Benchmark title").unwrap();
            drop(black_box(edit.commit().unwrap()));
        },
        Leaf::Numbers(package) => {
            let edit = package
                .edit_table_cells(
                    litchi_numbers::SheetSelector::index(0),
                    litchi_numbers::TableSelector::index(0),
                )
                .unwrap()
                .set(
                    litchi_numbers::table::At::new(2, 1),
                    litchi_numbers::table::cells::Input::number(7.5).unwrap(),
                )
                .unwrap();
            drop(black_box(edit.commit().unwrap()));
        },
    }
}

fn unified_set(kind: &str, file: &litchi_ops::File) {
    let result = match kind {
        "pages" => file.set("/body/p[1]", props! { "text" => "Benchmark paragraph" }),
        "keynote" => file.set(
            "/slide[1]/placeholder[title]",
            props! { "text" => "Benchmark title" },
        ),
        _ => file.set("/sheet[1]/table[1]/B3", props! { "value" => 7.5 }),
    };
    black_box(result.unwrap());
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
    let iterations = argument("--iterations", 40);
    for (kind, relative) in FIXTURES {
        let bytes = std::fs::read(test_data().join(relative)).unwrap();
        for scenario in ["open", "open_get", "set"] {
            let mut ratios = Vec::with_capacity(rounds * 2);
            let mut leaf_medians = Vec::new();
            let mut unified_medians = Vec::new();
            for round in 0..rounds {
                // ABBA: leaf, unified, unified, leaf.
                for (position, leg_name) in
                    ["leaf", "unified", "unified", "leaf"].iter().enumerate()
                {
                    let median = match (scenario, *leg_name) {
                        ("open", "leaf") => leg(iterations, || (), |()| leaf_open(kind, &bytes)),
                        ("open", _) => leg(
                            iterations,
                            || bytes.clone(),
                            |owned| {
                                black_box(litchi_ops::open_bytes(owned).unwrap());
                            },
                        ),
                        ("open_get", "leaf") => {
                            leg(iterations, || (), |()| leaf_open_get(kind, &bytes))
                        },
                        ("open_get", _) => leg(
                            iterations,
                            || bytes.clone(),
                            |owned| unified_open_get(kind, owned),
                        ),
                        ("set", "leaf") => leg(iterations, || leaf_package(kind, &bytes), leaf_set),
                        _ => leg(
                            iterations,
                            || litchi_ops::open_bytes(bytes.clone()).unwrap(),
                            |file| unified_set(kind, &file),
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
