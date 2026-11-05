//! Measures what the `litchi` process costs over the in-process call it wraps
//! (record 0256 §6).
//!
//! Three legs per fixture, interleaved A B C C B A per round so drift hits
//! every leg alike:
//!
//! - `in-process`: `litchi_ops::open`, `File::get`, and the node view
//!   serialized as pretty JSON into a buffer: the work `litchi get --json`
//!   does after argument parsing.
//! - `cli-get`: spawn `litchi get <file> <path> --json`, discard stdout, wait.
//! - `cli-version`: spawn `litchi --version`, the process and argument-parser
//!   floor.
//!
//! Prints one JSON object with the per-leg median, 10th and 90th percentile
//! in microseconds. Run it on a release build, pinned to one CPU:
//!
//! ```text
//! cargo build --release -p litchi-cli --bin litchi --example cli_get_cost
//! taskset -c 3 target/release/examples/cli_get_cost target/release/litchi 40 15
//! ```

use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::Instant;

use litchi_ops::node::ViewOptions;
use serde_json::ser::PrettyFormatter;

const CASES: &[(&str, &str)] = &[
    ("ooxml/docx/drawing.docx", "/body/p[2]"),
    ("ooxml/xlsx/SimpleMultiCell.xlsx", "/Sheet1/A1"),
    ("ooxml/pptx/sample.pptx", "/slide[1]"),
    ("ole/xls/Simple.xls", "/sheet[1]"),
];

fn in_process(path: &Path, node: &str) -> usize {
    let file = litchi_ops::open(path).expect("open");
    let node = file.get(node).expect("get");
    let view = node.view_with(ViewOptions::default());
    let mut sink = Vec::with_capacity(4096);
    let mut serializer = serde_json::Serializer::with_formatter(&mut sink, PrettyFormatter::new());
    if view.serialize_into(&mut serializer).is_err() {
        panic!("serialize");
    }
    sink.len()
}

fn spawn(binary: &Path, args: &[&str]) {
    let status = Command::new(binary)
        .args(args)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .expect("spawn litchi");
    assert!(status.success(), "litchi {args:?} failed");
}

fn time(iterations: usize, mut body: impl FnMut()) -> Vec<f64> {
    (0..iterations)
        .map(|_| {
            let start = Instant::now();
            body();
            start.elapsed().as_secs_f64() * 1e6
        })
        .collect()
}

fn summary(mut samples: Vec<f64>) -> serde_json::Value {
    samples.sort_by(f64::total_cmp);
    let at = |quantile: f64| {
        let index = ((samples.len() - 1) as f64 * quantile).round() as usize;
        (samples[index] * 10.0).round() / 10.0
    };
    serde_json::json!({
        "samples": samples.len(),
        "median_us": at(0.5),
        "p10_us": at(0.1),
        "p90_us": at(0.9),
    })
}

fn main() {
    let mut args = std::env::args().skip(1);
    let binary = PathBuf::from(
        args.next()
            .expect("usage: cli_get_cost <litchi> [iterations] [rounds]"),
    );
    let iterations: usize = args
        .next()
        .map_or(40, |value| value.parse().expect("iterations"));
    let rounds: usize = args
        .next()
        .map_or(15, |value| value.parse().expect("rounds"));
    let data = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");

    let mut report = serde_json::Map::new();
    for &(relative, node) in CASES {
        let path = data.join(relative);
        let file = path.to_str().expect("UTF-8 path");
        let get_args = ["get", file, node, "--json"];
        // Warm the page cache and the binary.
        in_process(&path, node);
        spawn(&binary, &get_args);
        spawn(&binary, &["--version"]);

        let (mut a, mut b, mut c) = (Vec::new(), Vec::new(), Vec::new());
        for _ in 0..rounds {
            a.extend(time(iterations, || {
                std::hint::black_box(in_process(&path, node));
            }));
            b.extend(time(iterations, || spawn(&binary, &get_args)));
            c.extend(time(iterations, || spawn(&binary, &["--version"])));
            c.extend(time(iterations, || spawn(&binary, &["--version"])));
            b.extend(time(iterations, || spawn(&binary, &get_args)));
            a.extend(time(iterations, || {
                std::hint::black_box(in_process(&path, node));
            }));
        }
        report.insert(
            format!("{relative} {node}"),
            serde_json::json!({
                "in-process": summary(a),
                "cli-get": summary(b),
                "cli-version": summary(c),
            }),
        );
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&report).expect("serialize report")
    );
}
