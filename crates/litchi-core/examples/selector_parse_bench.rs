//! Throughput probe for the path/selector grammar.
//!
//! Parses the harvested OfficeCLI corpus and a few typical inputs repeatedly
//! and prints the mean time per parse. Run with
//! `cargo run --release -p litchi-core --example selector_parse_bench`.
//! Numbers are machine-specific indications, not budgets.

use std::hint::black_box;
use std::time::Instant;

use litchi_core::selector::{Path, Query};

fn corpus(text: &'static str) -> Vec<&'static str> {
    text.lines()
        .filter(|line| !line.trim().is_empty() && !line.starts_with('#'))
        .collect()
}

fn time(label: &str, inputs: &[&str], rounds: usize, parse: impl Fn(&str) -> bool) {
    // Warm up caches and branch predictors once.
    for input in inputs {
        assert!(parse(input), "{input:?} failed to parse");
    }
    let start = Instant::now();
    for _ in 0..rounds {
        for input in inputs {
            black_box(parse(black_box(input)));
        }
    }
    let elapsed = start.elapsed();
    let parses = (rounds * inputs.len()) as f64;
    let bytes: usize = inputs.iter().map(|input| input.len()).sum();
    println!(
        "{label:<28} {:>8.1} ns/parse  {:>7.1} MB/s  ({} inputs x {rounds})",
        elapsed.as_nanos() as f64 / parses,
        (bytes * rounds) as f64 / elapsed.as_secs_f64() / 1e6,
        inputs.len(),
    );
}

fn main() {
    let paths = corpus(include_str!("../tests/fixtures/officecli_paths.txt"));
    let queries = corpus(include_str!("../tests/fixtures/officecli_selectors.txt"));
    let rounds = std::env::args()
        .nth(1)
        .and_then(|arg| arg.parse().ok())
        .unwrap_or(2_000);

    time("corpus paths", &paths, rounds, |input| {
        Path::parse(input).is_ok()
    });
    time("corpus selectors", &queries, rounds * 5, |input| {
        Query::parse(input).is_ok()
    });
    time(
        "/body/p[3]/r[1]",
        &["/body/p[3]/r[1]"],
        rounds * 500,
        |input| Path::parse(input).is_ok(),
    );
    time(
        "/slide[1]/shape[@name=Title]",
        &["/slide[1]/shape[@name=Title]"],
        rounds * 500,
        |input| Path::parse(input).is_ok(),
    );
    time("/Sheet1/A1:C5", &["/Sheet1/A1:C5"], rounds * 500, |input| {
        Path::parse(input).is_ok()
    });
    time(
        "row[Salary>5000 and ...]",
        &["row[Salary>5000 and Region=EMEA]"],
        rounds * 500,
        |input| Query::parse(input).is_ok(),
    );
}
