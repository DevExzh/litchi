//! Timing probe for the query engine over the test-only fake formats (M1
//! step R3, record 0010 §5).
//!
//! Prints the time to compile a plan, the walk time per visited node for
//! whole-document queries, and the time to the first match (`limit(1)`) on a
//! medium and a large document, where the first match sits at the same
//! position. Run with
//! `cargo run --release -p litchi-ops --features test-support --example query_walk_bench`.
//! Numbers are machine-specific indications over fake documents, not budgets
//! and not measurements of any real format adapter.

use std::hint::black_box;
use std::time::{Duration, Instant};

use litchi_ops::query::__fake::{Document, Family, compile_for, run};
use litchi_ops::query::Query;

/// The median of `samples` timings of `body`, each repeated `repeat` times.
fn median(samples: usize, repeat: usize, mut body: impl FnMut()) -> Duration {
    let mut times: Vec<Duration> = (0..samples)
        .map(|_| {
            let start = Instant::now();
            for _ in 0..repeat {
                body();
            }
            start.elapsed() / u32::try_from(repeat).unwrap_or(u32::MAX)
        })
        .collect();
    times.sort();
    times[times.len() / 2]
}

fn main() {
    println!("compile (median of 9 samples x 2000)");
    for (family, selector) in [
        (Family::Word, "paragraph[style=Heading1]"),
        (
            Family::Word,
            "paragraph[style=Normal] > run[font.latin!=Arial]",
        ),
        (Family::Word, "p:has(r[bold=true]):not(:empty), tbl tc > p"),
        (
            Family::Presentation,
            "shape[fill=#FF0000][width>=10cm]:contains(Revenue)",
        ),
        (Family::Workbook, "row[Region=West and Sales>1000]"),
    ] {
        let query = Query::parse(selector).expect("parses");
        let time = median(9, 2_000, || {
            black_box(compile_for(family, black_box(&query)).expect("compiles"));
        });
        println!("  {selector:<56} {:>8} ns", time.as_nanos());
    }

    println!("walk, whole document (median of 7 samples)");
    let documents = [
        (Family::Word, Document::word(20_000)),
        (Family::Presentation, Document::deck(2_000, 8)),
        (Family::Workbook, Document::workbook(20_000)),
    ];
    for (family, selector) in [
        (Family::Word, "p"),
        (Family::Word, "p[style=Heading1]"),
        (Family::Word, "r[bold=true]"),
        (Family::Word, "p:contains(revenue)"),
        (Family::Word, "body > p > r[bold=true]"),
        (Family::Word, "p:has(r[bold=true])"),
        (Family::Presentation, "shape[width>=10cm]"),
        (Family::Presentation, "textbox:contains(revenue)"),
        (Family::Workbook, "cell[value>1000]"),
        (Family::Workbook, "row[Sales>1000]"),
    ] {
        let document = &documents
            .iter()
            .find(|(candidate, _)| *candidate == family)
            .expect("a document per family")
            .1;
        let query = Query::parse(selector).expect("parses");
        let plan = compile_for(family, &query).expect("compiles");
        let outcome = run(document, &plan, usize::MAX).expect("runs");
        let time = median(7, 3, || {
            black_box(run(document, &plan, usize::MAX).expect("runs"));
        });
        println!(
            "  {selector:<32} {:>7} matches {:>8} visited {:>9.1} µs {:>6.1} ns/node",
            outcome.matches,
            outcome.visited,
            time.as_secs_f64() * 1e6,
            time.as_nanos() as f64 / outcome.visited.max(1) as f64,
        );
    }

    println!("first match, limit(1) (median of 9 samples x 1000)");
    for paragraphs in [1_000, 100_000] {
        let document = Document::word(paragraphs);
        let query = Query::parse("p[style=Heading1]").expect("parses");
        let plan = compile_for(Family::Word, &query).expect("compiles");
        let outcome = run(&document, &plan, 1).expect("runs");
        let time = median(9, 1_000, || {
            black_box(run(&document, &plan, 1).expect("runs"));
        });
        println!(
            "  p[style=Heading1] over {paragraphs:>6} paragraphs: {:>3} visited {:>8} ns",
            outcome.visited,
            time.as_nanos()
        );
    }
}
