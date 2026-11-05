//! Existing-path read cost before and after record 0319: a paragraph's
//! format, a body query and a root view of a real DOCX, through the public
//! `File` API only, so the same source builds on the base and the branch.
//!
//! Prints the median ns per iteration of each case over 15 rounds. Run with
//! `cargo run --release -p litchi-ops --features docx --example
//! docx_view_regression -- <file.docx> [paragraph path]`; record 0319 §6
//! has the results.

#[cfg(feature = "docx")]
use std::hint::black_box;
#[cfg(feature = "docx")]
use std::time::Instant;

#[cfg(feature = "docx")]
use litchi_ops::view::Mode;
#[cfg(feature = "docx")]
use litchi_ops::{OpenOptions, SourceMode};

#[cfg(feature = "docx")]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    samples[samples.len() / 2]
}

#[cfg(feature = "docx")]
fn time(name: &str, iterations: usize, mut leg: impl FnMut()) {
    for _ in 0..3 {
        for _ in 0..iterations {
            leg();
        }
    }
    let mut samples = Vec::new();
    for _ in 0..15 {
        let start = Instant::now();
        for _ in 0..iterations {
            leg();
        }
        samples.push(start.elapsed().as_nanos() as f64 / iterations as f64);
    }
    println!("{name}\t{:.0}", median(&mut samples));
}

#[cfg(feature = "docx")]
fn main() {
    let path = std::env::args().nth(1).expect("a docx path");
    let bytes = std::fs::read(&path).expect("read");
    let paragraph = std::env::args()
        .nth(2)
        .unwrap_or_else(|| "/body/p[3]".to_owned());
    let paragraph = paragraph.as_str();
    for (label, mode) in [("source", SourceMode::Auto), ("eager", SourceMode::Eager)] {
        let file =
            litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().source_mode(mode))
                .expect("open");
        let _ = file
            .get(paragraph)
            .unwrap()
            .view()
            .format()
            .unwrap()
            .count();
        time(&format!("{label}_paragraph_format"), 2000, || {
            let node = file.get(paragraph).unwrap();
            black_box(node.view().format().unwrap().count());
        });
        time(&format!("{label}_query_paragraphs_format"), 20, || {
            for node in file.query("paragraph").unwrap() {
                black_box(node.unwrap().view().format().unwrap().count());
            }
        });
        time(&format!("{label}_view_text"), 20, || {
            black_box(file.view(Mode::Text).unwrap().len());
        });
        time(&format!("{label}_root_children"), 2000, || {
            let root = file.get("/").unwrap();
            black_box(root.view().children().count());
        });
    }
}

#[cfg(not(feature = "docx"))]
fn main() {
    eprintln!("enable the docx feature");
}
