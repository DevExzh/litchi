//! Change 0747 sizing probe (scratch only; never committed as an example).
//!
//! Times `verify_source` under `Limits::default()` on the exact byte arrays a
//! measured source-backed XLSX one-edit publication audits, plus the parts of
//! that work a proof-carrying design could share or skip: quick-xml
//! tokenization alone, whole-input UTF-8 validation alone, and the prefix /
//! suffix comparison that would prove two payloads identical outside a
//! window.
//!
//! Usage: audit_probe_0747 <label> <original> <replacement> [<label> ...]
#![allow(clippy::unwrap_used, clippy::expect_used, clippy::print_stdout)]

use std::hint::black_box;
use std::time::Instant;

use quick_xml::Reader;
use quick_xml::events::Event;
use xml_minifier::audit::{Limits, verify_source, verify_source_replacement};

const BATCHES: usize = 41;
const TARGET_BATCH_NS: u128 = 20_000_000;

fn tokenize(bytes: &[u8]) -> usize {
    let xml = std::str::from_utf8(bytes).unwrap();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut events = 0usize;
    loop {
        match reader.read_event().unwrap() {
            Event::Eof => break,
            _ => events += 1,
        }
    }
    events
}

fn common_prefix(left: &[u8], right: &[u8]) -> usize {
    left.iter().zip(right).take_while(|(a, b)| a == b).count()
}

fn prove_outside_window(original: &[u8], replacement: &[u8]) -> (usize, usize) {
    let prefix = common_prefix(original, replacement);
    let limit = original.len().min(replacement.len()) - prefix;
    let suffix = original
        .iter()
        .rev()
        .zip(replacement.iter().rev())
        .take(limit)
        .take_while(|(a, b)| a == b)
        .count();
    (prefix, suffix)
}

/// Median nanoseconds per call over `BATCHES` batches sized to ~20 ms.
fn time<F: FnMut() -> usize>(mut f: F) -> (f64, f64, f64) {
    let mut reps = 1usize;
    loop {
        let start = Instant::now();
        for _ in 0..reps {
            black_box(f());
        }
        if start.elapsed().as_nanos() >= TARGET_BATCH_NS / 4 {
            break;
        }
        reps *= 2;
    }
    let mut per_call = Vec::with_capacity(BATCHES);
    for _ in 0..BATCHES {
        let start = Instant::now();
        for _ in 0..reps {
            black_box(f());
        }
        per_call.push(start.elapsed().as_nanos() as f64 / reps as f64);
    }
    per_call.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let p50 = per_call[BATCHES / 2];
    let p10 = per_call[BATCHES / 10];
    let p90 = per_call[BATCHES * 9 / 10];
    (p10, p50, p90)
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    assert!(args.len() % 3 == 0, "label original replacement triples");
    println!("label,side,bytes,measure,p10_ns,p50_ns,p90_ns,detail");
    for triple in args.chunks(3) {
        let label = &triple[0];
        let original = std::fs::read(&triple[1]).unwrap();
        let replacement = std::fs::read(&triple[2]).unwrap();
        for (side, bytes) in [("original", &original), ("replacement", &replacement)] {
            let report = verify_source(bytes, Limits::default()).expect("audit passes");
            let (p10, p50, p90) =
                time(|| verify_source(bytes, Limits::default()).unwrap().events());
            println!(
                "{label},{side},{},verify_source,{p10:.0},{p50:.0},{p90:.0},events={} attributes={}",
                bytes.len(),
                report.events(),
                report.attributes()
            );
            let (p10, p50, p90) = time(|| tokenize(bytes));
            println!(
                "{label},{side},{},tokenize_only,{p10:.0},{p50:.0},{p90:.0},events={}",
                bytes.len(),
                tokenize(bytes) + 1
            );
            let (p10, p50, p90) = time(|| std::str::from_utf8(bytes).unwrap().len());
            println!(
                "{label},{side},{},utf8_only,{p10:.0},{p50:.0},{p90:.0},",
                bytes.len()
            );
        }
        let proof = verify_source_replacement(&original, &replacement, Limits::default())
            .expect("pair passes");
        let (p10, p50, p90) = time(|| {
            verify_source(&original, Limits::default())
                .unwrap()
                .events()
                + verify_source(&replacement, Limits::default())
                    .unwrap()
                    .events()
        });
        println!(
            "{label},pair,{},two_verify_source,{p10:.0},{p50:.0},{p90:.0},",
            original.len()
        );
        let (p10, p50, p90) = time(|| {
            match verify_source_replacement(&original, &replacement, Limits::default()).unwrap() {
                xml_minifier::audit::ReplacementProof::Window { replacement, .. } => {
                    replacement.len()
                },
                _ => 0,
            }
        });
        println!(
            "{label},pair,{},verify_source_replacement,{p10:.0},{p50:.0},{p90:.0},{proof:?}",
            original.len()
        );
        let (prefix, suffix) = prove_outside_window(&original, &replacement);
        let (p10, p50, p90) = time(|| {
            let (prefix, suffix) = prove_outside_window(&original, &replacement);
            prefix + suffix
        });
        println!(
            "{label},pair,{},prefix_suffix_compare,{p10:.0},{p50:.0},{p90:.0},prefix={prefix} suffix={suffix} window_original={} window_replacement={}",
            original.len(),
            original.len() - prefix - suffix,
            replacement.len() - prefix - suffix
        );
    }
}
