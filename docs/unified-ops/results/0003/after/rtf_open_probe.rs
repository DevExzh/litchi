//! M0a S10 follow-up probe: the RTF open gap of `rtf_file_unified_open`.
//!
//! Rebuilds the harness's `rtf-semantic-plain-<shape>` corpus bytes, then
//! times `litchi::open_bytes` (unified) and
//! `litchi_rtf::RtfDocument::from_owned_bytes` (leaf) in one process, in
//! ABBA blocks, recording the elapsed time and the minor page faults of the
//! timed open only. Two loop shapes:
//! - `bare`: clone the input, time the open, drop the handle;
//! - `harness`: as the S8b harness sample loop does, also build the untimed
//!   projection (text digest and per-paragraph text) after the timer and
//!   drop the handle after it.
//!
//! usage: rtf-open-probe <tiny|medium|large> <bare|harness> <samples> <blocks>
//!        rtf-open-probe <tiny|medium|large> flat <samples>

use std::time::Instant;

fn paragraphs(shape: &str) -> usize {
    match shape {
        "tiny" => 24,
        "medium" => 200,
        "large" => 10_000,
        other => panic!("unknown shape {other}"),
    }
}

fn corpus(shape: &str) -> Vec<u8> {
    let mut source = String::from(r"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Arial;}}\f0\fs20 ");
    for index in 0..paragraphs(shape) {
        if index != 0 {
            source.push_str(r"\par ");
        }
        source.push_str(&format!("litchi-perf-baseline-rtf-semantic-v1-source-{index:05}"));
    }
    source.push('}');
    source.into_bytes()
}

fn minor_faults() -> u64 {
    let stat = std::fs::read_to_string("/proc/self/stat").expect("procfs");
    let after = &stat[stat.rfind(')').expect("comm") + 2..];
    after.split(' ').nth(7).expect("minflt").parse().expect("minflt")
}

fn median(values: &mut [u64]) -> u64 {
    values.sort_unstable();
    let n = values.len();
    (values[(n - 1) / 2] + values[n / 2]) / 2
}

fn run(leg: &str, archive: &[u8], harness: bool, samples: usize) -> (u64, u64) {
    let mut elapsed = Vec::with_capacity(samples);
    let mut faults = Vec::with_capacity(samples);
    for _ in 0..samples + 10 {
        let owned = archive.to_vec();
        let before = minor_faults();
        let (duration, after) = if leg == "unified" {
            let started = Instant::now();
            let file = litchi::open_bytes(owned).expect("unified open");
            let duration = started.elapsed();
            let after = minor_faults();
            if harness {
                let text = file.text().expect("text");
                let paragraphs: Vec<String> = file
                    .document()
                    .expect("document view")
                    .paragraphs()
                    .expect("paragraphs")
                    .iter()
                    .map(|paragraph| paragraph.text().expect("paragraph text"))
                    .collect();
                std::hint::black_box((&text, &paragraphs));
            }
            std::hint::black_box(&file);
            (duration, after)
        } else {
            let started = Instant::now();
            let document = litchi_rtf::RtfDocument::from_owned_bytes(owned).expect("leaf open");
            let duration = started.elapsed();
            let after = minor_faults();
            if harness {
                let text = document.text();
                let paragraphs: Vec<String> = document
                    .paragraphs_with_content()
                    .iter()
                    .map(|paragraph| paragraph.text())
                    .collect();
                std::hint::black_box((&text, &paragraphs));
            }
            std::hint::black_box(&document);
            (duration, after)
        };
        elapsed.push(u64::try_from(duration.as_nanos()).expect("ns"));
        faults.push(after - before);
    }
    let elapsed = &mut elapsed[10..];
    let faults = &mut faults[10..];
    (median(elapsed), median(faults))
}

/// Times `litchi_odf_common::detect::flat` on the corpus bytes: the flat-ODF
/// XML probe `litchi::open_bytes` runs before it checks the RTF header.
fn flat_probe(archive: &[u8], samples: usize) -> u64 {
    let mut elapsed = Vec::with_capacity(samples);
    for _ in 0..samples + 10 {
        let started = Instant::now();
        std::hint::black_box(litchi_odf_common::detect::flat(std::hint::black_box(archive)));
        elapsed.push(u64::try_from(started.elapsed().as_nanos()).expect("ns"));
    }
    median(&mut elapsed[10..])
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let shape = args[1].as_str();
    if args[2] == "flat" {
        let archive = corpus(shape);
        let samples: usize = args[3].parse().expect("samples");
        println!(
            "{{\"shape\":\"{shape}\",\"bytes\":{},\"flat_probe_p50_ns\":{}}}",
            archive.len(),
            flat_probe(&archive, samples)
        );
        return;
    }
    let harness = args[2] == "harness";
    let samples: usize = args[3].parse().expect("samples");
    let blocks: usize = args[4].parse().expect("blocks");
    let archive = corpus(shape);
    println!("corpus {shape}: {} bytes", archive.len());
    for block in 0..blocks {
        for leg in ["unified", "leaf", "leaf", "unified"] {
            let (p50, faults) = run(leg, &archive, harness, samples);
            println!(
                "{{\"shape\":\"{shape}\",\"loop\":\"{}\",\"block\":{block},\"leg\":\"{leg}\",\"p50_ns\":{p50},\"open_minor_faults_p50\":{faults}}}",
                if harness { "harness" } else { "bare" }
            );
        }
    }
}
