//! Change 0647 probe D: the callgrind isolation-pair body and the timing body.
//!
//! One iteration is the whole scenario: owned-source open, take the mutable
//! relationship seam, call `get_or_add` with a (type, target) pair the
//! collection already carries, and publish. Profiling `iterations = N` and
//! `iterations = N + M` and differencing the totals gives the instruction cost
//! of M scenarios with process start-up, the file read and the profiler's own
//! fixed cost cancelled out.
//!
//! `--seam-only` drops the `get_or_add` call and keeps everything else, which
//! is the baseline both legs share: the difference between the two modes on
//! one leg is the whole cost of the reusing call plus whatever the publication
//! then has to redo.
//!
//! With `--timing <samples>` the same body is timed instead of profiled: the
//! scenario runs once per sample and one line of nanoseconds per sample is
//! printed, with `--warmup` iterations discarded first.
//!
//! Usage: reuse_iters <fixture> <iterations> [--package|--part /partname]
//!                    [--seam-only] [--timing <samples>] [--warmup <n>]
use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use std::time::Instant;

struct Config {
    fixture: String,
    iterations: usize,
    part: Option<String>,
    seam_only: bool,
    timing: Option<usize>,
    warmup: usize,
}

fn parse() -> Config {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut positional: Vec<String> = Vec::new();
    let mut part = None;
    let mut seam_only = false;
    let mut timing = None;
    let mut warmup = 0usize;
    let mut index = 0usize;
    while index < args.len() {
        match args[index].as_str() {
            "--package" => index += 1,
            "--seam-only" => {
                seam_only = true;
                index += 1;
            },
            "--part" => {
                part = args.get(index + 1).cloned();
                index += 2;
            },
            "--timing" => {
                timing = args.get(index + 1).and_then(|value| value.parse().ok());
                index += 2;
            },
            "--warmup" => {
                warmup = args
                    .get(index + 1)
                    .and_then(|value| value.parse().ok())
                    .unwrap_or(0);
                index += 2;
            },
            other => {
                positional.push(other.to_string());
                index += 1;
            },
        }
    }
    Config {
        fixture: positional.first().cloned().expect("fixture"),
        iterations: positional
            .get(1)
            .and_then(|value| value.parse().ok())
            .unwrap_or(1),
        part,
        seam_only,
        timing,
        warmup,
    }
}

fn scenario(
    bytes: &[u8],
    uri: Option<&PackURI>,
    reltype: &str,
    target: &str,
    seam_only: bool,
) -> usize {
    let mut package = OpcPackage::from_vec(bytes.to_vec()).expect("open");
    match uri {
        None => {
            let rels = package.rels_mut();
            if !seam_only {
                let _ = rels.get_or_add(reltype, target);
            }
        },
        Some(uri) => {
            let part = package.get_part_mut(uri).expect("part");
            let rels = part.rels_mut();
            if !seam_only {
                let _ = rels.get_or_add(reltype, target);
            }
        },
    }
    PackageWriter::to_bytes(&package).expect("save").len()
}

fn main() {
    let config = parse();
    let bytes = std::fs::read(&config.fixture).expect("fixture");
    let uri = config
        .part
        .as_ref()
        .map(|name| PackURI::new(name).expect("partname"));

    let (reltype, target) = {
        let probe = OpcPackage::from_vec(bytes.clone()).expect("open");
        let rels = match &uri {
            None => probe.rels(),
            Some(uri) => probe.get_part(uri).expect("part").rels(),
        };
        let mut pairs: Vec<(String, String)> = rels
            .iter()
            .filter(|rel| !rel.is_external())
            .map(|rel| (rel.reltype().to_string(), rel.target_ref().to_string()))
            .collect();
        pairs.sort();
        pairs.first().cloned().expect("an internal relationship")
    };

    if let Some(samples) = config.timing {
        for _ in 0..config.warmup {
            let _ = scenario(&bytes, uri.as_ref(), &reltype, &target, config.seam_only);
        }
        let mut checksum = 0u64;
        for _ in 0..samples {
            let start = Instant::now();
            let published = scenario(&bytes, uri.as_ref(), &reltype, &target, config.seam_only);
            let elapsed = start.elapsed().as_nanos();
            checksum = checksum.wrapping_add(published as u64);
            println!("{elapsed}");
        }
        eprintln!("checksum={checksum}");
        return;
    }

    let mut checksum = 0u64;
    for _ in 0..config.iterations {
        checksum = checksum
            .wrapping_add(scenario(&bytes, uri.as_ref(), &reltype, &target, config.seam_only) as u64);
    }
    println!("iterations={} checksum={checksum}", config.iterations);
}
