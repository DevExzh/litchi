//! Change 0624: the deflate-scaling half of the gate measurement.
//!
//! This probe measures nothing about litchi. It measures the *ceiling* on what
//! per-member parallel deflate could win, by running the exact codec the
//! writers run — `flate2` 1.1 on the `zlib-rs` backend at
//! `Compression::default()` (level 6), raw deflate, one member per task —
//! serially and through a caller-owned Rayon pool of a stated width, over
//! member sets of a stated count and size. Nothing in the real writer can beat
//! these numbers, because the real writer also has to frame, account, audit and
//! emit.
//!
//! It is deliberately a separate crate from `probe0624`: it has no litchi
//! dependency, so a threshold measured here is a property of the codec and the
//! host, not of any litchi version.
//!
//! Usage:
//!   probe0624-deflate synth <count> <bytes> <kind> <warmup> <samples> <width..>
//!   probe0624-deflate files <dir> <warmup> <samples> <width..>
//!   probe0624-deflate sweep <kind> <warmup> <samples> <width..>
//!
//! `kind` is `xml` (worksheet-shaped markup, ~12:1 compressible) or `random`
//! (incompressible, the per-byte upper bound).

use std::env;
use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::time::Instant;

use flate2::Compression;
use flate2::write::DeflateEncoder;
use rayon::prelude::*;

const THRESHOLD: usize = 256 * 1024;

fn deflate_one(payload: &[u8]) -> usize {
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(payload).expect("deflate");
    encoder.finish().expect("finish").len()
}

fn deflate_serial(payloads: &[Vec<u8>]) -> usize {
    payloads.iter().map(|payload| deflate_one(payload)).sum()
}

fn deflate_pooled(pool: &rayon::ThreadPool, payloads: &[Vec<u8>]) -> usize {
    pool.install(|| {
        payloads
            .par_iter()
            .map(|payload| deflate_one(payload))
            .sum()
    })
}

fn percentile(sorted: &[u128], fraction: f64) -> u128 {
    if sorted.is_empty() {
        return 0;
    }
    let at = ((sorted.len() as f64 - 1.0) * fraction).round() as usize;
    sorted[at.min(sorted.len() - 1)]
}

struct Summary {
    p50: u128,
    p95: u128,
    p99: u128,
    mean: u128,
}

fn summarize(mut samples: Vec<u128>) -> Summary {
    samples.sort_unstable();
    Summary {
        p50: percentile(&samples, 0.50),
        p95: percentile(&samples, 0.95),
        p99: percentile(&samples, 0.99),
        mean: samples.iter().sum::<u128>() / samples.len().max(1) as u128,
    }
}

/// Worksheet-shaped markup: what an XLSX or DOCX regenerated member looks like.
fn xml_payload(index: usize, bytes: usize) -> Vec<u8> {
    let mut out = Vec::with_capacity(bytes + 128);
    out.extend_from_slice(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>",
    );
    let mut row = 0u64;
    while out.len() < bytes {
        row += 1;
        let _ = write!(
            out,
            "<row r=\"{row}\" spans=\"1:6\"><c r=\"A{row}\" t=\"inlineStr\"><is><t>member {index} row {row} value</t></is></c><c r=\"B{row}\"><v>{}</v></c></row>",
            row.wrapping_mul(2_654_435_761) % 1_000_000
        );
    }
    out.extend_from_slice(b"</sheetData></worksheet>");
    out
}

/// Incompressible bytes: the per-byte upper bound on deflate cost.
fn random_payload(index: usize, bytes: usize) -> Vec<u8> {
    let mut out = Vec::with_capacity(bytes + 8);
    let mut state = 0x243f_6a88_85a3_08d3u64 ^ (index as u64).wrapping_mul(0x9e37_79b9_7f4a_7c15);
    while out.len() < bytes {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        out.extend_from_slice(&state.to_le_bytes());
    }
    out.truncate(bytes);
    out
}

fn synth(count: usize, bytes: usize, kind: &str) -> Vec<Vec<u8>> {
    (0..count)
        .map(|index| match kind {
            "xml" => xml_payload(index, bytes),
            "random" => random_payload(index, bytes),
            other => panic!("unknown payload kind {other}"),
        })
        .collect()
}

fn load(dir: &Path) -> Vec<Vec<u8>> {
    let mut files: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("payload dir")
        .flatten()
        .map(|entry| entry.path())
        .filter(|path| path.is_file())
        .collect();
    files.sort();
    files
        .into_iter()
        .map(|path| std::fs::read(path).expect("payload"))
        .collect()
}

/// One scaling run: serial, then each width, with speedup, efficiency and the
/// Amdahl serial fraction. Superlinear and `S < 1` cells are labelled, never
/// fitted, as change 0088's harness and change 0615 both require.
fn scale(label: &str, payloads: &[Vec<u8>], warmup: u32, samples: u32, widths: &[usize]) {
    let total: usize = payloads.iter().map(Vec::len).sum();
    let over = payloads.iter().filter(|p| p.len() >= THRESHOLD).count();
    println!(
        "set\t{label}\tmembers={}\ttotal_bytes={total}\tmembers_ge_256KiB={over}\tmax_bytes={}",
        payloads.len(),
        payloads.iter().map(Vec::len).max().unwrap_or(0),
    );
    let mut sink = 0usize;
    for _ in 0..warmup {
        sink = sink.wrapping_add(deflate_serial(payloads));
    }
    let mut timings = Vec::with_capacity(samples as usize);
    for _ in 0..samples {
        let start = Instant::now();
        sink = sink.wrapping_add(deflate_serial(payloads));
        timings.push(start.elapsed().as_nanos());
    }
    let serial = summarize(timings);
    println!(
        "leg\t{label}\tserial\tp50={}\tmean={}\tp95={}\tp99={}",
        serial.p50, serial.mean, serial.p95, serial.p99
    );
    for &width in widths {
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(width)
            .build()
            .expect("pool");
        for _ in 0..warmup {
            sink = sink.wrapping_add(deflate_pooled(&pool, payloads));
        }
        let mut timings = Vec::with_capacity(samples as usize);
        for _ in 0..samples {
            let start = Instant::now();
            sink = sink.wrapping_add(deflate_pooled(&pool, payloads));
            timings.push(start.elapsed().as_nanos());
        }
        let parallel = summarize(timings);
        println!(
            "leg\t{label}\twidth{width}\tp50={}\tmean={}\tp95={}\tp99={}",
            parallel.p50, parallel.mean, parallel.p95, parallel.p99
        );
        let speedup = serial.p50 as f64 / parallel.p50.max(1) as f64;
        let efficiency = speedup / width as f64;
        let class = if width == 1 {
            "baseline-width"
        } else if speedup < 1.0 {
            "slowdown"
        } else if efficiency > 1.0 {
            "superlinear-out-of-model"
        } else {
            "valid"
        };
        let serial_fraction = if width > 1 && speedup >= 1.0 && efficiency <= 1.0 {
            let n = width as f64;
            format!("{:.4}", (1.0 / speedup - 1.0 / n) / (1.0 - 1.0 / n))
        } else {
            "out-of-model".to_owned()
        };
        println!(
            "scaling\t{label}\twidth={width}\tspeedup={speedup:.3}\tefficiency={efficiency:.3}\tamdahl_s={serial_fraction}\tclass={class}"
        );
    }
    eprintln!("sink {sink}");
}

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe0624-deflate synth <count> <bytes> <kind> <warmup> <samples> <width..> | files <dir> <warmup> <samples> <width..> | sweep <kind> <warmup> <samples> <width..>";
    match arguments.first().map(String::as_str) {
        Some("synth") => {
            let count: usize = arguments.get(1).expect(usage).parse().expect("count");
            let bytes: usize = arguments.get(2).expect(usage).parse().expect("bytes");
            let kind = arguments.get(3).expect(usage);
            let warmup: u32 = arguments.get(4).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(5).expect(usage).parse().expect("samples");
            let widths: Vec<usize> = arguments[6..]
                .iter()
                .map(|width| width.parse().expect("width"))
                .collect();
            scale(
                &format!("{kind}-{count}x{bytes}"),
                &synth(count, bytes, kind),
                warmup,
                samples,
                &widths,
            );
        },
        Some("files") => {
            let dir = PathBuf::from(arguments.get(1).expect(usage));
            let warmup: u32 = arguments.get(2).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(3).expect(usage).parse().expect("samples");
            let widths: Vec<usize> = arguments[4..]
                .iter()
                .map(|width| width.parse().expect("width"))
                .collect();
            let label = dir
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or("files")
                .to_owned();
            scale(&label, &load(&dir), warmup, samples, &widths);
        },
        Some("sweep") => {
            let kind = arguments.get(1).expect(usage);
            let warmup: u32 = arguments.get(2).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(3).expect(usage).parse().expect("samples");
            let widths: Vec<usize> = arguments[4..]
                .iter()
                .map(|width| width.parse().expect("width"))
                .collect();
            // The member sizes the corpus actually carries, and the ones the
            // 0587 survey's threshold names. 0.5 KiB is the `.rels` member the
            // preservation writer regenerates; 2 KiB an authored slide; 64 KiB
            // the largest member of the flagship 132-member workbook; 256 KiB
            // the survey's proposed floor; 1 MiB and 4 MiB the two members of
            // the largest real fixture.
            for &bytes in &[512usize, 2048, 16384, 65536, 262_144, 1_048_576, 4_194_304] {
                for &count in &[2usize, 4, 8, 16, 64, 161] {
                    if count * bytes > 512 * 1024 * 1024 {
                        continue;
                    }
                    scale(
                        &format!("{kind}-{count}x{bytes}"),
                        &synth(count, bytes, kind),
                        warmup,
                        samples,
                        &widths,
                    );
                }
            }
        },
        _ => {
            eprintln!("{usage}");
            std::process::exit(2);
        },
    }
}
