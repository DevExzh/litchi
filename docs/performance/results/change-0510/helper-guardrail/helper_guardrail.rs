#![allow(clippy::cast_precision_loss)]

use std::hint::black_box;
use std::time::Instant;

const LENGTHS: [usize; 3] = [49, 1024, 65_536];
const SAMPLES: usize = 50;

#[derive(Clone, Copy, Debug)]
enum Pattern {
    AsciiNoCr,
    SparseCrlf,
    DenseCrlf,
    AllCr,
    Unicode,
}

impl Pattern {
    const ALL: [Self; 5] = [
        Self::AsciiNoCr,
        Self::SparseCrlf,
        Self::DenseCrlf,
        Self::AllCr,
        Self::Unicode,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::AsciiNoCr => "ascii_no_cr",
            Self::SparseCrlf => "sparse_crlf",
            Self::DenseCrlf => "dense_crlf",
            Self::AllCr => "all_cr",
            Self::Unicode => "unicode",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Variant {
    Scalar,
    MemchrIter,
    MemmemIter,
    FirstCrThenScalar,
}

impl Variant {
    const ALL: [Self; 4] = [
        Self::Scalar,
        Self::MemchrIter,
        Self::MemmemIter,
        Self::FirstCrThenScalar,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::Scalar => "scalar",
            Self::MemchrIter => "memchr_iter",
            Self::MemmemIter => "memmem_iter",
            Self::FirstCrThenScalar => "first_cr_then_scalar",
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct Stats {
    min_ns: u64,
    p50_ns: u64,
    p95_ns: u64,
    max_ns: u64,
    checksum: usize,
}

fn normalized_xml10_decoded_len_scalar(raw: &[u8]) -> Result<usize, ()> {
    std::str::from_utf8(raw).map_err(|_| ())?;
    let mut length = raw.len();
    let mut index = 0;
    while index < raw.len() {
        if raw[index] == b'\r' {
            if raw.get(index + 1) == Some(&b'\n') {
                length -= 1;
                index += 2;
            } else {
                index += 1;
            }
        } else {
            index += 1;
        }
    }
    Ok(length)
}

fn normalized_xml10_decoded_len_memchr_iter(raw: &[u8]) -> Result<usize, ()> {
    std::str::from_utf8(raw).map_err(|_| ())?;
    let mut crlf = 0usize;
    for index in memchr::memchr_iter(b'\r', raw) {
        if raw.get(index + 1) == Some(&b'\n') {
            crlf += 1;
        }
    }
    Ok(raw.len() - crlf)
}

fn normalized_xml10_decoded_len_memmem_iter(raw: &[u8]) -> Result<usize, ()> {
    std::str::from_utf8(raw).map_err(|_| ())?;
    let crlf = memchr::memmem::find_iter(raw, b"\r\n").count();
    Ok(raw.len() - crlf)
}

fn normalized_xml10_decoded_len_first_cr_then_scalar(raw: &[u8]) -> Result<usize, ()> {
    std::str::from_utf8(raw).map_err(|_| ())?;
    let Some(start) = memchr::memchr(b'\r', raw) else {
        return Ok(raw.len());
    };
    let mut length = raw.len();
    let mut index = start;
    while index < raw.len() {
        if raw[index] == b'\r' {
            if raw.get(index + 1) == Some(&b'\n') {
                length -= 1;
                index += 2;
            } else {
                index += 1;
            }
        } else {
            index += 1;
        }
    }
    Ok(length)
}

fn run(variant: Variant, raw: &[u8]) -> usize {
    match variant {
        Variant::Scalar => normalized_xml10_decoded_len_scalar(raw).unwrap(),
        Variant::MemchrIter => normalized_xml10_decoded_len_memchr_iter(raw).unwrap(),
        Variant::MemmemIter => normalized_xml10_decoded_len_memmem_iter(raw).unwrap(),
        Variant::FirstCrThenScalar => {
            normalized_xml10_decoded_len_first_cr_then_scalar(raw).unwrap()
        },
    }
}

fn make_input(length: usize, pattern: Pattern) -> Vec<u8> {
    let mut output = Vec::with_capacity(length);
    match pattern {
        Pattern::AsciiNoCr => output.resize(length, b'a'),
        Pattern::SparseCrlf => {
            for index in 0..length {
                output.push(if index % 97 == 95 {
                    b'\r'
                } else if index % 97 == 96 {
                    b'\n'
                } else {
                    b'a'
                });
            }
        },
        Pattern::DenseCrlf => {
            for index in 0..length {
                output.push(if index % 2 == 0 { b'\r' } else { b'\n' });
            }
        },
        Pattern::AllCr => output.resize(length, b'\r'),
        Pattern::Unicode => {
            while output.len() + 2 <= length {
                output.extend_from_slice("é".as_bytes());
            }
            if output.len() < length {
                output.push(b'a');
            }
        },
    }
    assert_eq!(output.len(), length);
    assert!(std::str::from_utf8(&output).is_ok());
    output
}

fn iterations(length: usize) -> usize {
    match length {
        49 => 4096,
        1024 => 512,
        65_536 => 16,
        _ => unreachable!(),
    }
}

fn percentile(sorted: &[u64], numerator: usize, denominator: usize) -> u64 {
    let index = ((sorted.len() - 1) * numerator) / denominator;
    sorted[index]
}

fn measure(variant: Variant, raw: &[u8], repeats: usize) -> Stats {
    let mut samples = Vec::with_capacity(SAMPLES);
    let mut checksum = 0usize;
    let count = iterations(raw.len());
    for sample in 0..SAMPLES {
        let mut value = 0usize;
        let started = Instant::now();
        for _ in 0..count {
            value ^= black_box(run(variant, black_box(raw)));
        }
        let elapsed = started.elapsed().as_nanos() as u64;
        let per_call = (elapsed / count as u64).max(1);
        samples.push(per_call);
        checksum ^= value.wrapping_add(sample).wrapping_add(repeats);
    }
    samples.sort_unstable();
    Stats {
        min_ns: samples[0],
        p50_ns: percentile(&samples, 1, 2),
        p95_ns: percentile(&samples, 19, 20),
        max_ns: samples[SAMPLES - 1],
        checksum,
    }
}

fn main() {
    println!("kind=helper_guardrail version=1 samples={SAMPLES}");
    println!("raw,order,length,pattern,variant,iterations,min_ns,p50_ns,p95_ns,max_ns,checksum");

    for (order_name, order) in [
        ("forward", Variant::ALL),
        ("reverse", [
            Variant::FirstCrThenScalar,
            Variant::MemmemIter,
            Variant::MemchrIter,
            Variant::Scalar,
        ]),
    ] {
        for &length in &LENGTHS {
            for &pattern in &Pattern::ALL {
                let input = make_input(length, pattern);
                let expected = run(Variant::Scalar, &input);
                for &variant in &Variant::ALL {
                    assert_eq!(run(variant, &input), expected);
                }
                let count = iterations(length);
                for &variant in &order {
                    for _ in 0..10 {
                        black_box(run(variant, black_box(&input)));
                    }
                    let stats = measure(variant, &input, if order_name == "forward" { 0 } else { 1 });
                    println!(
                        "raw,{order_name},{length},{},{},{},{},{},{},{},{}",
                        pattern.name(),
                        variant.name(),
                        count,
                        stats.min_ns,
                        stats.p50_ns,
                        stats.p95_ns,
                        stats.max_ns,
                        stats.checksum,
                    );
                }
            }
        }
    }
}
