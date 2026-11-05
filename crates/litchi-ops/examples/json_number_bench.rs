//! Batch-payload JSON decode timing (record 0338 F1).
//!
//! The workspace builds `serde_json` with `float_roundtrip` (correctly
//! rounded decimal-to-`f64` conversion) since record 0338. This example
//! times `litchi_ops::batch::from_bytes` on the payloads that feature
//! affects, so two builds (with and without the feature) can be compared:
//!
//! - `seeds`: the OfficeCLI batch seeds of the batch-payload fuzz corpus
//!   (decoded one after another; mostly strings);
//! - `cells-10k-short`: 10,000 cell sets with short decimals (`12.5`);
//! - `cells-10k-long`: 10,000 cell sets with 17-significant-digit decimals,
//!   the slow case of a correctly rounded parse;
//! - `cells-10k-int`: 10,000 cell sets with integers (no float parse).
//!
//! Each line prints the median of `ROUNDS` timed decodes in nanoseconds and
//! a checksum of the decoded numbers, which differs between the builds only
//! where the old conversion was not correctly rounded.
//!
//! Run: `cargo run --release -p litchi-ops --example json_number_bench`.

use std::hint::black_box;
use std::time::Instant;

use litchi_ops::{PropValue, batch};

const ROUNDS: usize = 31;
const WARMUP: usize = 3;

fn median(mut values: Vec<u128>) -> u128 {
    values.sort_unstable();
    values[values.len() / 2]
}

fn checksum(items: &[litchi_ops::BatchItem<'_>]) -> u64 {
    let mut sum = 0u64;
    for item in items {
        if let Some(props) = &item.props {
            for (_, value) in props.iter() {
                if let PropValue::Number(number) = value {
                    sum = sum.rotate_left(5) ^ number.to_bits();
                }
            }
        }
    }
    sum
}

fn time(name: &str, payloads: &[Vec<u8>]) {
    let mut sum = 0u64;
    let mut items = 0usize;
    for payload in payloads {
        let decoded =
            batch::from_bytes(payload, &batch::SourceLimits::default()).expect("payload decodes");
        sum ^= checksum(&decoded);
        items += decoded.len();
    }
    for _ in 0..WARMUP {
        for payload in payloads {
            black_box(
                batch::from_bytes(black_box(payload), &batch::SourceLimits::default())
                    .expect("decodes"),
            );
        }
    }
    let mut samples = Vec::with_capacity(ROUNDS);
    for _ in 0..ROUNDS {
        let start = Instant::now();
        for payload in payloads {
            black_box(
                batch::from_bytes(black_box(payload), &batch::SourceLimits::default())
                    .expect("decodes"),
            );
        }
        samples.push(start.elapsed().as_nanos());
    }
    let bytes: usize = payloads.iter().map(Vec::len).sum();
    println!(
        "{name}: {} ns median, {items} items, {bytes} bytes, checksum {sum:016x}",
        median(samples)
    );
}

fn cells(count: usize, value: impl Fn(usize) -> String) -> Vec<u8> {
    let mut out = String::from("[");
    for index in 0..count {
        if index > 0 {
            out.push(',');
        }
        out.push_str(&format!(
            "{{\"command\":\"set\",\"path\":\"/Sheet1/A{}\",\"props\":{{\"value\":{},\"x\":{}}}}}",
            index + 1,
            value(index),
            value(index + 7)
        ));
    }
    out.push(']');
    out.into_bytes()
}

fn main() {
    let corpus = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("fuzz/corpus/batch_payload");
    let mut seeds: Vec<Vec<u8>> = std::fs::read_dir(&corpus)
        .expect("fuzz corpus")
        .filter_map(|entry| {
            let path = entry.ok()?.path();
            let name = path.file_name()?.to_str()?.to_owned();
            name.starts_with("officecli-")
                .then(|| std::fs::read(&path).expect("seed"))
        })
        .collect();
    seeds.sort();
    // A deterministic generator of decimals (no clock, no RNG crate).
    let digits = |index: usize, width: usize| -> String {
        let mut state = (index as u64).wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1;
        (0..width)
            .map(|_| {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                char::from(b'0' + (state % 10) as u8)
            })
            .collect()
    };
    let short = cells(10_000, |index| format!("{}.5", index % 1000));
    let long = cells(10_000, |index| {
        let text = digits(index, 17);
        format!("{}.{}e{}", &text[..1], &text[1..], index % 40)
    });
    let int = cells(10_000, |index| (index * 37).to_string());
    time("seeds", &seeds);
    time("cells-10k-short", &[short]);
    time("cells-10k-long", &[long]);
    time("cells-10k-int", &[int]);
}
