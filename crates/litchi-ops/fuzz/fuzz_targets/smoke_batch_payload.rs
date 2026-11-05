//! Stable-toolchain smoke run of the batch-payload targets (record 0334).
//!
//! Replays every seed in `corpus/batch_payload/` plus deterministic
//! mutations (every prefix and every single-bit flip) through
//! `batch_payload_harness`. With the `apply` feature, the seeds, their
//! prefixes and a strided sample of the bit flips also run through
//! `batch_apply_harness`. libFuzzer is not involved, so this runs anywhere
//! `cargo run` does:
//!
//! ```sh
//! cargo run --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml \
//!     --bin smoke_batch_payload [--features apply]
//! ```
#[cfg(feature = "apply")]
mod batch_apply_harness;
mod batch_payload_harness;

use std::path::Path;

fn main() {
    let directory = Path::new(env!("CARGO_MANIFEST_DIR")).join("corpus/batch_payload");
    let mut seeds: Vec<_> = std::fs::read_dir(&directory)
        .unwrap_or_else(|error| panic!("read {}: {error}", directory.display()))
        .map(|entry| entry.expect("corpus entry").path())
        .collect();
    seeds.sort();
    assert!(!seeds.is_empty(), "no seeds in {}", directory.display());

    let (mut inputs, mut decoded, mut items) = (0_usize, 0_usize, 0_usize);
    let mut applied = 0_usize;
    let mut run = |data: &[u8], apply: bool| {
        inputs += 1;
        if let Some(count) = batch_payload_harness::exercise(data) {
            decoded += 1;
            items += count;
        }
        #[cfg(feature = "apply")]
        if apply {
            applied += 1;
            batch_apply_harness::exercise(data);
        }
        #[cfg(not(feature = "apply"))]
        let _ = (apply, &mut applied);
    };
    for seed in &seeds {
        let mut data = std::fs::read(seed).expect("seed bytes");
        run(&data, true);
        for length in 0..data.len() {
            run(&data[..length], length % 7 == 0);
        }
        for index in 0..data.len() {
            for bit in 0..8 {
                data[index] ^= 1 << bit;
                run(&data, (index * 8 + bit) % 61 == 0);
                data[index] ^= 1 << bit;
            }
        }
    }
    println!(
        "batch payload smoke: {} seeds, {inputs} inputs, {decoded} decoded ({items} items), \
         {applied} executed",
        seeds.len()
    );
}
