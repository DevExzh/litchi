//! Paired host-vs-leaf cost probes for the W1-B Pages routes (record 0371).
//!
//! Each probe runs one route four times in ABBA order: open, one edit,
//! commit (host: apply) and serialization. Measure with callgrind on a
//! release build of this target, collecting only one side:
//!
//! ```sh
//! cargo test --release -p litchi-iwa --test parity --no-run
//! taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*probe_host*' \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   pages_cost_probes::<probe> --test-threads=1
//! # and '--toggle-collect=*probe_leaf*'
//! ```

use std::hint::black_box;

use crate::harness::{Fixture, TestResult};

/// Run `host` and `leaf` in ABBA order; both must succeed, and the leaf's
/// two outputs must agree on their length. The host's may not: a route that
/// mints random UUIDs compresses to a slightly different length each run.
pub fn abba(
    input: &Fixture,
    host: impl Fn(&[u8]) -> TestResult<Vec<u8>>,
    leaf: impl Fn(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult {
    probe_host(&host, &input.bytes)?;
    let b1 = probe_leaf(&leaf, &input.bytes)?;
    let b2 = probe_leaf(&leaf, &input.bytes)?;
    probe_host(&host, &input.bytes)?;
    if b1 != b2 {
        return Err(format!("the leaf route is not deterministic: {b1}/{b2} bytes").into());
    }
    Ok(())
}

#[inline(never)]
fn probe_host(route: &dyn Fn(&[u8]) -> TestResult<Vec<u8>>, input: &[u8]) -> TestResult<usize> {
    // Distinct bodies keep the linker from folding the two probe symbols.
    black_box(0x5a_u8);
    Ok(black_box(route(black_box(input))?).len())
}

#[inline(never)]
fn probe_leaf(route: &dyn Fn(&[u8]) -> TestResult<Vec<u8>>, input: &[u8]) -> TestResult<usize> {
    black_box(0xa5_u16);
    Ok(black_box(route(black_box(input))?).len())
}
