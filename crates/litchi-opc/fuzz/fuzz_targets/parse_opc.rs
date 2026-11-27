#![no_main]
#![allow(
    dead_code,
    reason = "the shared smoke-only boundary helpers are not fuzz entrypoints"
)]

use libfuzzer_sys::fuzz_target;
use std::hint::black_box;

#[path = "opc_harness.rs"]
mod opc_harness;

fuzz_target!(|data: &[u8]| {
    black_box(opc_harness::exercise(data));
});
