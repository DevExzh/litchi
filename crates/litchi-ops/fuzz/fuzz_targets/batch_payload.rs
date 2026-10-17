#![no_main]
//! Batch-payload ingress fuzz target; see `batch_payload_harness.rs`.
mod batch_payload_harness;
libfuzzer_sys::fuzz_target!(|data: &[u8]| {
    batch_payload_harness::exercise(data);
});
