#![no_main]
//! Batch execution fuzz target; see `batch_apply_harness.rs`. Needs the
//! fuzz crate's `apply` feature (DOCX, XLSX and PPTX adapters).
mod batch_apply_harness;
#[allow(dead_code, reason = "this target shares only the decode limits")]
mod batch_payload_harness;
libfuzzer_sys::fuzz_target!(|data: &[u8]| {
    batch_apply_harness::exercise(data);
});
