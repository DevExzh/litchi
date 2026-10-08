#![no_main]
mod sequential_ingress_harness;
use libfuzzer_sys::fuzz_target;
fuzz_target!(|data: &[u8]| {
    sequential_ingress_harness::exercise(data);
});
