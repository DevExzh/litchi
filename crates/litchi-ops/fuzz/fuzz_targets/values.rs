#![no_main]
mod values_harness;
use libfuzzer_sys::fuzz_target;
fuzz_target!(|data: &[u8]| {
    values_harness::exercise(data);
});
