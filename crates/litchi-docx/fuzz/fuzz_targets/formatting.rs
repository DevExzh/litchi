#![no_main]
mod formatting_harness;
use libfuzzer_sys::fuzz_target;
fuzz_target!(|data: &[u8]| {
    formatting_harness::exercise(data);
});
