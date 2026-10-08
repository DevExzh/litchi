#![no_main]
mod direct_properties_harness;
use libfuzzer_sys::fuzz_target;
fuzz_target!(|data: &[u8]| {
    direct_properties_harness::exercise(data);
});
