#![no_main]
mod paths_harness;
libfuzzer_sys::fuzz_target!(|data: &[u8]| {
    paths_harness::exercise(data);
});
