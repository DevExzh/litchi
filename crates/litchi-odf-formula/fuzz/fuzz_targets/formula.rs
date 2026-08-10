#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    litchi_odf_formula_fuzz::exercise(data);
});
