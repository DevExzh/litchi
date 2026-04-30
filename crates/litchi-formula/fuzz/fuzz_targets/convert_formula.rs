#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // Primary target: binary MTEF -> LaTeX (takes &[u8] directly).
    let _ = litchi_formula::mtef_to_latex(data);
});
