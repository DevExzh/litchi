#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = litchi_imgconv::emf::convert_emf_to_png(data, None, None);
});
