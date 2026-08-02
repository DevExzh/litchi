#![no_main]

use libfuzzer_sys::fuzz_target;

// Keep arbitrary packaged and flat OpenDocument input on the detector's
// bounded, non-panicking byte path. Malformed input is expected to return None.
fuzz_target!(|data: &[u8]| {
    let _ = litchi_odf::detect::bytes(data);
});
