#![no_main]

use libfuzzer_sys::fuzz_target;

// Drives raw bytes through litchi-core's format detection.
// Errors are expected on malformed input; we only want to ensure
// the detector does not panic, OOM, or hit UB on arbitrary bytes.
fuzz_target!(|data: &[u8]| {
    let _ = litchi_core::detection::odf::detect_odf_format(data);
});
