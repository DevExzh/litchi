#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(doc) = litchi_odf::Document::from_bytes(data.to_vec()) {
        let _ = doc.text();
    }
});
