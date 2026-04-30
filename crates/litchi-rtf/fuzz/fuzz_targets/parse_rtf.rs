#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(doc) = litchi_rtf::RtfDocument::from_bytes(data) {
        let _ = doc.text();
    }
});
