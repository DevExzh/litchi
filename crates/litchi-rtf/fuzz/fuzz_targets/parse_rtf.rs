#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = litchi_rtf::decompress(data);
    if let Ok(doc) = litchi_rtf::RtfDocument::from_bytes(data) {
        let _ = doc.text();
    }
});
