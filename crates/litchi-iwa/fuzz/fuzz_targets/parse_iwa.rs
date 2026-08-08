#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = litchi_iwa::detect::bytes(data);

    if let Ok(doc) = litchi_iwa::Document::from_bytes(data) {
        // The migration host temporarily retains text and media coverage.
        // Structured aggregation is fuzzed through the supported root
        // `litchi::iwork` coordinator in `crates/litchi/fuzz`.
        let _ = doc.text();
        let _ = doc.media_stats();
    }
});
