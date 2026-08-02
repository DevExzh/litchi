#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = litchi_iwa::detect::bytes(data);

    if let Ok(doc) = litchi_iwa::Document::from_bytes(data) {
        // Exercise downstream decoders: text extraction, structured data,
        // and media stats all walk the snappy + protobuf object graph.
        let _ = doc.text();
        let _ = doc.extract_structured_data();
        let _ = doc.media_stats();
    }
});
