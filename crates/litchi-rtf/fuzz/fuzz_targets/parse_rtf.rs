#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_rtf::{OpenOptions, Package, SourceLimits, transport};

// Keep both the fuzzer input and the parser's expansion/retention work
// bounded. Oversized inputs are skipped rather than truncated so every path
// below receives the same unchanged source slice.
const MAX_INPUT_BYTES: usize = 1 * 1024 * 1024;
const MAX_TOKENS: usize = 256 * 1024;
const MAX_OPAQUE_NODES: usize = 8 * 1024;
const MAX_OPAQUE_NODE_BYTES: usize = 64 * 1024;

fn parse_options() -> OpenOptions {
    let limits = SourceLimits::builder()
        .max_input_bytes(MAX_INPUT_BYTES)
        .expect("fuzzer input limit is representable")
        .max_events(MAX_TOKENS)
        .expect("fuzzer event limit is representable")
        .max_binary_bytes(MAX_INPUT_BYTES)
        .expect("fuzzer binary limit is representable")
        .max_aggregate_binary_bytes(MAX_INPUT_BYTES)
        .expect("fuzzer aggregate binary limit is representable")
        .max_decompressed_bytes(MAX_INPUT_BYTES)
        .expect("fuzzer decompression limit is representable")
        .max_opaque_nodes(MAX_OPAQUE_NODES)
        .expect("fuzzer opaque-node limit is representable")
        .max_opaque_node_bytes(MAX_OPAQUE_NODE_BYTES)
        .expect("fuzzer opaque-node byte limit is representable")
        .max_aggregate_opaque_bytes(MAX_INPUT_BYTES)
        .expect("fuzzer aggregate opaque-byte limit is representable")
        .build()
        .expect("fuzzer source profile is consistent");
    OpenOptions::new().source_limits(limits)
}

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }

    let _ = transport::decompress(data, transport::DecompressionLimits::new(MAX_INPUT_BYTES));
    if let Ok(doc) = Package::from_bytes_with_options(data, parse_options()) {
        let _ = doc.text();
        let mut paragraphs = doc.body().paragraphs();
        let selected = data
            .first()
            .copied()
            .map_or(0, |value| usize::from(value) % 32);
        if paragraphs.nth(selected).is_some() {
            let _ = paragraphs.next();
        }
    }
});
