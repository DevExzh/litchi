#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_rtf::{Document, read::Limits as ParseLimits, transport};

// Keep both the fuzzer input and the parser's expansion/retention work
// bounded. Oversized inputs are skipped rather than truncated so every path
// below receives the same unchanged source slice.
const MAX_INPUT_BYTES: usize = 1 * 1024 * 1024;
const MAX_TOKENS: usize = 256 * 1024;
const MAX_OPAQUE_NODES: usize = 8 * 1024;
const MAX_OPAQUE_NODE_BYTES: usize = 64 * 1024;

fn parse_limits() -> ParseLimits {
    ParseLimits::new()
        .with_max_source_bytes(MAX_INPUT_BYTES)
        .with_max_tokens(MAX_TOKENS)
        .with_max_binary_bytes(MAX_INPUT_BYTES)
        .with_max_total_binary_bytes(MAX_INPUT_BYTES)
        .with_max_decompressed_bytes(MAX_INPUT_BYTES)
        .with_max_opaque_nodes(MAX_OPAQUE_NODES)
        .with_max_opaque_node_bytes(MAX_OPAQUE_NODE_BYTES)
        .with_max_total_opaque_bytes(MAX_INPUT_BYTES)
}

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }

    let _ = transport::decompress_with_limits(data, transport::Limits::new(MAX_INPUT_BYTES));
    if let Ok(doc) = Document::from_bytes_with_limits(data, parse_limits()) {
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
