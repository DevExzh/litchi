//! Prints the digests of representative queries over the OOXML corpus
//! (`tests/support/query_corpus.rs`) as JSON, for comparing two builds and
//! for blessing `tests/golden/query_corpus/digests.json` (record 0262).
//!
//! ```text
//! cargo run --release -p litchi-ops --features "xlsx docx pptx xls xlsb" \
//!     --example query_corpus_digest -- [corpus root, default test-data]
//! ```

use std::path::PathBuf;

#[path = "../tests/support/query_corpus.rs"]
mod query_corpus;

fn main() {
    let root = std::env::args().nth(1).map_or_else(
        || PathBuf::from(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data")),
        PathBuf::from,
    );
    let digests = query_corpus::digests(&root);
    println!(
        "{}",
        serde_json::to_string_pretty(&digests).expect("digests serialize")
    );
}
