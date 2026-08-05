//! Minimal DOC writer MRE
//!
//! Generates the smallest possible Word 97-2003 (.doc) file using the library.
//! Useful to verify whether the produced OLE/DOC structure is valid in Microsoft Word.
//!
//! Run with:
//!   cargo run --example minimal_doc_mre --features doc --no-default-features

use litchi_doc::Writer;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Output path (default: minimal_mre.doc)
    let out = env::args()
        .nth(1)
        .unwrap_or_else(|| "minimal_mre.doc".to_string());

    // Build the tiniest document: a single paragraph with a paragraph mark
    let mut doc = Writer::new();
    doc.add_paragraph("Hello")?;

    // Save
    doc.save(&out)?;

    println!("Minimal DOC MRE written to: {}", out);
    println!(
        "Open this file in Microsoft Word. If it fails, the issue is likely in the writer stack."
    );

    Ok(())
}
