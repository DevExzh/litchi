//! Read a `.docx` file and print extracted text, paragraph count, and core
//! document properties.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-ooxml --example read_docx --all-features
//! cargo run -p litchi-ooxml --example read_docx --all-features -- path/to/file.docx
//! ```
//!
//! Default input: `test-data/ooxml/docx/documentProperties.docx` (relative to
//! the workspace root).

use std::env;

use litchi_ooxml::docx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() >= 2 {
        args[1].clone()
    } else {
        "test-data/ooxml/docx/documentProperties.docx".to_string()
    };

    println!("Opening DOCX: {}", path);
    let pkg = Package::open(&path)?;

    // ----- Core document content -----
    let doc = pkg.document()?;

    let paragraph_count = doc.paragraph_count()?;
    let table_count = doc.table_count()?;
    println!("Paragraph count: {}", paragraph_count);
    println!("Table count: {}", table_count);

    let text = doc.text()?;
    let preview_chars: usize = 500;
    if text.len() > preview_chars {
        // Walk char boundaries so we don't cut a UTF-8 sequence.
        let mut end = preview_chars;
        while end < text.len() && !text.is_char_boundary(end) {
            end += 1;
        }
        println!(
            "\n--- Document text (first {} bytes of {}) ---\n{}...",
            end,
            text.len(),
            &text[..end]
        );
    } else {
        println!("\n--- Document text ({} bytes) ---\n{}", text.len(), text);
    }

    // ----- Document properties (core.xml metadata) -----
    let props = pkg.properties();
    println!("\n--- Core document properties ---");
    println!("Title           : {:?}", props.title);
    println!("Subject         : {:?}", props.subject);
    println!("Creator         : {:?}", props.creator);
    println!("Keywords        : {:?}", props.keywords);
    println!("Description     : {:?}", props.description);
    println!("Last modified by: {:?}", props.last_modified_by);
    println!("Category        : {:?}", props.category);
    println!("Content status  : {:?}", props.content_status);
    println!("Language        : {:?}", props.language);
    println!("Created         : {:?}", props.created);
    println!("Modified        : {:?}", props.modified);

    // ----- Custom (app-defined) properties, if any -----
    let custom = pkg.custom_props();
    if !custom.is_empty() {
        println!("\n--- Custom properties ({}) ---", custom.len());
        for (name, value) in custom.iter() {
            println!("  {} = {:?}", name, value);
        }
    }

    Ok(())
}
