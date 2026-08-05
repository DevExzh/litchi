//! Read a `.docx` file and print extracted text, paragraph count, and core
//! document properties.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-docx --example read_docx
//! cargo run -p litchi-docx --example read_docx -- path/to/file.docx
//! ```
//!
//! Default input: `test-data/ooxml/docx/documentProperties.docx` (relative to
//! the workspace root).

use std::env;

use litchi_docx::Package;

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
    let props = pkg.props();
    println!("\n--- Core document properties ---");
    println!(
        "Title           : {:?}",
        props.and_then(|p| p.title.as_deref())
    );
    println!(
        "Subject         : {:?}",
        props.and_then(|p| p.subject.as_deref())
    );
    println!(
        "Creator         : {:?}",
        props.and_then(|p| p.creator.as_deref())
    );
    println!(
        "Keywords        : {:?}",
        props
            .and_then(|p| p.keywords.as_ref())
            .map(ToString::to_string)
    );
    println!(
        "Description     : {:?}",
        props.and_then(|p| p.description.as_deref())
    );
    println!(
        "Last modified by: {:?}",
        props.and_then(|p| p.last_modified_by.as_deref())
    );
    println!(
        "Category        : {:?}",
        props.and_then(|p| p.category.as_deref())
    );
    println!(
        "Content status  : {:?}",
        props.and_then(|p| p.content_status.as_deref())
    );
    println!(
        "Language        : {:?}",
        props.and_then(|p| p.language.as_deref())
    );
    println!(
        "Created         : {:?}",
        props.and_then(|p| p.created.as_ref())
    );
    println!(
        "Modified        : {:?}",
        props.and_then(|p| p.modified.as_ref())
    );

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
