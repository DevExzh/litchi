//! Read an Apple iWork document (`.pages`, `.numbers`, or `.key`) and print
//! a summary of its contents.
//!
//! # Run
//!
//! ```bash
//! cargo run -p litchi-iwa --example read_iwork -- /path/to/document.pages
//! ```
//!
//! Apple iWork test fixtures are not bundled with this checkout. Drop a
//! `.pages` / `.numbers` / `.key` file into
//! `test-data/iwa/{pages,numbers,keynote}/` (or anywhere on disk) and pass
//! its path on the command line.

use std::env;
use std::path::Path;

use litchi_iwa::Document;

const TEST_DATA_HINT: &str = "test-data/iwa/{pages,numbers,keynote}";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = match env::args().nth(1) {
        Some(p) => p,
        None => {
            eprintln!("usage: read_iwork <path-to-iwork-file>");
            eprintln!();
            eprintln!(
                "no path given. drop a .pages/.numbers/.key file into `{}` and pass its path,",
                TEST_DATA_HINT
            );
            eprintln!("or point at any iWork bundle on disk.");
            return Ok(());
        }
    };

    let path = Path::new(&path);
    if !path.exists() {
        eprintln!("file not found: {}", path.display());
        eprintln!(
            "iWork test fixtures are not committed; please supply a real document path."
        );
        return Ok(());
    }

    println!("opening: {}", path.display());
    let doc = Document::open(path)?;

    // High-level statistics.
    let stats = doc.stats();
    println!("--- document stats ---");
    println!("application:   {:?}", stats.application);
    println!("total objects: {}", stats.total_objects);
    println!("archives:      {}", stats.archives_count);
    println!("top message types: {}", stats.message_type_summary());

    // Plain-text extraction (truncated preview).
    println!("--- text preview ---");
    let text = doc.text()?;
    if text.is_empty() {
        println!("(no text extracted)");
    } else {
        let preview_end = text
            .char_indices()
            .nth(500)
            .map(|(i, _)| i)
            .unwrap_or(text.len());
        println!("{}", &text[..preview_end]);
        if preview_end < text.len() {
            println!("... ({} more chars)", text.len() - preview_end);
        }
    }

    // Structured data summary (tables / slides / sections).
    println!("--- structured summary ---");
    let structured = doc.extract_structured_data()?;
    println!("{}", structured.summary());

    Ok(())
}
