//! Read a legacy Word `.doc` file and print basic statistics.
//!
//! Run with:
//!     cargo run -p litchi-doc --example read_doc
//!     cargo run -p litchi-doc --example read_doc -- path/to/file.doc

#![allow(
    clippy::doc_markdown,
    clippy::map_unwrap_or,
    clippy::print_stdout,
    clippy::uninlined_format_args,
    reason = "the interactive example favors familiar terminal output and explanatory prose over library-code style constraints"
)]

use litchi_doc::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "test-data/ole/doc/Lists.doc".to_string());

    println!("Opening DOC: {}", path);

    let mut package = Package::open(&path)?;
    let doc = package.document()?;

    // ---- Basic stats ----
    let paragraph_count = doc.paragraph_count()?;
    let table_count = doc.table_count()?;
    let text = doc.text()?;
    let total_chars = text.chars().count();

    println!("\n=== Structure ===");
    println!("Paragraph count : {}", paragraph_count);
    println!("Table count     : {}", table_count);
    println!("Total characters: {}", total_chars);

    // ---- Text preview (first ~500 chars) ----
    println!("\n=== Text preview (first 500 chars) ===");
    let preview: String = text.chars().take(500).collect();
    println!("{}", preview);
    if total_chars > 500 {
        println!("... [truncated, {} more chars]", total_chars - 500);
    }

    // ---- Paragraph sample ----
    println!("\n=== First 5 paragraphs ===");
    let paragraphs = doc.paragraphs()?;
    for (i, para) in paragraphs.iter().take(5).enumerate() {
        let para_text = para.text().unwrap_or("");
        let snippet: String = para_text.chars().take(80).collect();
        println!("[{}] {}", i, snippet);
    }

    // ---- Tables (just headline counts) ----
    let tables = doc.tables()?;
    if !tables.is_empty() {
        println!("\n=== Tables ===");
        for (i, table) in tables.iter().enumerate() {
            let row_count = table.rows().map(|r| r.len()).unwrap_or(0);
            println!("Table {}: {} row(s)", i, row_count);
        }
    }

    // ---- FIB summary ----
    let fib = doc.fib();
    println!("\n=== FIB ===");
    println!("nFib (file format version): 0x{:04X}", fib.version());
    println!("Word version              : {}", fib.version_name());

    Ok(())
}
