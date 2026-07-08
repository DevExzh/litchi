//! Read an OpenDocument Text (`.odt`) file and print its text content.
//!
//! If a path argument is supplied, that file is opened. Otherwise the example
//! creates a small ODT in a tempfile via [`DocumentBuilder`] and reads it back.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-odf --example read_odt
//! cargo run -p litchi-odf --example read_odt -- path/to/file.odt
//! ```

use std::path::PathBuf;

use litchi_odf::{Document, DocumentBuilder};
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Either use a user-supplied path, or build a fresh ODT in a tempfile.
    let (path, _tempfile_guard): (PathBuf, Option<NamedTempFile>) = match std::env::args().nth(1) {
        Some(arg) => (PathBuf::from(arg), None),
        None => {
            println!("No path provided; creating a fresh ODT via DocumentBuilder...");
            let tmp = NamedTempFile::with_suffix(".odt")?;
            let mut builder = DocumentBuilder::new();
            builder.add_heading("litchi-odf example", 1)?;
            builder.add_paragraph("This document was created by the read_odt example.")?;
            builder.add_paragraph("It demonstrates a simple build-then-read round trip.")?;
            builder.add_bulleted_list(vec!["First bullet", "Second bullet", "Third bullet"])?;
            builder.add_heading("Conclusion", 2)?;
            builder.add_paragraph("Reading round-trips text content successfully.")?;
            // `save` consumes the builder, so use the tempfile path explicitly.
            let path = tmp.path().to_path_buf();
            builder.save(&path)?;
            (path, Some(tmp))
        },
    };

    println!("Opening: {}", path.display());
    let doc = Document::open(&path)?;

    // Full text extraction.
    let text = doc.text()?;
    println!("\n--- Full text ({} chars) ---", text.chars().count());
    println!("{}", text);

    // Per-paragraph view.
    let paragraphs = doc.paragraphs()?;
    println!("\n--- Paragraphs: {} ---", paragraphs.len());
    for (i, para) in paragraphs.iter().take(10).enumerate() {
        let body = para.text().unwrap_or_default();
        let style = para.style_name().unwrap_or("");
        println!("  [{}] style={:?} text={:?}", i + 1, style, body);
    }

    // Tempfile (if any) is dropped here, deleting the file automatically.
    Ok(())
}
