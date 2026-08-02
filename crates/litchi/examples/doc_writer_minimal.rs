//! Minimal DOC file test - single paragraph only

use litchi::doc::writer::DocWriter;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Hello World")?;
    writer.save("output/doc_minimal.doc")?;
    println!("✅ Minimal DOC saved to output/doc_minimal.doc");
    Ok(())
}
