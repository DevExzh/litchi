//! Minimal DOC file test - single paragraph only

use litchi_doc::builder::Builder;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut writer = Builder::new();
    writer.add_paragraph("Hello World")?;
    writer.build()?.save("output/doc_minimal.doc")?;
    println!("✅ Minimal DOC saved to output/doc_minimal.doc");
    Ok(())
}
