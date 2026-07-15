//! Create a Numbers spreadsheet without an input document or template.

use litchi_iwa::numbers::NumbersDocumentBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers <output.numbers>")?;
    NumbersDocumentBuilder::new().build()?.save(output)?;
    Ok(())
}
