//! Create a Keynote presentation without an input document or template.

use litchi_iwa::keynote::KeynoteDocumentBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote <output.key> [title] [subtitle]")?;
    let title = std::env::args()
        .nth(2)
        .unwrap_or_else(|| "Created from scratch".to_owned());
    let subtitle = std::env::args()
        .nth(3)
        .unwrap_or_else(|| "Built with litchi-iwa".to_owned());

    KeynoteDocumentBuilder::new()
        .title(title)
        .subtitle(subtitle)
        .build()?
        .save(output)?;
    Ok(())
}
