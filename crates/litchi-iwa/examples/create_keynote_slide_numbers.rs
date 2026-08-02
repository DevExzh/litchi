//! Create a Keynote presentation with native slide numbers and no input file.

use litchi_iwa::keynote::KeynoteDocumentBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_slide_numbers <output.key>")?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Native slide numbers")
        .subtitle("Created entirely by litchi-iwa")
        .slide_number_visible(true)
        .build()?;
    let layout = keynote.default_slide_layout()?;
    keynote.add_slide(layout)?;
    keynote.set_slide_title(1, "Second slide")?;
    keynote.set_slide_number_visible(1, true)?;
    keynote.save(output)?;
    Ok(())
}
