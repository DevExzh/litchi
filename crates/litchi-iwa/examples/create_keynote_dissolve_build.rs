//! Create a Keynote presentation with a native editable Dissolve Build In.

use litchi_iwa::keynote::{KeynoteBuildSettings, KeynoteDocumentBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_dissolve_build <output.key>")?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Dissolve build")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let drawable = keynote
        .slide_drawables(0)?
        .into_iter()
        .next()
        .ok_or("the initial slide has no drawable")?;
    keynote.add_slide_build(0, drawable.id.get(), KeynoteBuildSettings::dissolve_in())?;
    keynote.save(output)?;
    Ok(())
}
